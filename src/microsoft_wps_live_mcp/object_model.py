"""General Office object model access: no task whitelist and no Python eval/exec.

COM objects remain on the STA thread. Clients receive opaque references and can
compose get/set/call operations against the native software's public object model.
"""
from __future__ import annotations

import re
import uuid


def error(message):
    from .native import OfficeError
    return OfficeError(message)


def member_name(value):
    if not isinstance(value, str) or not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]*", value):
        raise error("Object members must be public COM names, for example Shapes or TextFrame2.")
    return value


def validate_steps(steps):
    if not isinstance(steps, list) or not 1 <= len(steps) <= 500:
        raise error("steps must contain 1 to 500 get/set/call operations.")
    allowed = {"action", "target", "member", "args", "kwargs", "value", "as", "read_only"}
    for step in steps:
        if not isinstance(step, dict) or set(step) - allowed:
            raise error("Each step accepts action, target, member, args, kwargs, value, as, read_only.")
        if step.get("action") not in ("get", "set", "call"):
            raise error("step.action must be get, set, or call.")
        if "read_only" in step and not isinstance(step["read_only"], bool):
            raise error("read_only must be a boolean.")
        member_name(step.get("member"))
        if "as" in step:
            alias = member_name(step["as"])
            if alias in ("document", "application") or alias.startswith("obj_"):
                raise error("Alias must not shadow document, application, or object references.")
        if step["action"] == "set" and "value" not in step:
            raise error("A set step requires value.")
        if not isinstance(step.get("args", []), list) or not isinstance(step.get("kwargs", {}), dict):
            raise error("args must be an array and kwargs must be an object.")
        for name in step.get("kwargs", {}):
            member_name(name)
    return any(s["action"] == "set" or (s["action"] == "call" and not s.get("read_only", False)) for s in steps)


_SHAPE_CREATE_MEMBERS = frozenset({
    "addshape", "addtextbox", "addconnector", "addpicture", "addtable",
    "addslide", "converttoshape", "duplicate", "group", "ungroup", "delete",
})
_FREEFORM_BUILD_MEMBERS = frozenset({"buildfreeform", "addnodes"})


class ObjectModel:
    def _api_objects(self):
        return self.__dict__.setdefault("object_references", {})

    def _forget_objects(self, document_id):
        objects = self._api_objects()
        for key in [k for k, v in objects.items() if v[0] == document_id]:
            del objects[key]

    def _api_target(self, record, target, aliases):
        if not isinstance(target, str) or len(target) > 1000:
            raise error("target must be document, application, an alias, or a returned object_id with optional dotted properties.")
        parts = target.split(".")
        root = parts.pop(0)
        if root == "document":
            obj = record.doc
        elif root == "application":
            obj = record.app
        elif root in aliases:
            obj = aliases[root]
        else:
            stored = self._api_objects().get(root)
            if not stored or stored[0] != record.id:
                raise error("Unknown object reference or reference belongs to another document. Inspect again after reopening.")
            obj = stored[1]
        for part in parts:
            obj = getattr(obj, member_name(part))
        return obj

    def _api_decode(self, record, value, aliases, depth=0):
        if depth > 20:
            raise error("Arguments are nested too deeply.")
        if isinstance(value, dict):
            if set(value) == {"$ref"}:
                return self._api_target(record, value["$ref"], aliases)
            if value == {"$missing": True}:
                import pythoncom
                return pythoncom.Missing
            raise error('Object arguments use {"$ref":"alias_or_object_id"}; omitted optional arguments use {"$missing":true}.')
        if isinstance(value, (list, tuple)):
            return tuple(self._api_decode(record, v, aliases, depth + 1) for v in value)
        if value is None or isinstance(value, (str, int, float, bool)):
            return value
        raise error("Unsupported COM argument type.")

    def _api_export(self, record, value, depth=0):
        from .native import json_value
        if depth > 10:
            return {"truncated": True}
        if hasattr(value, "_oleobj_"):
            identifier = "obj_" + uuid.uuid4().hex
            objects = self._api_objects()
            objects[identifier] = (record.id, value)
            while len(objects) > 1000:
                del objects[next(iter(objects))]
            try:
                typename = value._oleobj_.GetTypeInfo().GetDocumentation(-1)[0]
            except Exception:
                typename = "IDispatch"
            return {"object_id": identifier, "type": typename}
        if isinstance(value, (tuple, list)):
            return [self._api_export(record, v, depth + 1) for v in value[:10000]]
        if callable(value):
            return {"callable": True, "message": "Use action=call for this indexed property or method."}
        if isinstance(value, str) and len(value) > 100000:
            return {"text": value[:100000], "truncated": True, "total_characters": len(value)}
        return json_value(value)

    def api_describe(self, document, target="document", search="", limit=100):
        record = self.resolve(document)
        if not 1 <= limit <= 500:
            raise error("limit must be between 1 and 500.")
        obj = self._api_target(record, target, {})
        if not hasattr(obj, "_oleobj_"):
            return {"target": target, "value": self._api_export(record, obj), "members": []}
        try:
            info = obj._oleobj_.GetTypeInfo()
            members = []
            total = 0
            for index in range(info.GetTypeAttr().cFuncs):
                descriptor = info.GetFuncDesc(index)
                names = info.GetNames(descriptor.memid)
                if not names or names[0].startswith("_") or search.casefold() not in names[0].casefold():
                    continue
                total += 1
                if len(members) >= limit:
                    continue
                kinds = {1: "method", 2: "property_get", 4: "property_set", 8: "property_set_reference"}
                members.append({"name": names[0], "kind": kinds.get(descriptor.invkind, str(descriptor.invkind)),
                                "parameters": list(names[1:]), "optional_parameters": descriptor.cParamsOpt,
                                "documentation": info.GetDocumentation(descriptor.memid)[1]})
            return {**self._info(record), "target": target, "type": info.GetDocumentation(-1)[0],
                    "members": members, "total_matching_members": total, "truncated": total > limit}
        except Exception as exc:
            raise error(f"This COM object does not expose type information: {exc}. "
                        "Its documented public members may still be called using office_api_batch.") from exc

    def api_constants(self, document, search="", limit=100):
        record = self.resolve(document)
        if not 1 <= limit <= 500:
            raise error("limit must be between 1 and 500.")
        cache = self.__dict__.setdefault("constant_cache", {})
        key = (record.engine, record.kind)
        if key not in cache:
            values = {}
            libraries = [record.app._oleobj_.GetTypeInfo().GetContainingTypeLib()[0]]
            # Office shared enums (MsoAutoShapeType, arrows, connectors) live in
            # the referenced Office library, not the Word/Excel/PPT type library.
            try:
                from .com_dispatch import dispatch
                obj = dispatch(record.app.CommandBars)
                libraries.append(obj._oleobj_.GetTypeInfo().GetContainingTypeLib()[0])
            except Exception:
                pass
            for library in libraries:
                for index in range(library.GetTypeInfoCount()):
                    info = library.GetTypeInfo(index)
                    attributes = info.GetTypeAttr()
                    if attributes.typekind != 0:  # TKIND_ENUM
                        continue
                    group = info.GetDocumentation(-1)[0]
                    for j in range(attributes.cVars):
                        desc = info.GetVarDesc(j)
                        name = info.GetNames(desc.memid)[0]
                        values[name] = {"name": name, "value": desc.value, "enum": group}
            cache[key] = values
        result = [v for name, v in cache[key].items()
                  if search.casefold() in name.casefold() or search.casefold() in v["enum"].casefold()]
        return {"engine": record.engine, "kind": record.kind, "constants": result[:limit],
                "total_matching_constants": len(result), "truncated": len(result) > limit}

    def _flush_live_step(self, record, focus, step_pause_ms):
        view = self.live_checkpoint(record, focus, pause_ms=step_pause_ms)
        if view.get("state") == "unavailable":
            raise error("Live display became unavailable; the batch stopped after this native step.")

    def api_batch(self, document, steps, save=True, _aliases=None, _step_pause_ms=None):
        from .native import file_hash
        from pathlib import Path
        mutating = validate_steps(steps)
        record = self.resolve(document)
        aliases = {} if _aliases is None else _aliases
        backup = self._before_edit(record) if mutating else None
        results = []
        pending_shape = False
        pending_focus = None
        for index, step in enumerate(steps, 1):
            try:
                member_lower = step["member"].casefold()
                if record.live_enabled and pending_shape and member_lower in _SHAPE_CREATE_MEMBERS:
                    self._flush_live_step(record, pending_focus, _step_pause_ms)
                    pending_shape = False
                    pending_focus = None
                target = self._api_target(record, step.get("target", "document"), aliases)
                name = step["member"]
                if step["action"] == "get":
                    value = getattr(target, name)
                elif step["action"] == "set":
                    value = self._api_decode(record, step["value"], aliases)
                    setattr(target, name, value)
                else:
                    args = [self._api_decode(record, v, aliases) for v in step.get("args", [])]
                    kwargs = {k: self._api_decode(record, v, aliases) for k, v in step.get("kwargs", {}).items()}
                    value = getattr(target, name)(*args, **kwargs)
                    if name.casefold() in ("save", "saveas", "saveas2") and self._path(target) == self._path(record.doc):
                        self._info(record)
                        record.disk_hash = file_hash(Path(record.path)) if record.path else None
                if "as" in step:
                    aliases[step["as"]] = value
                results.append({"step": index, "result": self._api_export(record, value)})
                if record.live_enabled and (step["action"] == "set" or
                        (step["action"] == "call" and not step.get("read_only", False))):
                    if member_lower in _FREEFORM_BUILD_MEMBERS:
                        continue
                    if member_lower in _SHAPE_CREATE_MEMBERS:
                        pending_shape = True
                        pending_focus = self._focus_for_object(record, value) if hasattr(value, "_oleobj_") else None
                        if pending_focus is None:
                            pending_focus = self._focus_for_object(record, target)
                    elif pending_shape:
                        if pending_focus is None:
                            pending_focus = self._focus_for_object(record, target)
                    else:
                        focus = self._focus_for_object(record, value) if hasattr(value, "_oleobj_") else None
                        if focus is None:
                            focus = self._focus_for_object(record, target)
                        self._flush_live_step(record, focus, _step_pause_ms)
            except Exception as exc:
                record.autosave_paused = True
                record.autosave_state = {"state": "needs_inspection", "message": "A batch partially failed. Inspect, then explicitly save to resume automatic saving."}
                return {**self._info(record), "state": "partial_error", "failed_step": index,
                        "completed_steps": len(results), "results": results, "error": str(exc),
                        "disk_saved": False, "backup": backup, "inspect_before_retry": True,
                        "message": "Earlier steps may have changed the live document. The batch is not transactional. "
                                   "Inspect before retrying. No automatic replay or rollback was attempted."}
        if record.live_enabled and pending_shape:
            self._flush_live_step(record, pending_focus, _step_pause_ms)
        if mutating:
            output = self._result(record, {"completed_steps": len(results)}, save, backup)
        else:
            output = {**self._info(record), "disk_saved": False}
        return {**output, "state": "completed", "results": results}
