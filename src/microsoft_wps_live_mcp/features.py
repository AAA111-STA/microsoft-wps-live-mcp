"""Higher-level document workflows, shared by the Microsoft and WPS backends."""
from __future__ import annotations

import datetime as dt
import difflib
from pathlib import Path
import re
import sys


def fail(message):
    from .native import OfficeError
    return OfficeError(message)


def supported_value(getter):
    from .native import json_value
    try:
        return json_value(getter())
    except Exception as exc:
        return {"available": False, "reason": str(exc)}


class SuiteFeatures:
    def inspect_document(self, document):
        r = self.resolve(document)
        output = self._info(r)
        if r.kind == "word":
            names = ("Paragraphs", "Sections", "Tables", "InlineShapes", "Shapes", "Comments", "Revisions",
                     "Footnotes", "Endnotes", "Fields", "Bookmarks", "TablesOfContents", "OMaths")
            output["components"] = {name: supported_value(lambda n=name: int(getattr(r.doc, n).Count)) for name in names}
            output["pages"] = supported_value(lambda: int(r.doc.ComputeStatistics(2)))
            output["words"] = supported_value(lambda: int(r.doc.ComputeStatistics(0)))
            output["track_revisions"] = supported_value(lambda: bool(r.doc.TrackRevisions))
        elif r.kind == "excel":
            sheets = []
            for i in range(1, int(r.doc.Worksheets.Count) + 1):
                sheet = r.doc.Worksheets.Item(i)
                sheets.append({"index": i, "name": str(sheet.Name),
                               "used_range": supported_value(lambda: str(sheet.UsedRange.Address)),
                               "shapes": supported_value(lambda: int(sheet.Shapes.Count)),
                               "charts": supported_value(lambda: int(sheet.ChartObjects().Count)),
                               "tables": supported_value(lambda: int(sheet.ListObjects.Count)),
                               "pivot_tables": supported_value(lambda: int(sheet.PivotTables().Count))})
            output["worksheets"] = sheets
            output["names"] = supported_value(lambda: int(r.doc.Names.Count))
        else:
            output["slide_count"] = int(r.doc.Slides.Count)
            output["slide_size_points"] = {"width": float(r.doc.PageSetup.SlideWidth), "height": float(r.doc.PageSetup.SlideHeight)}
            output["designs"] = supported_value(lambda: int(r.doc.Designs.Count))
            output["sections"] = supported_value(lambda: int(r.doc.SectionProperties.Count))
            output["slides"] = [{"index": i, "shapes": int(r.doc.Slides.Item(i).Shapes.Count),
                                  "animations": supported_value(lambda: int(r.doc.Slides.Item(i).TimeLine.MainSequence.Count))}
                                 for i in range(1, min(int(r.doc.Slides.Count), 200) + 1)]
            output["slides_truncated"] = int(r.doc.Slides.Count) > 200
        return output

    def export_pdf(self, document, filename):
        from .native import file_hash
        r = self.resolve(document)
        path = Path(filename).expanduser()
        if not path.is_absolute() or path.suffix.lower() != ".pdf" or not path.parent.is_dir() or path.exists():
            raise fail("Use an unused absolute .pdf output path in an existing directory.")
        if sys.platform == "darwin":
            from .macos import export_pdf_in_app, get_app_name
            if not r.path:
                raise fail("Save the macOS document to a local path before exporting PDF.")
            if not bool(r.doc.Saved):
                r.doc.Save()
            if not export_pdf_in_app(get_app_name(r.engine, r.kind), Path(r.path), path.resolve()):
                raise fail("The macOS application did not produce a PDF file.")
        elif r.kind == "word":
            r.doc.ExportAsFixedFormat(str(path), 17)
        elif r.kind == "excel":
            r.doc.ExportAsFixedFormat(0, str(path))
        else:
            try:
                r.doc.SaveAs(str(path.resolve()), 32)
            except Exception:
                r.doc.ExportAsFixedFormat(str(path.resolve()), 2)
        if not path.is_file():
            raise fail("The native application did not produce a PDF file.")
        with path.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise fail("The exported file did not have a valid PDF header.")
        return {**self._info(r), "exported": True, "output": str(path), "sha256": file_hash(path),
                "size_bytes": path.stat().st_size, "source_disk_saved": False}

    def ppt_export_image(self, document, filename, slide=1, width=1920, height=1080):
        from .native import file_hash
        r = self.resolve(document, "ppt")
        path = Path(filename).expanduser()
        ext = path.suffix.lower()
        if not path.is_absolute() or ext not in (".png", ".jpg", ".jpeg") or not path.parent.is_dir() or path.exists():
            raise fail("Use an unused absolute .png or .jpg output path in an existing directory.")
        if not (160 <= int(width) <= 7680 and 90 <= int(height) <= 4320):
            raise fail("width must be 160..7680 and height must be 90..4320.")
        fmt = "PNG" if ext == ".png" else "JPG"
        if sys.platform == "darwin":
            from .macos import export_ppt_slide_image, get_app_name
            self._slide(r, int(slide))
            if not r.path:
                raise fail("Save the presentation to a local path before exporting an image on macOS.")
            if not bool(r.doc.Saved):
                r.doc.Save()
            if not export_ppt_slide_image(get_app_name(r.engine, "ppt"), Path(r.path), path.resolve(), slide=int(slide)):
                raise fail("The macOS application did not produce an image file.")
        else:
            page = self._slide(r, int(slide))
            page.Export(str(path.resolve()), fmt, int(width), int(height))
        if not path.is_file() or path.stat().st_size == 0:
            raise fail("The native application did not produce an image file.")
        return {**self._info(r), "exported": True, "output": str(path), "slide": int(slide),
                "format": fmt, "width": int(width), "height": int(height),
                "sha256": file_hash(path), "size_bytes": path.stat().st_size}

    def word_find(self, document, text, match_case=False, whole_word=False, wildcards=False, limit=100):
        r = self.resolve(document, "word")
        if not text or len(text) > 255 or not 1 <= limit <= 1000:
            raise fail("Search text must contain 1 to 255 characters; limit must be between 1 and 1000.")
        matches = []
        end = int(r.doc.Content.End)
        position = 0
        while position < end and len(matches) < limit:
            span = r.doc.Range(position, end)
            finder = span.Find
            finder.ClearFormatting()
            if not finder.Execute(FindText=text, MatchCase=match_case, MatchWholeWord=whole_word,
                                  MatchWildcards=wildcards, Forward=True, Wrap=0, Format=False):
                break
            start, finish = int(span.Start), int(span.End)
            matches.append({"start": start, "end": finish, "text": str(span.Text),
                            "context": str(r.doc.Range(max(0, start - 60), min(end, finish + 60)).Text)})
            position = max(position + 1, finish)
        return {**self._info(r), "matches": matches, "limit_reached": len(matches) == limit}

    def word_snapshot(self, document, action="take"):
        r = self.resolve(document, "word")
        snapshots = self.__dict__.setdefault("snapshots", {})
        old = snapshots.get(r.id)
        if action == "status":
            return {"document_id": r.id, "has_snapshot": old is not None,
                    "created_at": old["created_at"] if old else None,
                    "paragraph_count": len(old["paragraphs"]) if old else None}
        if action not in ("take", "diff"):
            raise fail("action must be take, status, or diff.")
        if int(r.doc.Paragraphs.Count) > 50000:
            raise fail("Snapshot is limited to 50000 paragraphs.")
        paragraphs = [str(r.doc.Paragraphs.Item(i).Range.Text) for i in range(1, int(r.doc.Paragraphs.Count) + 1)]
        if action == "take":
            snapshots[r.id] = {"created_at": dt.datetime.now(dt.timezone.utc).isoformat(), "paragraphs": paragraphs}
            return {**self._info(r), "snapshot_taken": True, "paragraph_count": len(paragraphs)}
        if old is None:
            raise fail("Take a snapshot before requesting a diff.")
        changes = []
        for tag, a, b, c, d in difflib.SequenceMatcher(None, old["paragraphs"], paragraphs, autojunk=False).get_opcodes():
            if tag != "equal":
                changes.append({"type": tag, "before_start": a + 1, "after_start": c + 1,
                                "before": old["paragraphs"][a:b], "after": paragraphs[c:d]})
        return {**self._info(r), "snapshot_created_at": old["created_at"], "changes": changes,
                "changed_blocks": len(changes), "snapshot_updated": False}

    def word_comments(self, document, action="list", index=None, start=None, end=None, text=None, save=True):
        r = self.resolve(document, "word")
        if action == "list":
            result = []
            for i in range(1, int(r.doc.Comments.Count) + 1):
                comment = r.doc.Comments.Item(i)
                result.append({"index": i, "author": str(comment.Author), "text": str(comment.Range.Text),
                               "start": int(comment.Scope.Start), "end": int(comment.Scope.End),
                               "resolved": supported_value(lambda: bool(comment.Done)),
                               "replies": supported_value(lambda: int(comment.Replies.Count))})
            return {**self._info(r), "comments": result}
        if action not in ("add", "reply", "resolve", "reopen", "delete"):
            raise fail("Unknown comment action.")
        if action in ("add", "reply") and not text:
            raise fail("text is required.")
        if action == "add":
            if start is None or end is None or not 0 <= start <= end <= int(r.doc.Content.End):
                raise fail("add requires valid start/end character offsets.")
            target = None
        else:
            if index is None or not 1 <= index <= int(r.doc.Comments.Count):
                raise fail("index must refer to an existing 1-based comment.")
            target = r.doc.Comments.Item(index)
        backup = self._before_edit(r)
        if action == "add":
            r.doc.Comments.Add(r.doc.Range(start, end), text)
        elif action == "reply":
            target.Replies.Add(target.Scope, text)
        elif action in ("resolve", "reopen"):
            target.Done = action == "resolve"
        else:
            target.Delete()
        return self._result(r, {"comment_action": action, "index": index}, save, backup)

    def word_revisions(self, document, action="list", enabled=None, indices=None, author=None, save=True):
        r = self.resolve(document, "word")
        if action == "list":
            result = []
            for i in range(1, int(r.doc.Revisions.Count) + 1):
                revision = r.doc.Revisions.Item(i)
                result.append({"index": i, "author": str(revision.Author), "type": int(revision.Type),
                               "text": str(revision.Range.Text), "start": int(revision.Range.Start),
                               "end": int(revision.Range.End), "date": str(revision.Date)})
            return {**self._info(r), "tracking": bool(r.doc.TrackRevisions), "revisions": result}
        if action not in ("track", "accept", "reject"):
            raise fail("action must be list, track, accept, or reject.")
        if action == "track" and enabled is None:
            raise fail("track requires an explicit enabled boolean.")
        total = int(r.doc.Revisions.Count)
        if indices is not None and any(isinstance(i, bool) or not isinstance(i, int) or not 1 <= i <= total for i in indices):
            raise fail("indices must refer to existing 1-based revisions.")
        selected = sorted(set(indices if indices is not None else range(1, total + 1)), reverse=True)
        backup = self._before_edit(r)
        affected = 0
        if action == "track":
            r.doc.TrackRevisions = enabled
        else:
            for index in selected:
                revision = r.doc.Revisions.Item(index)
                if author is None or str(revision.Author) == author:
                    revision.Accept() if action == "accept" else revision.Reject()
                    affected += 1
        return self._result(r, {"revision_action": action, "affected": affected,
                                "tracking": bool(r.doc.TrackRevisions)}, save, backup)

    def word_component(self, document, component, options, position="end", save=True):
        r = self.resolve(document, "word")
        valid = {"image", "equation", "footnote", "endnote", "bookmark", "field", "hyperlink",
                 "page_break", "section_break", "toc", "cross_reference"}
        if component not in valid:
            raise fail("Unknown component; use the general object-model API for other native components.")
        end = max(0, int(r.doc.Content.End) - 1)
        offset = 0 if position == "start" else end if position == "end" else position
        if isinstance(offset, bool) or not isinstance(offset, int) or not 0 <= offset <= end:
            raise fail("position must be start, end, or a valid character offset.")
        span = r.doc.Range(offset, offset)
        if component in ("footnote", "endnote", "equation") and not options.get("text"):
            raise fail("options.text is required.")
        if component == "image":
            image = Path(options.get("path", "")).expanduser()
            if not image.is_absolute() or not image.is_file():
                raise fail("options.path must be an existing absolute image path.")
        if component == "bookmark" and not re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{0,39}", options.get("name", "")):
            raise fail("Bookmark name must start with a letter and contain at most 40 letters, digits, or underscores.")
        backup = self._before_edit(r)
        if component == "image":
            picture = r.doc.InlineShapes.AddPicture(str(image), False, True, span)
            if options.get("width") is not None:
                picture.Width = float(options["width"])
            if options.get("height") is not None:
                picture.Height = float(options["height"])
        elif component in ("footnote", "endnote"):
            notes = r.doc.Footnotes if component == "footnote" else r.doc.Endnotes
            notes.Add(Range=span, Text=options["text"])
        elif component == "equation":
            # A separate paragraph avoids folding adjacent note/reference fields into
            # the equation range in WPS. Character offsets are UTF-16 code units.
            equation = options["text"]
            if r.engine == "wps":
                # WPS may return a range but fail to construct OMath in a document
                # containing foot/endnote fields. Build in a temporary native document
                # and transfer FormattedText, preserving an editable native equation.
                temporary = r.app.Documents.Add(Visible=False)
                try:
                    temporary = self._wait_document_ready(temporary)
                    source = temporary.Range(0, 0)
                    source.Text = equation
                    temporary.OMaths.Add(source)
                    math = temporary.OMaths.Item(1)
                    math.BuildUp()
                    if not options.get("inline", False):
                        span.InsertBefore("\r\r")
                        span = r.doc.Range(offset + 1, offset + 1)
                    span.FormattedText = math.Range.FormattedText
                finally:
                    temporary.Close(False)
            elif options.get("inline", False):
                span.InsertBefore(equation)
                equation_start = offset
            else:
                span.InsertBefore("\r" + equation + "\r")
                equation_start = offset + 1
            if r.engine != "wps":
                equation_end = equation_start + len(equation.encode("utf-16-le")) // 2
                math_range = r.doc.OMaths.Add(r.doc.Range(equation_start, equation_end))
                math_range.OMaths.Item(1).BuildUp()
        elif component == "bookmark":
            r.doc.Bookmarks.Add(options["name"], span)
        elif component == "field":
            if not options.get("code"):
                raise fail("options.code is required, for example PAGE or SEQ Figure.")
            r.doc.Fields.Add(Range=span, Type=-1, Text=options["code"], PreserveFormatting=True)
        elif component == "hyperlink":
            r.doc.Hyperlinks.Add(Anchor=span, Address=options.get("address", ""),
                                 SubAddress=options.get("subaddress", ""), TextToDisplay=options.get("text", options.get("address", "link")))
        elif component == "page_break":
            span.InsertBreak(7)
        elif component == "section_break":
            span.InsertBreak(int(options.get("type", 2)))
        elif component == "toc":
            r.doc.TablesOfContents.Add(Range=span, UseHeadingStyles=True,
                                       UpperHeadingLevel=int(options.get("upper_level", 1)),
                                       LowerHeadingLevel=int(options.get("lower_level", 3)))
        else:
            # Native enums: wdContentText=-1, wdPageNumber=7; caption types use
            # their actual localized label, not fabricated numeric ref types.
            span.InsertCrossReference(ReferenceType=options.get("reference_type", 2),
                                      ReferenceKind=options.get("reference_kind", -1),
                                      ReferenceItem=options.get("reference_item", 1),
                                      InsertAsHyperlink=options.get("as_hyperlink", True))
        return self._result(r, {"component": component, "position": offset}, save, backup)

    def word_undo(self, document, times=1, save=True):
        r = self.resolve(document, "word")
        if not 1 <= times <= 100:
            raise fail("times must be between 1 and 100 native undo entries.")
        backup = self._before_edit(r)
        undone = bool(r.doc.Undo(times))
        return self._result(r, {"undone": undone, "requested_entries": times}, save, backup)
