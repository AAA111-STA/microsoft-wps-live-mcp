/* The documented ONLYOFFICE plugin API applies commands in the live session. */
(function () {
  'use strict';
  let busy = false;
  const instance = crypto.randomUUID();
  let unconfirmed = null;
  const base = '/api/session/' + window.OFFICE_SESSION + '/';
  async function api(action, body) {
    const response = await fetch(base + action, {
      method: body === undefined ? 'GET' : 'POST',
      headers: {Authorization: 'Bearer ' + window.OFFICE_TOKEN, 'Content-Type': 'application/json',
                'X-Office-Bridge': instance},
      body: body === undefined ? undefined : JSON.stringify(body)
    });
    if (!response.ok) throw new Error('Bridge request failed: ' + response.status);
    return response.json();
  }
  function command(body) {
    return new Promise(resolve => {
      Asc.scope.officeCommand = body;
      Asc.plugin.callCommand(function () {
        try {
          const c = Asc.scope.officeCommand;
          if (c.op === 'word.replace') {
            const changed = Api.GetDocument().SearchAndReplace({searchString: c.search,
              replaceString: c.replace, matchCase: c.matchCase !== false});
            return {ok: !!changed, result: changed ? 'replaced' : 'No matching text found'};
          }
          if (c.op === 'word.append') {
            const paragraph = Api.CreateParagraph(); paragraph.AddText(c.text);
            Api.GetDocument().Push(paragraph);
            return {ok: true, result: 'paragraph appended'};
          }
          if (c.op === 'sheet.set' || c.op === 'sheet.read') {
            const sheet = c.sheet ? Api.GetSheet(c.sheet) : Api.GetActiveSheet();
            if (!sheet) throw new Error('Worksheet not found');
            const range = sheet.GetRange(c.range);
            if (c.op === 'sheet.read') return {ok: true, result: range.GetValue()};
            range.SetValue(c.values);
            return {ok: true, result: 'range updated'};
          }
          if (c.op === 'slide.set_text') {
            const slide = Api.GetPresentation().GetSlideByIndex(c.slide);
            if (!slide) throw new Error('Slide not found');
            const shape = slide.GetAllShapes()[c.shape];
            if (!shape) throw new Error('Shape not found');
            const content = shape.GetDocContent();
            if (!content) throw new Error('Shape has no editable text');
            content.RemoveAllElements();
            const paragraph = Api.CreateParagraph(); paragraph.AddText(c.text); content.Push(paragraph);
            return {ok: true, result: 'shape text updated'};
          }
          if (c.op === 'inspect') {
            if (typeof Api.GetDocument === 'function') {
              return {ok: true, result: {kind: 'word', text: Api.GetDocument().GetText().slice(0, 30000)}};
            }
            if (typeof Api.GetSheets === 'function') {
              return {ok: true, result: {kind: 'cell', sheets: Api.GetSheets().map(s => s.GetName())}};
            }
            const slides = Api.GetPresentation().GetAllSlides();
            return {ok: true, result: {kind: 'slide', slides: slides.map((slide, i) => ({
              index: i, shapes: slide.GetAllShapes().map((shape, j) => ({index: j,
                text: shape.GetDocContent() ? shape.GetDocContent().GetText().slice(0, 3000) : ''}))}))}};
          }
          throw new Error('Unsupported command');
        } catch (error) {return {ok: false, result: String(error)};}
      }, false, true, result => {
        if (result && typeof result === 'object' && typeof result.ok === 'boolean') resolve(result);
        else resolve({ok: false, result: 'Unconfirmed editor result: ' + JSON.stringify(result)});
      });
    });
  }
  async function poll() {
    if (busy) return;
    busy = true;
    try {
      if (unconfirmed) {
        await api('complete', unconfirmed);
        unconfirmed = null;
      }
      const response = await api('poll');
      if (response.command) {
        const result = await command(response.command.body);
        unconfirmed = {id: response.command.id, ...result};
        await api('complete', unconfirmed);
        unconfirmed = null;
      }
    } catch (error) { console.warn('Office bridge:', error.message); }
    finally {busy = false;}
  }
  window.Asc.plugin.init = function () {poll(); setInterval(poll, 600);};
  window.Asc.plugin.button = function () {};
})();
