/* GMAO leon: generic form dialog (add / edit records) and confirmation dialog.
 * Native <dialog> + native constraint validation; errors are shown inline once the user has
 * interacted (:user-invalid) and aria-invalid / aria-describedby are kept in sync with it. */
(function () {
  'use strict';
  const C = window.GMCore;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const icon = (name) => `<i data-lucide="${name}" aria-hidden="true"></i>`;
  const paint = () => window.lucide && window.lucide.createIcons();
  const iso = (s) => (s == null ? '' : new Date((s - 25569) * 86400000).toISOString().slice(0, 10));
  const hhmm = (t) => {
    if (t == null) return '';
    const m = Math.round(t * 1440);
    return `${String(Math.floor(m / 60) % 24).padStart(2, '0')}:${String(m % 60).padStart(2, '0')}`;
  };
  const opt = (o) => (typeof o === 'object' ? o : { value: o, label: o });
  let seq = 0;

  // ---------- validation state <-> ARIA (pattern: accessible-error-announcement) ----------
  function syncField(el, announce) {
    if (!el.matches || !el.matches('input, select, textarea') || el.type === 'hidden') return;
    const err = el.dataset.error ? document.getElementById(el.dataset.error) : null;
    const invalid = el.matches(':user-invalid');
    el.setAttribute('aria-invalid', invalid ? 'true' : 'false');
    if (!err) return;
    const ids = (el.getAttribute('aria-describedby') || '').split(/\s+/).filter((x) => x && x !== err.id);
    if (invalid) {
      err.textContent = el.validationMessage;
      ids.push(err.id);
      if (announce && !el._announced) {
        el._announced = true;
        const lbl = el.labels && el.labels[0] ? `${el.labels[0].textContent.trim()} : ` : '';
        setTimeout(() => ($('announcer').textContent = lbl + el.validationMessage), 500);
      }
    } else {
      err.textContent = '';
      el._announced = false;
    }
    if (ids.length) el.setAttribute('aria-describedby', ids.join(' '));
    else el.removeAttribute('aria-describedby');
  }

  // ---------- field rendering ----------
  function fieldHtml(f, v) {
    const id = `ff-${++seq}`;
    f._id = id;
    const req = f.required ? ' required' : '';
    const help = f.help ? `<p class="f-help" id="${id}-help">${esc(f.help)}</p>` : `<p class="f-help" id="${id}-help" hidden></p>`;
    const err = `<p class="f-err" id="${id}-err"></p>`;
    const desc = ` aria-describedby="${id}-help" data-error="${id}-err"`;
    const cls = `field-row${f.span ? ' span' : ''}`;
    const reqMark = f.required ? '<span class="req" aria-hidden="true">*</span>' : '';
    const attrs = [f.min != null ? `min="${f.min}"` : '', f.max != null ? `max="${f.max}"` : '', f.step != null ? `step="${f.step}"` : '', f.placeholder ? `placeholder="${esc(f.placeholder)}"` : '', f.autocomplete ? `autocomplete="${f.autocomplete}"` : 'autocomplete="off"', f.inputmode ? `inputmode="${f.inputmode}"` : ''].filter(Boolean).join(' ');
    if (f.type === 'checkbox') {
      return `<div class="${cls} check-row" data-field="${f.name}"><label class="check"><input type="checkbox" id="${id}" name="${f.name}"${v ? ' checked' : ''}${req}${desc}><span>${esc(f.label)}</span></label>${help}${err}</div>`;
    }
    if (f.type === 'checks' || f.type === 'radio') {
      const multi = f.type === 'checks';
      const vals = multi ? new Set(v || []) : new Set([v]);
      return `<fieldset class="${cls} choice-group${f.cards ? ' cards' : ''}" data-field="${f.name}" id="${id}"><legend>${esc(f.label)}${reqMark}</legend>${help}
        <div class="choices">${(f.options || []).map(opt).map((o, i) => `<label class="choice"><input type="${multi ? 'checkbox' : 'radio'}" name="${f.name}" value="${esc(o.value)}"${vals.has(o.value) ? ' checked' : ''}${!multi && f.required && i === 0 ? ' required' : ''}${i === 0 ? ` data-error="${id}-err"` : ''}>${o.icon ? icon(o.icon) : ''}<span>${esc(o.label)}${o.sub ? `<small>${esc(o.sub)}</small>` : ''}</span></label>`).join('')}</div>${err}</fieldset>`;
    }
    const label = `<label for="${id}">${esc(f.label)}${reqMark}${f.unit ? ` <span class="muted">(${esc(f.unit)})</span>` : ''}</label>`;
    let control;
    if (f.type === 'select') {
      control = `<select id="${id}" name="${f.name}"${req}${desc}>${f.placeholder !== false ? `<option value="">${esc(f.placeholder || 'Choisir…')}</option>` : ''}${(f.options || []).map(opt).map((o) => `<option value="${esc(o.value)}"${String(o.value) === String(v == null ? '' : v) ? ' selected' : ''}>${esc(o.label)}</option>`).join('')}</select>`;
    } else if (f.type === 'textarea') {
      control = `<textarea id="${id}" name="${f.name}" rows="${f.rows || 3}"${req}${desc} ${attrs}>${esc(v)}</textarea>`;
    } else {
      const type = f.type === 'number' ? 'text' : f.type || 'text';
      const val = f.type === 'date' ? iso(v) : f.type === 'time' ? hhmm(v) : v == null ? '' : f.type === 'number' ? String(v).replace('.', ',') : v;
      const list = f.list ? ` list="${id}-list"` : '';
      const num = f.type === 'number' ? ` inputmode="decimal" pattern="-?[0-9]+([.,][0-9]+)?" title="Nombre, par exemple 1,5" data-num="1"` : '';
      control = `<input type="${type}" id="${id}" name="${f.name}" value="${esc(val)}"${list}${num}${req}${desc} ${attrs}>${f.list ? `<datalist id="${id}-list">${f.list.map((x) => `<option value="${esc(x)}"></option>`).join('')}</datalist>` : ''}`;
    }
    return `<div class="${cls}" data-field="${f.name}">${label}${control}${help}${err}</div>`;
  }

  // ---------- reading values ----------
  function readValues(form, spec) {
    const out = {};
    for (const sec of spec.sections) for (const f of sec.fields) {
      if (f.type === 'checks') { out[f.name] = [...form.querySelectorAll(`input[name="${f.name}"]:checked`)].map((i) => i.value); continue; }
      if (f.type === 'radio') { const c = form.querySelector(`input[name="${f.name}"]:checked`); out[f.name] = c ? c.value : ''; continue; }
      const el = form.elements[f.name];
      if (!el) continue;
      if (f.type === 'checkbox') out[f.name] = el.checked;
      else if (f.type === 'date') out[f.name] = el.value ? C.toSerial(el.value) : null;
      else if (f.type === 'time') out[f.name] = el.value ? C.toTime(el.value) : null;
      else if (f.type === 'number') out[f.name] = el.value.trim() === '' ? null : C.num(el.value);
      else out[f.name] = el.value.trim();
    }
    return out;
  }

  // ---------- form dialog ----------
  let current = null;

  function open(spec) {
    const dlg = $('formDlg');
    const form = $('formEl');
    const values = Object.assign({}, spec.values || {});
    seq = 0;
    form.innerHTML = `
      <header class="modal-head">
        ${spec.icon ? `<span class="modal-ico">${icon(spec.icon)}</span>` : ''}
        <div><h2 id="formTitle">${esc(spec.title)}</h2>${spec.subtitle ? `<p>${esc(spec.subtitle)}</p>` : ''}</div>
        <button type="button" class="icon-btn" data-act="close" aria-label="Fermer">${icon('x')}</button>
      </header>
      <div class="modal-body">
        <div class="form-alert" id="formAlert" hidden></div>
        ${spec.sections.map((sec) => `<fieldset class="form-sec"${sec.id ? ` data-sec="${sec.id}"` : ''}${sec.hidden ? ' hidden' : ''}><legend>${sec.icon ? icon(sec.icon) : ''}${esc(sec.legend)}</legend>${sec.hint ? `<p class="sec-hint">${sec.hint}</p>` : ''}<div class="form-grid">${sec.fields.map((f) => fieldHtml(f, values[f.name])).join('')}</div></fieldset>`).join('')}
      </div>
      <footer class="modal-foot">
        <div class="discard" id="discardBox" hidden><span id="discardMsg">Les modifications ne sont pas enregistrées.</span><button type="button" class="btn danger-ghost" data-act="discard" aria-describedby="discardMsg">Abandonner les modifications</button></div>
        ${spec.onDelete ? `<button type="button" class="btn danger-ghost" data-act="delete">${icon('trash-2')}${esc(spec.deleteLabel || 'Supprimer')}</button>` : ''}
        <span class="grow"></span>
        <button type="button" class="btn" data-act="close">Annuler</button>
        <button type="submit" class="btn primary">${icon('check')}${esc(spec.submitLabel || 'Enregistrer')}</button>
      </footer>`;
    paint();
    const fields = new Map();
    spec.sections.forEach((sec) => sec.fields.forEach((f) => fields.set(f.name, f)));
    const api = {
      form,
      values: () => readValues(form, spec),
      el: (name) => form.querySelector(`[data-field="${name}"]`),
      input: (name) => form.elements[name],
      show(name, on) { const el = api.el(name); if (el) { el.hidden = !on; el.querySelectorAll('input,select,textarea').forEach((i) => (i.disabled = !on)); } },
      showSection(id, on) { const el = form.querySelector(`[data-sec="${id}"]`); if (el) { el.hidden = !on; el.disabled = !on; } },
      setHelp(name, text, tone) {
        const f = fields.get(name);
        const h = f && document.getElementById(`${f._id}-help`);
        if (!h) return;
        h.hidden = !text;
        h.textContent = text || '';
        h.className = `f-help${tone ? ` ${tone}` : ''}`;
      },
      setOptions(name, options) {
        const f = fields.get(name);
        if (!f) return;
        if (f.list) f.list = options;
        else f.options = options;
        const cur = readValues(form, { sections: [{ fields: [f] }] })[name];
        const holder = api.el(name);
        const tmp = document.createElement('div');
        tmp.innerHTML = fieldHtml(f, f.type === 'checks' ? (cur || []).filter((x) => options.map(opt).some((o) => o.value === x)) : cur);
        holder.replaceWith(tmp.firstElementChild);
        paint();
      },
      setValidity(name, msg) {
        const f = fields.get(name);
        if (!f) return;
        const el = f.type === 'checks' || f.type === 'radio' ? form.querySelector(`input[name="${name}"]`) : form.elements[name];
        if (el && el.setCustomValidity) el.setCustomValidity(msg || '');
      },
      alert(kind, html) {
        const a = $('formAlert');
        a.hidden = !html;
        a.className = `form-alert notice ${kind || ''}`;
        a.innerHTML = html ? `${icon(kind === 'warn' ? 'triangle-alert' : 'info')}<div>${html}</div>` : '';
        paint();
      },
    };
    const me = { spec, api, dirty: false, warned: false, closing: false };
    // Replacing a form that was never closed: drop its listener. A form that is closing keeps
    // its listener so its onClosed still runs when its (asynchronous) 'close' event arrives.
    if (current && !current.closing) current.dispose();
    current = me;
    // Each form only reacts to the close it requested itself.
    const onClose = () => {
      if (!me.closing) return;
      dlg.removeEventListener('close', onClose);
      if (current === me) {
        current = null;
        if (!dlg.open) form.innerHTML = '';
      }
      if (spec.onClosed) spec.onClosed(me.rv || '');
    };
    dlg.addEventListener('close', onClose);
    me.dispose = () => dlg.removeEventListener('close', onClose);
    me.close = (rv) => { me.closing = true; me.rv = rv || ''; dlg.close(rv); };
    const changed = (e) => {
      me.dirty = true;
      if (me.warned) { me.warned = false; $('discardBox').hidden = true; }
      if (spec.onChange) spec.onChange(api.values(), api, e && e.target && e.target.name);
    };
    form.oninput = changed;
    form.onchange = changed;
    form.onsubmit = async (e) => {
      e.preventDefault();
      if (spec.onChange) spec.onChange(api.values(), api, null);
      if (!form.checkValidity()) { form.reportValidity(); return; }
      const btn = form.querySelector('button[type="submit"]');
      btn.disabled = true;
      try {
        await spec.onSave(api.values(), api);
        me.dirty = false;
        me.close('saved');
      } catch (err) {
        api.alert('warn', esc(err.message));
        $('formAlert').scrollIntoView({ block: 'nearest' });
      } finally {
        btn.disabled = false;
      }
    };
    form.onclick = async (e) => {
      const b = e.target.closest('button[data-act]');
      if (!b) return;
      if (b.dataset.act === 'close') tryClose();
      else if (b.dataset.act === 'discard') { me.dirty = false; me.close(); }
      else if (b.dataset.act === 'delete') {
        const ok = await confirm({ title: spec.deleteTitle || 'Supprimer cet élément ?', text: spec.deleteText || 'Cette action est définitive.', confirmLabel: spec.deleteLabel || 'Supprimer', danger: true });
        if (!ok) return;
        try {
          await spec.onDelete();
          me.dirty = false;
          me.close('deleted');
        } catch (err) {
          api.alert('warn', esc(err.message));
        }
      }
    };
    if (spec.onChange) spec.onChange(api.values(), api, null);
    form.querySelectorAll('input, select, textarea').forEach((el) => el.setAttribute('aria-invalid', 'false'));
    me.dirty = false;
    if (!dlg.open) dlg.showModal();
    const first = form.querySelector('.modal-body input:not([type=hidden]):not(:disabled), .modal-body select:not(:disabled), .modal-body textarea:not(:disabled)');
    if (first) first.focus();
    form.querySelector('.modal-body').scrollTop = 0;
  }

  function tryClose() {
    const dlg = $('formDlg');
    if (current && current.dirty && !current.warned) {
      current.warned = true;
      $('discardBox').hidden = false;
      $('formEl').querySelector('[data-act="discard"]').focus();
      return;
    }
    if (current) current.close();
    else dlg.close();
  }

  // ---------- confirmation ----------
  function confirm({ title, text, confirmLabel = 'Confirmer', danger = false }) {
    const dlg = $('confirmDlg');
    $('confirmTitle').textContent = title;
    $('confirmText').textContent = text || '';
    const ok = $('confirmOk');
    ok.textContent = confirmLabel;
    ok.className = `btn ${danger ? 'danger' : 'primary'}`;
    dlg.returnValue = '';
    dlg.showModal();
    $('confirmCancel').focus();
    // Resolve on the button's submit (synchronous) or on any other close (Esc, click outside).
    return new Promise((resolve) => {
      const form = dlg.querySelector('form');
      let done = false;
      const settle = (v) => {
        if (done) return;
        done = true;
        form.removeEventListener('submit', onSubmit);
        resolve(v);
      };
      const onSubmit = (e) => settle(!!e.submitter && e.submitter.value === 'ok');
      form.addEventListener('submit', onSubmit);
      dlg.addEventListener('close', () => settle(dlg.returnValue === 'ok'), { once: true });
    });
  }

  function init() {
    const dlg = $('formDlg');
    // Esc / back gesture: warn once about unsaved changes, then let it close.
    dlg.addEventListener('cancel', (e) => {
      if (current && current.dirty && !current.warned) { e.preventDefault(); tryClose(); return; }
      if (current) current.closing = true;
    });
    document.addEventListener('blur', (e) => e.target && syncField(e.target, true), true);
    document.addEventListener('input', (e) => { if (e.target.hasAttribute && e.target.hasAttribute('aria-invalid')) syncField(e.target, false); });
    document.addEventListener('change', (e) => syncField(e.target, true), true);
    document.addEventListener('invalid', (e) => { e.target._announced = true; syncField(e.target, false); }, true);
  }

  window.GMForms = { open, confirm, init, isOpen: () => !!current };
})();
