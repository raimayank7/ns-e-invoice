/* NS bulk: many invoices (or e-way bills) at once from an Excel or CSV sheet, or a Tally XML export, made into ONE file for the portal.
   <section id="bulk" data-bulk="einv"> = e-invoice site, data-bulk="ewb" = e-way bill site (with the Part B vehicle file).
   The work (reading, checks, the files) is in bulkcore.js (NSBulkCore, tested in Node); this file is the screen.
   Our template's headings are English (it is a data file, like the printed bills). Nothing leaves the browser: the user uploads the file
   on the portal themselves. E-way bill keys and codes: NIC "EWB JSON ATTRIBUTES - Version (1.0.0621)" and "Vehicle Updation Attributes -
   Version (1.0.0621)", docs.ewaybillgst.gov.in › Bulk Generation Tools (read 5 Nov 2026). */
(function () {
  'use strict';
  var $ = NS.$, $$ = NS.$$, esc = NS.esc, t = NS.t;
  var C = window.NSBulkCore;
  var EWB_VERSION = '1.0.0621';
  var REASONS = [4, 1, 2, 3];   /* Part B: first time, break down, transhipment, others */
  var SHOW = 200;               /* documents drawn at once in the list (Show all draws the rest) */

  var box, mode, batch = [], skipped = [], lastRows = null, lastName = '', manual = null, sampleOn = false, showAll = false, onlyBad = false, lastReplies = [];
  function digits(s) { return String(s || '').replace(/\D/g, ''); }
  function key(s) { return String(s || '').trim().toLowerCase(); }

  /* ---------- screen ---------- */
  function aato() { var s = $('#bk-aato', box); return s ? s.value : ''; }
  function opt(v, l, cur) { return '<option value="' + esc(v) + '"' + (cur === v ? ' selected' : '') + '>' + esc(l) + '</option>'; }
  function render() {
    var ewb = mode === 'ewb', saved = ''; try { saved = localStorage.getItem('ns-aato') || ''; } catch (e) {}
    box.innerHTML =
      '<p class="bk-local">' + esc(t('bk.local')) + '</p>' +
      '<div class="bk-steps">' +
      '<div class="card"><h3>1. ' + esc(t('bk.s1')) + '</h3><div class="fg">' +
        '<div class="f wide"><label for="bk-biz">' + esc(t('bk.biz')) + '</label><select id="bk-biz"><option value="">' + esc(t('bk.loading')) + '</option></select><small class="hint">' + esc(t('bk.h.biz')) + ' <a href="' + hub() + '#biz" target="_blank" rel="noopener">' + esc(t('bk.addbiz')) + ' ›</a></small></div>' +
        '<div class="f wide"><label for="bk-aato">' + esc(t('ei.aato')) + '</label><select id="bk-aato">' + ['', 'lt5', '5to10', '10p'].map(function (k) { return opt(k, t('ei.a.' + (k || 'none')), saved); }).join('') + '</select>' + (ewb ? '<small class="hint">' + esc(t('bk.h.aatoe')) + '</small>' : '') + '</div>' +
        '<div class="f wide"><label class="chk"><input id="bk-round" type="checkbox" checked> ' + esc(t('l.round')) + '</label></div>' +
      '</div></div>' +
      '<div class="card"><h3>2. ' + esc(t('bk.s2b')) + '</h3><p class="hint">' + esc(t(ewb ? 'bk.h.s2e' : 'bk.h.s2')) + ' ' + esc(t('bk.h.tally')) + '</p>' +
        '<div class="row"><label class="btn primary" for="bk-file">' + esc(t('bk.choose2')) + '</label><input id="bk-file" type="file" multiple accept=".xlsx,.csv,.txt,.xml,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,text/csv,text/xml,application/xml" hidden>' +
        '<button type="button" class="btn" id="bk-tpl">' + esc(t('bk.tpl')) + '</button>' +
        '<button type="button" class="btn" id="bk-sample">' + esc(t('bk.sample')) + '</button></div>' +
        '<div class="bk-drop" id="bk-drop" tabindex="-1">' + esc(t('bk.drop')) + '</div>' +
        '<div class="bk-out" id="bk-out" aria-live="polite"></div></div>' +
      '<div class="card"><h3>3. ' + esc(t('bk.s3')) + '</h3><p class="hint">' + esc(t(ewb ? 'bk.h.s3e' : 'bk.h.s3')) + ' ' + esc(t('bk.h.s3many')) + '</p>' +
        '<div class="row"><label class="btn" for="bk-reply">' + esc(t('bk.loadmany')) + '</label><input id="bk-reply" type="file" multiple accept=".json,.xlsx,.csv,application/json" hidden></div>' +
        '<div class="bk-out" id="bk-rep" aria-live="polite"></div></div>' +
      (ewb ? partBHtml() : '') +
      '</div>';
    $('#bk-tpl', box).addEventListener('click', function () { saveBlob(NS.xlsx.write(C.template(mode)), ewb ? 'e-way-bill-template.xlsx' : 'e-invoice-template.xlsx'); });
    $('#bk-sample', box).addEventListener('click', trySample);
    $('#bk-biz', box).addEventListener('change', function () { sampleOn = false; });
    $('#bk-file', box).addEventListener('change', function (e) { var f = Array.prototype.slice.call(e.target.files); e.target.value = ''; if (f.length) readFiles(f); });
    $('#bk-reply', box).addEventListener('change', function (e) { var f = Array.prototype.slice.call(e.target.files); e.target.value = ''; if (f.length) readReplies(f); });
    var drop = $('#bk-drop', box);
    ['dragenter', 'dragover'].forEach(function (k) { drop.addEventListener(k, function (e) { e.preventDefault(); drop.classList.add('on'); }); });
    ['dragleave', 'drop'].forEach(function (k) { drop.addEventListener(k, function () { drop.classList.remove('on'); }); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); var f = Array.prototype.slice.call(e.dataTransfer.files || []); if (f.length) readFiles(f); });
    var a = $('#bk-aato', box); a.addEventListener('change', function () { try { localStorage.setItem('ns-aato', a.value); } catch (e) {} if (lastRows) redo(); });
    $('#bk-round', box).addEventListener('change', function () { if (lastRows) redo(); });
    if (ewb) wirePartB();
    fillBiz();
  }
  function hub() { return (window.NS_CFG && NS_CFG.vault ? new URL(NS_CFG.vault, location.href).origin : 'https://my.niveshsahayak.in') + '/'; }
  var bizList = [];
  function fillBiz() {
    var draft = null; try { draft = (JSON.parse(localStorage.getItem('ns-draft-invoice') || 'null') || {}).doc; } catch (e) {}
    NSStore.ready().then(function () { return NSStore.list('biz'); }).then(function (all) {
      bizList = all.filter(function (b) { return b.name; }).sort(function (a, b) { return a.name.localeCompare(b.name); });
      if (draft && draft.seller && draft.seller.name && !bizList.some(function (b) { return key(b.name) === key(draft.seller.name); })) bizList.unshift(Object.assign({ id: '_form' }, draft.seller));
      var sel = $('#bk-biz', box);
      sel.innerHTML = bizList.length ? bizList.map(function (b) { return opt(b.id, b.name + (b.gstin ? ' · ' + b.gstin : '') + (b.id === '_form' ? ' (' + t('bk.fromform') + ')' : ''), ''); }).join('') : '<option value="">' + esc(t('bk.nobiz')) + '</option>';
      var x = mode === 'ewb' && seller(), pl = $('#pb-place', box);
      if (x && pl && !pl.value) { pl.value = x.city || ''; $('#pb-state', box).value = x.state || ''; }
    }).catch(function () { var sel = $('#bk-biz', box); if (sel) sel.innerHTML = '<option value="">' + esc(t('bk.nobiz')) + '</option>'; });
  }
  function seller() {
    if (sampleOn) return C.sample(mode).seller;
    var sel = $('#bk-biz', box), id = sel && sel.value, b = bizList.filter(function (x) { return x.id === id; })[0];
    if (!b) return null;
    var s = JSON.parse(JSON.stringify(b)); delete s.id; delete s.created; delete s.updated;
    ['name', 'addr', 'city', 'pin', 'state', 'gstin', 'phone', 'email'].forEach(function (k) { if (s[k] == null) s[k] = ''; });
    return s;
  }
  function saveBlob(blob, name) {
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
  }
  function download(name, data) { saveBlob(new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' }), name); }

  /* ---------- reading the files ---------- */
  function trySample() {
    var s = C.sample(mode);
    sampleOn = true; manual = null;
    lastRows = { kind: 'sheet', rows: s.rows }; lastName = 'sample';
    redo();
  }
  function readFiles(files) {
    var out = $('#bk-out', box);
    sampleOn = false; manual = null;
    if (!seller()) { out.innerHTML = '<p class="warn">' + esc(t('bk.needbiz')) + '</p>'; $('#bk-biz', box).focus(); return; }
    out.innerHTML = '<p>' + esc(t('bk.reading')) + '</p>';
    var xml = files.filter(function (f) { return /\.xml$/i.test(f.name) || /xml/.test(f.type); });
    lastName = files.map(function (f) { return f.name; }).join(', ');
    if (xml.length) {
      if (!window.NSTallyIn) { out.innerHTML = '<p class="warn">' + esc(t('bk.badfile')) + '</p>'; return; }
      Promise.all(xml.map(function (f) { return f.arrayBuffer().then(function (b) { return NSTallyIn.decode(new Uint8Array(b)); }); })).then(function (texts) {
        lastRows = { kind: 'tally', texts: texts }; redo();
      }).catch(function () { out.innerHTML = '<p class="warn">' + esc(t('bk.tl.bad')) + '</p>'; });
      return;
    }
    NS.xlsx.read(files[0]).then(function (rows) { lastRows = { kind: 'sheet', rows: rows }; redo(); })
      .catch(function (e) { out.innerHTML = '<p class="warn">' + esc(t(e && e.message === 'xls' ? 'bk.xls' : 'bk.badfile')) + '</p>'; });
  }
  function opts() { return { mode: mode, aato: aato(), round: $('#bk-round', box).checked, map: manual }; }
  /* read (again) what was chosen: after a new file, a change of turnover / round off, or the Columns screen */
  function redo() {
    var out = $('#bk-out', box), s = seller(), o = opts();
    if (!s) { out.innerHTML = '<p class="warn">' + esc(t('bk.needbiz')) + '</p>'; return; }
    var r = lastRows.kind === 'tally' ? C.fromTally(lastRows.texts, s, o) : C.fromSheet(lastRows.rows, s, o);
    if (r.err) {
      if (r.head && lastRows.kind === 'sheet') return columns(r.head);
      out.innerHTML = '<p class="warn">' + esc(r.err) + '</p>'; return;
    }
    skipped = r.skipped || [];
    out.innerHTML = '<p>' + esc(t('bk.checking', { n: r.docs.length })) + '</p>';
    var have = Promise.resolve({});
    if (!sampleOn && window.NSStore) have = NSStore.list('doc').then(C.haveIndex).catch(function () { return {}; });
    have.then(function (H) {
      o.have = H;
      return NS.eachSeq(r.docs, function (x) { C.checkOne(x, o); }, { label: t('bk.checkl') });
    }).then(function (res) {
      batch = res.cancelled ? r.docs.filter(function (x) { return x.bad; }) : r.docs;
      showAll = false; onlyBad = false;
      showBatch(r);
    }).catch(function (e) { out.innerHTML = '<p class="warn">' + esc(t('bk.badfile')) + '</p>'; if (window.console) console.error(e); });
  }
  /* the Columns screen: our headings were not found, so the user says which column is which */
  function columns(H) {
    var out = $('#bk-out', box), cols = C.colsFor(mode), heads = H.heads || [];
    var sel = function (k) { var cur = H.map[k]; return '<select data-col="' + k + '"><option value="">' + esc(t('bk.col.none')) + '</option>' + heads.map(function (h, i) { return '<option value="' + i + '"' + (cur === i ? ' selected' : '') + '>' + esc(String(h || '').slice(0, 40) || t('bk.col.n', { n: i + 1 })) + '</option>'; }).join('') + '</select>'; };
    out.innerHTML = '<p class="warn">' + esc(t('bk.col.need')) + '</p><div class="fg bk-cols">' +
      cols.map(function (c) { return '<div class="f"><label>' + esc(c[1]) + (['no', 'date', 'desc'].indexOf(c[0]) >= 0 ? ' *' : '') + '</label>' + sel(c[0]) + '</div>'; }).join('') +
      '</div><p class="row"><button type="button" class="btn primary" id="bk-colgo">' + esc(t('bk.col.go')) + '</button></p>';
    $('#bk-colgo', box).addEventListener('click', function () {
      manual = {}; $$('[data-col]', out).forEach(function (s) { manual[s.getAttribute('data-col')] = s.value; });
      redo();
    });
  }
  function showBatch(r) {
    var out = $('#bk-out', box), ok = batch.filter(function (x) { return !x.bad.length; }), ewb = mode === 'ewb';
    if (!batch.length) { out.innerHTML = '<p class="warn">' + esc(t('bk.none')) + '</p>' + skippedHtml(); return; }
    var list = onlyBad ? batch.filter(function (x) { return x.bad.length; }) : batch, shown = showAll ? list : list.slice(0, SHOW);
    out.innerHTML = (sampleOn ? '<p class="bk-smp">' + esc(t('bk.smpnote')) + '</p>' : '') +
      (r && r.company ? '<p class="hint">' + esc(t('bk.tl.co', { c: r.company })) + '</p>' : '') +
      '<p><b>' + esc(t('bk.read', { n: batch.length, ok: ok.length, bad: batch.length - ok.length })) + '</b></p>' + skippedHtml() +
      '<p class="row"><button type="button" class="btn sm" id="bk-errx">' + esc(t('bk.errx')) + '</button>' +
      (batch.length > ok.length ? '<label class="chk"><input id="bk-onlybad" type="checkbox"' + (onlyBad ? ' checked' : '') + '> ' + esc(t('bk.onlybad')) + '</label>' : '') + '</p>' +
      '<div class="tbl"><table class="bk-t"><thead><tr><th>' + esc(t('bk.c.no')) + '</th><th>' + esc(t('l.date')) + '</th><th>' + esc(t('bk.c.to')) + '</th><th class="n">' + esc(t('bk.c.val')) + '</th><th>' + esc(t('bk.c.st')) + '</th></tr></thead><tbody>' +
      shown.map(function (x) {
        var d = x.d;
        return '<tr class="' + (x.bad.length ? 'bad' : 'ok') + '"><td class="mono">' + esc(d.no || '—') + (x.line ? '<br><small class="soft">' + esc(t('bk.rown', { r: x.line })) + '</small>' : '') + '</td><td>' + esc(d.date ? NS.fmtDate(d.date) : '—') + '</td><td>' + esc(d.buyer.name || '—') + '</td><td class="n">' + NS.money(x.tt.grand) + '</td><td>' +
          (x.bad.length ? '<details' + (batch.length <= 20 ? ' open' : '') + '><summary>' + esc(t('bk.fixn', { n: x.bad.length })) + '</summary><ul>' + x.bad.map(function (p) { return '<li>' + esc(p) + '</li>'; }).join('') + '</ul></details>' : esc(t('bk.ready'))) +
          (x.note.length ? '<details class="soft"><summary>' + esc(t('bk.notes', { n: x.note.length })) + '</summary><ul>' + x.note.map(function (p) { return '<li>' + esc(p) + '</li>'; }).join('') + '</ul></details>' : '') + '</td></tr>';
      }).join('') + '</tbody></table></div>' +
      (shown.length < list.length ? '<p><button type="button" class="btn sm" id="bk-all">' + esc(t('bk.showall', { n: list.length })) + '</button></p>' : '') +
      (ok.length ? (ewb || sampleOn ? '' : '<p><label class="chk"><input id="bk-save" type="checkbox" checked> ' + esc(t('bk.save')) + '</label></p>') +
        '<p class="row"><button type="button" class="btn primary" id="bk-make">' + esc(t(ewb ? 'bk.makeewb' : 'bk.make', { n: ok.length })) + '</button></p>' +
        (batch.length > ok.length ? '<p class="hint">' + esc(t('bk.leftout')) + '</p>' : '') : '<p class="hint">' + esc(t('bk.fixfirst')) + '</p>');
    var mk = $('#bk-make', box); if (mk) mk.addEventListener('click', function () { make(ok); });
    $('#bk-errx', box).addEventListener('click', function () { var sh = [[ewb ? 'E-way bills' : 'Invoices', C.problemSheet(batch)]]; if (skipped.length) sh.push(['Rows left out', C.skippedSheet(skipped)]); saveBlob(NS.xlsx.write(sh), (ewb ? 'e-way-bill' : 'e-invoice') + '-check-' + NS.today() + '.xlsx'); });
    var all = $('#bk-all', box); if (all) all.addEventListener('click', function () { showAll = true; showBatch(r); });
    var ob = $('#bk-onlybad', box); if (ob) ob.addEventListener('change', function () { onlyBad = ob.checked; showBatch(r); });
  }
  function skippedHtml() {
    if (!skipped.length) return '';
    return '<details class="soft"><summary>' + esc(t('bk.skipn', { n: skipped.length })) + '</summary><ul>' + skipped.slice(0, 100).map(function (s) { return '<li>' + esc((s.r ? t('bk.rown', { r: s.r }) + ': ' : '') + s.why) + '</li>'; }).join('') + '</ul></details>';
  }
  function make(ok) {
    var ewb = mode === 'ewb', data = C.file(ok, mode), name = (sampleOn ? 'sample-' : '') + (ewb ? 'e-way-bills-' : 'e-invoices-') + NS.today() + '.json';
    download(name, data);
    var msg = t(ewb ? 'bk.madeewb' : 'bk.made', { file: name, n: ok.length });
    var save = $('#bk-save', box);
    (save && save.checked ? saveAll(ok) : Promise.resolve(null)).then(function (n) {
      NS.toast(msg);
      var p = document.createElement('p'); p.className = 'ok'; p.textContent = msg + (n != null ? ' ' + t('bk.savedn', { n: n }) : '');
      $('#bk-out', box).appendChild(p);
    });
  }
  /* the invoices go into My data like ones typed in the invoice maker; one already saved with the same number, year and business is updated */
  function saveAll(ok) {
    return NSStore.list('doc').then(function (all) {
      return NS.eachSeq(ok, function (x) {
        var d = x.d, have = all.filter(function (r) { return (r.type || 'invoice') === 'invoice' && key(r.no) === key(d.no) && NS.fy(r.date) === NS.fy(d.date) && key(r.seller && r.seller.name) === key(d.seller.name); })[0];
        var rec = JSON.parse(JSON.stringify(d)); rec.total = x.tt.grand;
        rec.items.forEach(function (it) { delete it._r; });
        if (have) { rec.id = have.id; rec.created = have.created; if (have.irn) rec.irn = have.irn; rec.status = have.status || rec.status; }
        return NSStore.put('doc', rec).then(function (r) { x.id = r.id; });
      }, { label: t('bk.savingl') });
    }).then(function () { NS.folderWrite(false); NS.checkReminder(); return ok.length; }).catch(function (e) { NS.toast(t('nosave', { e: e.message }), 'bad'); return null; });
  }

  /* ---------- the portal's replies: JSON (any shape) or Excel/CSV result sheets, many files at once ---------- */
  function readReplies(files) {
    var out = $('#bk-rep', box);
    out.innerHTML = '<p>' + esc(t('bk.reading')) + '</p>';
    var bad = [];
    Promise.all(files.map(function (f) {
      var p = /\.json$/i.test(f.name) || f.type === 'application/json' ? f.text().then(function (s) { return NS.einv.findAll(JSON.parse(s.replace(/^﻿/, '')), []); })
        : NS.xlsx.read(f).then(function (rows) { return NS.einv.findAll(C.replyFromRows(rows), []); });
      return p.catch(function () { bad.push(f.name); return []; });
    })).then(function (lists) {
      var found = [].concat.apply([], lists);
      if (!found.length) { out.innerHTML = '<p class="warn">' + esc(t('ei.noreply')) + '</p>'; return; }
      var all = window.NSStore ? NSStore.list('doc').catch(function () { return []; }) : Promise.resolve([]);
      return all.then(function (docs) { return attach(found, docs, bad); });
    });
  }
  function attach(found, all, bad) {
    var out = $('#bk-rep', box), ewb = mode === 'ewb', R = C.replies(found, batch, all);
    lastReplies = R.rows;
    NS.eachSeq(R.save, function (rec) { return NSStore.put('doc', rec); }, { label: t('bk.savingl') }).then(function () {
      var have = R.rows.filter(function (r) { return r.id; }).length, shown = R.rows.slice(0, SHOW);
      out.innerHTML = (bad.length ? '<p class="warn">' + esc(t('bk.repbad', { f: bad.join(', ') })) + '</p>' : '') +
        '<p><b>' + esc(t('bk.got', { n: R.rows.length, k: have })) + '</b></p>' +
        '<p class="row"><button type="button" class="btn sm" id="bk-repx">' + esc(t('bk.repx')) + '</button></p>' +
        '<div class="tbl"><table class="bk-t"><thead><tr><th>' + esc(t('bk.c.no')) + '</th><th>IRN</th><th>' + esc(t('ei.ack')) + '</th><th>' + esc(t('ei.ewbno')) + '</th><th></th></tr></thead><tbody>' +
        shown.map(function (r) {
          return '<tr><td class="mono">' + esc(r.no || '—') + '</td><td class="mono">' + esc(r.irn ? r.irn.slice(0, 16) + '…' : '—') + '</td><td class="mono">' + esc(r.ack || '—') + '</td><td class="mono">' + esc(r.ewb || '—') + (r.ewbTill ? '<br><small class="soft">' + esc(r.ewbTill) + '</small>' : '') + '</td><td>' +
            (r.id ? '<a href="?open=' + encodeURIComponent(r.id) + '#tool">' + esc(t('bk.open')) + '</a>' : '<span class="soft">' + esc(t('bk.notsaved')) + '</span>') + '</td></tr>';
        }).join('') + '</tbody></table></div>' +
        (shown.length < R.rows.length ? '<p class="hint">' + esc(t('bk.repmore', { n: R.rows.length - shown.length })) + '</p>' : '') +
        (R.rows.length > have && !ewb ? '<p class="hint">' + esc(t('bk.h.notsaved')) + '</p>' : '');
      $('#bk-repx', box).addEventListener('click', function () { saveBlob(NS.xlsx.write([['Replies', C.replySheet(lastReplies)]]), (ewb ? 'e-way-bill' : 'e-invoice') + '-numbers-' + NS.today() + '.xlsx'); });
      if (ewb) R.rows.forEach(function (r) { if (/^\d{12}$/.test(r.ewb)) addPartB({ ewb: r.ewb }); });
      if (R.save.length) NS.folderWrite(false);
    }).catch(function (e) { NS.toast(t('nosave', { e: e.message }), 'bad'); });
  }

  /* ---------- Part B: the vehicle file (Vehicle Updation, version 1.0.0621) ---------- */
  function partBHtml() {
    return '<div class="card" id="pb"><h3>4. ' + esc(t('bk.pb')) + '</h3><p class="hint">' + esc(t('bk.h.pb')) + '</p>' +
      '<div class="fg"><div class="f"><label for="pb-place">' + esc(t('bk.pbplace')) + '</label><input id="pb-place" dir="auto"></div>' +
      '<div class="f"><label for="pb-state">' + esc(t('l.state')) + '</label><select id="pb-state"></select></div>' +
      '<div class="f"><label for="pb-reason">' + esc(t('bk.pbreason')) + '</label><select id="pb-reason">' + REASONS.map(function (k) { return opt(String(k), t('bk.r.' + k), '4'); }).join('') + '</select></div>' +
      '<div class="f"><label for="pb-remark">' + esc(t('bk.pbremark')) + '</label><input id="pb-remark" maxlength="25" dir="auto"></div></div>' +
      '<div class="tbl"><table class="bk-t pb-t"><thead><tr><th>' + esc(t('ei.ewbno')) + '</th><th>' + esc(t('ei.mode')) + '</th><th>' + esc(t('ei.veh')) + '</th><th>' + esc(t('ei.tdoc')) + '</th><th>' + esc(t('ei.tdate')) + '</th><th></th></tr></thead><tbody id="pb-rows"></tbody></table></div>' +
      '<p class="row"><button type="button" class="btn" id="pb-add">' + esc(t('bk.pbadd')) + '</button><button type="button" class="btn primary" id="pb-make">' + esc(t('bk.pbmake')) + '</button></p>' +
      '<div class="bk-out" id="pb-out" aria-live="polite"></div></div>';
  }
  function addPartB(v) {
    var tb = $('#pb-rows', box); if (!tb) return;
    if (v.ewb && $$('[data-pb="ewb"]', tb).some(function (i) { return i.value === v.ewb; })) return;
    var first = $$('tr', tb)[0];
    if (first && !$('[data-pb="ewb"]', first).value && !$$('tr', tb)[1]) first.remove();
    var tr = document.createElement('tr'), n = tb.children.length;
    tr.innerHTML = '<td><input data-pb="ewb" inputmode="numeric" maxlength="12" dir="ltr" aria-label="' + esc(t('ei.ewbno')) + ' ' + (n + 1) + '" value="' + esc(v.ewb || '') + '"></td>' +
      '<td><select data-pb="mode" aria-label="' + esc(t('ei.mode')) + '">' + [1, 2, 3, 4].map(function (k) { return opt(String(k), t('ei.m.' + k), '1'); }).join('') + '</select></td>' +
      '<td><input data-pb="veh" maxlength="15" dir="ltr" placeholder="PB10AB1234" aria-label="' + esc(t('ei.veh')) + '"></td>' +
      '<td><input data-pb="doc" maxlength="16" dir="ltr" aria-label="' + esc(t('ei.tdoc')) + '"></td>' +
      '<td><input data-pb="date" type="date" aria-label="' + esc(t('ei.tdate')) + '"></td>' +
      '<td><button type="button" class="link" data-pb-del aria-label="' + esc(t('bk.pbdel')) + '">✕</button></td>';
    tb.appendChild(tr);
  }
  function wirePartB() {
    var st = $('#pb-state', box);
    st.innerHTML = '<option value="">' + esc(t('choose.state')) + '</option>' + NS.STATES.filter(function (s) { return s[0] !== '96'; }).map(function (s) { return '<option value="' + s[0] + '">' + s[1] + ' (' + s[0] + ')</option>'; }).join('');
    var s = seller(); if (s) { st.value = s.state || ''; $('#pb-place', box).value = s.city || ''; }
    $('#bk-biz', box).addEventListener('change', function () { var x = seller(); if (x) { st.value = x.state || ''; $('#pb-place', box).value = x.city || ''; } });
    addPartB({});
    $('#pb-add', box).addEventListener('click', function () { addPartB({}); var i = $$('[data-pb="ewb"]', box); i[i.length - 1].focus(); });
    $('#pb-rows', box).addEventListener('click', function (e) { var b = e.target.closest('[data-pb-del]'); if (b) b.closest('tr').remove(); });
    $('#pb-make', box).addEventListener('click', makePartB);
  }
  function makePartB() {
    var out = $('#pb-out', box), s = seller(), bad = [], list = [];
    if (!s || !NS.gstinCheck(s.gstin).ok || !s.gstin) bad.push(t('ev.ewbseller'));
    var place = $('#pb-place', box).value.trim(), state = $('#pb-state', box).value, reason = +$('#pb-reason', box).value, remark = $('#pb-remark', box).value.trim();
    if (place.length < 3) bad.push(t('bk.v.place'));
    if (!state) bad.push(t('ev.state'));
    if (reason === 3 && !remark) bad.push(t('bk.v.remark'));
    $$('#pb-rows tr', box).forEach(function (tr, i) {
      var g = function (k) { return $('[data-pb="' + k + '"]', tr).value.trim(); };
      var ewb = digits(g('ewb')), m = +g('mode'), veh = g('veh').toUpperCase().replace(/[\s-]/g, ''), doc = g('doc'), dt = g('date');
      if (!ewb && !veh && !doc) return;
      if (!/^\d{12}$/.test(ewb)) bad.push(t('bk.v.ewb', { n: i + 1 }));
      if (m === 1 && !veh) bad.push(t('bk.v.veh', { n: i + 1 }));
      if (veh && !/^[A-Z]{2}\d{1,2}[A-Z]{0,3}\d{4}$|^[A-Z]{3}\d{4}$|^\d{2}BH\d{4}[A-Z]{1,2}$|^TR[A-Z0-9]{6,13}$|^(DF|TM|BP|NP)[A-Z0-9]{5,13}$/.test(veh)) bad.push(t('bk.v.vehf', { n: i + 1 }));
      if ((m === 2 || m === 3) && !(doc && dt)) bad.push(t('bk.v.doc', { n: i + 1 }));
      if (m === 4 && !veh && !(doc && dt)) bad.push(t('bk.v.ship', { n: i + 1 }));
      var dmy = dt ? dt.split('-').reverse().join('/') : '';
      list.push({ ewbno: +ewb, transMode: m, vehicleType: 'R', vehicleNo: veh, docNo: doc, docDate: dmy, fromPlace: place.slice(0, 50), fromState: +state, reason: reason, remark: remark.slice(0, 25) });
    });
    if (!list.length) bad.push(t('bk.v.none'));
    if (bad.length) { out.innerHTML = '<p><b>' + esc(t('ev.fix')) + '</b></p><ul class="warn">' + bad.map(function (b) { return '<li>' + esc(b) + '</li>'; }).join('') + '</ul>'; return; }
    var name = 'e-way-bill-part-b-' + NS.today() + '.json';
    download(name, { version: EWB_VERSION, userGstin: s.gstin, vehicleUpdts: list });
    out.innerHTML = '<p class="ok">' + esc(t('bk.pbmade', { file: name, n: list.length })) + '</p>';
  }

  function init() {
    box = $('[data-bulk]');
    if (!box || !NS.einv || !NS.xlsx || !C) return;
    mode = box.getAttribute('data-bulk') === 'ewb' ? 'ewb' : 'einv';
    render();
    document.addEventListener('ns-lang', function () { render(); if (batch.length) showBatch(); });
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init); else init();
})();
