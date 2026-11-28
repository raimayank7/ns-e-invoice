/* NSBulkCore: the work behind the bulk screens of e-invoice.niveshsahayak.in and eway-bill.niveshsahayak.in (engine/bulk.js), with no
   screen of its own, so Node tests can run it (scripts/test_bulkeinv.js). Needs NS (core.js) and NS.einv (einv.js) for the per-document
   checks and the JSON itself; NSTallyIn (tallyin.js) only for Tally files.

   In: an Excel/CSV sheet with one row per item line (rows with the same document number make one document), or Tally XML
   (Day Book / Sales vouchers, with or without the Masters export; also Marg, Vyapar, myBillBook "Export to Tally" files).
   Out: documents with their problems (must fix before the file is made) and notes (worth a look), then ONE file for the portal:
   e-invoice = array of INV-01 v1.1 invoices (IRP bulk upload), e-way bill = {version, billLists} (NIC bulk tool 1.0.0621).
   Rules used here besides einv.js check() (read 28 Nov 2026):
   - GST rate list by the document's date: NS.gstRatesFor (28% left on 1 Feb 2026, Notification 19/2025-CT(Rate)).
   - HSN digits: 6 digits on invoices when the turnover (AATO) is above Rs 5 crore (Notification 78/2020-CT); the e-invoice IRP needs
     at least 6 for everyone it covers; the e-way bill portal takes 4 digits only from businesses up to Rs 5 crore.
   - PIN code and state: India Post's first two PIN digits belong to postal circles; a PIN outside the state's circle is a note (the
     portals check PIN against state and refuse a mismatch; our table is a coarse check, so it never blocks).
   - The same document number may be reported once in a financial year (IRP error 2150 "Duplicate IRN"): a number that already has an
     IRN in My data is a problem; a number used twice in one file with another buyer or date is a problem. */
(function (G) {
  'use strict';
  var NS = G.NS;
  function t(k, v) { return NS.t(k, v); }
  function num(v) { return NS.num(v); }
  function r2(n) { return NS.r2(n); }
  function norm(s) { return String(s || '').toLowerCase().replace(/[^a-z0-9%]/g, ''); }
  function digits(s) { return String(s || '').replace(/\D/g, ''); }
  function key(s) { return String(s || '').trim().toLowerCase(); }
  /* Indian and plain numbers: "1,23,456.50", "Rs. 250", "₹ 250/-", "18%" */
  function amt(v) { return String(v == null ? '' : v).replace(/^\s*(rs\.?|inr|₹)\s*/i, '').replace(/[₹,\s]/g, '').replace(/\/-$/, '').replace(/%$/, ''); }

  var EWB_VERSION = '1.0.0621';
  var SUBS = [[1, 'Supply'], [3, 'Export'], [4, 'Job Work'], [5, 'For Own Use'], [8, 'Others'], [9, 'SKD/CKD/Lots'], [10, 'Line Sales'], [11, 'Recipient Not Known'], [12, 'Exhibition or Fairs']];
  /* outward sub supply type -> document types the portal accepts (NIC Validations, Table 2) */
  var DOCS_FOR = { 1: ['INV', 'BIL'], 3: ['INV', 'BIL'], 4: ['CHL'], 5: ['CHL'], 8: ['CHL', 'OTH'], 9: ['INV', 'BIL', 'CHL'], 10: ['CHL'], 11: ['CHL', 'OTH'], 12: ['CHL'] };
  var TO_SELF = { 5: 1, 10: 1, 11: 1, 12: 1 };   /* goods stay with the sender: "To GSTIN" = own GSTIN */
  var MODES = { road: 1, rail: 2, air: 3, ship: 4, 'road cum ship': 4 };

  /* columns: key, heading in the template, example, aliases (headings written in other ways, also Tally / Busy / Marg register exports) */
  var COMMON = [
    ['no', 'Invoice No', 'INV/26-27/101', 'document no|doc no|bill no|invoice number|docno|voucher no|vch no|voucher number|inv no|bill number|vch/bill no|vchbill no|invoice no.|bill no.'],
    ['date', 'Invoice Date', '05/11/2026', 'document date|doc date|bill date|docdate|date|voucher date|vch date|inv date'],
    ['bgstin', 'Buyer GSTIN', '27AAPFU0939F1ZV', 'customer gstin|gstin of buyer|to gstin|recipient gstin|gstin|gstin/uin|party gstin|gst no|gstin uin of recipient|buyer gst no|party gstin no|gstin no'],
    ['bname', 'Buyer Name', 'Shree Traders', 'customer name|buyer legal name|party name|to name|recipient name|legal name|particulars|party|customer|buyer|receiver name'],
    ['baddr', 'Buyer Address', '12 Market Road', 'customer address|to address|address|party address|buyer address 1|address 1'],
    ['bcity', 'Buyer City', 'Pune', 'customer city|to place|city|place|town|location'],
    ['bpin', 'Buyer PIN', '411001', 'customer pin|to pincode|pincode|pin code|pin|party pincode|zip'],
    ['bstate', 'Buyer State Code', '27', 'customer state|to state|state code|state|party state|buyer state'],
    ['pos', 'Place of Supply', '27', 'pos|place of supply state|supply state'],
    ['desc', 'Item Description', 'Cotton yarn 40s', 'item|product|product name|description|item name|stock item|name of item|goods description|particulars of goods|item details|item description'],
    ['hsn', 'HSN/SAC', '52051200', 'hsn|hsn code|sac|hsn sac|hsn/sac code|hsnsac'],
    ['qty', 'Quantity', '100', 'qty|billed qty|quantity billed|nos'],
    ['unit', 'Unit', 'KGS', 'uqc|uom|units|unit of measure|per'],
    ['rate', 'Rate', '250', 'price|unit price|rate per unit|item rate|rate rs'],
    ['disc', 'Discount %', '0', 'discount|disc|disc %|discount percent'],
    ['gst', 'GST %', '5', 'gst rate|tax rate|rate of tax|igst %|gst|gst rate %|tax %|integrated tax rate']
  ];
  var EINV = [['sup', 'Supply Type', 'B2B', 'supply type code|suptyp'], ['rc', 'Reverse Charge', 'N', 'reverse charge y/n|regrev']];
  var EWB = [
    ['dtype', 'Document Type', 'INV', 'doc type|doctype'], ['sub', 'Sub Supply Type', 'Supply', 'sub type|subsupplytype|reason for movement'], ['subdesc', 'Sub Supply Description', '', 'sub supply desc'],
    ['dist', 'Distance km', '0', 'distance|approximate distance|km|distance in km'], ['mode', 'Transport Mode', 'Road', 'mode|trans mode|mode of transport'], ['veh', 'Vehicle No', 'MH12AB1234', 'vehicle|vehicle number|truck no|lorry no'],
    ['vtype', 'Vehicle Type', 'R', 'odc'], ['tid', 'Transporter ID', '', 'transporter gstin|transin|transporter id gstin'], ['tname', 'Transporter Name', '', 'transporter'],
    ['tdoc', 'Transport Doc No', '', 'lr no|rr no|transport document no|lr|gr no|bilty no'], ['tdate', 'Transport Doc Date', '', 'lr date|transport document date|gr date'],
    ['shaddr', 'Ship To Address', '', 'delivery address|dispatch to address|consignee address'], ['shcity', 'Ship To City', '', 'delivery city|ship to place|consignee city'], ['shpin', 'Ship To PIN', '', 'delivery pin|ship to pincode|consignee pin'], ['shstate', 'Ship To State Code', '', 'delivery state|ship to state|consignee state']
  ];
  var HELP = {
    no: 'Up to 16 letters, digits, / or -. Repeat it on every row of the same invoice (one row per item).',
    date: 'DD/MM/YYYY. Not a future date.',
    sup: 'B2B (sale to a registered business), SEZWP / SEZWOP (SEZ with / without IGST), EXPWP / EXPWOP (export with / without IGST), DEXP (deemed export).',
    rc: 'Y if the buyer pays the tax under reverse charge, else N.',
    bgstin: '15 characters. Leave empty for an export (and in an e-way bill for a buyer without GSTIN).',
    bname: 'Legal name as registered.', baddr: 'Building, street, area.', bcity: 'At least 3 letters.', bpin: '6 digits. Exports: the e-invoice uses 999999 by itself.',
    bstate: 'Two-digit GST state code (03 Punjab, 07 Delhi, 27 Maharashtra...) or the state name. Exports: 96.',
    pos: 'State code of the place of supply. Empty = the buyer\'s state.',
    desc: 'What you sold.', hsn: 'HSN or SAC code. E-invoice: 6 or 8 digits. E-way bill: 4, 6 or 8 digits (6 when your turnover is above Rs 5 crore).',
    qty: 'Number.', unit: 'GST unit code: NOS, PCS, KGS, LTR, MTR, BOX, SET, PAC... (other words become OTH).',
    rate: 'Price per unit before GST. Commas are fine (1,25,000.50).', disc: 'Per cent off the line, if any.', gst: 'The GST rate on the date of the invoice: 0, 0.1, 0.25, 1.5, 3, 5, 12, 18 or 40 (28 only before 1 Feb 2026).',
    dtype: 'INV tax invoice, BIL bill of supply, CHL delivery challan, OTH others.',
    sub: 'Why the goods move: Supply, Export, Job Work, For Own Use, Others, SKD/CKD/Lots, Line Sales, Recipient Not Known, Exhibition or Fairs.',
    subdesc: 'Only for Sub Supply Type "Others": say what it is.',
    dist: 'Kilometres, up to 4000. 0 = the portal works it out from the PIN codes.',
    mode: 'Road, Rail, Air or Ship. Leave empty with no vehicle to make Part A only (add the vehicle later as Part B).',
    veh: 'Like MH12AB1234. Road needs a vehicle number or a transporter ID.', vtype: 'R regular, O over-dimensional cargo.',
    tid: '15-character GSTIN or TRANSIN of the transporter.', tname: 'Transporter\'s name.', tdoc: 'LR, RR or airway bill number (needed for rail and air).', tdate: 'DD/MM/YYYY.',
    shaddr: 'Only when the goods go to another place than the buyer address (bill to - ship to). For an export: the port.', shcity: 'City of that place.', shpin: '6-digit PIN of that place.', shstate: 'State code of that place.'
  };
  function colsFor(mode) {
    if (mode !== 'ewb') return COMMON.concat(EINV);
    return COMMON.concat(EWB).map(function (c) { return c[0] === 'no' ? ['no', 'Document No', 'INV/26-27/101', c[3] + '|invoice no'] : c[0] === 'date' ? ['date', 'Document Date', c[2], c[3] + '|invoice date'] : c; });
  }

  /* India Post: first two digits of the PIN -> GST state codes of that postal circle (coarse; border towns may differ) */
  var PIN2 = { 11: '07', 12: '06', 13: '06', 14: '03', 15: '03', 16: '03 04', 17: '02', 18: '01', 19: '01 38', 20: '09', 21: '09', 22: '09', 23: '09', 24: '09 05', 25: '09', 26: '09 05', 27: '09', 28: '09',
    30: '08', 31: '08', 32: '08', 33: '08', 34: '08', 36: '24 26', 37: '24', 38: '24', 39: '24 26', 40: '27 30', 41: '27', 42: '27', 43: '27', 44: '27', 45: '23', 46: '23', 47: '23', 48: '23', 49: '22',
    50: '36 37', 51: '37 36', 52: '37', 53: '37 34', 56: '29', 57: '29', 58: '29', 59: '29', 60: '33 34', 61: '33', 62: '33', 63: '33', 64: '33', 67: '32 34', 68: '32 31', 69: '32',
    70: '19', 71: '19', 72: '19', 73: '19 11', 74: '19 35', 75: '21', 76: '21', 77: '21', 78: '18', 79: '12 13 14 15 16 17 18', 80: '10', 81: '10 20', 82: '10 20', 83: '20 10', 84: '10', 85: '10 20' };
  /* true = fits, false = PIN of another state, null = cannot tell */
  function pinFits(pin, st) {
    pin = digits(pin);
    if (!/^[1-9]\d{5}$/.test(pin) || !st || st === '96' || st === '97') return null;
    var ok = PIN2[+pin.slice(0, 2)];
    return ok == null ? null : ok.split(' ').indexOf(st) >= 0;
  }

  /* ---------- values ---------- */
  function headerMap(row, cols) {
    var map = {};
    (row || []).forEach(function (h, i) {
      var n = norm(h); if (!n) return;
      cols.forEach(function (c) {
        if (map[c[0]] != null) return;
        if (n === norm(c[1]) || c[3].split('|').some(function (a) { return norm(a) === n; })) map[c[0]] = i;
      });
    });
    return map;
  }
  /* the heading row: the first of the top 15 rows with at least 4 of our columns; manual = { field: column } from the Columns screen */
  function findHead(rows, cols, manual) {
    var at = 0;
    while (at < Math.min(rows.length, 15) && Object.keys(headerMap(rows[at], cols)).length < 4) at++;
    if (at >= Math.min(rows.length, 15)) at = 0;
    var map = headerMap(rows[at] || [], cols);
    if (manual) Object.keys(manual).forEach(function (k) { if (manual[k] === '' || manual[k] == null) delete map[k]; else map[k] = +manual[k]; });
    var missing = ['no', 'date', 'desc'].filter(function (k) { return map[k] == null && !(k === 'desc' && map.hsn != null); });
    return { at: at, map: map, heads: rows[at] || [], missing: missing };
  }
  function isoDate(v) { var s = isoDate0(v), m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(s), x = m && new Date(+m[1], +m[2] - 1, +m[3]); return m && x.getMonth() === +m[2] - 1 && x.getDate() === +m[3] ? s : ''; }
  function isoDate0(v) {
    v = String(v || '').trim();
    var m;
    if (/^\d{5}(\.\d+)?$/.test(v)) return NS.xlsx ? NS.xlsx.serialDate(v) : new Date(Date.UTC(1899, 11, 30) + Math.round(+v) * 864e5).toISOString().slice(0, 10);
    if (/^\d{8}$/.test(v)) return v.slice(0, 4) + '-' + v.slice(4, 6) + '-' + v.slice(6);   /* Tally yyyymmdd */
    if ((m = v.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
    if ((m = v.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})(\s.*)?$/))) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    var mon = 'jan feb mar apr may jun jul aug sep oct nov dec'.split(' ');
    if ((m = v.toLowerCase().match(/^(\d{1,2})[\s\-\/]([a-z]{3})[a-z]*[\s\-\/,]+(\d{2,4})$/)) && mon.indexOf(m[2]) >= 0) return (m[3].length === 2 ? '20' + m[3] : m[3]) + '-' + ('0' + (mon.indexOf(m[2]) + 1)).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    return '';
  }
  function stateCode(v) {
    v = String(v || '').trim();
    if (!v) return '';
    if (/^\d{1,2}$/.test(v)) { var c = ('0' + v).slice(-2); return c === '99' ? '96' : NS.stateName(c) ? c : ''; }
    var m = v.match(/\((\d{2})\)/) || v.match(/^(\d{2})\s*[-–]/); if (m && NS.stateName(m[1])) return m[1];
    var n = norm(v.replace(/&/g, 'and')), hit = NS.STATES.filter(function (s) { return norm(s[1]) === n; })[0];
    return hit ? hit[0] : /^(other|foreign|outside|export)/i.test(v) ? '96' : '';
  }
  function subCode(v) {
    v = String(v || '').trim();
    if (!v) return 1;
    if (/^\d+$/.test(v)) return +v;
    var n = norm(v), hit = SUBS.filter(function (s) { return norm(s[1]) === n || norm(s[1]).indexOf(n) === 0; })[0];
    return hit ? hit[0] : 0;
  }
  function modeCode(v) { v = String(v || '').trim().toLowerCase(); return !v ? 0 : /^\d$/.test(v) ? +v : MODES[v] || (/ship/.test(v) ? 4 : 0); }

  /* a blank document in the shape the invoice maker (docs.js) saves, so einv.js can make its JSON */
  function blank(no, seller, o) {
    return { type: 'invoice', kind: 'tax', no: no, date: '', due: '', pos: '', rc: false, ref: '',
      seller: JSON.parse(JSON.stringify(seller)), buyer: { name: '', addr: '', city: '', pin: '', state: '', gstin: '', phone: '' },
      ship: null, irn: null, items: [], round: o.round !== false, comp: true, notes: '', terms: '', accent: '#0B3B5B', status: 'unpaid', copyLabel: 'orig', paper: 'a4', design: 'classic', copies: 1,
      einv: { sup: 'B2B', aato: o.aato || '', dist: '0', mode: '', veh: '', vtype: 'R', tid: '', tname: '', tdoc: '', tdate: '', withEwb: false } };
  }

  /* ---------- a sheet -> documents (not yet checked). Returns { docs, rows, skipped: [{r, why}], err, head } ---------- */
  function fromSheet(rows, seller, o) {
    o = o || {};
    var mode = o.mode === 'ewb' ? 'ewb' : 'einv', cols = colsFor(mode), H = findHead(rows || [], cols, o.map);
    if (H.missing.length) return { err: t('bk.nohead'), head: H };
    var map = H.map, at = H.at;
    var get = function (r, k) { return map[k] == null ? '' : String(r[map[k]] == null ? '' : r[map[k]]).trim(); };
    var docs = [], byNo = {}, skipped = [], lastNo = '';
    rows.slice(at + 1).forEach(function (r, i) {
      var line = at + i + 2, no = get(r, 'no').toUpperCase().replace(/\s+/g, '');
      if (!r || !r.some(function (c) { return String(c == null ? '' : c).trim(); })) return;
      if (/^(example|e\.g\.)/i.test(no) || /^(total|grand ?total|sub ?total)$/i.test(no) || /^(total|grand total|sub total)\b/i.test(get(r, 'bname'))) { skipped.push({ r: line, why: t('bk.sk.total') }); return; }
      if (!no) { if (get(r, 'desc') || get(r, 'rate')) skipped.push({ r: line, why: t('bk.sk.nono') }); return; }
      var d = byNo[no];
      if (!d) {
        d = byNo[no] = { line: line, prob: [], note: [], src: 'sheet', d: blank(no, seller, o) };
        docs.push(d);
        var x = d.d;
        x.date = isoDate(get(r, 'date')); x.pos = stateCode(get(r, 'pos')); x.rc = /^y/i.test(get(r, 'rc'));
        x.buyer = { name: get(r, 'bname'), addr: get(r, 'baddr'), city: get(r, 'bcity'), pin: digits(get(r, 'bpin')), state: stateCode(get(r, 'bstate')), gstin: get(r, 'bgstin').toUpperCase().replace(/\s/g, ''), phone: '' };
        if (!x.buyer.state && /^\d{2}/.test(x.buyer.gstin) && NS.stateName(x.buyer.gstin.slice(0, 2))) x.buyer.state = x.buyer.gstin.slice(0, 2);
        var e = x.einv;
        e.sup = (get(r, 'sup') || 'B2B').toUpperCase(); e.dist = digits(get(r, 'dist')) || '0'; e.veh = get(r, 'veh').toUpperCase().replace(/[\s-]/g, ''); e.vtype = (get(r, 'vtype') || 'R').toUpperCase().slice(0, 1);
        e.tid = get(r, 'tid').toUpperCase(); e.tname = get(r, 'tname'); e.tdoc = get(r, 'tdoc'); e.tdate = isoDate(get(r, 'tdate'));
        if (get(r, 'dist') && !/^\d+$/.test(get(r, 'dist').replace(/\s*km$/i, ''))) d.prob.push(t('bk.dist', { r: line, v: get(r, 'dist') }));
        if (!x.date) d.prob.push(t('bk.date', { r: line }));
        if (get(r, 'bstate') && !x.buyer.state) d.prob.push(t('bk.state', { r: line, v: get(r, 'bstate') }));
        if (mode === 'ewb') {
          var sub = subCode(get(r, 'sub')), dt = (get(r, 'dtype') || '').toUpperCase().slice(0, 3);
          if (!sub || !DOCS_FOR[sub]) d.prob.push(t('bk.sub', { r: line, v: get(r, 'sub') }));
          else {
            if (!dt) dt = DOCS_FOR[sub][0];
            if (DOCS_FOR[sub].indexOf(dt) < 0) d.prob.push(t('bk.dtype', { r: line, v: dt, ok: DOCS_FOR[sub].join(', ') }));
          }
          if (get(r, 'shaddr') || get(r, 'shpin')) x.ship = { name: x.buyer.name, addr: get(r, 'shaddr'), city: get(r, 'shcity'), pin: digits(get(r, 'shpin')), state: stateCode(get(r, 'shstate')) };
          e.sub = sub; e.dtype = dt; e.subDesc = get(r, 'subdesc');
          if (sub === 8 && !e.subDesc) d.prob.push(t('bk.subdesc', { r: line }));
          x.kind = dt === 'INV' ? 'tax' : 'supply';
          if (sub === 3) { x.buyer.gstin = ''; e.sup = 'EXPWP'; }
          if (TO_SELF[sub]) { x.buyer.gstin = seller.gstin; if (!x.buyer.name) x.buyer.name = seller.name; }
          var mc = modeCode(get(r, 'mode'));
          if (get(r, 'mode') && !mc) d.prob.push(t('bk.mode', { r: line, v: get(r, 'mode') }));
          e.mode = mc ? String(mc) : '';
          if (mc === 1 && !e.veh && !e.tid) d.prob.push(t('bk.road', { r: line }));
          if ((mc === 2 || mc === 3) && !(e.tdoc && e.tdate)) d.prob.push(t('bk.raildoc', { r: line }));
        } else if (['B2B', 'SEZWP', 'SEZWOP', 'EXPWP', 'EXPWOP', 'DEXP'].indexOf(e.sup) < 0) d.prob.push(t('bk.sup', { r: line, v: e.sup }));
        if (/^EXP/.test(e.sup)) { x.buyer.gstin = ''; x.buyer.state = '96'; if (!x.pos) x.pos = '96'; }
      } else {
        var dd = isoDate(get(r, 'date')), g2 = get(r, 'bgstin').toUpperCase().replace(/\s/g, '');
        if (dd && dd !== d.d.date) d.prob.push(t('bk.twodates', { r: line }));
        else if (g2 && d.d.buyer.gstin && g2 !== d.d.buyer.gstin && mode !== 'ewb') d.prob.push(t('bk.dupno', { r: line, no: no }));
        else if (no !== lastNo && !d.apart) { d.apart = 1; d.note.push(t('bk.apart', { r: line })); }
      }
      lastNo = no;
      var g = amt(get(r, 'gst'));
      d.d.items.push({ desc: get(r, 'desc'), hsn: digits(get(r, 'hsn')), qty: amt(get(r, 'qty')) || '1', unit: get(r, 'unit') || 'Nos', rate: amt(get(r, 'rate')), disc: amt(get(r, 'disc')), gst: g === '' ? 0 : num(g), _r: line });
    });
    return { docs: docs, rows: rows.length - at - 1, skipped: skipped, head: H };
  }

  /* ---------- Tally XML (one or more files: Day Book / Sales register, and Masters for the buyers' addresses) -> documents ---------- */
  function fromTally(texts, seller, o) {
    o = o || {};
    var T = G.NSTallyIn; if (!T) return { err: 'NSTallyIn missing' };
    var P = T.plan(texts, { seller: seller });
    var par = {}; (P.parties || []).forEach(function (p) { par[key(p.name)] = p; });
    var docs = [], skipped = [], kinds = {};
    (P.docs || []).forEach(function (src) {
      if (src.type !== 'invoice') { if (src.type === 'note' && src.side === 'sale') kinds.note = (kinds.note || 0) + 1; return; }
      var no = String(src.vno || src.no || '').toUpperCase().replace(/\s+/g, '');
      var x = blank(no, seller, o), p = par[key(src.buyer && src.buyer.name)] || {};
      x.date = src.date; x.kind = src.kind === 'supply' ? 'supply' : 'tax'; x.pos = src.pos || '';
      var addr = String(p.addr || '').split(/\s*,\s*/).filter(Boolean), last = addr.length > 1 ? addr[addr.length - 1].replace(/\b\d{6}\b/, '').replace(/[-,\s]+$/, '').trim() : '';
      x.buyer = { name: src.buyer.name || '', addr: addr.length > 1 ? addr.slice(0, -1).join(', ') : addr.join(', '), city: last, pin: digits(p.pin), state: src.buyer.state || p.state || '', gstin: String(src.buyer.gstin || p.gstin || '').toUpperCase(), phone: '' };
      x.items = (src.items || []).map(function (it) { return { desc: it.desc, hsn: digits(it.hsn), qty: String(it.qty || 1), unit: it.unit || 'Nos', rate: String(it.rate || 0), disc: it.disc || '', gst: +it.gst || 0 }; });
      x.round = o.round !== false;
      if (o.mode === 'ewb') { x.einv.sub = 1; x.einv.dtype = x.kind === 'tax' ? 'INV' : 'BIL'; }
      var d = { line: 0, prob: [], note: [], src: 'tally', d: x };
      if (!no) d.prob.push(t('bk.tl.nono'));
      if (!p.addr && !p.pin) d.note.push(t('bk.tl.nomaster', { p: x.buyer.name }));
      if ((src.charges || []).length) d.note.push(t('bk.tl.chg', { v: src.charges.map(function (c) { return c.name; }).join(', ') }));
      docs.push(d);
    });
    if (kinds.note) skipped.push({ r: 0, why: t('bk.tl.notes', { n: kinds.note }) });
    if (!docs.length && !(P.docs || []).length) return { err: t('bk.tl.none') };
    return { docs: docs, rows: docs.length, skipped: skipped, company: P.company };
  }

  /* ---------- the checks of one document (einv.js check() + the bulk ones above) ---------- */
  var HAVE_EXP = /^EXP/;
  function checkOne(x, o) {
    o = o || {};
    var mode = o.mode === 'ewb' ? 'ewb' : 'einv', d = x.d, aato = d.einv.aato = o.aato || d.einv.aato || '';
    x.tt = totals(d);
    var prob = x.prob.slice(), note = x.note.slice();
    if (mode === 'ewb' && d.kind !== 'tax' && x.tt.rows.some(function (r) { return r.g; })) prob.push(t('bk.zero', { v: d.einv.dtype }));
    /* rate list of the document's date (a tax invoice only; a bill of supply has no GST) */
    if (d.kind === 'tax' && d.date) {
      var ok = NS.gstRatesFor(d.date);
      x.tt.rows.forEach(function (r, i) {
        if (ok.indexOf(r.g) < 0 && [0, 0.1, 0.25, 1, 1.5, 3, 5, 6, 7.5, 12, 18, 28, 40].indexOf(r.g) >= 0) prob.push(t('bk.ratedate', { n: i + 1, g: r.g, d: NS.fmtDate ? NS.fmtDate(d.date) : d.date }));
        else if (r.g === 12 && d.date >= '2025-09-22' && !/^99/.test(r.it.hsn)) note.push(t('bk.r12', { n: i + 1 }));
      });
    }
    /* HSN digits for the e-way bill by turnover (the e-invoice check already wants 6) */
    if (mode === 'ewb' && (aato === '5to10' || aato === '10p')) x.tt.rows.forEach(function (r, i) { if (digits(r.it.hsn).length === 4) prob.push(t('bk.hsn6e', { n: i + 1 })); });
    /* PIN and state */
    [[d.buyer, 'lg.billto'], [d.ship, 'bk.shipto'], [d.seller, 'lg.biz']].forEach(function (p) {
      if (!p[0] || (p[0] === d.buyer && HAVE_EXP.test(d.einv.sup))) return;
      if (pinFits(p[0].pin, p[0].state) === false) note.push(t('bk.pinst', { w: t(p[1]), pin: digits(p[0].pin), st: NS.stateName(p[0].state) }));
    });
    /* the buyer's GSTIN belongs to the buyer's state */
    var g = d.buyer.gstin;
    if (g && /^\d{2}/.test(g) && d.buyer.state && g.slice(0, 2) !== d.buyer.state && !HAVE_EXP.test(d.einv.sup) && !(mode === 'ewb' && g === d.seller.gstin)) prob.push(t('bk.gststate', { g: g, st: NS.stateName(d.buyer.state) }));
    if (mode === 'ewb' && g && !NS.gstinCheck(g).ok) prob.push(t('bk.gstin', { r: x.line || '-', v: g }));
    /* already reported (My data) */
    var h = o.have && o.have[sigKey(d)];
    if (h && mode !== 'ewb' && h.irn && h.irn.irn) prob.push(t('bk.hasirn', { no: d.no }));
    else if (h && mode === 'ewb' && h.irn && h.irn.ewb) note.push(t('bk.hasewb', { no: d.no, e: h.irn.ewb }));
    var c = NS.einv.check(d, x.tt, mode);
    x.bad = prob.concat(c.bad.map(function (p) { return p[1]; }));
    x.note = note.concat(c.note);
    delete x.apart;
    return x;
  }
  function sigKey(d) { return [key(d.no), NS.fy(d.date), key(d.seller && d.seller.gstin)].join('|'); }
  /* saved invoices -> { sigKey: record } for the "already has an IRN" check */
  function haveIndex(all) {
    var m = {};
    (all || []).forEach(function (r) { if ((r.type || 'invoice') === 'invoice' && r.no && r.date) { var k = sigKey(r), p = m[k]; if (!p || (r.irn && !p.irn)) m[k] = r; } });
    return m;
  }
  /* the same arithmetic as the invoice maker (docs.js): GST per line on the value after discount, total rounded if ticked */
  function totals(d) {
    var origin = d.seller.state, dest = d.pos || d.buyer.state || origin;
    var tt = { rows: [], taxable: 0, cgst: 0, sgst: 0, igst: 0, intra: !origin || !dest || origin === dest, sum: {} };
    d.items.forEach(function (it) {
      var q = num(it.qty), rate = num(it.rate);
      if (!it.desc && !rate) return;
      var taxable = r2(q * rate - q * rate * num(it.disc) / 100), g = num(it.gst), tx = r2(taxable * g / 100);
      tt.rows.push({ it: it, taxable: taxable, g: g, tax: tx, amount: r2(taxable + tx) });
      tt.taxable += taxable;
      if (tt.intra) { tt.cgst += tx / 2; tt.sgst += tx / 2; } else tt.igst += tx;
    });
    tt.taxable = r2(tt.taxable); tt.cgst = r2(tt.cgst); tt.sgst = r2(tt.sgst); tt.igst = r2(tt.igst);
    var total = r2(tt.taxable + tt.cgst + tt.sgst + tt.igst);
    tt.roundOff = d.round ? r2(Math.round(total) - total) : 0;
    tt.grand = r2(total + tt.roundOff);
    return tt;
  }
  /* sheet or Tally -> checked documents, in one go (tests; the screen checks in steps with NS.eachSeq) */
  function read(rows, seller, o) {
    var r = fromSheet(rows, seller, o);
    if (!r.err) r.docs.forEach(function (x) { checkOne(x, o); });
    return r;
  }

  /* ---------- the files ---------- */
  function file(ok, mode) {
    if (mode === 'ewb') return { version: EWB_VERSION, billLists: ok.map(function (x) { return NS.einv.ewaybill(x.d, x.tt).billLists[0]; }) };
    return ok.map(function (x) { return NS.einv.einvoice(x.d, x.tt)[0]; });
  }
  /* the list of problems as a sheet (Excel / CSV): one row per problem or note */
  function problemSheet(docs) {
    var out = [['Document No', 'Date', 'Buyer', 'Buyer GSTIN', 'Total', 'Status', 'Kind', 'What to fix or check', 'Sheet row']];
    docs.forEach(function (x) {
      var d = x.d, base = [d.no, d.date ? d.date.split('-').reverse().join('/') : '', d.buyer.name, d.buyer.gstin, x.tt ? x.tt.grand : ''];
      if (!x.bad.length && !x.note.length) out.push(base.concat(['Ready', '', '', x.line || '']));
      x.bad.forEach(function (p) { out.push(base.concat(['Needs fixing', 'Problem', p, x.line || ''])); });
      x.note.forEach(function (p) { out.push(base.concat([x.bad.length ? 'Needs fixing' : 'Ready', 'Note', p, x.line || ''])); });
    });
    return out;
  }
  function skippedSheet(skipped) { return [['Sheet row', 'Why it was left out']].concat((skipped || []).map(function (s) { return [s.r || '', s.why]; })); }

  /* ---------- the portal's replies: many files at once, matched to the documents of this batch and to My data ---------- */
  /* found = NS.einv.findAll(...) results of all files; batch = documents of the file made here; all = saved docs. Returns rows + jobs (records to save) */
  function replies(found, batch, all) {
    var seen = {}, rows = [], save = [];
    found.forEach(function (r) {
      var q = r.qr ? NS.einv.qrInfo(r.qr) : null, no = String(r.no || (q && q.DocNo) || '').toUpperCase().replace(/\s+/g, ''), sg = q && q.SellerGstin;
      var info = { no: no, irn: String(r.irn || (q && q.Irn) || '').trim(), ack: String(r.ack || '').trim(), ackDt: String(r.ackDt || (q && q.IrnDt) || '').trim(), ewb: r.ewb ? String(r.ewb).trim() : '', ewbDt: String(r.ewbDt || '').trim(), ewbTill: String(r.ewbTill || '').trim(), qr: r.qr || '', id: '', inBatch: false };
      var k = no + '|' + (sg || ''), was = no && seen[k];
      if (was) { Object.keys(info).forEach(function (f) { if (info[f] && !was[f]) was[f] = info[f]; }); info = was; }
      else if (no) seen[k] = info;
      info.inBatch = (batch || []).some(function (x) { return x.d.no === no && (!sg || x.d.seller.gstin === sg); });
      var rec = no && (all || []).filter(function (d) { return (d.type || 'invoice') === 'invoice' && String(d.no).toUpperCase().replace(/\s+/g, '') === no && (!sg || (d.seller && d.seller.gstin) === sg); })
        .sort(function (a, b) { return (b.updated || 0) - (a.updated || 0); })[0];
      if (rec) {
        var irn = rec.irn || {};
        ['irn', 'ack', 'ackDt', 'qr', 'ewb', 'ewbDt', 'ewbTill'].forEach(function (f) { if (info[f]) irn[f] = info[f]; });
        irn.sig = [String(rec.no).toUpperCase(), rec.date, rec.seller.gstin, rec.buyer.gstin, totals(rec).grand].join('|');
        rec.irn = irn; info.id = rec.id; if (save.indexOf(rec) < 0) save.push(rec);
      }
      if (!was) rows.push(info);
    });
    return { rows: rows, save: save };
  }
  function replySheet(rows) {
    return [['Document No', 'IRN', 'Ack No', 'Ack Date', 'E-way bill No', 'E-way bill Date', 'Valid till', 'In this batch', 'Saved in My data']].concat(rows.map(function (r) {
      return [r.no, r.irn, r.ack, r.ackDt, r.ewb, r.ewbDt, r.ewbTill, r.inBatch ? 'Yes' : '', r.id ? 'Yes' : ''];
    }));
  }
  /* the portal's Excel / CSV result sheet -> objects findAll understands */
  function replyFromRows(rows) {
    var at = rows.findIndex(function (r) { return (r || []).some(function (c) { return /^(irn|ewbno|ewaybillno|ewaybillnumber|ewbnumber|eway billno)$/.test(norm(c)); }); });
    if (at < 0) return [];
    var h = rows[at].map(function (c) { return norm(c).replace('number', 'no').replace('date', 'dt'); });
    return rows.slice(at + 1).map(function (r) { var o = {}; h.forEach(function (k, i) { if (k) o[k] = r[i]; }); return o; });
  }

  /* ---------- the template (with a "How to fill" sheet) and the sample ---------- */
  function template(mode) {
    var cols = colsFor(mode), ks = cols.map(function (c) { return c[0]; });
    var head = cols.map(function (c) { return c[1]; }), ex1 = cols.map(function (c) { return c[2]; }), ex2 = ex1.slice();
    ex2[ks.indexOf('desc')] = 'Cotton yarn 60s'; ex2[ks.indexOf('hsn')] = '52052400'; ex2[ks.indexOf('rate')] = '310'; ex2[ks.indexOf('qty')] = '40';
    var help = [['Column', 'What to type']].concat(cols.map(function (c) { return [c[1], HELP[c[0]] || '']; }),
      [['', ''], ['One row per item', 'Rows with the same ' + cols[0][1] + ' make one ' + (mode === 'ewb' ? 'e-way bill' : 'invoice') + '. Buyer details are read from its first row.'],
       ['Delete the examples', 'Replace the two example rows (INV/26-27/101) with your own.'],
       ['Your own export', 'A sales register from Tally, Busy or Marg works too if its headings are like these (Voucher No, Date, Party GSTIN, Item, HSN, Qty, Rate, GST %...). If a heading is not found, the site asks which column it is.'],
       ['Save', 'Keep it as Excel (.xlsx), or File > Save As > CSV UTF-8.']]);
    return [[mode === 'ewb' ? 'E-way bills' : 'Invoices', [head, ex1, ex2]], ['How to fill', help, [24, 110], 'wrap']];
  }
  var SAMPLE_SELLER = { name: 'Sample Textiles (sample)', addr: '45 Industrial Area A', city: 'Ludhiana', pin: '141003', state: '03', gstin: '03AABCS1234F1Z1', phone: '9800000000', email: '' };
  /* five made-up invoices (seven item rows) of a Punjab business; the last one has a wrong GSTIN and 28% to show the problem list */
  function sample(mode, today) {
    var day = function (back) { var x = new Date(Date.parse(today || NS.today()) - back * 864e5); return ('0' + x.getUTCDate()).slice(-2) + '/' + ('0' + (x.getUTCMonth() + 1)).slice(-2) + '/' + x.getUTCFullYear(); };
    var cols = colsFor(mode), ks = cols.map(function (c) { return c[0]; });
    var R = [
      { no: 'SMP/26-27/001', date: day(3), bgstin: '27AAACT5678K1ZS', bname: 'Tulsi Garments (sample)', baddr: '8 Station Road', bcity: 'Pune', bpin: '411001', bstate: '27', desc: 'Cotton yarn 40s', hsn: '52051200', qty: '200', unit: 'KGS', rate: '245', gst: '5', veh: 'PB10AB1234', mode: 'Road', dist: '1650' },
      { no: 'SMP/26-27/001', desc: 'Cotton yarn 60s', hsn: '52052400', qty: '50', unit: 'KGS', rate: '310', gst: '5' },
      { no: 'SMP/26-27/002', date: day(2), bgstin: '07AAFCM9012Q1Z0', bname: 'Mehra Fabrics (sample)', baddr: '22 Chandni Chowk', bcity: 'Delhi', bpin: '110006', bstate: '07', desc: 'Knitted fabric', hsn: '60062200', qty: '120', unit: 'MTR', rate: '180', disc: '2', gst: '5', veh: 'PB10CD5678', mode: 'Road', dist: '310' },
      { no: 'SMP/26-27/003', date: day(2), bgstin: '03AAKCT2468B1ZL', bname: 'Kaur Trading Co (sample)', baddr: '3 Ghumar Mandi', bcity: 'Ludhiana', bpin: '141001', bstate: '03', desc: 'Woollen shawl', hsn: '62142010', qty: '40', unit: 'PCS', rate: '1,250', gst: '5', veh: 'PB10EF9012', mode: 'Road', dist: '8' },
      { no: 'SMP/26-27/003', desc: 'Packing service', hsn: '998549', qty: '1', unit: 'NOS', rate: '1500', gst: '18' },
      { no: 'SMP/26-27/004', date: day(1), bgstin: '29AAGCR3456L1ZU', bname: 'Raman Exports (sample)', baddr: '14 MG Road', bcity: 'Bengaluru', bpin: '560001', bstate: '29', desc: 'Silk saree', hsn: '50072010', qty: '25', unit: 'PCS', rate: '4200', gst: '5', veh: 'PB10GH3456', mode: 'Road', dist: '2600' },
      { no: 'SMP/26-27/005', date: day(1), bgstin: '24AAHFP7890M1ZK', bname: 'Patel Stores (sample, wrong GSTIN)', baddr: '9 Relief Road', bcity: 'Ahmedabad', bpin: '380001', bstate: '24', desc: 'Cotton bedsheet', hsn: '6304', qty: '60', unit: 'PCS', rate: '450', gst: '28', veh: 'PB-10 X', mode: 'Road', dist: '1250' }
    ];
    return { seller: JSON.parse(JSON.stringify(SAMPLE_SELLER)), rows: [cols.map(function (c) { return c[1]; })].concat(R.map(function (r) { return ks.map(function (k) { return r[k] == null ? '' : r[k]; }); })) };
  }

  var API = { colsFor: colsFor, findHead: findHead, fromSheet: fromSheet, fromTally: fromTally, checkOne: checkOne, read: read, totals: totals, file: file,
    problemSheet: problemSheet, skippedSheet: skippedSheet, replies: replies, replySheet: replySheet, replyFromRows: replyFromRows, haveIndex: haveIndex,
    template: template, sample: sample, pinFits: pinFits, isoDate: isoDate, stateCode: stateCode, EWB_VERSION: EWB_VERSION, SUBS: SUBS };
  G.NSBulkCore = API;
})(typeof window !== 'undefined' ? window : this);
