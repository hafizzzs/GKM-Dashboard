/**
 * ============================================================================
 *  WEB APP DASHBOARD 7 STAGE – IMPORTASI KACANG  |  PT BGR Logistik Indonesia
 *  File server. Pasangannya: KacangDashboard_UI.html
 * ----------------------------------------------------------------------------
 *  Semua logika dibungkus di namespace KacangDash agar tidak bentrok dengan
 *  script lain di project yang sama. Fungsi global yang dibuat hanya:
 *    KacangDash, doGet, kcgGetData, kcgRefresh
 *  Catatan: satu project Apps Script hanya boleh punya SATU doGet. Jika project
 *  ini sudah punya doGet lain, pasang kedua file ini di project Apps Script
 *  terpisah (script.google.com) dan isi CFG.SPREADSHEET_ID.
 * ============================================================================
 */

var KacangDash = (function () {
  'use strict';
  const CFG = {
    SPREADSHEET_ID: '',                 // kosong = Google Sheet tempat script ini terpasang. Isi ID bila script dibuat di project terpisah.
    CACHE_SECONDS: 300,                 // data di-cache 5 menit supaya web app terbuka cepat
    TZ: 'Asia/Jakarta',
    CUTOFF: null,                       // null = hari ini. Contoh tanggal tetap: new Date(2026, 10, 6)
    SUMMARY: { sheet: 'Summary Report', headerRow: 10 },
    // Sheet status per container. Tambahkan cabang lain bila sheet-nya sudah tersedia
    // dengan header yang sama (NO BL, NO CONT, STATUS CONTAINER, KETERANGAN, ...).
    CONTAINER_SOURCES: [
      { sheet: 'Import Range Status Container', headerRow: 7, cabang: 'JAKARTA' }
      // , { sheet: 'Status Container SBY', headerRow: 7, cabang: 'SURABAYA' }
    ],
    OUT: { dash: 'Dashboard', db: 'DB_Stage', dq: 'Data_Quality', ref: 'Ref_Stage' },
    WILAYAH: [['JAKARTA', 'BLI-DKI & SNJL'], ['SEMARANG', 'BLI-SMRG'], ['SURABAYA', 'BLI-SBY'], ['MEDAN', 'BLI-MDN']],
    EXCLUDE_BL_STATUS: ['DIALIHKAN KE TBA', 'DIALIHKAN'],
    KARANTINA_KEYWORDS: ['KARANTINA', 'LAB', 'SAMPLE', 'SAMPEL', 'TILA', 'UJI'],
    // SLA mengacu SOP di Summary Report: Uji Lab H+1–H+3, SPPB H+4, Delivery H+5 s.d. H+7 (H = ATA)
    SLA: {
      SPPB_FROM_ATA: 4,          // stage 2–3 lewat SLA bila belum SPPB > H+4
      DELIVERY_START_FROM_ATA: 5,// stage 4 lewat SLA bila belum dikirim > H+5
      DELIVERED_FROM_ATA: 7,     // stage 5 lewat SLA bila belum selesai bongkar > H+7
      RETURN_DEPO_DAYS: 2,       // stage 6 lewat SLA bila container kosong belum balik depo > 2 hari
      DO_WARN_DAYS: 3            // peringatan DO expired ≤ 3 hari
    },
    // Jenis komoditas. Sumber dibaca berurutan:
    //  (1) kolom "Komoditas" di Summary Report, (2) kolom "Komoditas" di sheet master BL di bawah,
    //  (3) kata kunci di keterangan status BL, (4) DEFAULT_KOMODITAS.
    KOMODITAS: [
      { id: 'TANAH', label: 'Kacang Tanah', keys: ['TANAH', 'PEANUT', 'GROUNDNUT'] },
      { id: 'HIJAU', label: 'Kacang Hijau', keys: ['HIJAU', 'MUNG', 'GREEN GRAM'] }
    ],
    KOMODITAS_MASTER: { sheet: 'Database Order Konfirmasi', headerRow: 2 },
    DEFAULT_KOMODITAS: '',      // '' = "Belum ditentukan". Isi 'TANAH' bila semua BL tanpa keterangan adalah kacang tanah.
    ACTION_LIST_MAX: 30,
    MAX_ROWS: 5000               // batas baris yang dibaca dari sheet sumber
  };

  const STAGES = [
    { code: 1, name: '1. On Vessel', short: 'On Vessel', color: '#5B8DEF' },
    { code: 2, name: '2. At CY', short: 'At CY', color: '#F39C3D' },
    { code: 3, name: '3. At Karantina', short: 'At Karantina', color: '#E8C547' },
    { code: 4, name: '4. Menunggu Penjaluran', short: 'Menunggu Penjaluran', color: '#9B6FD1' },
    { code: 5, name: '5. On Delivery', short: 'On Delivery', color: '#2BB3C0' },
    { code: 6, name: '6. Delivered', short: 'Delivered', color: '#4CAF50' },
    { code: 7, name: '7. At Depo', short: 'At Depo', color: '#7F8C8D' }
  ];
  const STAGE_EXCL = '0. Dialihkan (di luar scope)';
  const SEV_RANK = { 'Tinggi': 3, 'Sedang': 2, 'Rendah': 1 };
  const THEME = { navy: '#0B2E59', navy2: '#16457F', gold: '#F2A900', grey: '#F3F5F8', line: '#D0D7E2', text: '#1F2937', muted: '#6B7280', red: '#C0392B', green: '#2E7D32' };


  // ============================================================================
  //  READ SOURCES
  // ============================================================================
  function readSources_(ss) {
    const sh = ss.getSheetByName(CFG.SUMMARY.sheet);
    if (!sh) throw new Error('Sheet "' + CFG.SUMMARY.sheet + '" tidak ditemukan.');
    const nS = Math.min(sh.getLastRow(), CFG.SUMMARY.headerRow + CFG.MAX_ROWS) - CFG.SUMMARY.headerRow + 1;
    const sumVals = sh.getRange(CFG.SUMMARY.headerRow, 1, nS, Math.min(sh.getLastColumn(), 30)).getValues();
    const containers = CFG.CONTAINER_SOURCES.map(s => {
      const c = ss.getSheetByName(s.sheet);
      if (!c) return { cabang: s.cabang, sheet: s.sheet, values: null };
      const n = Math.min(c.getLastRow(), s.headerRow + CFG.MAX_ROWS) - s.headerRow + 1;
      return { cabang: s.cabang, sheet: s.sheet, values: n > 0 ? c.getRange(s.headerRow, 1, n, c.getLastColumn()).getValues() : null };
    });
    let master = null;
    const km = CFG.KOMODITAS_MASTER && ss.getSheetByName(CFG.KOMODITAS_MASTER.sheet);
    if (km && km.getLastRow() >= CFG.KOMODITAS_MASTER.headerRow) {
      const n = Math.min(km.getLastRow(), CFG.KOMODITAS_MASTER.headerRow + CFG.MAX_ROWS) - CFG.KOMODITAS_MASTER.headerRow + 1;
      master = km.getRange(CFG.KOMODITAS_MASTER.headerRow, 1, n, km.getLastColumn()).getValues();
    }
    return { summary: sumVals, containers: containers, master: master };
  }

  // ============================================================================
  //  ENGINE (pure functions – tanpa akses spreadsheet)
  // ============================================================================
  const SUM_COLS = {
    cabang: ['CABANG'], bl: ['NO BL'], eta: ['ETA'], ata: ['ATA'], aju: ['NO AJU'], fcl: ['FCL'],
    telex: ['TELEX'], dor: ['DO RELEASE'], sp2mp: ['SP2MP'], sppb: ['SPPB'], tglSppb: ['TGL SPPB'],
    status: ['STATUS'], tglTila: ['TGL TILA', 'TGL PELEPASAN KARANTINA'],
    kom: ['KOMODITAS', 'JENIS KACANG', 'JENIS KOMODITAS', 'JENIS BARANG']
  };
  const MASTER_COLS = { bl: ['NO. BL', 'NO BL'], kom: ['KOMODITAS', 'JENIS KACANG', 'JENIS KOMODITAS', 'JENIS BARANG'] };

  /** Cocokkan teks bebas ke id komoditas ('' bila tidak dikenali). */
  function komOf_(text) {
    const t = norm_(text);
    if (!t) return '';
    const k = CFG.KOMODITAS.find(x => x.keys.some(w => t.indexOf(w) >= 0) || norm_(x.label) === t || x.id === t);
    return k ? k.id : '';
  }
  function komLabel_(id) { const k = CFG.KOMODITAS.find(x => x.id === id); return k ? k.label : 'Belum ditentukan'; }
  const CONT_COLS = {
    bl: ['NO BL'], cont: ['NO CONT'], status: ['STATUS CONTAINER'], ket: ['KETERANGAN'],
    muat: ['TANGGAL MUAT', 'TGL MUAT'], bongkar: ['TANGGAL BONGKAR', 'TGL BONGKAR'], tglSppb: ['TGL SPPB'],
    expDo: ['EXPIRED DO'], cust: ['NAMA PERUSAHAAN'], kota: ['WILAYAH'], terminal: ['TERMINAL'], pelaksana: ['PELAKSANA'],
    tglSelesai: ['TGL SELESAI BONGKAR'], tglDepo: ['TGL BALIK DEPO']
  };

  function norm_(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toUpperCase(); }

  function mapCols_(header, spec) {
    const h = header.map(norm_), out = {};
    Object.keys(spec).forEach(k => {
      let idx = -1;
      spec[k].forEach(a => { if (idx < 0) idx = h.indexOf(a); });
      if (idx < 0) spec[k].forEach(a => { if (idx < 0) idx = h.findIndex(x => x.indexOf(a) === 0); });
      out[k] = idx;
    });
    return out;
  }

  function dateOnly_(d) { return new Date(d.getFullYear(), d.getMonth(), d.getDate()); }

  function toDate_(v) {
    if (v === null || v === undefined || v === '') return null;
    if (Object.prototype.toString.call(v) === '[object Date]') return isNaN(v) ? null : dateOnly_(v);
    if (typeof v === 'number' && v > 20000 && v < 80000) return dateOnly_(new Date(Math.round((v - 25569) * 864e5)));
    const s = String(v).trim();
    let m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
    if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
    m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
    return null;
  }

  function days_(a, b) { return Math.round((a - b) / 864e5); }  // a - b dalam hari

  /** Koreksi tanggal yang terbaca terbalik (mm/dd vs dd/mm) bila di luar rentang wajar. */
  function fixSwap_(d, lo, hi) {
    if (!d) return { d: null, swapped: false };
    if ((lo && d < lo) || (hi && d > hi)) {
      if (d.getDate() <= 12) {
        const s = new Date(d.getFullYear(), d.getDate() - 1, d.getMonth() + 1);
        if ((!lo || s >= lo) && (!hi || s <= hi)) return { d: s, swapped: true };
      }
    }
    return { d: d, swapped: false };
  }

  function isYes_(v) {
    if (Object.prototype.toString.call(v) === '[object Date]') return true;
    if (v === true) return true;
    const s = norm_(v);
    return s.indexOf('✅') >= 0 || s.indexOf('✔') >= 0 || ['YA', 'YES', 'Y', 'TRUE', 'OK', 'SUDAH', 'DONE', 'V'].indexOf(s) >= 0;
  }

  function hasAny_(text, words) { const t = norm_(text); return words.some(w => t.indexOf(w) >= 0); }

  function wilayah_(cabang) {
    const f = CFG.WILAYAH.find(w => w[0] === cabang);
    return f ? f[1] : cabang;
  }

  function buildModel_(src, cutoff) {
    const dq = [];
    // ---------- BL master dari Summary Report
    const sv = src.summary, sc = mapCols_(sv[0], SUM_COLS);
    const bls = {}, blOrder = [];
    for (let i = 1; i < sv.length; i++) {
      const r = sv[i], bl = norm_(r[sc.bl]);
      if (!bl) continue;
      if (bls[bl]) { dq.push(dqRow_('Sedang', 'Duplikasi BL', norm_(r[sc.cabang]), bl, '', 'No BL muncul lebih dari sekali di Summary Report', 'Hapus/gabungkan baris duplikat')); continue; }
      const fcl = Number(r[sc.fcl]) || 0;
      const b = {
        bl: bl, cabang: norm_(r[sc.cabang]), eta: toDate_(r[sc.eta]), ata: toDate_(r[sc.ata]), fcl: fcl,
        telex: isYes_(r[sc.telex]), dor: isYes_(r[sc.dor]), sp2mp: sc.sp2mp >= 0 && isYes_(r[sc.sp2mp]),
        tglSppb: toDate_(r[sc.tglSppb]), sppbFlag: isYes_(r[sc.sppb]),
        tglTila: sc.tglTila >= 0 ? toDate_(r[sc.tglTila]) : null,
        status: String(r[sc.status] == null ? '' : r[sc.status]).trim(), nCont: 0,
        komCol: sc.kom >= 0 ? komOf_(r[sc.kom]) : ''
      };
      b.sppb = !!(b.tglSppb || b.sppbFlag);
      b.excluded = EXCL_(b.status);
      if (Math.abs(fcl - Math.round(fcl)) > 1e-6) dq.push(dqRow_('Rendah', 'FCL tidak bulat', b.cabang, bl, '', 'FCL tertulis ' + fcl, 'Koreksi menjadi ' + Math.round(fcl)));
      bls[bl] = b; blOrder.push(bl);
    }

    // ---------- Komoditas per BL
    const masterKom = {};
    if (src.master && src.master.length > 1) {
      const mc = mapCols_(src.master[0], MASTER_COLS);
      if (mc.bl >= 0 && mc.kom >= 0) src.master.slice(1).forEach(r => { const k = komOf_(r[mc.kom]); if (k) masterKom[norm_(r[mc.bl])] = k; });
    }
    blOrder.forEach(id => {
      const b = bls[id];
      const fromKet = komOf_(b.status);
      if (b.komCol) { b.kom = b.komCol; b.komSrc = 'kolom'; }
      else if (masterKom[id]) { b.kom = masterKom[id]; b.komSrc = 'master'; }
      else if (fromKet) { b.kom = fromKet; b.komSrc = 'keterangan'; }
      else { b.kom = CFG.DEFAULT_KOMODITAS || ''; b.komSrc = CFG.DEFAULT_KOMODITAS ? 'default' : ''; }
      if (!b.komSrc) dq.push(dqRow_('Sedang', 'Komoditas kosong', b.cabang, id, '', 'Jenis kacang (tanah/hijau) belum diisi', 'Isi kolom Komoditas di ' + CFG.KOMODITAS_MASTER.sheet));
      else if (b.komSrc === 'keterangan') dq.push(dqRow_('Rendah', 'Komoditas dari keterangan', b.cabang, id, '', 'Jenis kacang hanya terbaca dari teks keterangan status', 'Isi kolom Komoditas agar tidak hilang saat status diupdate'));
    });
    const komFill_ = b => { if (b.kom === undefined) { b.kom = masterKom[b.bl] || CFG.DEFAULT_KOMODITAS || ''; b.komSrc = b.kom ? 'master' : ''; } };

    // ---------- Container records
    const recs = [], seenKey = {}, cabWithSrc = {};
    src.containers.forEach(cs => {
      if (!cs.values) { dq.push(dqRow_('Sedang', 'Sumber tidak ditemukan', cs.cabang, '', '', 'Sheet "' + cs.sheet + '" tidak ditemukan/kosong', 'Cek nama sheet di CFG.CONTAINER_SOURCES')); return; }
      cabWithSrc[cs.cabang] = true;
      const cv = cs.values, cc = mapCols_(cv[0], CONT_COLS);
      for (let i = 1; i < cv.length; i++) {
        const r = cv[i], bl = norm_(r[cc.bl]);
        if (!bl) continue;
        const cont = norm_(r[cc.cont]);
        let b = bls[bl];
        if (!b) {
          dq.push(dqRow_('Sedang', 'BL tidak terdaftar', cs.cabang, bl, cont, 'BL ada di sheet container tapi tidak ada di Summary Report', 'Tambahkan BL ke Database Order Konfirmasi'));
          b = { bl: bl, cabang: cs.cabang, eta: null, ata: null, fcl: 0, telex: false, dor: false, sppb: false, tglSppb: null, tglTila: null, status: '', excluded: false, nCont: 0, orphan: true };
          komFill_(b);
          bls[bl] = b; blOrder.push(bl);
        }
        b.nCont++;
        const g = c => (c >= 0 ? r[c] : '');
        const rec = baseRec_(b, cs.cabang);
        rec.cont = cont; rec.qty = 1; rec.cust = String(g(cc.cust) || '').trim(); rec.kota = String(g(cc.kota) || '').trim();
        rec.statusLap = norm_(g(cc.status)); rec.ket = String(g(cc.ket) || '').trim();
        rec.terminal = String(g(cc.terminal) || ''); rec.pelaksana = String(g(cc.pelaksana) || '');
        const lo = b.ata ? new Date(b.ata.getTime() - 864e5) : null, hi = new Date(cutoff.getTime() + 864e5);
        const mu = fixSwap_(toDate_(g(cc.muat)), lo, hi), bo = fixSwap_(toDate_(g(cc.bongkar)), lo, hi);
        rec.muat = mu.d; rec.bongkar = bo.d;
        rec.expDo = toDate_(g(cc.expDo));
        rec.tglSelesai = toDate_(g(cc.tglSelesai)); rec.tglDepo = toDate_(g(cc.tglDepo));
        if (!rec.tglSppbC) rec.tglSppbC = toDate_(g(cc.tglSppb));
        rec.source = 'Per container (' + cs.sheet + ')';
        if (mu.swapped || bo.swapped) addFlag_(rec, 'Rendah', 'Format tanggal muat/bongkar terbaca terbalik (dd/mm) – sudah dikoreksi otomatis');
        if (!cont) addFlag_(rec, 'Rendah', 'No container kosong');
        else {
          const key = bl + '|' + cont;
          if (seenKey[key]) addFlag_(rec, 'Sedang', 'Duplikasi No Container dalam BL yang sama');
          seenKey[key] = true;
        }
        recs.push(rec);
      }
    });

    // ---------- BL tanpa data container → level BL
    blOrder.forEach(id => {
      const b = bls[id];
      if (b.nCont > 0 || b.orphan) return;
      const rec = baseRec_(b, b.cabang);
      rec.qty = Math.round(b.fcl); rec.cont = ''; rec.statusLap = '';
      rec.source = 'Level BL (Summary Report)';
      if (cabWithSrc[b.cabang]) addFlag_(rec, 'Sedang', 'BL belum ada di sheet status container');
      recs.push(rec);
    });

    // ---------- Count vs FCL
    blOrder.forEach(id => {
      const b = bls[id];
      if (b.nCont > 0 && b.fcl && b.nCont !== Math.round(b.fcl))
        dq.push(dqRow_('Sedang', 'Jumlah container ≠ FCL', b.cabang, b.bl, '', 'Tercatat ' + b.nCont + ' container, FCL ' + b.fcl, 'Lengkapi / koreksi daftar container'));
    });

    // ---------- Stage & SLA
    recs.forEach(r => classify_(r, cutoff));

    recs.forEach(r => r.flags.forEach(f => dq.push(dqRow_(f.sev, 'Flag container', r.cabang, r.bl, r.cont, f.msg, rekomendasi_(f.msg)))));
    dq.sort((a, b) => SEV_RANK[b.sev] - SEV_RANK[a.sev] || (a.cabang > b.cabang ? 1 : -1));
    return { cutoff: cutoff, recs: recs, bls: bls, blOrder: blOrder, dq: dq, agg: aggregate_(recs, bls, blOrder, cutoff) };
  }

  function EXCL_(status) { const s = norm_(status); return CFG.EXCLUDE_BL_STATUS.some(x => s.indexOf(x) === 0); }

  function baseRec_(b, cabang) {
    return {
      cabang: cabang, wilayah: wilayah_(cabang), bl: b.bl, b: b, kom: b.kom || '', eta: b.eta, ata: b.ata, tglSppbC: b.tglSppb,
      cust: '', kota: '', statusLap: '', ket: '', terminal: '', pelaksana: '', muat: null, bongkar: null, expDo: null,
      tglSelesai: null, tglDepo: null, flags: [], stage: '', stageCode: 0, sub: '', aging: null, sla: '', source: ''
    };
  }

  function addFlag_(r, sev, msg) { if (!r.flags.some(f => f.msg === msg)) r.flags.push({ sev: sev, msg: msg }); }

  function classify_(r, cut) {
    const b = r.b, S = CFG.SLA;
    const arrived = !!(b.ata && b.ata <= cut);
    const karantinaKet = hasAny_(b.status, CFG.KARANTINA_KEYWORDS);
    let st = r.statusLap;

    if (b.ata && b.ata > cut) addFlag_(r, 'Sedang', 'ATA melebihi tanggal cut-off – cek ETA/ATA');
    if (!b.ata && b.eta && b.eta < cut && !b.excluded) addFlag_(r, 'Sedang', 'ETA sudah lewat ' + days_(cut, b.eta) + ' hari, ATA belum dicatat');

    // status efektif
    if (!st) st = norm_(b.status) === 'DONE' ? 'BALIK DEPO' : (arrived ? 'CY' : 'ON VESSEL');
    if (st === 'ON VESSEL' && arrived) { addFlag_(r, 'Sedang', 'Status container masih ON VESSEL padahal sudah ATA'); st = 'CY'; }
    if (r.tglDepo) st = 'BALIK DEPO';
    else if (r.tglSelesai && st !== 'BALIK DEPO') st = 'SELESAI BONGKAR';

    const known = ['ON VESSEL', 'CY', 'ANTRI MUAT', 'MENUNGGU PENJALURAN', 'KARANTINA', 'TITIP DEPO', 'OUT CY', 'ON DELIVERY', 'ANTRI BONGKAR', 'BONGKAR', 'PROSES BONGKAR', 'SELESAI BONGKAR', 'BALIK DEPO'];
    if (known.indexOf(st) < 0) addFlag_(r, 'Sedang', 'Status container tidak dikenal: "' + st + '"');

    const ket = norm_(r.ket);
    let code;
    if (b.excluded) code = 0;
    else if (st === 'BALIK DEPO') code = 7;
    else if (st === 'SELESAI BONGKAR' || (['BONGKAR', 'PROSES BONGKAR'].indexOf(st) >= 0 && ket.indexOf('SELESAI') >= 0)) code = 6;
    else if (['BONGKAR', 'ANTRI BONGKAR', 'PROSES BONGKAR'].indexOf(st) >= 0) code = 5;
    else if (st === 'TITIP DEPO' || (['OUT CY', 'ON DELIVERY'].indexOf(st) >= 0 && ket.indexOf('TITIP') >= 0)) code = 4;
    else if (['OUT CY', 'ON DELIVERY'].indexOf(st) >= 0) code = 5;
    else if (st === 'ON VESSEL') code = 1;
    else { // di pelabuhan
      if (b.sppb && !(st === 'KARANTINA' && !b.tglTila)) code = 4;
      else if (st === 'KARANTINA' || karantinaKet) code = 3;
      else code = 2;
    }
    r.stageCode = code;
    r.stage = code === 0 ? STAGE_EXCL : STAGES[code - 1].name;

    // sub-status
    if (code === 1) r.sub = b.status || 'Dalam pelayaran';
    else if (code === 2) r.sub = !b.telex ? 'Menunggu Telex Release' : (!b.dor ? 'Menunggu DO Release' : 'Menunggu SPPB / PIB');
    else if (code === 3) r.sub = b.status || 'Pemeriksaan karantina, menunggu TILA';
    else if (code === 4) r.sub = (st === 'OUT CY' || st === 'ON DELIVERY' || st === 'TITIP DEPO') ? 'Isi – dititipkan di depo, menunggu slot bongkar' : 'Di CY – SPPB terbit, menunggu penjaluran truk';
    else if (code === 5) {
      if (ket.indexOf('JAMUR') >= 0) r.sub = 'Dialihkan ke gudang lain (kargo berjamur)';
      else if (st === 'ANTRI BONGKAR' || ket.indexOf('ANTRI') >= 0) r.sub = 'Tiba di gudang customer – antri bongkar';
      else if (st === 'BONGKAR' || st === 'PROSES BONGKAR') r.sub = 'Tiba di gudang customer – proses bongkar';
      else if (ket.indexOf('BUKA CONT') >= 0) r.sub = 'Tiba di gudang customer – menunggu buka container';
      else r.sub = 'Dalam perjalanan ke customer';
    } else if (code === 6) r.sub = 'Selesai bongkar – container kosong menuju depo';
    else if (code === 7) r.sub = 'Siklus selesai';
    else r.sub = b.status;

    // consistency flags
    if (code >= 4 && code <= 6 && !b.sppb) addFlag_(r, 'Tinggi', 'Container keluar/penjaluran tanpa SPPB tercatat');
    if (hasAny_(b.status, ['PENJALURAN']) && !b.sppb) addFlag_(r, 'Sedang', 'Keterangan "Menunggu Penjaluran" tetapi SPPB belum terbit');
    if (ket.indexOf('JAMUR') >= 0 || (ket.indexOf('DIALIHKAN') >= 0 && code !== 0)) addFlag_(r, 'Tinggi', 'Exception kargo: ' + r.ket);
    if (code === 5 && r.source.indexOf('Per container') === 0 && !r.muat && r.sub.indexOf('perjalanan') >= 0) addFlag_(r, 'Rendah', 'Tanggal muat kosong untuk container On Delivery');

    // aging & SLA
    const dwell = b.ata ? days_(cut, b.ata) : null;
    r.sla = 'On Track';
    if (code === 0) { r.sla = '–'; }
    else if (code === 1) { r.aging = null; if (!b.ata && b.eta && b.eta < cut) r.sla = 'ETA Lewat'; }
    else if (code === 2 || code === 3) { r.aging = dwell; if (dwell !== null && dwell > S.SPPB_FROM_ATA) r.sla = 'Lewat SLA'; }
    else if (code === 4) { r.aging = dwell; if (dwell !== null && dwell > S.DELIVERY_START_FROM_ATA) r.sla = 'Lewat SLA'; }
    else if (code === 5) { r.aging = dwell; if (dwell !== null && dwell > S.DELIVERED_FROM_ATA) r.sla = 'Lewat SLA'; }
    else if (code === 6) { const d = r.bongkar || r.tglSelesai; r.aging = d ? days_(cut, d) : null; if (r.aging !== null && r.aging > S.RETURN_DEPO_DAYS) r.sla = 'Lewat SLA'; }
    else if (code === 7) { r.sla = 'Selesai'; }

    if (code >= 2 && code <= 5 && r.expDo) {
      const left = days_(r.expDo, cut);
      if (left < 0) addFlag_(r, 'Tinggi', 'DO sudah expired ' + (-left) + ' hari – risiko demurrage/detention');
      else if (left <= S.DO_WARN_DAYS) addFlag_(r, 'Tinggi', 'DO expired dalam ' + left + ' hari');
    }
    r.sev = r.flags.reduce((m, f) => SEV_RANK[f.sev] > SEV_RANK[m] ? f.sev : m, r.flags.length ? 'Rendah' : '');
  }

  function rekomendasi_(msg) {
    const m = msg.toUpperCase();
    if (m.indexOf('ON VESSEL PADAHAL') >= 0) return 'Update status container menjadi CY pada hari ATA';
    if (m.indexOf('ATA MELEBIHI') >= 0) return 'Isi ATA hanya setelah kapal sandar; gunakan ETA untuk rencana';
    if (m.indexOf('ETA SUDAH LEWAT') >= 0) return 'Konfirmasi ke pelayaran, isi ATA atau revisi ETA';
    if (m.indexOf('PENJALURAN') >= 0) return 'Gunakan "Menunggu Penjaluran" hanya setelah SPPB terbit';
    if (m.indexOf('TANPA SPPB') >= 0) return 'Cek SPPB; lengkapi Tgl SPPB di Database Custom Clearance';
    if (m.indexOf('EXCEPTION') >= 0) return 'Eskalasi ke PIC operasi & customer; catat rencana penyelesaian';
    if (m.indexOf('DO ') >= 0) return 'Percepat penarikan / perpanjang DO ke pelayaran';
    if (m.indexOf('TERBALIK') >= 0) return 'Seragamkan format tanggal dd/mm/yyyy di sheet sumber';
    if (m.indexOf('NO CONTAINER KOSONG') >= 0) return 'Lengkapi nomor container';
    if (m.indexOf('DUPLIKASI') >= 0) return 'Hapus baris duplikat';
    if (m.indexOf('TIDAK DIKENAL') >= 0) return 'Gunakan daftar status baku (lihat Ref_Stage)';
    if (m.indexOf('BELUM ADA DI SHEET') >= 0) return 'Input daftar container BL ini di sheet status container';
    if (m.indexOf('TANGGAL MUAT') >= 0) return 'Isi tanggal muat / gate out';
    return '';
  }

  function dqRow_(sev, kat, cabang, bl, cont, isu, rek) { return { sev: sev, kat: kat, cabang: cabang, bl: bl, cont: cont, isu: isu, rek: rek }; }

  function aggregate_(recs, bls, blOrder, cut) {
    const cabs = CFG.WILAYAH.map(w => w[0]);
    recs.forEach(r => { if (cabs.indexOf(r.cabang) < 0) cabs.push(r.cabang); });
    const pipe = {}, outlook = {};
    cabs.forEach(c => { pipe[c] = [0, 0, 0, 0, 0, 0, 0, 0]; outlook[c] = [0, 0, 0, 0, 0, 0]; });
    let lewat = 0, flagged = 0;
    recs.forEach(r => {
      pipe[r.cabang][r.stageCode] += r.qty;
      if (r.sla === 'Lewat SLA' || r.sla === 'ETA Lewat') lewat += r.qty;
      if (r.flags.some(f => f.sev !== 'Rendah')) flagged += r.qty;
      if (r.stageCode === 1) {
        const e = r.eta; let k;
        if (!e) k = 5; else { const d = days_(e, cut); k = d < 0 ? 0 : d <= 3 ? 1 : d <= 7 ? 2 : d <= 14 ? 3 : 4; }
        outlook[r.cabang][k] += r.qty;
      }
    });
    // BL progress
    const blRows = blOrder.map(id => {
      const b = bls[id], cnt = [0, 0, 0, 0, 0, 0, 0, 0];
      recs.filter(r => r.bl === id).forEach(r => { cnt[r.stageCode] += r.qty; });
      const active = cnt.slice(1).reduce((a, x) => a + x, 0);
      return { b: b, cnt: cnt, active: active, pct: active ? (cnt[6] + cnt[7]) / active : 0 };
    });
    // Action list (grup BL + stage + sub)
    const grp = {};
    recs.forEach(r => {
      const imp = r.sla === 'Lewat SLA' || r.sla === 'ETA Lewat' || r.flags.some(f => f.sev !== 'Rendah');
      if (!imp || r.stageCode === 0 || r.stageCode === 7) return;
      const k = r.bl + '|' + r.stageCode + '|' + r.sub;
      if (!grp[k]) grp[k] = { cabang: r.cabang, bl: r.bl, stage: r.stage, sub: r.sub, qty: 0, aging: r.aging, sla: r.sla, sev: '', isu: [] };
      const g = grp[k]; g.qty += r.qty;
      if (r.aging !== null && (g.aging === null || r.aging > g.aging)) g.aging = r.aging;
      r.flags.filter(f => f.sev !== 'Rendah').forEach(f => { if (g.isu.indexOf(f.msg) < 0) g.isu.push(f.msg); });
      const s = r.sla === 'Lewat SLA' ? 'Tinggi' : (r.sev || 'Sedang');
      [s, r.sev].forEach(x => { if (x && (!g.sev || SEV_RANK[x] > SEV_RANK[g.sev])) g.sev = x; });
    });
    const actions = Object.keys(grp).map(k => grp[k])
      .sort((a, b) => SEV_RANK[b.sev] - SEV_RANK[a.sev] || (b.aging || 0) - (a.aging || 0) || b.qty - a.qty);
    return { cabs: cabs, pipe: pipe, outlook: outlook, lewat: lewat, flagged: flagged, blRows: blRows, actions: actions };
  }


  // ==========================================================================
  //  API WEB APP
  // ==========================================================================
  const CACHE_KEY = 'KCG_DASH_V2';
  const SLA_CODE = { '': 0, '–': 0, 'On Track': 1, 'Lewat SLA': 2, 'ETA Lewat': 3, 'Selesai': 4 };
  const SEV_CODE = { '': 0, 'Rendah': 1, 'Sedang': 2, 'Tinggi': 3 };

  function openSs_() {
    if (CFG.SPREADSHEET_ID) return SpreadsheetApp.openById(CFG.SPREADSHEET_ID);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error('Spreadsheet tidak ditemukan. Isi CFG.SPREADSHEET_ID di KacangDashboard_Server.gs.');
    return ss;
  }

  function ymd_(d) { return d ? Utilities.formatDate(d, CFG.TZ, 'yyyy-MM-dd') : ''; }

  /** Ubah model menjadi JSON ringkas untuk browser. */
  function pack_(m) {
    const cabIds = CFG.WILAYAH.map(w => w[0]);
    m.recs.forEach(r => { if (cabIds.indexOf(r.cabang) < 0) cabIds.push(r.cabang); });
    const SHORT = { JAKARTA: 'DKI', SEMARANG: 'SMG', SURABAYA: 'SBY', MEDAN: 'MDN' };
    const komIds = CFG.KOMODITAS.map(k => k.id).concat(['']);
    const blIdx = {}, bls = [];
    m.blOrder.forEach(id => {
      const b = m.bls[id];
      blIdx[id] = bls.length;
      bls.push([b.bl, cabIds.indexOf(b.cabang), Math.round(b.fcl * 100) / 100, ymd_(b.eta), ymd_(b.ata), b.status || '', ymd_(b.tglSppb), komIds.indexOf(b.kom || ''), b.komSrc || '']);
    });
    const recs = m.recs.map(r => [cabIds.indexOf(r.cabang), blIdx[r.bl], r.cont || '', r.qty, r.stageCode, r.sub || '',
      r.aging === null ? -1 : r.aging, SLA_CODE[r.sla] || 0, SEV_CODE[r.sev || ''] || 0, r.flags.map(f => f.msg).join('; '), r.cust || '', r.ket || '']);
    return {
      v: 1, cutoff: ymd_(m.cutoff), refreshed: Utilities.formatDate(new Date(), CFG.TZ, "yyyy-MM-dd'T'HH:mm"),
      koms: komIds.map(k => ({ id: k, label: komLabel_(k) })),
      cabs: cabIds.map(c => ({ id: c, label: wilayah_(c), short: SHORT[c] || c })),
      sla: CFG.SLA, bls: bls, recs: recs
    };
  }

  function getJson(force) {
    const cache = CacheService.getScriptCache();
    if (!force) { const hit = cache.get(CACHE_KEY); if (hit) return hit; }
    const ss = openSs_();
    const model = buildModel_(readSources_(ss), dateOnly_(CFG.CUTOFF || new Date()));
    const json = JSON.stringify(pack_(model));
    if (json.length < 95000) { try { cache.put(CACHE_KEY, json, CFG.CACHE_SECONDS); } catch (e) {} }
    return json;
  }

  return { getJson: getJson, buildModel: buildModel_, pack: pack_ };
})();

/** Halaman web app. */
function doGet(e) {
  let boot;
  try { boot = KacangDash.getJson(false); }
  catch (err) { boot = JSON.stringify({ error: String(err && err.message || err) }); }
  const t = HtmlService.createTemplateFromFile('KacangDashboard_UI');
  t.boot = boot.replace(/</g, '\\u003c').split(String.fromCharCode(8232)).join('\\u2028').split(String.fromCharCode(8233)).join('\\u2029');
  return t.evaluate()
    .setTitle('Kacang Impor - Monitoring Container')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes');
}

/** Dipanggil browser (data dari cache bila ada). */
function kcgGetData() { return KacangDash.getJson(false); }

/** Dipanggil tombol refresh (hitung ulang, abaikan cache). */
function kcgRefresh() { return KacangDash.getJson(true); }
