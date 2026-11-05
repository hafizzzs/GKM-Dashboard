/**
 * DASHBOARD MONITORING RAW SUGAR – MV. THE RAINBOW 50.000 MT
 * Receiving → Re-bagging → Delivery (PG Rajawali II)
 *
 * Cara pakai singkat:
 * 1. Buka spreadsheet rekap > Extensions > Apps Script.
 * 2. Tempel file ini sebagai Code.gs, tambahkan file HTML bernama "Index" (isi dari Index.html).
 * 3. Deploy > New deployment > Web app > Execute as: Me, Who has access: sesuai kebutuhan.
 * 4. Buka URL web app di iPhone (Safari) > Share > Add to Home Screen.
 *
 * Semua parameter bisnis ada di CONFIG di bawah. Ubah di sini, tidak perlu menyentuh rumus lain.
 */

const CONFIG = {
  // Kosongkan bila script dibuat dari menu Extensions > Apps Script di spreadsheet rekap (bound script).
  // Isi dengan ID spreadsheet bila script berdiri sendiri (standalone).
  SPREADSHEET_ID: '',

  // Kata kunci nama tab. Pencarian mengabaikan spasi, tanda baca dan huruf besar/kecil,
  // dan cukup salah satu kata kunci yang terkandung di nama tab.
  SHEETS: {
    RCV: ['PEMBONGKARAN GULA', 'PEMBONGKARAN'],                 // Receiving
    JB:  ['TIMBANG JBMOBIL', 'TIMBANG JB', 'JBMOBIL', 'JUMBO BAG', 'PENGADAAN JB'], // Pengadaan jumbo bag (opsional)
    RBG: ['HASIL REBAGGING', 'REBAGGING', 'RE-BAGGING'],        // Re-bagging
    OUT: ['PENGELUARAN GULA', 'PENGELUARAN'],                   // Delivery keluar TPS
    PG:  ['PENERIMAAN GULA', 'PENERIMAAN']                      // Penerimaan di PG
  },
  // Isi manual bila tab pengadaan jumbo bag tidak ada (jumlah JB yang sudah diadakan). 0 = pakai tab.
  JB_PROCURED_MANUAL: 0,

  // Baris pertama data (di bawah header)
  START_ROW: { RCV: 14, JB: 12, RBG: 14, OUT: 14, PG: 14 },

  // Ringkasan stowage per palka di tab Receiving (nama palka di kolom E, stowage di kolom G)
  PALKA_RANGE: 'E4:G8',

  PARTY_KG: 50000000,          // 50.000 MT
  SLA_START: '2026-10-27',     // Hari ke-1 SLA (mulai bongkar)
  SLA_DAYS: 50,                // SLA pengiriman ke PG
  TARGET_JB_PER_DAY: 1000,     // Target re-bagging per hari
  SUSUT_MAX: 0.002,            // Toleransi susut 0,2%
  DUP_WINDOW_MIN: 45,          // Dua timbangan truk yang sama dengan gross & tare identik dalam X menit = double input
  SUSPECT_WINDOW_MIN: 120,     // Netto sama pada truk yang sama dalam X menit = suspect (siklus normal ±3 jam)
  TD_JEMTIM_TOL_KG: 500,       // Toleransi selisih timbang duduk vs jembatan timbang
  CACHE_SECONDS: 300,          // Cache hasil hitung 5 menit agar dashboard cepat
  TZ: 'Asia/Jakarta'
};

/* ============================ WEB APP ============================ */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Monitoring MV. The Rainbow')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-status-bar-style', 'black-translucent')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Dipanggil dari halaman. force = true untuk melewati cache. */
function getDashboard(force) {
  const cache = CacheService.getScriptCache();
  const key = 'dash_v1';
  if (!force) {
    const hit = cache.get(key);
    if (hit) return hit;
  }
  const json = JSON.stringify(buildDashboard_());
  try { cache.put(key, json, CONFIG.CACHE_SECONDS); } catch (e) { /* payload > 100KB: lewati cache */ }
  return json;
}

/** Jalankan manual dari editor untuk mengecek hasil di log. */
function testDashboard() {
  const d = buildDashboard_();
  Logger.log(JSON.stringify(d.kpi, null, 2));
  Logger.log(JSON.stringify(d.concerns, null, 2));
}

/* ============================ KALKULASI ============================ */

function buildDashboard_() {
  const ss = CONFIG.SPREADSHEET_ID ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  const today = dayKey_(new Date());

  const rcv = readReceiving_(ss);
  const jbProc = readJbProcurement_(ss);
  const rbg = readRebagging_(ss);
  const out = readOutbound_(ss);
  const pg = readPg_(ss);

  /* ---------- RECEIVING ---------- */
  const counted = rcv.rows.filter(r => r.counted);
  const rcvTotal = sum_(counted, 'netto');
  const flags = detectDoubleInput_(rcv.rows);
  const dupKg = flags.filter(f => f.level === 'double' && f.counted).reduce((a, f) => a + f.netto, 0);
  const suspectKg = flags.filter(f => f.level === 'suspect' && f.counted).reduce((a, f) => a + f.netto, 0);
  const latentKg = flags.filter(f => f.level === 'double' && !f.counted).reduce((a, f) => a + f.netto, 0);
  const base = rcvTotal - dupKg;                       // volume terkoreksi = dasar rebagging & delivery

  const rcvDaily = groupDaily_(counted, r => r.day, r => ({ truk: 1, kg: r.netto }));
  const palka = rcv.palka.map(p => {
    const kg = sum_(counted.filter(r => norm_(r.palka) === norm_(p.name)), 'netto');
    return { name: p.name, plan: p.plan, kg: kg, pct: p.plan ? kg / p.plan : 0, truk: counted.filter(r => norm_(r.palka) === norm_(p.name)).length };
  });

  /* ---------- REBAGGING ---------- */
  const rbgJb = sum_(rbg, 'jb');
  const rbgKg = sum_(rbg, 'kg');
  const kgPerJb = rbgJb ? rbgKg / rbgJb : 1115;
  const rbgDays = unique_(rbg.map(r => r.day)).filter(d => d !== today);
  const rbgFull = rbg.filter(r => r.day !== today);
  const rbgAvg = rbgDays.length ? sum_(rbgFull, 'jb') / rbgDays.length : 0;
  const rbgCut = rbg.length ? maxDay_(rbg.map(r => r.day)) : today;
  const slaEnd = addDays_(CONFIG.SLA_START, CONFIG.SLA_DAYS - 1);
  const daysLeft = Math.max(diffDays_(rbgCut, slaEnd), 1);
  const rbgRemJb = Math.max(Math.ceil((base - rbgKg) / kgPerJb), 0);
  const rbgDaysNeed = rbgAvg ? rbgRemJb / rbgAvg : null;
  const rbgDaysNeedTarget = rbgRemJb / CONFIG.TARGET_JB_PER_DAY;
  const rbgDaily = rbg.map(r => ({ day: r.day, jb: r.jb, kg: r.kg, hit: r.jb >= CONFIG.TARGET_JB_PER_DAY }));

  /* ---------- DELIVERY ---------- */
  const outDated = out.filter(r => r.day);
  const shipDays = unique_(outDated.map(r => r.day)).filter(d => d !== today);
  const shipFullJb = sum_(outDated.filter(r => r.day !== today), 'jb');
  const shipAvgJb = shipDays.length ? shipFullJb / shipDays.length : 0;
  const shipCut = shipDays.length ? maxDay_(shipDays) : today;
  const shippedKg = out.reduce((a, r) => a + (r.netto > 0 ? r.netto : r.td), 0);
  const shippedJb = sum_(out, 'jb');
  const shipRemJb = Math.max(Math.ceil((base - shippedKg) / kgPerJb), 0);
  const shipDaysLeft = Math.max(diffDays_(shipCut, slaEnd), 1);
  const shipDaysNeed = shipAvgJb ? shipRemJb / shipAvgJb : null;
  const avgTruckKg = out.length ? shippedKg / out.length : 36000;
  const outDaily = groupDaily_(outDated, r => r.day, r => ({ truk: 1, jb: r.jb, kg: r.netto > 0 ? r.netto : r.td }));

  // Join TPS ↔ PG via No. SJ
  const pgBySj = {};
  pg.forEach(p => { if (p.sj) pgBySj[p.sj] = p; });
  const outBySj = {};
  out.forEach(o => { if (o.sj) outBySj[o.sj] = o; });

  const trips = [];
  const issues = [];
  out.forEach(o => {
    const p = pgBySj[o.sj];
    const muat = o.netto > 0 ? o.netto : o.td;
    let status, selisih = null, pct = null;
    if (!p) status = 'transit';
    else if (!(p.tare > 0)) status = 'tare';
    else {
      selisih = p.netto - (p.muat || muat);
      pct = selisih / (p.muat || muat);
      status = pct < -CONFIG.SUSUT_MAX ? 'over' : (selisih < 0 ? 'susut' : (selisih > 0 ? 'lebih' : 'ok'));
    }
    trips.push({ sj: o.sj, day: o.day, plat: o.plat, tujuan: o.tujuan, pg: p ? p.tujuan : null, jb: o.jb, muat: muat, terima: p ? p.netto : null, selisih: selisih, pct: pct, status: status });

    if (o.netto <= 0) issues.push(issue_('Belum ditimbang jemtim TPS', o.sj, o.plat, 'Netto TPS 0, timbang duduk ' + fmt_(o.td) + ' kg'));
    if (!o.day) issues.push(issue_('Tanggal kirim kosong', o.sj, o.plat, 'Isi tanggal kirim di tab Pengeluaran'));
    if (o.netto > 0 && o.td > 0 && Math.abs(o.td - o.netto) > CONFIG.TD_JEMTIM_TOL_KG)
      issues.push(issue_('Selisih timbang duduk vs jemtim', o.sj, o.plat, 'TD ' + fmt_(o.td) + ' vs jemtim ' + fmt_(o.netto) + ' kg'));
    if (p) {
      if (p.tujuan && o.tujuan && norm_(p.tujuan) !== norm_(o.tujuan))
        issues.push(issue_('Tujuan TPS ≠ PG', o.sj, o.plat, 'TPS: ' + o.tujuan + ' · PG: ' + p.tujuan));
      if (p.jb && o.jb && p.jb !== o.jb)
        issues.push(issue_('Jumlah JB beda', o.sj, o.plat, 'TPS ' + o.jb + ' JB · PG ' + p.jb + ' JB'));
      if (p.plat && norm_(p.plat) !== norm_(o.plat))
        issues.push(issue_('Plat beda tulis', o.sj, o.plat, 'PG menulis ' + p.plat));
      if (!(p.tare > 0) && p.gross > 0)
        issues.push(issue_('Tare PG belum diinput', o.sj, o.plat, 'Netto PG terbaca = gross ' + fmt_(p.gross) + ' kg'));
    }
  });
  pg.forEach(p => { if (p.sj && !outBySj[p.sj]) issues.push(issue_('SJ PG tidak ada di TPS', p.sj, p.plat, 'Cek nomor SJ')); });

  const done = trips.filter(t => ['over', 'susut', 'lebih', 'ok'].indexOf(t.status) >= 0);
  const muatDone = sum_(done, 'muat');
  const terimaDone = sum_(done, 'terima');
  const susutKg = done.filter(t => t.selisih < 0).reduce((a, t) => a + t.selisih, 0);
  const lebihKg = done.filter(t => t.selisih > 0).reduce((a, t) => a + t.selisih, 0);
  const susutRate = muatDone ? susutKg / muatDone : 0;
  const tolKg = CONFIG.PARTY_KG * CONFIG.SUSUT_MAX;
  const projSusut = -susutRate * CONFIG.PARTY_KG;
  const perPg = {};
  done.forEach(t => {
    const k = t.pg || t.tujuan || '-';   // dikelompokkan per PG yang menimbang
    perPg[k] = perPg[k] || { name: k, truk: 0, muat: 0, terima: 0, susut: 0, lebih: 0 };
    perPg[k].truk++; perPg[k].muat += t.muat; perPg[k].terima += t.terima;
    if (t.selisih < 0) perPg[k].susut += t.selisih; else perPg[k].lebih += t.selisih;
  });
  const pgList = Object.keys(perPg).map(k => { const x = perPg[k]; x.net = x.terima - x.muat; x.pct = x.muat ? x.net / x.muat : 0; return x; });
  const transit = trips.filter(t => t.status === 'transit');
  const tarePending = trips.filter(t => t.status === 'tare');
  const pgDaily = groupDaily_(pg.filter(p => p.day && p.tare > 0), p => p.day, p => ({ truk: 1, kg: p.netto }));

  /* ---------- JUMBO BAG ---------- */
  const jbFound = !!jbProc;
  const jbIn = jbFound ? sum_(jbProc, 'jb') : 0;
  const jbStock = jbIn - rbgJb;
  const jbNeed = Math.ceil(base / kgPerJb);

  /* ---------- KPI & CONCERN ---------- */
  const slaDay = Math.min(diffDays_(CONFIG.SLA_START, today) + 1, 999);
  const kpi = {
    today: today,
    slaStart: CONFIG.SLA_START, slaEnd: slaEnd, slaDays: CONFIG.SLA_DAYS, slaDay: slaDay,
    party: CONFIG.PARTY_KG,
    rcv: { total: rcvTotal, truk: counted.length, dupKg: dupKg, suspectKg: suspectKg, latentKg: latentKg, base: base, vsParty: rcvTotal - CONFIG.PARTY_KG, lastDay: counted.length ? maxDay_(counted.map(r => r.day)) : null },
    rbg: { jb: rbgJb, kg: rbgKg, kgPerJb: kgPerJb, avg: rbgAvg, target: CONFIG.TARGET_JB_PER_DAY, hitDays: rbgDaily.filter(d => d.hit).length, days: rbgDaily.length,
           cut: rbgCut, remJb: rbgRemJb, finish: rbgDaysNeed == null ? null : addDays_(rbgCut, Math.ceil(rbgDaysNeed)),
           finishTarget: addDays_(rbgCut, Math.ceil(rbgDaysNeedTarget)), needPace: rbgRemJb / daysLeft, pct: base ? rbgKg / base : 0 },
    out: { truk: out.length, jb: shippedJb, kg: shippedKg, avgJb: shipAvgJb, cut: shipCut, remJb: shipRemJb,
           finish: shipDaysNeed == null ? null : addDays_(shipCut, Math.ceil(shipDaysNeed)),
           delay: shipDaysNeed == null ? null : diffDays_(slaEnd, addDays_(shipCut, Math.ceil(shipDaysNeed))),
           needPace: shipRemJb / shipDaysLeft, needTruk: (shipRemJb / shipDaysLeft) * kgPerJb / avgTruckKg, pct: base ? shippedKg / base : 0,
           transit: transit.length, transitKg: sum_(transit, 'muat') },
    pg: { truk: done.length, muat: muatDone, terima: terimaDone, net: terimaDone - muatDone, susut: susutKg, lebih: lebihKg, rate: susutRate,
          tolKg: tolKg, proj: projSusut, use: tolKg ? projSusut / tolKg : 0, over: trips.filter(t => t.status === 'over').length,
          tarePending: tarePending.length, pct: base ? terimaDone / base : 0 },
    jb: { found: jbFound, in: jbIn, used: rbgJb, stock: jbFound ? jbStock : null, cover: jbFound ? jbStock / CONFIG.TARGET_JB_PER_DAY : null, need: jbNeed, short: jbFound ? Math.max(jbNeed - jbIn, 0) : null }
  };

  const concerns = buildConcerns_(kpi, flags, issues, pgList);

  return {
    generated: Utilities.formatDate(new Date(), CONFIG.TZ, 'dd MMM yyyy HH:mm'),
    kpi: kpi,
    concerns: concerns,
    rcv: { daily: rcvDaily, palka: palka, flags: flags },
    rbg: { daily: rbgDaily },
    out: { daily: outDaily, pgDaily: pgDaily, perPg: pgList,
           worst: done.filter(t => t.selisih < 0).sort((a, b) => a.pct - b.pct).slice(0, 15),
           transit: transit, tare: tarePending, issues: issues }
  };
}

function buildConcerns_(k, flags, issues, pgList) {
  const c = [];
  const dbl = flags.filter(f => f.level === 'double' && f.counted);
  if (dbl.length) c.push({ sev: 'high', area: 'Receiving', title: dbl.length + ' double input menggelembungkan penerimaan', value: '+' + fmtT_(k.rcv.dupKg) + ' t', note: 'Total ' + fmtT_(k.rcv.total) + ' t → terkoreksi ' + fmtT_(k.rcv.base) + ' t' });
  const lat = flags.filter(f => f.level === 'double' && !f.counted);
  if (lat.length) c.push({ sev: 'mid', area: 'Receiving', title: lat.length + ' baris duplikat belum terhitung (palka kosong)', value: fmtT_(k.rcv.latentKg) + ' t', note: 'Hapus. Bila palka diisi, total akan menggelembung' });
  const sus = flags.filter(f => f.level === 'suspect');
  if (sus.length) c.push({ sev: 'mid', area: 'Receiving', title: sus.length + ' suspect double, cek tiket timbang', value: fmtT_(k.rcv.suspectKg) + ' t', note: 'Netto sama, truk sama, jarak < ' + CONFIG.SUSPECT_WINDOW_MIN + ' menit (siklus normal ±3 jam)' });

  if (k.out.delay != null && k.out.delay > 0) c.push({ sev: 'high', area: 'Delivery', title: 'Pengiriman ke PG telat ' + k.out.delay + ' hari dari SLA', value: fmtD_(k.out.finish), note: 'Pace ' + fmt_(k.out.avgJb) + ' JB/hari, butuh ' + fmt_(k.out.needPace) + ' JB/hari (±' + Math.ceil(k.out.needTruk) + ' truk)' });
  if (k.rbg.finish && k.rbg.finish > k.slaEnd) c.push({ sev: 'high', area: 'Rebagging', title: 'Rebagging selesai melewati SLA', value: fmtD_(k.rbg.finish), note: 'Rata-rata ' + fmt_(k.rbg.avg) + ' JB/hari (' + Math.round(k.rbg.avg / k.rbg.target * 100) + '% target), butuh ' + fmt_(k.rbg.needPace) + ' JB/hari' });
  if (!k.jb.found) c.push({ sev: 'low', area: 'Jumbo Bag', title: 'Tab pengadaan jumbo bag tidak ditemukan', value: '-', note: 'Isi CONFIG.JB_PROCURED_MANUAL atau nama tab di CONFIG.SHEETS.JB' });
  else if (k.jb.cover < 14) c.push({ sev: k.jb.cover < 7.5 ? 'high' : 'mid', area: 'Jumbo Bag', title: 'Stok jumbo bag kosong ±' + k.jb.cover.toFixed(1).replace('.', ',') + ' hari', value: fmt_(k.jb.stock) + ' JB', note: 'Kekurangan pengadaan ' + fmt_(k.jb.short) + ' JB' });

  const useP = Math.round(k.pg.use * 100);
  c.push({ sev: k.pg.use > 0.8 ? 'high' : (k.pg.use > 0.5 ? 'mid' : 'ok'), area: 'Susut', title: 'Proyeksi susut ' + useP + '% dari toleransi 0,2%', value: '(' + fmt_(-k.pg.susut) + ') kg', note: 'Rate susut ' + pct_(k.pg.rate) + ' · proyeksi ' + fmtT_(k.pg.proj) + ' t dari maks ' + fmtT_(k.pg.tolKg) + ' t' });
  if (k.pg.over) c.push({ sev: 'mid', area: 'Susut', title: k.pg.over + ' truk susut > 0,2% per truk', value: k.pg.over + ' truk', note: 'Lihat daftar di tab Delivery' });
  pgList.forEach(p => { if (p.pct > CONFIG.SUSUT_MAX * 0.75 && p.truk >= 5) c.push({ sev: 'mid', area: 'Susut', title: p.name + ' konsisten timbang lebih berat', value: '+' + pct_(p.pct), note: 'Indikasi beda kalibrasi timbangan, bisa menutupi susut riil' }); });
  if (k.pg.tarePending) c.push({ sev: 'mid', area: 'Data', title: k.pg.tarePending + ' truk tare PG belum diinput', value: k.pg.tarePending + ' truk', note: 'Summary penerimaan PG menggelembung' });
  const oth = issues.filter(i => i.type !== 'Tare PG belum diinput');
  if (oth.length) c.push({ sev: 'low', area: 'Data', title: oth.length + ' ketidaksesuaian data TPS vs PG', value: oth.length + ' item', note: 'Plat, tujuan, jumlah JB, tanggal' });
  const order = { high: 0, mid: 1, low: 2, ok: 3 };
  return c.sort((a, b) => order[a.sev] - order[b.sev]);
}

/* ---------- Double input receiving ---------- */
function detectDoubleInput_(rows) {
  const out = [];
  const seen = {};
  const push = (r, pair, level, reason) => {
    const key = r.row;
    if (seen[key] && (seen[key].level === 'double' || level === 'suspect')) return;
    const f = { row: r.row, pair: pair.row, plat: r.plat, sj: r.sj, pairSj: pair.sj, netto: r.netto, gross: r.gross, tare: r.tare, palka: r.palka, time: r.ts ? Utilities.formatDate(r.ts, CONFIG.TZ, 'dd/MM HH:mm') : '', counted: r.counted, level: level, reason: reason };
    if (seen[key]) out.splice(out.indexOf(seen[key]), 1);
    seen[key] = f; out.push(f);
  };
  // 1. Nomor SJ identik
  const bySj = {};
  rows.forEach(r => { const s = r.sjFull; if (!s) return; if (bySj[s]) push(r, bySj[s], 'double', 'No. SJ identik'); else bySj[s] = r; });
  // 2–4. Per truk
  const byPlat = {};
  rows.forEach(r => { (byPlat[r.platKey] = byPlat[r.platKey] || []).push(r); });
  Object.keys(byPlat).forEach(p => {
    const list = byPlat[p].slice().sort((a, b) => (a.ts || 0) - (b.ts || 0));
    for (let i = 0; i < list.length; i++) {
      for (let j = 0; j < i; j++) {
        const a = list[j], b = list[i];
        const gapMin = (a.ts && b.ts) ? Math.abs(b.ts - a.ts) / 60000 : null;
        if (a.sjNo && a.sjNo === b.sjNo) {
          const why = (a.gross === b.gross && a.tare === b.tare) ? ', gross & tare identik'
            : (a.tare === b.tare ? ', tare sama, gross beda ' + fmt_(Math.abs(b.gross - a.gross)) + ' kg (indikasi salah ketik)' : ', timbangan berbeda');
          push(b, a, 'double', 'Nomor urut SJ sama (' + a.sjNo + ')' + why);
        } else if (a.gross === b.gross && a.tare === b.tare && gapMin != null && gapMin <= CONFIG.DUP_WINDOW_MIN) {
          push(b, a, 'double', 'Gross & tare identik, jarak ' + Math.round(gapMin) + ' menit');
        } else if (a.netto === b.netto && gapMin != null && gapMin <= CONFIG.SUSPECT_WINDOW_MIN && a.counted && b.counted) {
          push(b, a, 'suspect', 'Netto sama, jarak ' + Math.round(gapMin) + ' menit' + (a.gross === b.gross && a.tare === b.tare ? ', gross & tare identik' : ''));
        }
      }
    }
  });
  return out.sort((a, b) => a.row - b.row);
}

/* ============================ PEMBACA SHEET ============================ */

function sheet_(ss, key, optional) {
  const keys = [].concat(CONFIG.SHEETS[key]).map(k => String(k).toUpperCase().replace(/[^A-Z0-9]/g, ''));
  const all = ss.getSheets();
  const nm = s => s.getName().toUpperCase().replace(/[^A-Z0-9]/g, '');
  let sh = null;
  keys.some(k => (sh = all.filter(s => nm(s).indexOf(k) === 0)[0] || all.filter(s => nm(s).indexOf(k) >= 0)[0] || null));
  if (!sh && !optional) throw new Error('Tab untuk "' + [].concat(CONFIG.SHEETS[key])[0] + '" tidak ditemukan. Tab yang ada: ' +
    all.map(s => '"' + s.getName() + '"').join(', ') + '. Tambahkan nama tab yang benar di CONFIG.SHEETS.' + key + '.');
  return sh;
}

function values_(ss, key, ncol, optional) {
  const sh = sheet_(ss, key, optional);
  if (!sh) return null;
  const start = CONFIG.START_ROW[key];
  const last = sh.getLastRow();
  if (last < start) return [];
  return sh.getRange(start, 1, last - start + 1, ncol).getValues().map((v, i) => { v.row = start + i; return v; });
}

function readReceiving_(ss) {
  // A No | B Tanggal Shift | C No Polisi | D Palka | E Gudang | F No SJ | G Driver | H Gross | I Tare | J Netto | K Waktu | L Petugas | M Remarks | N - | O Sort (datetime)
  const sh = sheet_(ss, 'RCV');
  const pv = sh.getRange(CONFIG.PALKA_RANGE).getValues();
  const palka = pv.filter(r => r[0]).map(r => ({ name: String(r[0]).trim(), plan: num_(r[2]) }));
  const names = palka.map(p => norm_(p.name));
  const rows = values_(ss, 'RCV', 15).filter(v => String(v[2]).trim()).map(v => {
    const sj = String(v[5] || '').trim();
    const ts = v[14] instanceof Date ? v[14] : joinDateTime_(v[1], v[10]);
    return {
      row: v.row, day: dayKey_(v[1]), plat: String(v[2]).trim(), platKey: norm_(String(v[2]).split(/P\.\d/)[0]),
      palka: String(v[3] || '').trim(), counted: names.indexOf(norm_(v[3])) >= 0,
      sj: sj, sjFull: sj.replace(/\s/g, '').indexOf('/') > 0 && sj.split('/')[1].trim() ? sj.replace(/\s/g, '') : '',
      sjNo: sj.split('/')[0].replace(/\D/g, ''),
      gross: num_(v[7]), tare: num_(v[8]), netto: num_(v[9]), ts: ts
    };
  });
  return { rows: rows, palka: palka };
}

function readJbProcurement_(ss) {
  // A No | B Tanggal | C No Polisi | D No SJ | E Driver | F Jumlah JB
  if (CONFIG.JB_PROCURED_MANUAL > 0) return [{ day: '', jb: CONFIG.JB_PROCURED_MANUAL }];
  const v = values_(ss, 'JB', 6, true);           // opsional: dashboard tetap jalan tanpa tab ini
  if (!v) return null;
  return v.filter(r => num_(r[5]) > 0).map(r => ({ day: dayKey_(r[1]), jb: num_(r[5]) }));
}

function readRebagging_(ss) {
  // A No | B Tanggal | C Jumlah Bag | D Realisasi (Kg)
  return values_(ss, 'RBG', 4).filter(v => dayKey_(v[1]) && num_(v[2]) > 0)
    .map(v => ({ day: dayKey_(v[1]), jb: num_(v[2]), kg: num_(v[3]) }));
}

function readOutbound_(ss) {
  // A No | B Tgl Kirim | C Plat | D Supir | E Telp | F No SJ | G Milik | H Tujuan | I JB | J Timbang Duduk | K Gross | L Tare | M Netto
  return values_(ss, 'OUT', 13).filter(v => String(v[2]).trim() && String(v[5]).trim()).map(v => ({
    day: dayKey_(v[1]), plat: String(v[2]).trim(), sj: sjKey_(v[5]), tujuan: String(v[7] || '').trim(),
    jb: num_(v[8]), td: num_(v[9]), netto: num_(v[12])
  }));
}

function readPg_(ss) {
  // A No | B Tgl | C Plat | D Sopir | E No SJ | F Milik | G Tujuan | H JB | I Gross | J Tare | K Netto | L Netto Muat
  return values_(ss, 'PG', 12).filter(v => String(v[4]).trim()).map(v => {
    const gross = num_(v[8]), tare = num_(v[9]);
    return { day: dayKey_(v[1]), plat: String(v[2]).trim(), sj: sjKey_(v[4]), tujuan: String(v[6] || '').trim(), jb: num_(v[7]),
             gross: gross, tare: tare, netto: tare > 0 ? gross - tare : num_(v[10]), muat: num_(v[11]) };
  });
}

/* ============================ UTIL ============================ */

const BULAN_ = { januari: 0, februari: 1, maret: 2, april: 3, mei: 4, juni: 5, juli: 6, agustus: 7, september: 8, oktober: 9, november: 10, desember: 11 };

function dayKey_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, CONFIG.TZ, 'yyyy-MM-dd');
  if (typeof v === 'string' && v.trim()) {
    const m = v.toLowerCase().match(/(\d{1,2})\s*[- ]?\s*([a-z]+)\s*[- ]?\s*(\d{4})/);       // "Minggu,01 november 2026"
    if (m && BULAN_[m[2]] != null) return m[3] + '-' + pad_(BULAN_[m[2]] + 1) + '-' + pad_(+m[1]);
    const n = v.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);                              // 01/11/2026
    if (n) return n[3] + '-' + pad_(+n[2]) + '-' + pad_(+n[1]);
  }
  return '';
}
function joinDateTime_(d, t) {
  if (!(d instanceof Date)) return null;
  const x = new Date(d.getTime());
  if (t instanceof Date) x.setHours(t.getHours(), t.getMinutes(), t.getSeconds());
  return x;
}
function addDays_(key, n) { const p = key.split('-'); const d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2] + n)); return d.toISOString().slice(0, 10); }
function diffDays_(a, b) { const pa = a.split('-'), pb = b.split('-'); return Math.round((Date.UTC(+pb[0], +pb[1] - 1, +pb[2]) - Date.UTC(+pa[0], +pa[1] - 1, +pa[2])) / 864e5); }
function maxDay_(arr) { return arr.filter(String).sort().pop(); }
function unique_(arr) { return arr.filter((x, i) => x && arr.indexOf(x) === i); }
function pad_(n) { return (n < 10 ? '0' : '') + n; }
function num_(v) { if (typeof v === 'number') return v; const s = String(v || '').replace(/\s/g, ''); if (!s) return 0; const n = Number(s.replace(/\./g, '').replace(',', '.')); return isNaN(n) ? 0 : n; }
function norm_(s) { return String(s || '').replace(/\s/g, '').toUpperCase(); }
function sjKey_(s) { return String(s || '').replace(/\s/g, '').replace(/^'/, ''); }
function sum_(arr, k) { return arr.reduce((a, r) => a + (Number(r[k]) || 0), 0); }
function issue_(type, sj, plat, note) { return { type: type, sj: sj, plat: plat, note: note }; }
function groupDaily_(arr, keyFn, valFn) {
  const m = {};
  arr.forEach(r => { const k = keyFn(r); if (!k) return; const v = valFn(r); m[k] = m[k] || { day: k }; Object.keys(v).forEach(x => { m[k][x] = (m[k][x] || 0) + v[x]; }); });
  return Object.keys(m).sort().map(k => m[k]);
}
function fmt_(n) { return Math.round(n).toString().replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function fmtT_(kg) { return (kg / 1000).toFixed(1).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }
function pct_(x) { return (x * 100).toFixed(3).replace('.', ',') + '%'; }
function fmtD_(key) { if (!key) return '-'; const b = ['Jan','Feb','Mar','Apr','Mei','Jun','Jul','Agu','Sep','Okt','Nov','Des']; const p = key.split('-'); return +p[2] + ' ' + b[+p[1] - 1] + ' ' + p[0].slice(2); }
