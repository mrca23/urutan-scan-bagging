// Urutan Scan Bagging - semua proses di browser.
// Kolom dikenali dari ISI data (pola nilai), nama kolom hanya bonus kecil.
(function () {
  'use strict';

  var $ = function (id) { return document.getElementById(id); };
  var state = { header: [], rows: [], map: {}, groups: [], index: {}, jenisAktif: 0, view: 'bag' };

  var PERAN = [
    { key: 'awb', label: 'No. Waybill', wajib: true },
    { key: 'bag', label: 'No. Bagging', wajib: true },
    { key: 'waktu', label: 'Waktu Scan', wajib: true },
    { key: 'jenis', label: 'Jenis Scan', wajib: false },
    { key: 'petugas', label: 'Discan oleh', wajib: false },
    { key: 'station', label: 'Station Scan', wajib: false }
  ];

  var KODE_RE = /^[A-Z]{1,5}\d{6,16}$/;
  var KATA_SCAN = ['unpack', 'bagging', 'pack', 'incoming', 'outgoing', 'kirim', 'terima', 'sampai', 'bongkar',
    'muat', 'pickup', 'pick up', 'delivery', 'ttd', 'arrival', 'departure', 'bermasalah', 'retur', 'inventory',
    'loading', 'unloading', 'tanda terima', 'scan'];

  // ---------- util ----------
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function str(v) {
    if (v == null) return '';
    if (v instanceof Date) return waktuKey(v) || '';
    return String(v).trim();
  }
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  // Normalisasi waktu ke 'YYYY-MM-DD HH:MM:SS' (string ini sekaligus kunci urut).
  function waktuKey(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date) {
      if (isNaN(v)) return null;
      return v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate()) + ' ' +
        pad(v.getHours()) + ':' + pad(v.getMinutes()) + ':' + pad(v.getSeconds());
    }
    if (typeof v === 'number') {
      if (v < 20000 || v > 80000 || v % 1 === 0) return null; // serial Excel bertanggal+jam
      var d = XLSX.SSF.parse_date_code(v);
      if (!d) return null;
      return d.y + '-' + pad(d.m) + '-' + pad(d.d) + ' ' + pad(d.H) + ':' + pad(d.M) + ':' + pad(Math.floor(d.S));
    }
    var s = String(v).trim(), m;
    m = s.match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (m) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]) + ' ' + pad(+m[4]) + ':' + m[5] + ':' + (m[6] || '00');
    m = s.match(/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{4})[ T](\d{1,2}):(\d{2})(?::(\d{2}))?/);
    if (m) return m[3] + '-' + pad(+m[2]) + '-' + pad(+m[1]) + ' ' + pad(+m[4]) + ':' + m[5] + ':' + (m[6] || '00');
    return null;
  }
  function namaCocok(h, re) { return re.test(String(h || '').toLowerCase()) ? 1 : 0; }

  // ---------- baca file ----------
  function bacaFile(file) {
    setStatus('Membaca ' + file.name + ' ...');
    var reader = new FileReader();
    reader.onload = function (e) {
      try {
        var wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array', cellDates: true });
        var terbaik = null;
        wb.SheetNames.forEach(function (n) {
          var aoa = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null, blankrows: false });
          if (!terbaik || aoa.length > terbaik.length) terbaik = aoa;
        });
        siapkan(terbaik || [], file.name);
      } catch (err) {
        setStatus('Gagal membaca file: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  function siapkan(aoa, namaFile) {
    // baris header = baris pertama yang terisi >= 3 sel
    var hi = aoa.findIndex(function (r) { return r.filter(function (c) { return str(c) !== ''; }).length >= 3; });
    if (hi < 0) return setStatus('File kosong atau format tidak dikenali.', 'error');
    var lebar = aoa.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    state.header = [];
    for (var j = 0; j < lebar; j++) state.header.push(str(aoa[hi][j]) || ('Kolom ' + (j + 1)));
    state.rows = aoa.slice(hi + 1).filter(function (r) { return r.some(function (c) { return str(c) !== ''; }); });
    state.namaFile = namaFile;
    if (!state.rows.length) return setStatus('Tidak ada baris data di bawah header.', 'error');
    state.map = deteksiKolom(state.header, state.rows);
    tampilKolom();
    proses();
  }

  // ---------- deteksi kolom dari isi ----------
  function profil(header, rows) {
    var n = rows.length;
    return header.map(function (h, j) {
      var vals = [], waktu = [], kode = 0, awalJ = 0, awalL = 0, huruf = 0, vocab = 0;
      rows.forEach(function (r) {
        var v = r[j], s = str(v).toUpperCase();
        if (s === '') { waktu.push(null); return; }
        vals.push(s);
        var wk = waktuKey(v);
        waktu.push(wk);
        if (KODE_RE.test(s)) { kode++; if (s[0] === 'J') awalJ++; if (s[0] === 'L' && s[1] !== 'S') awalL++; }
        if (/^[A-Z][A-Z .'\-]*$/.test(s)) huruf++;
        var low = s.toLowerCase();
        if (KATA_SCAN.some(function (k) { return low.indexOf(k) >= 0; })) vocab++;
      });
      var c = vals.length || 1, uniq = new Set(vals).size;
      return {
        j: j, h: h, isi: vals.length / n, uniq: uniq, uniqRatio: uniq / c,
        kode: kode / c, awalJ: awalJ / c, awalL: awalL / c, huruf: huruf / c, vocab: vocab / c,
        waktu: waktu, waktuFrac: waktu.filter(Boolean).length / c, contoh: vals[0] || ''
      };
    });
  }

  function pilih(prof, skor, dipakai) {
    var best = null, bestSkor = 0;
    prof.forEach(function (p) {
      if (dipakai.indexOf(p.j) >= 0) return;
      var s = skor(p);
      if (s > bestSkor) { bestSkor = s; best = p.j; }
    });
    if (best != null) dipakai.push(best);
    return best;
  }

  function deteksiKolom(header, rows) {
    var prof = profil(header, rows), pakai = [], map = {};
    // No. Waybill: pola kode (huruf+angka), hampir semua unik, biasanya diawali J
    map.awb = pilih(prof, function (p) {
      if (p.kode < 0.85 || p.isi < 0.8) return 0;
      return p.kode * 2 + p.uniqRatio * 3 + p.awalJ + namaCocok(p.h, /waybill|awb|resi/);
    }, pakai);
    // No. Bagging: pola kode, berulang (1 bagging berisi banyak AWB), jumlah nilai beda paling banyak
    map.bag = pilih(prof, function (p) {
      if (p.kode < 0.8 || p.uniq < 1) return 0;
      if (p.uniqRatio > 0.9 && p.uniq > 3) return 0;
      return p.kode + p.awalL * 2 + Math.min(p.uniq, 40) / 20 + namaCocok(p.h, /bag|karung/) * 3;
    }, pakai);
    // Waktu Scan: kolom tanggal+jam paling unik; Waktu Upload biasanya lebih lambat & berkelompok
    var kolomWaktu = prof.filter(function (p) { return p.waktuFrac >= 0.9 && p.isi >= 0.8; });
    map.waktu = pilih(prof, function (p) {
      if (kolomWaktu.indexOf(p) < 0) return 0;
      var lebihAwal = 0, banding = 0;
      kolomWaktu.forEach(function (q) {
        if (q === p) return;
        p.waktu.forEach(function (w, i) { if (w && q.waktu[i]) { banding++; if (w <= q.waktu[i]) lebihAwal++; } });
      });
      return 1 + p.uniqRatio * 2 + (banding ? lebihAwal / banding : 0) +
        namaCocok(p.h, /scan/) - namaCocok(p.h, /upload|kirim|input/);
    }, pakai);
    // Jenis Scan: teks, sedikit variasi, isinya kata-kata jenis scan
    map.jenis = pilih(prof, function (p) {
      if (p.vocab < 0.8 || p.uniq > 30 || p.kode > 0.2 || p.waktuFrac > 0.2) return 0;
      return p.vocab * 3 - p.uniq / 30 + namaCocok(p.h, /jenis|tipe|type/) * 2;
    }, pakai);
    // Opsional (hanya untuk info): petugas & station
    map.petugas = pilih(prof, function (p) {
      if (p.huruf < 0.9 || p.uniq > 60 || p.isi < 0.8) return 0;
      return namaCocok(p.h, /discan|oleh|petugas|operator/) * 3;
    }, pakai);
    map.station = pilih(prof, function (p) {
      if (p.isi < 0.8 || p.uniq > 20 || p.kode > 0.2) return 0;
      return namaCocok(p.h, /station|stasiun|lokasi scan/) * 3;
    }, pakai);
    return map;
  }

  function tampilKolom() {
    var body = PERAN.map(function (r) {
      var opsi = '<option value="">' + (r.wajib ? '- pilih -' : '- tidak ada -') + '</option>' +
        state.header.map(function (h, j) {
          return '<option value="' + j + '"' + (state.map[r.key] === j ? ' selected' : '') + '>' + esc(h) + '</option>';
        }).join('');
      var j = state.map[r.key];
      var contoh = j == null ? '' : str((state.rows.find(function (x) { return str(x[j]) !== ''; }) || [])[j]);
      return '<tr><td>' + esc(r.label) + (r.wajib ? '' : ' <small>(opsional)</small>') + '</td>' +
        '<td><select data-peran="' + r.key + '">' + opsi + '</select></td>' +
        '<td class="contoh">' + esc(contoh) + '</td></tr>';
    }).join('');
    $('kolomBody').innerHTML = body;
    $('kolomBox').classList.remove('tersembunyi');
  }

  // ---------- olah urutan ----------
  function proses() {
    var m = state.map, kurang = PERAN.filter(function (r) { return r.wajib && m[r.key] == null; });
    if (kurang.length) {
      $('kolomBox').open = true;
      ['cariBox', 'listBox'].forEach(function (id) { $(id).classList.add('tersembunyi'); });
      return setStatus('Kolom ' + kurang.map(function (r) { return r.label; }).join(', ') +
        ' tidak terdeteksi. Pilih manual di tabel kolom.', 'error');
    }
    var data = [], tanpaWaktu = 0;
    state.rows.forEach(function (r, i) {
      var awb = str(r[m.awb]).toUpperCase();
      if (!awb) return;
      var w = waktuKey(r[m.waktu]);
      if (!w) { tanpaWaktu++; return; }
      data.push({
        awb: awb, bag: str(r[m.bag]).toUpperCase() || '(TANPA BAGGING)', waktu: w, baris: i,
        jenis: m.jenis == null ? 'Semua scan' : (str(r[m.jenis]) || '(kosong)'),
        petugas: m.petugas == null ? '' : str(r[m.petugas]),
        station: m.station == null ? '' : str(r[m.station])
      });
    });
    if (!data.length) return setStatus('Tidak ada baris dengan No. Waybill dan waktu scan yang valid.', 'error');

    var perJenis = {};
    data.forEach(function (d) { (perJenis[d.jenis] = perJenis[d.jenis] || []).push(d); });
    state.index = {};
    state.groups = Object.keys(perJenis).sort().map(function (jenis, gi) {
      var list = perJenis[jenis].sort(function (a, b) {
        return a.waktu < b.waktu ? -1 : a.waktu > b.waktu ? 1 : a.baris - b.baris;
      });
      var bags = [], byId = {};
      list.forEach(function (d, i) {
        d.urutGlobal = i + 1;
        var b = byId[d.bag];
        if (!b) { b = byId[d.bag] = { id: d.bag, items: [] }; bags.push(b); }
        b.items.push(d);
        d.urutBag = b.items.length;
      });
      bags.forEach(function (b, bi) {
        b.no = bi + 1;
        b.mulai = b.items[0].waktu;
        b.selesai = b.items[b.items.length - 1].waktu;
        b.items.forEach(function (d) { d.noBag = b.no; d.bagObj = b; });
      });
      list.forEach(function (d) { (state.index[d.awb] = state.index[d.awb] || []).push({ g: gi, d: d }); });
      return { jenis: jenis, list: list, bags: bags };
    });
    state.jenisAktif = 0;

    var info = data.length + ' scan, ' + state.groups.reduce(function (s, g) { return s + g.bags.length; }, 0) +
      ' bagging dari ' + state.namaFile + '.';
    if (tanpaWaktu) info += ' ' + tanpaWaktu + ' baris dilewati (waktu scan kosong).';
    setStatus(info, 'ok');
    ['cariBox', 'listBox'].forEach(function (id) { $(id).classList.remove('tersembunyi'); });
    $('hasil').innerHTML = '';
    renderTabs();
    renderList();
    $('awbInput').focus();
  }

  // ---------- tampilan daftar ----------
  function renderTabs() {
    $('tabsJenis').innerHTML = state.groups.length < 2 ? '' : state.groups.map(function (g, i) {
      return '<button class="sekunder' + (i === state.jenisAktif ? ' aktif' : '') + '" data-jenis="' + i + '">' +
        esc(g.jenis) + ' (' + g.list.length + ')</button>';
    }).join('');
  }

  function jam(w) { return w.slice(11); }

  function renderList() {
    var g = state.groups[state.jenisAktif];
    var unikAwb = new Set(g.list.map(function (d) { return d.awb; })).size;
    $('ringkas').innerHTML =
      kotak('Jenis scan', esc(g.jenis)) + kotak('Total bagging', g.bags.length) +
      kotak('Total AWB', unikAwb + (unikAwb !== g.list.length ? ' <small>(' + g.list.length + ' scan)</small>' : '')) +
      kotak('Scan pertama', esc(g.list[0].waktu)) + kotak('Scan terakhir', esc(g.list[g.list.length - 1].waktu));

    $('viewBag').innerHTML = g.bags.map(function (b) {
      return '<details class="bag" id="bag-' + b.no + '"><summary>' +
        '<span class="no">Bagging ke-' + b.no + '</span><span class="id">' + esc(b.id) + '</span>' +
        '<span class="meta">' + b.items.length + ' AWB &middot; ' + esc(b.mulai) + ' s/d ' + esc(jam(b.selesai)) + '</span>' +
        '</summary><div class="tabel"><table><thead><tr><th class="num">Waybill ke</th><th>No. Waybill</th><th>Waktu Scan</th>' +
        '<th class="num">Urutan total</th>' + (state.map.petugas != null ? '<th>Discan oleh</th>' : '') +
        '</tr></thead><tbody>' + b.items.map(function (d) {
          return '<tr data-row="' + d.urutGlobal + '"><td class="num">' + d.urutBag + '</td><td class="mono">' + esc(d.awb) +
            '</td><td>' + esc(d.waktu) + '</td><td class="num">' + d.urutGlobal + '</td>' +
            (state.map.petugas != null ? '<td>' + esc(d.petugas) + '</td>' : '') + '</tr>';
        }).join('') + '</tbody></table></div></details>';
    }).join('');

    $('viewAwb').innerHTML = '<div class="scroll"><table><thead><tr><th class="num">No</th><th>No. Waybill</th>' +
      '<th class="num">Bagging ke</th><th>No. Bagging</th><th class="num">Waybill ke</th><th>Waktu Scan</th></tr></thead><tbody>' +
      g.list.map(function (d) {
        return '<tr data-row="' + d.urutGlobal + '"><td class="num">' + d.urutGlobal + '</td><td class="mono">' + esc(d.awb) +
          '</td><td class="num">' + d.noBag + '</td><td class="mono">' + esc(d.bag) + '</td><td class="num">' + d.urutBag +
          '</td><td>' + esc(d.waktu) + '</td></tr>';
      }).join('') + '</tbody></table></div>';
    setView(state.view);
  }

  function kotak(label, isi) { return '<div><small>' + label + '</small><strong>' + isi + '</strong></div>'; }

  function setView(v) {
    state.view = v;
    document.querySelectorAll('[data-view]').forEach(function (b) { b.classList.toggle('aktif', b.dataset.view === v); });
    $('viewBag').classList.toggle('tersembunyi', v !== 'bag');
    $('viewAwb').classList.toggle('tersembunyi', v !== 'awb');
    $('aksiBag').classList.toggle('tersembunyi', v !== 'bag');
  }

  // ---------- cari ----------
  function cari() {
    var daftar = $('awbInput').value.toUpperCase().split(/[\s,;]+/).filter(Boolean);
    daftar = daftar.filter(function (a, i) { return daftar.indexOf(a) === i; });
    if (!daftar.length) { $('hasil').innerHTML = ''; return; }
    var html = [];
    daftar.forEach(function (awb) {
      var hits = state.index[awb];
      if (!hits) {
        html.push('<div class="temu tidak"><div class="awb">' + esc(awb) + '</div>' +
          '<div class="besar">Tidak ditemukan di file ini</div></div>');
        return;
      }
      hits.forEach(function (h) {
        var d = h.d, g = state.groups[h.g], b = d.bagObj;
        html.push('<div class="temu"><div class="awb">' + esc(awb) +
          (state.groups.length > 1 || hits.length > 1 ? ' &middot; ' + esc(g.jenis) : '') + '</div>' +
          '<div class="besar">Masuk bagging ke-<b>' + d.noBag + '</b> dan waybill ke-<b>' + d.urutBag + '</b></div>' +
          '<div class="detail"><span>No. Bagging: <b>' + esc(b.id) + '</b></span>' +
          '<span>Waybill ke-' + d.urutBag + ' dari ' + b.items.length + ' di bagging ini</span>' +
          '<span>Urutan ke-' + d.urutGlobal + ' dari ' + g.list.length + ' scan ' + esc(g.jenis) + '</span>' +
          '<span>Waktu scan: ' + esc(d.waktu) + '</span>' +
          (d.petugas ? '<span>Discan oleh: ' + esc(d.petugas) + '</span>' : '') +
          (d.station ? '<span>Station: ' + esc(d.station) + '</span>' : '') +
          '</div><a href="#" data-lihat="' + h.g + ':' + d.urutGlobal + '">Lihat di daftar</a></div>');
      });
    });
    $('hasil').innerHTML = html.join('');
    var pertama = $('hasil').querySelector('[data-lihat]');
    if (pertama) sorot(pertama.dataset.lihat, false);
  }

  function sorot(kunci, gulir) {
    var p = kunci.split(':'), gi = +p[0], urut = +p[1];
    if (gi !== state.jenisAktif) { state.jenisAktif = gi; renderTabs(); renderList(); }
    var d = state.groups[gi].list[urut - 1];
    document.querySelectorAll('tr.hl').forEach(function (tr) { tr.classList.remove('hl'); });
    var wadah = state.view === 'bag' ? $('viewBag') : $('viewAwb');
    if (state.view === 'bag') { var det = $('bag-' + d.noBag); if (det) det.open = true; }
    var tr = wadah.querySelector('tr[data-row="' + urut + '"]');
    if (!tr) return;
    tr.classList.add('hl');
    if (gulir) tr.scrollIntoView({ behavior: 'smooth', block: 'center' });
  }

  // ---------- event ----------
  function setStatus(msg, cls) { var s = $('status'); s.textContent = msg; s.className = 'status' + (cls ? ' ' + cls : ''); }

  $('file').addEventListener('change', function (e) { if (e.target.files[0]) bacaFile(e.target.files[0]); e.target.value = ''; });
  var drop = $('drop');
  ['dragenter', 'dragover'].forEach(function (ev) { drop.addEventListener(ev, function (e) { e.preventDefault(); drop.classList.add('aktif'); }); });
  ['dragleave', 'drop'].forEach(function (ev) { drop.addEventListener(ev, function () { drop.classList.remove('aktif'); }); });
  drop.addEventListener('drop', function (e) {
    e.preventDefault();
    if (e.dataTransfer.files[0]) bacaFile(e.dataTransfer.files[0]);
  });
  $('kolomBody').addEventListener('change', function (e) {
    var sel = e.target.closest('select[data-peran]');
    if (!sel) return;
    state.map[sel.dataset.peran] = sel.value === '' ? null : +sel.value;
    tampilKolom();
    proses();
  });
  $('btnCari').addEventListener('click', cari);
  $('awbInput').addEventListener('keydown', function (e) {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); cari(); }
  });
  $('hasil').addEventListener('click', function (e) {
    var a = e.target.closest('[data-lihat]');
    if (!a) return;
    e.preventDefault();
    sorot(a.dataset.lihat, true);
  });
  $('tabsJenis').addEventListener('click', function (e) {
    var b = e.target.closest('[data-jenis]');
    if (!b) return;
    state.jenisAktif = +b.dataset.jenis;
    renderTabs();
    renderList();
  });
  document.querySelectorAll('[data-view]').forEach(function (b) {
    b.addEventListener('click', function () { setView(b.dataset.view); });
  });
  $('bukaSemua').addEventListener('click', function () { document.querySelectorAll('details.bag').forEach(function (d) { d.open = true; }); });
  $('tutupSemua').addEventListener('click', function () { document.querySelectorAll('details.bag').forEach(function (d) { d.open = false; }); });
})();
