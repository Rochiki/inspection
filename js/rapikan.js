/* =====================================================================
   RAPIKAN TANPA AI — aturan bahasa & pemilihan klausul (berjalan di perangkat)
   Bagian yang paling sering diedit tim:
   - DICT : kamus kata lapangan -> kata baku  (tambah baris  kata: "pengganti",)
   - ACR  : singkatan yang ditulis huruf kapital
   - T    : daftar topik. Tiap topik punya kw (kata kunci), klausul iso/smk/reg, tingkat temuan.
            Nomor klausul harus ada di js/data-klausul.js; regulasi dicocokkan dari awal namanya.
   ===================================================================== */
/* ---------------- RAPIKAN TANPA AI (berjalan di perangkat, tanpa akun & tanpa internet) ---------------- */
const RAPI = (() => {
  /* 1. Kamus bahasa lapangan -> bahasa baku */
  const DICT = {
    yg: "yang", dg: "dengan", dgn: "dengan", tdk: "tidak", gak: "tidak", ga: "tidak", gk: "tidak", nggak: "tidak", ngga: "tidak", nggk: "tidak",
    enggak: "tidak", engga: "tidak", kagak: "tidak", gada: "tidak ada", blm: "belum", blum: "belum", blon: "belum", sdh: "sudah", udh: "sudah", udah: "sudah",
    krn: "karena", karna: "karena", soalnya: "karena", utk: "untuk", bnyk: "banyak", byk: "banyak", jd: "jadi", kl: "jika", klo: "jika", kalo: "jika",
    tp: "tetapi", tapi: "tetapi", dll: "dan lain-lain", dsb: "dan sebagainya", org: "orang", pake: "menggunakan", pakai: "menggunakan", makai: "menggunakan",
    dipake: "digunakan", dipakai: "digunakan", kerjaan: "pekerjaan", bener: "benar", ditaruh: "diletakkan", ditaro: "diletakkan", naruh: "meletakkan",
    naro: "meletakkan", ancur: "hancur", bolong: "berlubang", sobek: "robek", copot: "lepas", kecopot: "terlepas", mampet: "tersumbat", kececer: "tercecer",
    numpuk: "menumpuk", ngumpul: "berkumpul", ketutup: "tertutup", ketutupan: "tertutup", kebuka: "terbuka", abis: "habis", expired: "kedaluwarsa",
    exp: "kedaluwarsa", kadaluarsa: "kedaluwarsa", kadaluwarsa: "kedaluwarsa", kedaluarsa: "kedaluwarsa", nyala: "menyala", ngerokok: "merokok",
    ngrokok: "merokok", nyetir: "mengemudi", ngebut: "melaju dengan kecepatan tinggi", masi: "masih", msh: "masih", lg: "lagi", aja: "saja", doang: "saja",
    cuma: "hanya", cuman: "hanya", sblm: "sebelum", sblum: "sebelum", stlh: "setelah", spt: "seperti", kyk: "seperti", kayak: "seperti", tsb: "tersebut",
    pd: "pada", dr: "dari", dlm: "dalam", thd: "terhadap", kpd: "kepada", dpt: "dapat", krg: "kurang", lbh: "lebih", sgt: "sangat", jln: "jalan",
    kndrn: "kendaraan", olie: "oli", pdhl: "padahal", trus: "lalu", ngerjain: "mengerjakan", dikerjain: "dikerjakan", benerin: "memperbaiki",
    dibenerin: "diperbaiki", ngelas: "mengelas", ngangkat: "mengangkat", manjat: "memanjat", colokan: "stop kontak", gede: "besar", dikit: "sedikit",
    gimana: "bagaimana", kenapa: "mengapa", sampe: "sampai", ampe: "sampai", bikin: "membuat", dibikin: "dibuat", skrg: "sekarang", skg: "sekarang",
    tgl: "tanggal", hrs: "harus", msti: "harus", mesti: "harus", bbrp: "beberapa", bberapa: "beberapa", smua: "semua", tiap: "setiap", tmpt: "tempat",
    brg: "barang", brang: "barang", dibiarin: "dibiarkan", ditinggal: "ditinggalkan", nempel: "menempel", kendor: "kendur", karatan: "berkarat",
    kepake: "terpakai", ngeces: "menetes", netes: "menetes", ngalir: "mengalir", rembesan: "rembesan", nyangkut: "tersangkut", kejepit: "terjepit",
    kesetrum: "tersengat listrik", nyetrum: "menyengat listrik", ketabrak: "tertabrak", kesandung: "tersandung", kepleset: "terpeleset", kejatuhan: "tertimpa",
    diiket: "diikat", nancep: "tertancap", nancap: "tertancap", diturunin: "diturunkan", dinaikin: "dinaikkan", dimatiin: "dimatikan",
    dinyalain: "dinyalakan", dibersihin: "dibersihkan", diberesin: "dibereskan", dirapiin: "dirapikan", beresin: "membereskan", rapiin: "merapikan",
    bersihin: "membersihkan", matiin: "mematikan", nyalain: "menyalakan", turunin: "menurunkan", naikin: "menaikkan", ngobrol: "berbincang", bareng: "bersama",
    ganti: "mengganti", nutup: "menutup", ditutupin: "ditutup", nyimpen: "menyimpan", disimpen: "disimpan", nyender: "bersandar", disenderin: "disandarkan", ngiket: "mengikat", iket: "ikat", kotor2: "kotor", lobang: "lubang", bocorin: "membocorkan", gelap2: "gelap"
  };
  const ACR = ["apd", "apar", "p3k", "p2k3", "b3", "lb3", "tps", "loto", "sop", "ik", "jsa", "jsea", "sds", "msds", "ldkb", "k3", "k3l", "qshe", "hse", "she",
    "mcu", "sio", "sia", "silo", "pjk3", "sim", "stnk", "kir", "ppe", "ptw", "pic", "ho", "bbm", "ipal", "ows", "mcb", "elcb", "rcbo", "pln", "lpg", "ac",
    "hp", "ert", "hiradc", "ga", "hrga", "fifo", "ppm"];
  const ACR_SET = new Set(ACR.filter(a => a !== "ga")); /* "ga" di catatan lapangan hampir selalu berarti "tidak" */
  const PROPER = { hitachi: "Hitachi", foton: "Foton", kpc: "KPC", hexindo: "Hexindo" };

  function fixWord(tok, changes) {
    const low = tok.toLowerCase();
    let m = low.match(/^([a-z]+)2$/);                                   /* rak2 -> rak-rak */
    if (m) {
      const b = DICT[m[1]] || m[1]; const p = b.match(/^(di|ter|ber|ke)([a-z]{3,})$/);
      const out = p ? b + "-" + p[2] : b + "-" + b; changes.push([tok, out]); return out;
    }
    if (tok.length >= 2 && tok === tok.toUpperCase() && /[A-Z]/.test(tok)) return tok;   /* sudah ditulis kapital (mis. GA, PIC) -> biarkan */
    if (DICT[low]) { changes.push([tok, DICT[low]]); return DICT[low]; }
    if (ACR_SET.has(low)) { const up = low.toUpperCase(); if (tok !== up) changes.push([tok, up]); return up; }
    m = low.match(/^([a-z0-9]+?)nya$/);                                 /* apdnya -> APD-nya */
    if (m && ACR_SET.has(m[1])) { const out = m[1].toUpperCase() + "-nya"; changes.push([tok, out]); return out; }
    if (PROPER[low]) { if (tok !== PROPER[low]) changes.push([tok, PROPER[low]]); return PROPER[low]; }
    return tok;
  }

  function bahasa(input) {
    const changes = [];
    let s = String(input || "").replace(/\r/g, "");
    if (!s.trim()) return { text: "", changes };
    s = s.replace(/\s*&\s*/g, " dan ");
    s = s.replace(/([A-Za-z\u00C0-\u024F-]+)\s+(banget|bgt|bngt)(?![A-Za-z])/gi, (all, w) => { changes.push([all, "sangat " + w]); return "sangat " + w; });
    s = s.replace(/\b(dicampur|tercampur|campur|digabung|bersama|bareng)\s+sama\b/gi, (all, w) => { changes.push([all, w + " dengan"]); return w + " dengan"; });
    s = s.replace(/[A-Za-z\u00C0-\u024F0-9]+/g, t => fixWord(t, changes));
    /* gabung baris menjadi kalimat */
    s = s.split(/\n+/).map(x => x.trim()).filter(Boolean).map(x => /[.!?:;]$/.test(x) ? x : x + ".").join(" ");
    s = s.replace(/[ \t]+/g, " ")
      .replace(/\s+([,.;:!?])/g, "$1")
      .replace(/([,;:])(?=[A-Za-z])/g, "$1 ")
      .replace(/\.(?=[A-Za-z]{2,})/g, ". ")
      .replace(/\.{2,}/g, ".").replace(/!{2,}/g, "!").replace(/\?{2,}/g, "?")
      .replace(/\btidak tidak\b/gi, "tidak").trim();
    if (!/[.!?]$/.test(s)) s += ".";
    s = s.replace(/(^|[.!?]\s+)([a-z])/g, (a, p, c) => p + c.toUpperCase());
    const seen = new Set(); const uniq = [];
    changes.forEach(([a, b]) => { const k = a.toLowerCase() + ">" + b; if (a !== b && !seen.has(k)) { seen.add(k); uniq.push([a, b]); } });
    return { text: s, changes: uniq };
  }

  /* 2. Referensi dari KB yang sudah ada di aplikasi */
  const isoTitle = (std, k) => { const src = std.startsWith("ISO 45001") ? KB.iso45 : std.startsWith("ISO 14001") ? KB.iso14 : KB.iso9; const h = src.find(x => x[0] === k); return h ? h[1] : ""; };
  const REG = key => { const h = KB.regs.find(([n]) => n.startsWith(key)); return h ? h[0] : null; };
  const H = "ISO 45001:2018", E = "ISO 14001:2026", Q = "ISO 9001:2026";

  /* 3. Aturan topik. Urutan = prioritas bila skor sama (yang lebih spesifik di atas) */
  const T = [
    { id: "loto", label: "Isolasi energi / LOTO", kw: ["loto", "lock out", "lockout", "tag out", "tagout", "gembok", "isolasi energi", "master switch", "battery disconnect", "kunci kontak", "mesin hidup", "unit menyala"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /(tanpa|tidak ada)\s+(loto|lock|gembok|tag)|(tidak|belum)\s+\w*(pasang|masang|terap|pakai|guna)\w*\s+(loto|lock|gembok|tag)|unit menyala|mesin hidup/,
      judul: "Isolasi energi (LOTO) tidak diterapkan", syarat: "ketentuan isolasi energi berbahaya melalui penerapan lock out tag out sebelum perbaikan atau perawatan",
      bahaya: "Energi berbahaya (mekanik, hidrolik, listrik) yang tidak terisolasi", risiko: "Hal ini berpotensi menyebabkan unit atau komponen bergerak/menyala tiba-tiba sehingga pekerja dapat terjepit, tergilas, atau tersengat listrik.",
      iso: [[H, "8.1.2"], [H, "8.1.1"]], smk: ["6.5.8", "6.1.5"], reg: ["UU No. 1 Tahun 1970", "Permenaker No. 38 Tahun 2016"],
      segera: "Hentikan pekerjaan, pasang gembok dan tag pada titik isolasi energi, lalu verifikasi zero energy sebelum melanjutkan.",
      ca: [["Rekayasa", "Sediakan lock box, gembok personal, dan tag LOTO di setiap workshop dan mobile workshop"], ["Administratif", "Terapkan verifikasi LOTO oleh pengawas sebelum pekerjaan perbaikan unit dimulai"]],
      cegah: "Masukkan verifikasi LOTO ke checklist harian pengawas di seluruh cabang dan site.", bukti: "Foto penerapan LOTO dan checklist verifikasi terisi.", pic: "Workshop / Service",
      rc: ["Metode", "penerapan LOTO belum menjadi langkah wajib yang diverifikasi sebelum pekerjaan."] },
    { id: "support", label: "Penyangga unit / bekerja di bawah beban", kw: ["jack stand", "jackstand", "dongkrak", "ganjal", "diganjal", "wheel chock", "chock", "cribbing", "blocking", "kolong unit", "di bawah unit", "boom", "bucket", "arm", "attachment"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /(tanpa|tidak (ada|di)?)\s*(ganjal|jack stand|jackstand|chock|cribbing)|tidak diganjal|hanya (menggunakan )?dongkrak|kolong unit|di bawah unit|(boom|bucket|arm|attachment) (tidak|belum) (di)?turun/,
      judul: "Unit/komponen tidak ditopang dengan aman", syarat: "prosedur kerja aman saat perbaikan unit alat berat, termasuk penopangan dengan jack stand/cribbing, pengganjalan roda, dan penurunan attachment ke tanah",
      bahaya: "Unit atau komponen berat yang dapat turun/bergerak (energi gravitasi)", risiko: "Hal ini berpotensi menyebabkan pekerja tertimpa atau terjepit unit/komponen dengan konsekuensi cedera berat hingga fatal.",
      iso: [[H, "8.1.2"], [H, "8.1.1"]], smk: ["6.1.3", "6.2.1"], reg: ["UU No. 1 Tahun 1970", "Permenaker No. 38 Tahun 2016"],
      segera: "Hentikan pekerjaan di bawah unit, turunkan attachment ke tanah, pasang jack stand/cribbing dan ganjal roda sebelum melanjutkan.",
      ca: [["Rekayasa", "Sediakan jack stand/cribbing berkapasitas sesuai berat unit dan ganjal roda di setiap area servis"], ["Administratif", "Cantumkan langkah penopangan unit dalam IK/JSEA servis dan lakukan briefing sebelum kerja"]],
      cegah: "Lakukan inspeksi rutin kelayakan jack stand dan cribbing serta sosialisasi ke seluruh teknisi lapangan.", bukti: "Foto unit tertopang dengan benar dan daftar hadir briefing.", pic: "Workshop / Service",
      rc: ["Metode", "langkah penopangan unit belum ditegakkan dalam prosedur kerja di lapangan."] },
    { id: "lifting", label: "Pesawat angkat & alat bantu angkat", kw: ["crane", "hoist", "chain block", "sling", "shackle", "lifting", "forklift", "overhead crane", "jib crane", "gantry", "webbing", "wire rope", "hook", "safety latch", "beban tergantung", "angkat beban"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /(putus|aus|sobek|robek|terkelupas|tertekuk|berkarat|rusak).{0,40}(sling|webbing|wire rope|hook|shackle)|(sling|webbing|wire rope|hook|shackle).{0,40}(putus|aus|sobek|robek|tertekuk|rusak)|tanpa safety latch|safety latch (tidak ada|rusak|hilang)|di bawah beban|beban tergantung/,
      judul: "Alat angkat/lifting gear tidak layak", syarat: "persyaratan K3 pesawat angkat dan alat bantu angkat, meliputi kelayakan, riksa uji, dan operator yang kompeten",
      bahaya: "Beban terangkat yang dapat jatuh atau berayun", risiko: "Hal ini berpotensi menyebabkan beban jatuh sehingga pekerja dapat cedera berat atau fatal dan unit/komponen rusak.",
      iso: [[H, "8.1.2"], [H, "7.2"]], smk: ["6.5.3", "6.5.7", "12.5.1"], reg: ["Permenaker No. 8 Tahun 2020"],
      segera: "Tarik lifting gear yang rusak dari pemakaian dan beri tanda \"JANGAN DIGUNAKAN\".",
      ca: [["Administratif", "Lakukan riksa uji pesawat angkat dan lifting gear oleh PJK3 serta pasang penanda hasil pemeriksaan"], ["Administratif", "Terapkan inspeksi sebelum pakai (pre-use) lifting gear dengan checklist dan color code"]],
      cegah: "Susun register seluruh pesawat angkat dan lifting gear per cabang beserta jadwal riksa uji dan lisensi operator.", bukti: "Laporan riksa uji PJK3 dan foto lifting gear pengganti.", pic: "Workshop / Service",
      rc: ["Mesin/Peralatan", "belum ada pemeriksaan berkala dan pengendalian lifting gear yang rusak."] },
    { id: "height", label: "Bekerja di ketinggian", kw: ["ketinggian", "tangga", "perancah", "scaffold", "harness", "full body harness", "lanyard", "atas unit", "atas kabin", "naik ke atas", "memanjat", "handrail", "pegangan tangga", "jatuh dari"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /tanpa (harness|pengaman|handrail)|tidak (menggunakan|memakai) (full body )?harness|(tangga|perancah|scaffold|handrail).{0,30}(patah|rusak|goyang|tidak stabil)|jatuh dari/,
      judul: "Pekerjaan di ketinggian tanpa pengaman memadai", syarat: "persyaratan K3 pekerjaan pada ketinggian, termasuk akses kerja yang aman, alat pelindung jatuh, dan izin kerja",
      bahaya: "Jatuh dari ketinggian", risiko: "Hal ini berpotensi menyebabkan pekerja jatuh dengan konsekuensi cedera berat hingga fatal.",
      iso: [[H, "8.1.2"]], smk: ["6.1.5", "6.1.6"], reg: ["Permenaker No. 9 Tahun 2016"],
      segera: "Hentikan pekerjaan di ketinggian sampai akses aman dan alat pelindung jatuh tersedia dan digunakan.",
      ca: [["Rekayasa", "Sediakan akses kerja aman (platform atau tangga berpagar) untuk pekerjaan di atas unit"], ["APD", "Sediakan dan wajibkan full body harness dengan titik tambat yang memadai"], ["Administratif", "Terapkan izin kerja ketinggian dan pelatihan bagi pekerja terkait"]],
      cegah: "Petakan seluruh pekerjaan di ketinggian di workshop dan site, lalu tetapkan pengendalian standarnya.", bukti: "Foto akses kerja/harness terpasang dan izin kerja ketinggian terisi.", pic: "Workshop / Service",
      rc: ["Metode", "pengendalian pekerjaan di ketinggian belum ditetapkan dan diawasi."] },
    { id: "listrik", label: "Instalasi listrik", kw: ["listrik", "kabel", "stop kontak", "panel", "mcb", "elcb", "rcbo", "steker", "grounding", "sambungan kabel", "roll kabel", "extension", "korsleting", "tersengat listrik"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /(terkelupas|terbuka|telanjang|meleleh|gosong|percikan|korsleting|tersengat listrik|tergenang|basah)/,
      judul: "Instalasi listrik tidak aman", syarat: "persyaratan K3 listrik di tempat kerja, meliputi kondisi kabel, panel, pengaman, dan pemeriksaan berkala",
      bahaya: "Bahaya listrik (sengatan listrik, korsleting)", risiko: "Hal ini berpotensi menyebabkan pekerja tersengat listrik atau kebakaran akibat korsleting.",
      iso: [[H, "8.1.2"]], smk: ["6.5.1", "6.5.4"], reg: ["Permenaker No. 12 Tahun 2015"],
      segera: "Putuskan aliran listrik pada titik tersebut, beri tanda bahaya, dan larang penggunaan sampai diperbaiki.",
      ca: [["Rekayasa", "Ganti kabel/komponen listrik yang rusak dan pasang penutup panel serta pengaman arus bocor"], ["Administratif", "Lakukan pemeriksaan instalasi listrik berkala oleh personel kompeten dan catat hasilnya"]],
      cegah: "Jadwalkan pemeriksaan instalasi listrik tahunan di seluruh cabang dan larang sambungan kabel tidak standar.", bukti: "Foto instalasi setelah perbaikan dan laporan pemeriksaan listrik.", pic: "GA",
      rc: ["Mesin/Peralatan", "belum ada pemeriksaan dan pemeliharaan instalasi listrik secara berkala."] },
    { id: "b3", label: "Limbah B3 / TPS", kw: ["limbah b3", "lb3", "tps", "oli bekas", "filter bekas", "majun", "aki bekas", "baterai bekas", "lampu bekas", "drum oli bekas", "limbah", "logbook limbah", "manifest"],
      aspek: ["Lingkungan"], kat: "Compliance", sev: "Major",
      crit: /(tercampur|dibuang ke|dibakar|tidak ada tps|tanpa tps|melebihi masa simpan|lebih dari \d+ hari|di luar tps|tanpa izin)/,
      judul: "Pengelolaan limbah B3 belum sesuai", syarat: "ketentuan pengelolaan limbah B3, meliputi pemilahan, pengemasan, simbol dan label, penyimpanan di TPS, serta pencatatan logbook",
      bahaya: "Limbah B3 yang tidak terkendali", risiko: "Hal ini berpotensi mencemari lingkungan dan menimbulkan ketidaksesuaian terhadap kewajiban pengelolaan limbah B3 yang dapat dikenai sanksi administratif.",
      iso: [[E, "8.1"], [E, "6.1.3"]], smk: ["9.2.3"], reg: ["Permen LHK No. 6 Tahun 2021", "Permen LH No. 14 Tahun 2013", "PP No. 22 Tahun 2021"],
      segera: "Pindahkan limbah B3 ke TPS, kemas dalam wadah tertutup sesuai jenis, dan beri simbol serta label.",
      ca: [["Administratif", "Pilah dan kemas limbah B3 sesuai jenis serta pasang simbol dan label pada setiap kemasan"], ["Administratif", "Catat keluar-masuk limbah B3 di logbook dan pantau masa simpan di TPS"]],
      cegah: "Lakukan pemeriksaan TPS limbah B3 bulanan di seluruh cabang dan site dengan checklist yang sama.", bukti: "Foto TPS/kemasan berlabel dan salinan logbook limbah B3.", pic: "Workshop / Service",
      rc: ["Manajemen", "pengendalian limbah B3 di area kerja belum diawasi secara rutin."] },
    { id: "spill", label: "Tumpahan oli / bahan bakar", kw: ["tumpah", "tumpahan", "ceceran", "tercecer", "oli", "rembes", "rembesan", "bocor", "spill", "spill kit", "bbm", "solar", "minyak", "grease", "coolant", "drip tray", "menetes"],
      aspek: ["Lingkungan", "K3"], kat: "Operational", sev: "Minor",
      major: /(bocor|rembes|tumpah|menetes|tanpa drip tray|tidak ada spill kit)/,
      crit: /(selokan|drainase|saluran air|saluran|sungai|parit|meresap ke tanah|tanah terbuka)/,
      judul: "Ceceran oli/bahan bakar tidak terkendali", syarat: "persyaratan pengendalian tumpahan dan pencegahan pencemaran lingkungan",
      bahaya: "Tumpahan oli/bahan bakar (permukaan licin, pencemaran tanah dan air)", risiko: "Hal ini berpotensi mencemari tanah dan saluran air serta menyebabkan pekerja terpeleset.",
      iso: [[E, "8.1"], [E, "8.2"]], smk: ["9.1.4"], reg: ["PP No. 22 Tahun 2021"],
      segera: "Bersihkan tumpahan menggunakan spill kit/absorbent dan kelola majun terkontaminasi sebagai limbah B3.",
      ca: [["Rekayasa", "Pasang drip tray atau secondary containment di titik potensi tetesan dan tumpahan"], ["Administratif", "Sediakan spill kit lengkap di area kerja dan latih pekerja cara penanganan tumpahan"]],
      cegah: "Terapkan inspeksi harian kebocoran unit/peralatan dan kelengkapan spill kit di seluruh workshop.", bukti: "Foto area setelah dibersihkan dan drip tray/spill kit terpasang.", pic: "Workshop / Service",
      rc: ["Metode", "pencegahan dan penanganan tumpahan belum diterapkan secara konsisten."] },
    { id: "kimia", label: "Bahan kimia berbahaya", kw: ["bahan kimia", "kimia", "thinner", "cat", "chemical", "sds", "msds", "ldkb", "solvent", "degreaser", "cairan pembersih", "aerosol", "tanpa label", "tidak berlabel", "jerigen"],
      aspek: ["K3", "Lingkungan"], kat: "Operational", sev: "Minor",
      major: /(tanpa label|tidak berlabel|tidak ada (m)?sds|tidak ada ldkb|bocor|tumpah|dekat (api|sumber panas))/,
      judul: "Pengendalian bahan kimia belum sesuai", syarat: "persyaratan pengendalian bahan kimia berbahaya, meliputi label, SDS/LDKB, dan penyimpanan yang aman",
      bahaya: "Paparan bahan kimia berbahaya dan potensi kebakaran", risiko: "Hal ini berpotensi menyebabkan iritasi, keracunan, atau kebakaran karena informasi bahaya dan cara penanganan tidak tersedia.",
      iso: [[H, "8.1.2"]], smk: ["9.3.1", "9.3.2", "9.3.3"], reg: ["Kepmenaker No. KEP.187/MEN/1999"],
      segera: "Beri label sementara pada wadah tanpa identitas dan pindahkan bahan kimia ke tempat penyimpanan yang sesuai.",
      ca: [["Administratif", "Sediakan SDS/LDKB berbahasa Indonesia di lokasi penyimpanan dan pemakaian bahan kimia"], ["Rekayasa", "Simpan bahan kimia di lemari/area khusus berventilasi dengan secondary containment"]],
      cegah: "Susun daftar bahan kimia per lokasi beserta SDS dan lakukan pemeriksaan label secara berkala.", bukti: "Foto wadah berlabel, lokasi penyimpanan, dan SDS yang tersedia.", pic: "Warehouse / Part",
      rc: ["Manajemen", "daftar dan pengendalian bahan kimia per lokasi belum tersedia."] },
    { id: "apar", label: "APAR & proteksi kebakaran", kw: ["apar", "pemadam", "fire extinguisher", "hydrant", "hidran", "alarm kebakaran", "smoke detector", "sprinkler", "kebakaran"],
      aspek: ["K3"], kat: "Compliance", sev: "Major",
      crit: /(tidak ada apar|tanpa apar|apar (kosong|habis))/,
      judul: "APAR/sarana proteksi kebakaran tidak siap pakai", syarat: "persyaratan pemasangan, pemeriksaan, dan pemeliharaan APAR serta sarana proteksi kebakaran",
      bahaya: "Kebakaran yang tidak dapat dipadamkan pada tahap awal", risiko: "Hal ini berpotensi menghambat pemadaman awal saat terjadi kebakaran sehingga kerugian dan korban dapat meluas.",
      iso: [[H, "8.2"]], smk: ["6.7.6", "6.7.7"], reg: ["Permenakertrans No. PER.04/MEN/1980", "Kepmenaker No. KEP.186/MEN/1999"],
      segera: "Ganti atau isi ulang APAR yang tidak layak dan pastikan akses ke APAR tidak terhalang.",
      ca: [["Administratif", "Lakukan inspeksi APAR bulanan (tekanan, segel, masa isi ulang) dan tempel kartu inspeksi"], ["Rekayasa", "Pasang APAR sesuai jenis dan jumlah kebutuhan area, lengkap dengan tanda dan ketinggian pemasangan yang sesuai"]],
      cegah: "Buat register APAR per cabang dengan jadwal isi ulang dan pemeriksaan yang terpantau.", bukti: "Foto APAR setelah perbaikan dan kartu inspeksi terisi.", pic: "GA",
      rc: ["Manajemen", "pemeriksaan dan pengisian ulang APAR belum terjadwal dan terpantau."] },
    { id: "evak", label: "Jalur evakuasi & keadaan darurat", kw: ["evakuasi", "jalur darurat", "pintu darurat", "assembly point", "titik kumpul", "emergency exit", "tangga darurat", "exit", "denah evakuasi", "nomor darurat"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /(pintu darurat|emergency exit|exit|tangga darurat).{0,30}(terkunci|digembok|terhalang|tertutup barang)/,
      judul: "Sarana tanggap darurat tidak memadai", syarat: "persyaratan sarana tanggap darurat, meliputi jalur evakuasi yang bebas hambatan, tanda, dan titik kumpul",
      bahaya: "Evakuasi terhambat saat keadaan darurat", risiko: "Hal ini berpotensi menghambat evakuasi saat keadaan darurat sehingga pekerja dapat terjebak.",
      iso: [[H, "8.2"]], smk: ["6.7.5", "6.7.7"], reg: ["Kepmenaker No. KEP.186/MEN/1999", "Permen PU No. 26/PRT/M/2008"],
      segera: "Bebaskan jalur/pintu darurat dari hambatan dan pastikan dapat dibuka dari dalam tanpa kunci.",
      ca: [["Rekayasa", "Pasang tanda jalur evakuasi, penerangan darurat, dan denah evakuasi di titik strategis"], ["Administratif", "Masukkan pemeriksaan jalur evakuasi ke inspeksi rutin dan lakukan simulasi evakuasi berkala"]],
      cegah: "Standarkan pemeriksaan sarana tanggap darurat di seluruh cabang dan contact office.", bukti: "Foto jalur/pintu darurat yang bebas hambatan dan tanda yang terpasang.", pic: "GA",
      rc: ["Manajemen", "pemeriksaan sarana tanggap darurat belum dilakukan secara rutin."] },
    { id: "kendaraan", label: "Kendaraan operasional & lalu lintas", kw: ["kendaraan", "mobil", "truk", "truck", "service car", "mobile workshop", "pickup", "pick up", "rem", "ban", "sabuk pengaman", "seat belt", "spion", "parkir", "kecepatan tinggi", "sim", "stnk", "kir", "muatan", "mengemudi", "pengemudi", "driver", "hauling"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /(rem (blong|tidak berfungsi)|ban gundul|tanpa sabuk|tidak (memakai|menggunakan) (sabuk|seat belt)|muatan tidak diikat|kecepatan tinggi|hp saat mengemudi|mengantuk)/,
      judul: "Kendaraan operasional tidak laik/pengoperasian tidak aman", syarat: "persyaratan kelaikan kendaraan dan tata cara berkendara yang aman",
      bahaya: "Kecelakaan lalu lintas kendaraan operasional", risiko: "Hal ini berpotensi menyebabkan kecelakaan lalu lintas dengan konsekuensi cedera pada pengemudi, penumpang, atau pihak lain.",
      iso: [[H, "8.1.2"]], smk: ["6.5.1", "12.5.1"], reg: ["UU No. 22 Tahun 2009"],
      segera: "Hentikan pengoperasian kendaraan sampai kondisi yang tidak laik diperbaiki.",
      ca: [["Administratif", "Terapkan inspeksi kendaraan sebelum berangkat (P2H) dengan checklist dan tindak lanjut perbaikan"], ["Administratif", "Pastikan pengemudi memiliki SIM dan izin mengemudi internal yang berlaku"]],
      cegah: "Pantau kepatuhan P2H dan perilaku berkendara melalui data telematika secara berkala.", bukti: "Foto kendaraan setelah perbaikan dan P2H terisi.", pic: "Mobile Workshop / Field Service",
      rc: ["Metode", "pemeriksaan kendaraan sebelum pengoperasian belum dilakukan secara konsisten."] },
    { id: "apd", label: "Alat pelindung diri", kw: ["apd", "helm", "safety shoes", "sepatu safety", "sepatu", "kacamata", "goggle", "sarung tangan", "glove", "masker", "respirator", "ear plug", "earplug", "ear muff", "rompi", "vest", "face shield", "coverall", "wearpack", "safety helmet"],
      aspek: ["K3"], kat: "Operational", sev: "Minor",
      major: /(tidak (memakai|menggunakan)|tanpa (apd|helm|sarung tangan|kacamata|sepatu)|tidak ada (apd|helm|sarung tangan|kacamata))/,
      judul: "APD tidak layak/tidak digunakan", syarat: "ketentuan penyediaan, penggunaan, dan pemeliharaan APD dalam kondisi layak pakai",
      bahaya: "APD yang tidak memberikan perlindungan", risiko: "Hal ini berpotensi menyebabkan cedera pada pekerja karena APD tidak memberikan perlindungan yang memadai.",
      iso: [[H, "8.1.2"]], smk: ["6.1.6", "6.1.7"], reg: ["Permenakertrans No. PER.08/MEN/VII/2010"],
      segera: "Tarik APD yang rusak dari penyimpanan/pemakaian dan ganti dengan APD layak pakai.",
      ca: [["Administratif", "Pisahkan dan tandai APD rusak, lalu musnahkan sesuai prosedur agar tidak dipakai kembali"], ["Administratif", "Lakukan inspeksi kondisi APD secara berkala dengan checklist dan catat hasilnya"]],
      cegah: "Tetapkan jadwal inspeksi dan penggantian APD berkala di seluruh cabang, termasuk stok minimum.", bukti: "Foto APD/penyimpanan setelah perbaikan dan catatan inspeksi APD.", pic: "GA",
      rc: ["Manajemen", "pemeriksaan dan penggantian APD belum dilakukan secara berkala."] },
    { id: "gas", label: "Tabung gas & bejana tekan", kw: ["tabung gas", "oksigen", "asetilen", "acetylene", "lpg", "kompresor", "air receiver", "tangki", "regulator gas", "flashback arrestor", "tabung"],
      aspek: ["K3"], kat: "Operational", sev: "Major",
      crit: /(bocor|tidak diikat|tanpa rantai|tanpa flashback|roboh|dekat (api|sumber panas))/,
      judul: "Tabung gas/bejana tekan tidak aman", syarat: "persyaratan K3 bejana tekanan, termasuk penyimpanan tabung, pengaman, dan pemeriksaan berkala",
      bahaya: "Tekanan tinggi dan gas mudah terbakar", risiko: "Hal ini berpotensi menyebabkan kebocoran gas, kebakaran, atau ledakan.",
      iso: [[H, "8.1.2"]], smk: ["6.5.1", "6.5.3"], reg: ["Permenaker No. 37 Tahun 2016"],
      segera: "Tegakkan dan ikat tabung dengan rantai, jauhkan dari sumber panas, dan tutup katup bila tidak digunakan.",
      ca: [["Rekayasa", "Sediakan rak tabung dengan rantai pengikat dan pasang flashback arrestor pada peralatan las potong"], ["Administratif", "Lakukan pemeriksaan berkala bejana tekan dan kompresor oleh pihak berwenang"]],
      cegah: "Buat register bejana tekan dan tabung gas per cabang lengkap dengan jadwal pemeriksaan.", bukti: "Foto tabung/rak setelah perbaikan dan catatan pemeriksaan.", pic: "Workshop / Service",
      rc: ["Mesin/Peralatan", "penyimpanan dan pemeriksaan tabung gas belum distandarkan."] },
    { id: "tools", label: "Tools, peralatan & mesin", kw: ["tools", "tool", "kunci", "obeng", "palu", "gerinda", "grinda", "bor", "mesin bubut", "guarding", "pelindung mesin", "tang", "impact", "pneumatic", "press", "mesin las", "welding", "las", "toolbox"],
      aspek: ["K3"], kat: "Operational", sev: "Minor",
      major: /(rusak|retak|patah|aus|tanpa (guard|pelindung)|pelindung (lepas|hilang|tidak ada)|dimodifikasi|kabel terkelupas)/,
      judul: "Tools/peralatan kerja tidak layak", syarat: "persyaratan pemeriksaan, pemeliharaan, dan pengendalian peralatan kerja agar aman digunakan",
      bahaya: "Peralatan kerja rusak atau tanpa pelindung", risiko: "Hal ini berpotensi menyebabkan pekerja terluka, tertusuk serpihan, atau terjepit saat menggunakan peralatan.",
      iso: [[H, "8.1.2"]], smk: ["6.5.1", "6.5.7"], reg: ["Permenaker No. 38 Tahun 2016"],
      segera: "Tarik tools yang rusak dari pemakaian dan beri tanda tidak boleh digunakan.",
      ca: [["Administratif", "Lakukan pemeriksaan kondisi tools berkala dengan checklist dan penandaan periode inspeksi"], ["Substitusi", "Ganti tools yang rusak/aus dengan tools standar sesuai spesifikasi pekerjaan"]],
      cegah: "Terapkan program pemeriksaan tools di seluruh cabang dengan rekap kondisi yang dipantau.", bukti: "Foto tools pengganti dan checklist pemeriksaan tools.", pic: "Workshop / Service",
      rc: ["Mesin/Peralatan", "pemeriksaan dan penggantian tools belum dilakukan secara berkala."] },
    { id: "p3k", label: "P3K", kw: ["p3k", "kotak obat", "first aid", "obat", "petugas p3k"],
      aspek: ["K3"], kat: "Operational", sev: "Minor",
      judul: "Sarana P3K tidak lengkap", syarat: "ketentuan penyediaan kotak P3K beserta isi dan petugas P3K",
      bahaya: "Pertolongan pertama tidak dapat diberikan", risiko: "Hal ini berpotensi memperburuk kondisi korban karena pertolongan pertama tidak dapat segera diberikan.",
      iso: [[H, "8.2"]], smk: ["6.8.1", "6.8.2"], reg: ["Permenakertrans No. PER.15/MEN/VIII/2008"],
      segera: "Lengkapi isi kotak P3K dan singkirkan obat yang kedaluwarsa.",
      ca: [["Administratif", "Lakukan pemeriksaan isi kotak P3K bulanan dengan daftar isi standar"], ["Administratif", "Tunjuk dan latih petugas P3K sesuai jumlah pekerja per area"]],
      cegah: "Standarkan isi dan jadwal pemeriksaan kotak P3K di seluruh lokasi.", bukti: "Foto kotak P3K lengkap dan kartu pemeriksaan.", pic: "HRGA",
      rc: ["Manajemen", "pemeriksaan isi P3K belum terjadwal."] },
    { id: "rambu", label: "Rambu & marka K3", kw: ["rambu", "safety sign", "signage", "tanda bahaya", "papan peringatan", "marka", "garis kuning", "line marking", "barikade", "barricade", "safety line"],
      aspek: ["K3"], kat: "Operational", sev: "Minor",
      judul: "Rambu/marka K3 tidak memadai", syarat: "ketentuan pemasangan rambu dan marka K3 sesuai standar",
      bahaya: "Bahaya yang tidak terkomunikasikan", risiko: "Hal ini berpotensi menyebabkan pekerja atau tamu tidak mengenali bahaya di area tersebut.",
      iso: [[H, "8.1.2"], [H, "7.4.2"]], smk: ["6.4.4"], reg: ["Permenaker No. 5 Tahun 2018"],
      segera: "Pasang rambu atau barikade sementara di area tersebut.",
      ca: [["Administratif", "Pasang rambu dan marka K3 permanen sesuai hasil identifikasi bahaya area"], ["Administratif", "Masukkan kondisi rambu dan marka ke checklist inspeksi rutin"]],
      cegah: "Susun peta kebutuhan rambu per area kerja di seluruh cabang.", bukti: "Foto rambu/marka yang telah terpasang.", pic: "GA",
      rc: ["Manajemen", "kebutuhan rambu belum dipetakan dari identifikasi bahaya."] },
    { id: "gudang", label: "Penyimpanan material & part", kw: ["rak", "penyimpanan", "part", "sparepart", "spare part", "gudang", "stok", "identifikasi", "fifo", "kardus", "palet", "pallet", "tumpukan"],
      aspek: ["Mutu"], kat: "Operational", sev: "Minor",
      major: /(miring|roboh|hampir jatuh|melebihi kapasitas|terlalu tinggi|tidak stabil)/,
      judul: "Penyimpanan material/part belum sesuai", syarat: "ketentuan penyimpanan, identifikasi, dan preservasi material/part yang aman",
      bahaya: "Material jatuh/roboh dan kerusakan part", risiko: "Hal ini berpotensi menyebabkan material jatuh menimpa pekerja serta part rusak atau salah ambil.",
      iso: [[Q, "8.5.4"], [Q, "8.5.2"]], smk: ["9.2.1"], reg: [],
      segera: "Rapikan dan stabilkan tumpukan, turunkan barang berat ke rak bawah, dan beri identitas pada part.",
      ca: [["Rekayasa", "Pasang label kapasitas rak dan pengaman agar barang tidak jatuh"], ["Administratif", "Terapkan penataan dan identifikasi part (lokasi rak, label, FIFO) dengan pemeriksaan berkala"]],
      cegah: "Standarkan tata letak dan identifikasi penyimpanan part di seluruh gudang cabang.", bukti: "Foto rak/penyimpanan setelah ditata.", pic: "Warehouse / Part",
      rc: ["Metode", "standar penataan penyimpanan belum diterapkan."] },
    { id: "hk", label: "Housekeeping", kw: ["housekeeping", "berantakan", "tidak tertata", "tidak rapi", "menumpuk", "kotor", "licin", "menghalangi", "jalur jalan", "lantai", "5r", "5s", "sampah"],
      aspek: ["K3"], kat: "Operational", sev: "Minor",
      major: /(licin|menghalangi|tersandung|terpeleset|jalur jalan)/,
      judul: "Housekeeping area kerja kurang baik", syarat: "ketentuan kebersihan dan kerapian tempat kerja (housekeeping)",
      bahaya: "Lingkungan kerja tidak tertata (tersandung, terpeleset, kebakaran)", risiko: "Hal ini berpotensi menyebabkan pekerja tersandung atau terpeleset serta menghambat akses keadaan darurat.",
      iso: [[H, "8.1.2"]], smk: ["6.4.3"], reg: ["Permenaker No. 5 Tahun 2018"],
      segera: "Rapikan dan bersihkan area, singkirkan barang yang tidak diperlukan dari jalur jalan.",
      ca: [["Administratif", "Terapkan program 5R dengan penanggung jawab area dan jadwal pemeriksaan"], ["Rekayasa", "Sediakan tempat penyimpanan dan marka area agar barang memiliki lokasi tetap"]],
      cegah: "Lakukan audit 5R berkala antar-area dengan skor yang dipantau manajemen cabang.", bukti: "Foto area setelah dirapikan dan hasil audit 5R.", pic: "Workshop / Service",
      rc: ["Manajemen", "tanggung jawab kerapian area belum ditetapkan dan diawasi."] },
    { id: "sampah", label: "Sampah domestik", kw: ["tempat sampah", "sampah", "pemilahan", "organik", "anorganik"],
      aspek: ["Lingkungan"], kat: "Operational", sev: "Minor",
      judul: "Pengelolaan sampah belum sesuai", syarat: "ketentuan pengelolaan dan pemilahan sampah, termasuk pemisahan sampah yang mengandung B3",
      bahaya: "Sampah tidak terpilah", risiko: "Hal ini berpotensi mencemari lingkungan dan mencampur sampah yang mengandung B3 dengan sampah biasa.",
      iso: [[E, "8.1"]], smk: [], reg: ["UU No. 18 Tahun 2008", "Permen LHK No. 9 Tahun 2024"],
      segera: "Pilah sampah yang tercampur dan pisahkan sampah yang mengandung B3 (baterai, lampu, kemasan kimia).",
      ca: [["Rekayasa", "Sediakan tempat sampah terpilah berlabel, termasuk wadah khusus sampah mengandung B3"], ["Administratif", "Sosialisasikan cara pemilahan sampah kepada pekerja dan petugas kebersihan"]],
      cegah: "Pantau kepatuhan pemilahan sampah melalui inspeksi rutin.", bukti: "Foto tempat sampah terpilah yang terpasang.", pic: "GA",
      rc: ["Material", "sarana pemilahan sampah belum tersedia."] },
    { id: "lingkerja", label: "Lingkungan kerja & ergonomi", kw: ["pencahayaan", "penerangan", "lampu mati", "gelap", "redup", "panas", "ventilasi", "pengap", "bising", "kebisingan", "debu", "ergonomi", "kursi", "meja kerja", "angkat manual", "mengangkat manual"],
      aspek: ["K3"], kat: "Operational", sev: "Minor",
      judul: "Kondisi lingkungan kerja belum memenuhi syarat", syarat: "ketentuan K3 lingkungan kerja (faktor fisika dan ergonomi) serta pemantauannya",
      bahaya: "Faktor fisika/ergonomi lingkungan kerja", risiko: "Hal ini berpotensi menyebabkan kelelahan, gangguan kesehatan, atau kesalahan kerja.",
      iso: [[H, "8.1.2"], [H, "9.1.1"]], smk: ["7.2.1", "7.2.2", "9.1.1"], reg: ["Permenaker No. 5 Tahun 2018"],
      segera: "Lakukan perbaikan sementara (penerangan tambahan/ventilasi) dan batasi durasi paparan.",
      ca: [["Rekayasa", "Perbaiki sumber masalah (penerangan, ventilasi, peredam, alat bantu angkat) sesuai hasil pengukuran"], ["Administratif", "Lakukan pengukuran lingkungan kerja oleh pihak berkompeten dan tindak lanjuti hasilnya"]],
      cegah: "Jadwalkan pengukuran lingkungan kerja berkala di seluruh cabang.", bukti: "Foto setelah perbaikan dan hasil pengukuran lingkungan kerja.", pic: "GA",
      rc: ["Lingkungan", "kondisi lingkungan kerja belum dipantau secara berkala."] },
    { id: "emisi", label: "Air limbah, emisi & genset", kw: ["genset", "asap hitam", "emisi", "cerobong", "cucian unit", "washing bay", "oil trap", "oil water separator", "ows", "septic", "ipal", "air limbah"],
      aspek: ["Lingkungan"], kat: "Compliance", sev: "Major",
      judul: "Pengendalian air limbah/emisi belum sesuai", syarat: "kewajiban pengendalian air limbah dan emisi sesuai baku mutu yang berlaku",
      bahaya: "Pencemaran air/udara", risiko: "Hal ini berpotensi melampaui baku mutu lingkungan dan menimbulkan temuan dari instansi lingkungan hidup.",
      iso: [[E, "8.1"], [E, "9.1.1"]], smk: [], reg: ["PP No. 22 Tahun 2021"],
      segera: "Bersihkan/perbaiki unit pengolah (oil trap/OWS) dan hentikan pembuangan yang tidak terkendali.",
      ca: [["Rekayasa", "Perbaiki dan rawat oil trap/OWS atau sistem pengolah sesuai kapasitas"], ["Administratif", "Lakukan pemantauan dan pengujian berkala sesuai persetujuan lingkungan"]],
      cegah: "Masukkan pemeliharaan sarana pengolah limbah ke jadwal pemeliharaan fasilitas cabang.", bukti: "Foto sarana setelah perbaikan dan hasil uji laboratorium.", pic: "GA",
      rc: ["Mesin/Peralatan", "pemeliharaan sarana pengolah limbah belum terjadwal."] },
    { id: "lisensi", label: "Lisensi, sertifikat & kompetensi", kw: ["sio", "sia", "lisensi", "sertifikat", "riksa uji", "pjk3", "silo", "kompetensi", "tidak kompeten", "belum dilatih", "pelatihan", "training", "induksi"],
      aspek: ["K3"], kat: "Compliance", sev: "Major",
      judul: "Lisensi/sertifikat belum dipenuhi", syarat: "ketentuan lisensi personel dan sertifikat kelayakan peralatan sesuai peraturan",
      bahaya: "Peralatan/personel tidak terverifikasi kelayakannya", risiko: "Hal ini berpotensi menyebabkan pengoperasian peralatan oleh personel atau dengan peralatan yang belum terverifikasi aman, serta ketidakpatuhan regulasi.",
      iso: [[H, "7.2"], [H, "9.1.2"]], smk: ["12.5.1", "6.5.3"], reg: [],
      segera: "Hentikan pengoperasian oleh personel/peralatan yang belum berlisensi sampai dipenuhi.",
      ca: [["Administratif", "Daftarkan personel ke pelatihan dan sertifikasi yang dipersyaratkan"], ["Administratif", "Jadwalkan riksa uji peralatan dan pantau masa berlaku dalam register lisensi"]],
      cegah: "Pantau masa berlaku lisensi dan sertifikat melalui sistem pengingat terpusat.", bukti: "Salinan lisensi/sertifikat yang berlaku.", pic: "QSHE",
      rc: ["Manajemen", "masa berlaku lisensi dan sertifikat belum dipantau."] },
    { id: "dokumen", label: "Dokumen & catatan", kw: ["dokumen", "sop", "ik", "instruksi kerja", "prosedur", "formulir", "form", "checklist", "ceklis", "logbook", "catatan", "record", "jsa", "jsea", "hiradc", "izin kerja", "permit", "ptw", "tidak diisi", "belum diisi"],
      aspek: ["Mutu", "K3"], kat: "Administration", sev: "Minor",
      judul: "Dokumen/catatan belum dipenuhi", syarat: "ketentuan pengendalian informasi terdokumentasi (tersedia, terbaru, dan terisi lengkap)",
      bahaya: "Pekerjaan tanpa acuan dan catatan yang valid", risiko: "Hal ini berpotensi menyebabkan pekerjaan dilakukan tanpa acuan yang benar dan bukti pelaksanaan tidak dapat ditelusuri.",
      iso: [[Q, "7.5"], [H, "7.5.3"]], smk: ["4.1.3", "4.1.4"], reg: [],
      segera: "Lengkapi dokumen/catatan yang kosong dan tarik versi dokumen yang sudah tidak berlaku.",
      ca: [["Administratif", "Sediakan dokumen versi terbaru di lokasi kerja dan tetapkan penanggung jawab pengisian"], ["Administratif", "Lakukan verifikasi kelengkapan catatan secara berkala oleh atasan langsung"]],
      cegah: "Masukkan pemeriksaan kelengkapan dokumen ke audit internal berkala.", bukti: "Salinan dokumen/catatan yang telah dilengkapi.", pic: "QSHE",
      rc: ["Manajemen", "pengendalian dan verifikasi dokumen di lapangan belum berjalan."] }
  ];
  const PERMIT = /(izin kerja|permit|ptw)/;
  const CRIT_ALL = /(fatal|meninggal|luka berat|amputasi|patah tulang|nyaris celaka.{0,20}(tertimpa|terjepit|tergilas))/;

  const norm = s => " " + String(s || "").toLowerCase().replace(/[^a-z0-9\s]/g, " ").replace(/\s+/g, " ") + " ";
  const escRe = s => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const RX = new Map();
  function hit(txt, k) {
    if (!RX.has(k)) RX.set(k, new RegExp("(^|[^a-z0-9])" + escRe(k) + "(?=$|[^a-z0-9]|nya)"));
    return RX.get(k).test(txt);
  }
  function match(text) {
    const t = norm(text);
    const res = [];
    T.forEach((top, idx) => {
      const kws = top.kw.filter(k => hit(t, k));
      if (!kws.length) return;
      const score = kws.reduce((a, k) => a + (k.includes(" ") ? 2 : 1), 0);
      res.push({ top, kws, score, idx });
    });
    res.sort((a, b) => b.score - a.score || a.idx - b.idx);
    return res;
  }
  function sevOf(top, t) {
    if (top.crit && top.crit.test(t)) return "Critical";
    if (CRIT_ALL.test(t)) return "Critical";
    if (top.major && top.major.test(t)) return "Major";
    return top.sev;
  }
  const rank = { Minor: 0, Major: 1, Critical: 2 };

  function refs(tops, t) {
    const iso = [], smk = [], reg = [];
    tops.forEach(({ top }) => {
      top.iso.forEach(([std, k]) => { if (!iso.some(x => x.standar === std && x.klausul === k)) iso.push({ standar: std, klausul: k, judul: isoTitle(std, k), alasan: "Dipilih otomatis dari kata kunci: " + top.label }); });
      top.smk.forEach(k => { if (!smk.some(x => x.kode === k)) { const h = KB.crit.find(c => c[0] === k); if (h) smk.push({ kode: k, ringkas: h[1] }); } });
      top.reg.forEach(key => { const n = REG(key); if (n && !reg.some(x => x.peraturan === n)) reg.push({ peraturan: n, pasal: "", relevansi: top.label, perlu_verifikasi: false }); });
    });
    if (PERMIT.test(t) && !smk.some(x => x.kode === "6.1.5")) { const h = KB.crit.find(c => c[0] === "6.1.5"); smk.unshift({ kode: "6.1.5", ringkas: h[1] }); }
    return { klausul_iso: iso.slice(0, 3), kriteria_smk3: smk.slice(0, 3), regulasi: reg.slice(0, 3) };
  }
  function picFor(p) {
    const list = S.settings.pics || [];
    if (list.includes(p)) return p;
    const lo = p.toLowerCase().split(/[\s/]+/)[0];
    return list.find(x => x.toLowerCase().includes(lo)) || p;
  }
  const lcFirst = s => /^([A-Z0-9]{2,}|Hitachi|Foton|KPC|Hexindo)/.test(s) ? s : s.charAt(0).toLowerCase() + s.slice(1);
  const words = (s, n) => { const w = String(s || "").trim().split(/\s+/).filter(Boolean); return w.length > n ? w.slice(0, n).join(" ") : w.join(" "); };

  /* 4. Susun satu entri temuan / good practice dari catatan kasar */
  function susun({ jenis, area, lokasi, raw }) {
    const good = jenis === "good";
    const b = bahasa(raw);
    const lok = bahasa(lokasi).text.replace(/\.$/, "");
    const t = norm(b.text + " " + lok);
    const m = match(b.text + " " + lok);
    const tops = m.slice(0, 1);
    if (tops.length) { const sec = match(b.text).find(x => x.top !== tops[0].top && x.score >= 2); if (sec) tops.push(sec); }
    const main = tops[0] && tops[0].top;
    const body = b.text.replace(/\.$/, "");
    const where = `Pada area ${area || "-"}${lok ? `, tepatnya di ${lcFirst(lok)},` : ","}`;
    const f = {};
    if (main) {
      const r = refs(tops, t);
      Object.assign(f, r);
      f.kategori_aspek = [...new Set(tops.flatMap(x => x.top.aspek))].slice(0, 3);
      f.kategori_temuan = main.kat;
      const lokShort = lok && lok.split(/\s+/).length <= 5 ? " di " + lcFirst(lok) : "";
      f.judul = words(good ? `Praktik baik: ${main.label.toLowerCase()}${lokShort}` : main.judul + lokShort, 12);
      f.bahaya = main.bahaya;
      if (good) {
        f.deskripsi = `${where} teramati praktik baik berikut: ${lcFirst(body)}. Praktik ini mendukung pemenuhan ${main.syarat}.`;
        f.potensi_risiko = ""; f.severity = ""; f.stop_work = false; f.alasan_severity = "";

      } else {
        let sev = "Minor", why = main;
        tops.forEach(({ top }) => { const s = sevOf(top, t); if (rank[s] > rank[sev]) { sev = s; why = top; } });
        f.severity = sev; f.stop_work = sev === "Critical";
        f.alasan_severity = `Ditentukan otomatis dari kata kunci topik "${why.label}" (${sev === "Critical" ? "potensi cedera berat/fatal, pelanggaran regulasi, atau pencemaran signifikan" : sev === "Major" ? "potensi cedera ringan, peringatan regulator, atau gangguan operasional" : "dampak minimal"}) — konfirmasi oleh reviewer.`;
        f.deskripsi = `${where} ditemukan kondisi berikut: ${lcFirst(body)}. Kondisi tersebut belum memenuhi ${main.syarat}. ${main.risiko}`;
        f.potensi_risiko = main.risiko.replace(/^Hal ini berpotensi /, "Berpotensi ");
      }
      f.catatan_konfirmasi = `Dirapikan otomatis tanpa AI dari kata kunci: ${tops.flatMap(x => x.kws).slice(0, 6).join(", ")}. Periksa kembali kalimat, klausul/regulasi, dan tingkat temuan sebelum dikirim.`;
    } else {
      f.judul = words(body.charAt(0).toUpperCase() + body.slice(1), 10);
      f.deskripsi = good ? `${where} teramati praktik baik berikut: ${lcFirst(body)}.` : `${where} ditemukan kondisi berikut: ${lcFirst(body)}.`;
      f.klausul_iso = good ? [] : [{ standar: H, klausul: "6.1.2.1", judul: isoTitle(H, "6.1.2.1"), alasan: "" }];
      f.kriteria_smk3 = []; f.regulasi = [];
      f.catatan_konfirmasi = "Kata kunci catatan belum dikenali, jadi hanya bahasanya yang dirapikan. Tentukan klausul, regulasi, tingkat temuan, dan tindak lanjut secara manual.";
      if (!good) f.severity = "Minor";
    }
    return { fields: f, changes: b.changes, topics: tops.map(x => x.top.label) };
  }

  /* 5. Tambah rujukan ke entri yang sudah ada (dipakai tombol di editor) */
  function tambahRujukan(w) {
    const txt = [w.judul, w.deskripsi, w.raw, w.lokasi].join(" ");
    const all = match(txt); if (!all.length) return 0;
    const m = [all[0], ...all.slice(1).filter(x => x.score >= 2).slice(0, 1)];
    const r = refs(m, norm(txt)); let n = 0;
    w.klausul_iso = w.klausul_iso || []; w.kriteria_smk3 = w.kriteria_smk3 || []; w.regulasi = w.regulasi || [];
    r.klausul_iso.forEach(x => { if (!w.klausul_iso.some(y => y.standar === x.standar && y.klausul === x.klausul)) { w.klausul_iso.push(x); n++; } });
    r.kriteria_smk3.forEach(x => { if (!w.kriteria_smk3.some(y => y.kode === x.kode)) { w.kriteria_smk3.push(x); n++; } });
    r.regulasi.forEach(x => { if (!w.regulasi.some(y => y.peraturan === x.peraturan)) { w.regulasi.push(x); n++; } });
    return n;
  }
  return { bahasa, susun, tambahRujukan, match };
})();

