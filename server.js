// server.js — Express/Nunjucks etiket yazdırma uygulaması (users.json DESTEKLİ)
const express = require("express");
const session = require("express-session");
const multer = require("multer");
const path = require("path");
const fs = require("fs");
const nunjucks = require("nunjucks");
const xlsx = require("xlsx");
const csv = require("csvtojson");
const qrcode = require("qrcode");
const iconv = require("iconv-lite");

const app = express();
const PORT = process.env.PORT || 3000;
const UPLOAD_FOLDER = "uploads";
const ALLOWED_EXTENSIONS = [".csv", ".xlsx"];

// 🚀 users.json DOSYASINI YÜKLEME ------------------------------------------
let users = [];
try {
    const usersPath = path.join(__dirname, "users.json");
    const usersData = fs.readFileSync(usersPath, "utf8");
    users = JSON.parse(usersData);
    console.log(`✅ ${users.length} kullanıcı users.json dosyasından yüklendi.`);
} catch (error) {
    console.error("❌ users.json yükleme hatası:", error.message);
}
// -------------------------------------------------------------------------

// Yardımcı fonksiyon: Metin kısaltma
const truncateText = (text, maxLength = 80) => {
    if (!text) return "";
    text = String(text).trim();
    return text.length > maxLength
        ? text.substring(0, maxLength) + "..."
        : text;
};

// ------------------------------------------ ETİKET ŞABLONLARI ------------------------------------------

// Bütçe kısaltma mantığını tekrar eden koddan kaçınmak için yardımcı fonksiyon.
const getButceKisa = (butce) => {
    butce = (butce || "").toUpperCase();
    if (butce === 'GENEL BÜTÇE') return 'GB';
    if (butce === 'DÖNER SERMAYE') return 'DS';
    return truncateText(butce, 5);
};

// Şablon 1: Detaylı Biyomedikal Etiketi (Tablo Düzeni)
const generateLabel1Html = (row, truncateText) => {
    const qrData = row.QR_KOD || "BOS";
    const qrUrl = `/qrcode/${encodeURIComponent(qrData)}`;

    const kurumAdi = row.KURUM_ADI || "KURUM ADI";
    const cihazTanimi = truncateText(row.TASINIR_TANIMI || row.URUN_ADI || "CİHAZ TANIMI", 400);
    const marka = truncateText(row.MARKA || "MARKA", 40);
    const model = truncateText(row.MODEL || "MODEL", 40);
    const seriNo = truncateText(row.SERI_NO || "SERİ NO", 40);
    const butceKisa = getButceKisa(row.BUTCE_TURU);
    const sicil = truncateText(row.SICIL_NO || "SİCİL", 40);
    const yil = row.EDINME_YILI || "YIL";
    const kunye = truncateText(row.KUNYE_NO || "KÜNYE", 12);

    return `
<div class="etiket-kutu etiket-kutu-1">
    <table class="etiket-tablo">
        <tbody>
            <tr>
                <td class="baslik-blok" colspan="4">
                    <span class="logo-konteyner">
                        <img src="/images/logo.svg" alt="Kurum Logosu" class="logo-resmi">
                    </span>
                    <span class="kurum">BURSA İL SAĞLIK MÜDÜRLÜĞÜ</span>
                    <span class="kurum-adi">${kurumAdi.substring(22)}</span>
                </td>
            </tr>

            <tr>
                <td class="cihaz-tanimi" colspan="3">
                    ${cihazTanimi}
                </td>
                <td class="qr-cell" rowspan="4">
                    <img src="${qrUrl}" alt="QR Kod">
                </td>
            </tr>

            <tr>
                <td class="alt-bilgi-deger">${marka}</td>
                <td class="alt-bilgi-deger">${model}</td>
                <td class="alt-bilgi-deger">${seriNo}</td>
            </tr>

            <tr>
                <td class="alt-bilgi-deger butce">
                    ${butceKisa}
                </td>
                <td class="sicil" colspan="2">
                    ${sicil}
                </td>
            </tr>

            <tr>
                <td class="alt-bilgi-deger">${yil}</td>
                <td class="kunye" colspan="2">
                    ${kunye}
                </td>
            </tr>
        </tbody>
    </table>
</div>
`;
};

// Şablon 2: Basit Etiket (46mm x 32mm Dikey Düzen)
const generateLabel2Html = (row, truncateText) => {
    const kurumAdi = truncateText(row.KURUM_ADI || "KURUM ADI", 80);
    const qrData = row.QR_KOD || "BOS";
    const qrUrl = `/qrcode/${encodeURIComponent(qrData)}`;
    const cihazTanimi = truncateText(row.TASINIR_TANIMI || row.URUN_ADI || "CİHAZ TANIMI", 80);
    const kunyeNo = row.KUNYE_NO || "KÜNYE NO";

    return `
<div class="etiket-kutu etiket-kutu-2">
    <p class="kurum-adi-2">${kurumAdi.substring(22)}</p>

    <div class="qr-cell-2">
        <img src="${qrUrl}" alt="QR Kod">
    </div>

    <div class="veri-alani-2">
        <p class="kunye-no-2">${kunyeNo}</p>
        <p class="urun-adi-2">${cihazTanimi}</p>
    </div>
</div>
`;
};

// Şablon 3: Sade Ayniyat Etiketi (Marka/Model bilgisi olmayan versiyon)
const generateLabel3Html = (row, truncateText) => {
    // Şablon 1'e çok benzer, ancak farklı bir stil veya veri seti için kullanılabilir.
    const qrData = row.QR_KOD || "BOS";
    const qrUrl = `/qrcode/${encodeURIComponent(qrData)}`;

    const kurumAdi = row.KURUM_ADI || "KURUM ADI";
    const cihazTanimi = truncateText(row.TASINIR_TANIMI || row.URUN_ADI || "CİHAZ TANIMI", 80);
    const seriNo = truncateText(row.SERI_NO || "SERİ NO", 40); // Seri No'yu dahil ettik
    const butceKisa = getButceKisa(row.BUTCE_TURU);
    const sicil = truncateText(row.SICIL_NO || "SİCİL", 40);
    const yil = row.EDINME_YILI || "YIL";
    const kunye = truncateText(row.KUNYE_NO || "KÜNYE", 12);

    return `
<div class="etiket-kutu etiket-kutu-3">
    <table class="etiket-tablo">
        <tbody>
            <tr>
                <td class="baslik-blok" colspan="4">
                    <span class="logo-konteyner">
                        <img src="/images/logo.svg" alt="Kurum Logosu" class="logo-resmi">
                    </span>
                    <span class="kurum">BURSA İL SAĞLIK MÜDÜRLÜĞÜ</span>
                    <span class="kurum-adi">${kurumAdi.substring(22)}</span>
                </td>
            </tr>

            <tr>
                <td class="cihaz-tanimi" colspan="3">
                    ${cihazTanimi}
                </td>
                <td class="qr-cell" rowspan="4">
                    <img src="${qrUrl}" alt="QR Kod">
                </td>
            </tr>

            <tr>
                <td class="alt-bilgi-deger butce">
                    ${butceKisa}
                </td>
                <td class="sicil" colspan="2">
                    ${sicil}
                </td>
            </tr>

            <tr>
                <td class="alt-bilgi-deger">${yil}</td>
                <td class="kunye" colspan="2">
                    ${kunye}
                </td>
            </tr>
        </tbody>
    </table>
</div>
`;
};

// Şablon 4: Yatay Şerit Etiket (Sol: QR/Künye - Sağ: Logo/Kurum/Bilgi)
const generateLabel4Html = (row, truncateText) => {
    const qrUrl = `/qrcode/${encodeURIComponent(row.QR_KOD || "BOS")}`;
    return `
<div class="etiket-kutu etiket-kutu-4">
    <div class="k4-sol">
        <img src="${qrUrl}" alt="QR" class="k4-qr">
        <div class="k4-kunye-text">${row.KUNYE_NO || ""}</div>
    </div>

    <div class="k4-sag">
        <div class="k4-ust-blok">
            <div class="k4-baslik">
                ${(row.KURUM_ADI || "").trim().split(/\s+/).slice(-3).join(" ")} 
            </div>
            <img src="/images/logox.svg" alt="Logo" class="k4-logo-sag">
        </div>
        <div class="k4-cihaz">${truncateText(row.TASINIR_TANIMI || row.URUN_ADI || "", 200)}</div>
        <div class="k4-alt"><b>SİCİL:</b> ${row.SICIL_NO || ""}</div>
    </div>
</div>
`;
};

// Şablon 5: Hurda Detay Odaklı (Hurda (HEK) Etiketi Tarzı)
const generateLabel5Html = (row, truncateText) => {
    const qrUrl = `/qrcode/${encodeURIComponent(row.QR_KOD || "BOS")}`;
    return `
<div class="etiket-kutu etiket-kutu-5">
    <div class="k5-ust">HEK (HURDA) TAKİP ETİKETİ</div>
    <div class="k5-orta">
        <div class="k5-qr"><img src="${qrUrl}"></div>
        <div class="k5-detay">
            <div><b>MARKA:</b> ${truncateText(row.MARKA || "-", 20)}</div>
            <div><b>MODEL:</b> ${truncateText(row.MODEL || "-", 20)}</div>
            <div><b>S/N:</b> ${truncateText(row.SERI_NO || "-", 20)}</div>
        </div>
    </div>
    <div class="k5-alt">${row.KUNYE_NO || "KÜNYE YOK"}</div>
</div>
`;
};

// Şablon 6: Minimalist Kare (Sadece QR ve Temel Kod)
const generateLabel6Html = (row, truncateText) => {
    const qrUrl = `/qrcode/${encodeURIComponent(row.QR_KOD || "BOS")}`;
    return `
<div class="etiket-kutu etiket-kutu-6">
    <div class="k6-qr-wrap">
        <img src="${qrUrl}">
        <div class="k6-kunye">${row.KUNYE_NO || ""}</div>
    </div>
    <div class="k6-yan-yazi">${truncateText("HİZMET ALIM", 200)}</div>
</div>
`;
};

// Şablonlar listesi
const labelGenerators = {
    template1: generateLabel1Html,
    template2: generateLabel2Html,
    template3: generateLabel3Html,
    template4: generateLabel4Html,
    template5: generateLabel5Html,
    template6: generateLabel6Html,// Yeni şablon eklendi.
};

// ----------------------------------------------------------------------------------------------------------


// Nunjucks Template Engine — **TEK CONFIG**
const env = nunjucks.configure(path.join(__dirname, "views"), {
    autoescape: true,
    express: app,
    noCache: process.env.NODE_ENV !== "production",
});

// 📌 Filtreler — **Kaybolmasın diye tek yerde tanımlı**
env.addFilter("tojson", (obj) => JSON.stringify(obj));

env.addFilter("substr", (str, start, len) => {
    return (str || "").substr(start, len);
});

// Global değişken
env.addGlobal("url_for", (route) => `/${route}`);


// Express ayarları
app.set("view engine", "html");
app.use(express.static(path.join(__dirname, "public")));
app.use(express.urlencoded({ extended: true }));
app.use(
    session({
        secret: process.env.SESSION_SECRET || "super_guvenli_anahtar",
        resave: false,
        saveUninitialized: true,
        cookie: { maxAge: 24 * 60 * 60 * 1000 },
    })
);

// uploads klasörü yoksa oluştur
if (!fs.existsSync(UPLOAD_FOLDER)) fs.mkdirSync(UPLOAD_FOLDER);

// Multer ayarları
const storage = multer.diskStorage({
    destination: (_, __, cb) => cb(null, UPLOAD_FOLDER),
    filename: (_, file, cb) => cb(null, Date.now() + "-" + file.originalname),
});

const upload = multer({
    storage,
    fileFilter: (_, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        if (ALLOWED_EXTENSIONS.includes(ext)) cb(null, true);
        else cb(new Error("Sadece CSV ve XLSX dosyaları desteklenir."));
    },
});

// middleware
app.use((req, res, next) => {
    res.locals.messages = req.session.messages || [];
    req.session.messages = [];
    res.locals.session = req.session;
    next();
});

// Login kontrol
const loginRequired = (req, res, next) => {
    if (req.session.logged_in) return next();
    req.session.messages = [
        { category: "danger", message: "Lütfen giriş yapın." },
    ];
    res.redirect("/login");
};

// ROUTES ----------------------------------------------------
app.get("/login", (req, res) => {
    res.render("login.html", {
        title: "Kullanıcı Girişi",
        messages: res.locals.messages,
    });
});

app.post("/login", (req, res) => {
    const { username, password } = req.body;

    const user = users.find(u => u.username === username && u.password === password);

    if (user) {
        req.session.logged_in = true;
        req.session.username = user.username;

        req.session.messages.push({
            category: "success",
            message: `Hoş geldiniz, ${user.username}!`,
        });

        res.redirect("/upload");
    } else {
        req.session.messages.push({
            category: "danger",
            message: "Kullanıcı adı veya parola hatalı.",
        });
        res.redirect("/login");
    }
});

app.get("/", (_, res) => res.redirect("/login"));
app.get("/logout", (req, res) =>
    req.session.destroy(() => res.redirect("/login"))
);

// Upload sayfası
app.get("/upload", loginRequired, (req, res) => {
    res.render("upload.html", {
        title: "Dosya Yükleme",
        messages: res.locals.messages,
    });
});

app.post("/upload", loginRequired, upload.single("file"), async (req, res) => {
    if (!req.file) {
        req.session.messages.push({
            category: "danger",
            message: "Lütfen bir dosya seçin.",
        });
        return res.redirect("/upload");
    }

    const filePath = req.file.path;
    const fileExt = path.extname(req.file.originalname).toLowerCase();

    try {
        let data;

        if (fileExt === ".xlsx") {
            const wb = xlsx.readFile(filePath);
            const ws = wb.Sheets[wb.SheetNames[0]];
            data = xlsx.utils.sheet_to_json(ws);
        } else if (fileExt === ".csv") {
            const fileBuffer = fs.readFileSync(filePath);
            const csvString = iconv.decode(fileBuffer, "win1254");

            const firstLine = csvString.split(/\r?\n/)[0];
            let delimiter = ",";

            if (firstLine.includes("\t") && !firstLine.includes(",")) delimiter = "\t";
            else if (firstLine.includes(";") && !firstLine.includes(",")) delimiter = ";";
            else if (firstLine.includes("\t") && firstLine.includes(",")) delimiter = "\t";

            data = await csv({ delimiter }).fromString(csvString);

            const normalizeKey = (k) =>
                k
                    .replace(/\uFEFF/g, "")
                    .replace(/\s+/g, "_")
                    .replace(/[İIı]/g, "I")
                    .replace(/[Şş]/g, "S")
                    .replace(/[Ğğ]/g, "G")
                    .replace(/[Üü]/g, "U")
                    .replace(/[Öö]/g, "O")
                    .replace(/[Çç]/g, "C")
                    .toUpperCase();

            data = data.map((row) => {
                const fixed = {};
                for (const [key, value] of Object.entries(row)) {
                    fixed[normalizeKey(key)] = value;
                }
                return fixed;
            });
        }

        fs.unlink(filePath, () => { });

        if (!data || data.length === 0) {
            req.session.messages.push({
                category: "warning",
                message: "Dosya yüklendi ancak veri yok.",
            });
            return res.redirect("/upload");
        }

        req.session.data_table = data;
        req.session.columns = Object.keys(data[0]);

        req.session.messages.push({
            category: "success",
            message: `${req.file.originalname} başarıyla yüklendi. (${data.length} satır okundu)`,
        });

        res.redirect("/table_view");
    } catch (err) {
        console.error("Yükleme hatası:", err);
        req.session.messages.push({
            category: "danger",
            message: `Hata: ${err.message}`,
        });
        res.redirect("/upload");
    }
});

// Veri tablosu
app.get("/table_view", loginRequired, (req, res) => {
    res.render("table_view.html", {
        title: "Veri Tablosu",
        columns: req.session.columns || [],
        data: req.session.data_table || [],
        template_set: true,
        messages: res.locals.messages,
    });
});

// Satır seçimi → yazdırma (GÜNCELLENDİ)
app.post("/table", loginRequired, (req, res) => {
    const selected = Array.isArray(req.body.selected_rows)
        ? req.body.selected_rows.map((i) => parseInt(i))
        : [parseInt(req.body.selected_rows || -1)].filter((x) => x >= 0);

    const templateType = req.body.template_type || 'template1';
    req.session.template_type = templateType;

    const allData = req.session.data_table || [];
    const selectedData = selected.map((i) => allData[i]).filter(Boolean);

    if (selectedData.length === 0) {
        req.session.messages.push({
            category: "warning",
            message: "Hiç satır seçilmedi.",
        });
        return res.redirect("/table_view");
    }

    req.session.selected_data_for_print = selectedData;

    req.session.messages.push({
        category: "success",
        message: `${selectedData.length} satır yazdırmaya hazırlandı.`,
    });

    res.redirect("/print_preview");
});

// QR Code
app.get("/qrcode/:data", async (req, res) => {
    try {
        const buffer = await qrcode.toBuffer(req.params.data, {
            errorCorrectionLevel: "H",
            scale: 6,
        });
        res.type("png").send(buffer);
    } catch (err) {
        res.status(500).send("QR Code oluşturulamadı.");
    }
});

// Printer status endpoint
app.get("/printer_status", (req, res) => {
    // Brother yazıcının Windows adını tam yazmalısın (Örn: "Brother PT-P900W")
    // 'Get-Printer' komutu yazıcıyı bulamazsa hata döndürür.
    exec(
        `powershell "Get-Printer | Where-Object {$_.Name -like '*P900*'} | Select-Object -ExpandProperty PrinterStatus"`,
        (err, stdout) => {
            if (err || !stdout.trim()) {
                return res.json({ online: false });
            }
            const status = stdout.trim();
            // 3 genelde "Hazır/Idle" demektir. 
            // 128 veya farklı değerler "Çevrimdışı" veya "Hata" olabilir.
            res.json({ online: (status === "3" || status === "0") });
        }
    );
});

// Yazdırma önizleme (GÜNCELLENDİ)
app.get("/print_preview", loginRequired, (req, res) => {
    const data = req.session.selected_data_for_print || [];
    const templateType = req.session.template_type || 'template1';

    if (data.length === 0) {
        req.session.messages.push({
            category: "warning",
            message: "Yazdırılacak veri yok.",
        });
        return res.redirect("/table_view");
    }

    // Şablon seçimine göre generator fonksiyonunu belirle.
    // Eğer şablonType, tanımlı listemizde yoksa, default olarak template1 kullanılır.
    const labelGenerator = labelGenerators[templateType] || labelGenerators['template1'];

    const labels = data
        .map((row) => {
            return labelGenerator(row, truncateText);
        })
        .join("\n");

    res.render("print_preview.html", {
        title: "Yazdırma Önizlemesi",
        labels,
        messages: res.locals.messages,
    });
});

// 404
app.use((req, res) =>
    res.status(404).render("404.html", { title: "Sayfa bulunamadı" })
);

// Sunucu
app.listen(PORT, () =>
    console.log(`➡ Sunucu çalışıyor: http://localhost:${PORT}`)
);
