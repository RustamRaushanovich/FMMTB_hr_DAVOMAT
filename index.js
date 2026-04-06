const { Telegraf, Markup, session, Scenes } = require('telegraf');
const fs = require('fs');
const moment = require('moment-timezone');
const cron = require('node-cron');
const path = require('path');
const http = require('http');

// RENDER HEALTH CHECK & KEEP-ALIVE SERVER
const PORT = process.env.PORT || 10000;
const server = http.createServer((req, res) => {
    if (req.url === '/health' || req.url === '/') {
        res.writeHead(200, { 'Content-Type': 'text/plain' });
        res.end('HR BOT IS ONLINE 24/7\n');
    } else {
        res.writeHead(404);
        res.end();
    }
});

server.listen(PORT, '0.0.0.0', () => {
    console.log(`🚀 Health check server is running on port ${PORT}`);
    // Self-ping every 14 minutes to prevent Render from sleeping
    setInterval(() => {
        const url = process.env.RENDER_EXTERNAL_URL;
        if (url) {
            http.get(url, (res) => {
                console.log(`🕒 Keep-alive ping sent to ${url}. Status: ${res.statusCode}`);
            }).on('error', (err) => {
                console.error('Keep-alive ping error:', err.message);
            });
        }
    }, 14 * 60 * 1000); 
});

const TOKEN = process.env.BOT_TOKEN || '8754716546:AAHkFMWqdPf2qi0axCTa8XSkqWtVZzghhZM';
// Adminlar ro'yxati (ID'larni raqam va string ko'rinishida tekshirish uchun)
const ADMINS = [65002404, 786314811, 5310405293]; 
const bot = new Telegraf(TOKEN);

// Adminlikni tekshirish uchun yordamchi funksiya
const isAdmin = (id) => ADMINS.map(String).includes(String(id));


const WORK_START = "09:00"; 
const WORK_END = "18:00";   
const LOGO_PATH = path.join(__dirname, 'logo.png');


const DB_PATH = './db.json';
const STAFF_LIST = [
    { id: 1, name: "Dolimov Sherzod Abdumutalovich", dept: "Rahbariyat" },
    { id: 2, name: "Teshaboev Muxiddin Maribovich", dept: "Rahbariyat" },
    { id: 3, name: "Mamajonov Vohidjon Maxmudovich", dept: "Rahbariyat" },
    { id: 4, name: "Kasimova Dilafruz Shavkatovna", dept: "Rahbariyat" },
    { id: 5, name: "Nishanov Azizxon Miramirovich", dept: "Rahbariyat" },
    { id: 6, name: "G'aniyev Abdujabbor Jumanazarovich", dept: "Inson resurslari" },
    { id: 7, name: "Turg'unboyev Boburjon Baxodirjon o'g'li", dept: "Inson resurslari" },
    { id: 8, name: "Ismatullayeva Shoiraxon Xikmatullayevna", dept: "Inson resurslari" },
    { id: 9, name: "Saydullayev Abutolib Muhammadumar o'g'li", dept: "Inson resurslari" },
    { id: 10, name: "Ergashev Azizbek Yigitaliyevich", dept: "Yuridik xizmat" },
    { id: 11, name: "Abdullayeva Iroda Yuldashaliyevna", dept: "Yuridik xizmat" },
    { id: 12, name: "Mahmudov Umidjon Nematovich", dept: "Nazorat-tahlil" },
    { id: 13, name: "Nuraliyev Axrorbek Muxammadaliyevich", dept: "Nazorat-tahlil" },
    { id: 14, name: "Akbarova Nodira Isroil qizi", dept: "Nazorat-tahlil" },
    { id: 15, name: "Jalolov Odiljon Soliyevich", dept: "Matbuot kotibi" },
    { id: 16, name: "Axmedov Nodirjon To'ychiyevich", dept: "Korrupsiyaga qarshi" },
    { id: 17, name: "Ikromov Mansur Ibragimovich", dept: "Korrupsiyaga qarshi" },
    { id: 18, name: "Usmonov Nizomiddin Sirojiddinovich", dept: "Pedagoglarni baholash" },
    { id: 19, name: "Sidiqov Doniyorbek", dept: "Pedagoglarni baholash" },
    { id: 20, name: "Isroilov Xasanboy Ibroximjon o'g'li", dept: "Pedagoglarni baholash" },
    { id: 21, name: "Madg'oziyev Abror Raximovich", dept: "Umumiy o'rta ta'lim" },
    { id: 22, name: "Ne'matjonov Lochinbek", dept: "Umumiy o'rta ta'lim" },
    { id: 23, name: "Em Yuliya Yuryevna", dept: "Umumiy o'rta ta'lim" },
    { id: 24, name: "Sodiqov Jo'rabek Odilovich", dept: "Metodik ta'minlash" },
    { id: 25, name: "Mamayusupov Sodirjon Ne'matjonovich", dept: "Metodik ta'minlash" },
    { id: 26, name: "A'zamov Dilshodjon Muxtorovich", dept: "Metodik ta'minlash" },
    { id: 27, name: "Latipova Maftuna Qaxramon qizi", dept: "Metodik ta'minlash" },
    { id: 28, name: "Ahmadaliyeva Mohidil Ma'rufjonovna", dept: "Xorijiy tillar" },
    { id: 29, name: "Xamdamova Xayotxon Xaydaraliyevna", dept: "Xorijiy tillar" },
    { id: 30, name: "Ortiqov G'ulomjon Solijonovich", dept: "O'quvchilar bilimi" },
    { id: 31, name: "Xamzayeva Shirinoy Abdurasulovna", dept: "O'quvchilar bilimi" },
    { id: 32, name: "Raximova Odinaxon Abdukarimovna", dept: "O'quvchilar bilimi" },
    { id: 33, name: "X Hakimova Feruzaxon G'ulomovna", dept: "O'quvchilar bilimi" },
    { id: 34, name: "Ergasheva Dildora Madaminovna", dept: "Maktabgacha ta'lim" },
    { id: 35, name: "Sultonov Oybek Maxammatjonovich", dept: "Maktabgacha ta'lim" },
    { id: 36, name: "Obidova Gulmiraxon", dept: "Maktabgacha ta'lim" },
    { id: 37, name: "Utamboyeva Muxabbat Abduxalimovna", dept: "Pedagog kadrlar" },
    { id: 38, name: "Sultonova Xolida Basirovna", dept: "Pedagog kadrlar" },
    { id: 39, name: "Xadiyatullayev Qaxramon Adxamovich", dept: "Litsenziyalash" },
    { id: 40, name: "Ortiqov Asilbek Akramjon o'g'li", dept: "Akkreditatsiyalash" },
    { id: 41, name: "Jo'rayev Jaxongir Nurmuhammadovich", dept: "AKT" },
    { id: 42, name: "Badalov Akbarali Yuldashaliyevich", dept: "AKT" },
    { id: 43, name: "Vaxobov Azamatjon Ulug'bek o'g'li", dept: "AKT" },
    { id: 44, name: "Turdiyev Rustamjon Raushanovich", dept: "Tarbiyaviy ishlar" },
    { id: 45, name: "Usmonova Dilfuzaxon", dept: "Tarbiyaviy ishlar" },
    { id: 46, name: "Qo'chqorova Gulnoraxon Qo'chqorovna", dept: "Tarbiyaviy ishlar" },
    { id: 47, name: "Azimov Abrorjon Maxsudaliyevich", dept: "Tarbiyaviy ishlar" },
    { id: 48, name: "Eshmatov Axror Maxmutali o'g'li", dept: "Psixologik" },
    { id: 49, name: "Esonaliyev Shoxrux Ne'matjon o'g'li", dept: "Kasbga yo'naltirish" },
    { id: 50, name: "Sultanov Qahramon Mamiraliyevich", dept: "Moliya" },
    { id: 51, name: "Ergashev Azizbek Nabiyevich", dept: "Moliya" },
    { id: 52, name: "Pulatova Mavludaxon Xolmatovna", dept: "Moliya" },
    { id: 53, name: "Ergashev Soxibjon Abdurayimovich", dept: "Moliya" },
    { id: 54, name: "Daminov Davronbek Shokirovich", dept: "Buxgalteriya" },
    { id: 55, name: "Shokirov Farrux Baxodirovich", dept: "Buxgalteriya" },
    { id: 56, name: "Umarov Azizbek", dept: "Infratuzilma" },
    { id: 57, name: "Saidov Ikromjon Idrisxonovich", dept: "Infratuzilma" },
    { id: 58, name: "Yo'ldashov Sherzodbek", dept: "Infratuzilma" },
    { id: 59, name: "Sattorov Barhayotjon Xaydarali o'g'li", dept: "Infratuzilma" },
    { id: 60, name: "Xoliqnazarov Rustamjon Sotvoldiyevich", dept: "Davlat xususiy" },
    { id: 61, name: "Po'latov Soxibbek Xasanovich", dept: "Davlat xususiy" },
    { id: 62, name: "Toshpo'latov Sarvar Soibjonovich", dept: "Davlat xususiy" },
    { id: 63, name: "Mirzakarimov Dilshodjon Baxodirjon o'g'li", dept: "Sog'lom ovqatlanish" },
    { id: 64, name: "Rustamov Bobir Taxirovich", dept: "Sog'lom ovqatlanish" },
    { id: 65, name: "Turdimatov Ma'ripjon Raxmatovich", dept: "Sog'lom ovqatlanish" }
];

const FIXED_HOLIDAYS = ["01.01", "08.03", "21.03", "09.05", "01.09", "01.10", "08.12"];
const WEEKDAYS = ["Yakshanba", "Dushanba", "Seshanba", "Chorshanba", "Payshanba", "Juma", "Shanba"];
const MOTTO_MORNING = ["Har bir yangi kun — yangi zafarlar uchun imkoniyatdir! ✨", "Ishingizga muhabbat — muvaffaqiyat garovidir. 💫", "Omad doim sizga yor bo'lsin! 🌟"];
const MOTTO_EVENING = ["Bugun juda yaxshi ishladingiz! 🚀🌠", "Dam olishingiz mazmunli bo'lsin! 💆‍♂️💤", "Ertaga yangi kuch bilan kutib qolamiz! ✨"];

let db = { users: {}, logs: [], settings: { lat: 40.3878, lon: 71.7910, radius: 200, holidays: [] } };
if (fs.existsSync(DB_PATH)) {
    const loaded = JSON.parse(fs.readFileSync(DB_PATH));
    db = { ...db, ...loaded, settings: { ...db.settings, ...loaded.settings } };
}
function saveDb() { fs.writeFileSync(DB_PATH, JSON.stringify(db, null, 2)); }

function isHoliday(date) {
    const d = date.format("DD.MM");
    const full = date.format("DD.MM.YYYY");
    // 5-kunlik ish haftasi: Shanba (6) va Yakshanba (0) dam olish kunlari
    return FIXED_HOLIDAYS.includes(d) || db.settings.holidays.includes(full) || date.day() === 0 || date.day() === 6;
}

const registerWizard = new Scenes.WizardScene('REG_SCENE',
    (ctx) => {
        const depts = [...new Set(STAFF_LIST.map(s => s.dept))];
        ctx.reply("🏢 <b>Bo'limni tanlang:</b>", { parse_mode: 'HTML', ...Markup.keyboard([...depts.map(d => [d]), ["❌ Bekor qilish"]]).resize() });
        return ctx.wizard.next();
    },
    (ctx) => {
        if (ctx.message.text === "❌ Bekor qilish") return ctx.scene.leave();
        ctx.wizard.state.dept = ctx.message.text;
        const names = STAFF_LIST.filter(s => s.dept === ctx.wizard.state.dept).map(s => s.name);
        if (names.length === 0) return ctx.reply("Xato bo'lim!");
        ctx.reply("👤 <b>Ism-sharifingizni tanlang:</b>", { parse_mode: 'HTML', ...Markup.keyboard([...names.map(n => [n]), ["⬅️ Orqaga"]]).resize() });
        return ctx.wizard.next();
    },
    (ctx) => {
        if (ctx.message.text === "⬅️ Orqaga") return ctx.wizard.back();
        const staff = STAFF_LIST.find(s => s.name === ctx.message.text);
        if (!staff) return ctx.reply("Xato ism!");
        db.users[ctx.from.id] = { staff_id: staff.id, name: staff.name, dept: staff.dept, vacations: [] };
        saveDb(); showMenu(ctx, false, "🎉 Ro'yxatdan muvaffaqiyatli o'tdingiz!");
        return ctx.scene.leave();
    }
);

const locationScene = new Scenes.WizardScene('LOC_SCENE',
    (ctx) => {
        ctx.wizard.state.type = ctx.session.pendingType;
        const label = ctx.wizard.state.type === 'kirish' ? "📍 KELDIM (Kirish tasdiqlash)" : "🚪 KETDIM (Chiqish tasdiqlash)";
        ctx.replyWithHTML(`📍 <b>Lokatsiyangizni yuboring:</b>\n\nPastdagi "${label}" tugmasini bosing.`, 
            Markup.keyboard([[Markup.button.locationRequest(label)], ["🏠 Bekor qilish"]]).resize());
        return ctx.wizard.next();
    },
    (ctx) => {
        if (ctx.message.text === "🏠 Bekor qilish") { showMenu(ctx); return ctx.scene.leave(); }
        if (!ctx.message.location) return ctx.reply("Iltimos, faqat lokatsiya tugmasini bosing.");
        const type = ctx.wizard.state.type;
        const now = moment().tz("Asia/Tashkent");
        const date = now.format("YYYY-MM-DD");
        const timeNow = now.format("HH:mm");
        if (isHoliday(now)) { showMenu(ctx, false, "😴 Dam olish kunida davomat olinmaydi."); return ctx.scene.leave(); }
        const dist = getDistance(ctx.message.location.latitude, ctx.message.location.longitude, db.settings.lat, db.settings.lon);
        if (dist > db.settings.radius) { showMenu(ctx, false, "⚠️ Masofa xato (Ishxonadan uzoq)"); return ctx.scene.leave(); }

        let statusHeader = "";
        if (type === 'kirish') {
            if (timeNow <= WORK_START) statusHeader = `✅ <b>A’LO! VAQTIDA KELDINGIZ!</b>\n\nDisiplina — muvaffaqiyat garovidir! ⚡️💎`;
            else statusHeader = `⚠️ <b>KECHIKDINGIZ!</b>\n\nErtaga vaqtliroq kelishga harakat qiling! 😊🕒`;
        } else {
            const hasKirish = db.logs.find(l => l.uid == ctx.from.id && l.date === date && l.type === 'kirish');
            if (!hasKirish) { ctx.reply("⚠️ Siz hali bugun 'KELDIM'ni qayd etmagansiz!"); showMenu(ctx); return ctx.scene.leave(); }
            if (timeNow < WORK_END) statusHeader = `🛑 <b>VAQTIDAN OLDIN KETISH?!</b>\n\nIshingizni yakunlab, keyin ketishni tavsiya qilamiz! 👔📉`;
            else statusHeader = `🏠 <b>DAM OLISHINGIZ XAYRLI O'TSIN!</b>\n\nBugungi mehnatingiz uchun tashakkur! 😊🌙`;
        }

        db.logs.push({ uid: ctx.from.id, type, date, time: now.format("HH:mm:ss") }); saveDb();
        showMenu(ctx, type === 'chiqish', statusHeader);
        return ctx.scene.leave();
    }
);

const hududWizard = new Scenes.WizardScene('HUDUD_SCENE',
    (ctx) => {
        ctx.replyWithHTML("📂 <b>Hududdagi vazifa turini tanlang:</b>", Markup.keyboard([["🏫 O'rganish (Topshiriq)", "📄 O'rganish (Taqdimnoma)"],["👥 Fuqaro murojaati", "🏢 Kuratorlik hududida"],["🏠 Bosh menyu"]]).resize());
        return ctx.wizard.next();
    },
    (ctx) => {
        if (ctx.message.text === "🏠 Bosh menyu") { showMenu(ctx); return ctx.scene.leave(); }
        const now = moment().tz("Asia/Tashkent");
        db.logs.push({ uid: ctx.from.id, type: 'special', status: ctx.message.text, date: now.format("YYYY-MM-DD"), time: now.format("HH:mm:ss") });
        saveDb(); showMenu(ctx, false, `✅ <b>${ctx.message.text}</b> qayd etildi.`);
        return ctx.scene.leave();
    }
);

const vacationWizard = new Scenes.WizardScene('VAC_SCENE',
    (ctx) => { ctx.replyWithHTML("🌴 <b>Ta'til turini tanlang:</b>", Markup.keyboard([["☀️ Mehnat ta'tili", "👤 O'z hisobidan ta'til"], ["🏠 Bosh menyu"]]).resize()); return ctx.wizard.next(); },
    (ctx) => {
        if (ctx.message.text === "🏠 Bosh menyu") { showMenu(ctx); return ctx.scene.leave(); }
        ctx.wizard.state.type = ctx.message.text;
        ctx.replyWithHTML("📅 <b>Boshlanish sanasi (KK.OO.YYYY):</b>"); return ctx.wizard.next();
    },
    (ctx) => { ctx.wizard.state.start = ctx.message.text; ctx.replyWithHTML("📅 <b>Tugash sanasi (KK.OO.YYYY):</b>"); return ctx.wizard.next(); },
    (ctx) => {
        const user = db.users[ctx.from.id];
        if (!user.vacations) user.vacations = [];
        user.vacations.push({ start: ctx.wizard.state.start, end: ctx.message.text, type: ctx.wizard.state.type });
        saveDb(); showMenu(ctx, false, `✅ <b>Ta'til qayd etildi!</b>`);
        return ctx.scene.leave();
    }
);

const stage = new Scenes.Stage([registerWizard, vacationWizard, hududWizard, locationScene]);
bot.use(session());
bot.use(stage.middleware());

// Global registration check middleware
bot.use((ctx, next) => {
    if (ctx.message && ctx.message.text === '/start') return next();
    if (ctx.session && ctx.session.__scenes && ctx.session.__scenes.current) return next();
    
    const u = db.users[ctx.from.id];
    if ((!u || !u.staff_id) && ctx.from.id !== bot.botInfo?.id) {
        return ctx.scene.enter('REG_SCENE');
    }
    return next();
});

bot.start((ctx) => { 
    const u = db.users[ctx.from.id]; 
    if (!u || !u.staff_id) return ctx.scene.enter('REG_SCENE'); 
    showMenu(ctx); 
});

async function showMenu(ctx, isExit = false, statusHeader = "") {
    const user = db.users[ctx.from.id]; if (!user) return; 
    const now = moment().tz("Asia/Tashkent");
    const dateStr = now.format("DD.MM.YYYY");
    const weekDay = WEEKDAYS[now.day()];
    const timeStr = now.format("HH:mm");
    const mottoList = isExit ? MOTTO_EVENING : MOTTO_MORNING;
    const randomMotto = mottoList[Math.floor(Math.random() * mottoList.length)];
    let caption = "";
    if (statusHeader) caption += `${statusHeader}\n\n`;
    if (isExit) { caption += `🏠 <b>Oila bag'riga omon boring, ${user.name}!</b>\n\n🗓 Bugun: <b>${dateStr} (${weekDay})</b>\n🕒 Vaqt: <b>${timeStr}</b>\n\n✨ <i>${randomMotto}</i>\n\nXayrli o'tsin! 👋`; }
    else { caption += `🏙 <b>Assalomu alaykum, ${user.name}!</b>\n\n🗓 Bugun: <b>${dateStr} (${weekDay})</b>\n🕒 Vaqt: <b>${timeStr}</b>\n\n✨ <i>${randomMotto}</i>\n\nHozirni qayd eting: 👇`; }
    let keyboard = [
        ["📍 KELDIM (Kirish)", "🚪 KETDIM (Chiqish)"],
        ["✈️ Hizmat safari", "📂 Hududlarda"],
        ["📝 Boshliq topshirig'i", "🤒 Kasal bo'ldim"],
        ["🌴 Ta'til / O'z hisobi", "📊 Statistika"]
    ];

    if (isAdmin(ctx.from.id)) {
        keyboard.push(["⚙️ Admin Panel"]);
    }

    const menuMarkup = Markup.keyboard(keyboard).resize();
    try {
        if (fs.existsSync(LOGO_PATH)) { await ctx.replyWithPhoto({ source: LOGO_PATH }, { caption, parse_mode: 'HTML', ...menuMarkup }); }
        else { ctx.replyWithHTML(caption, menuMarkup); }
    } catch (e) { ctx.replyWithHTML(caption, menuMarkup); }
}

bot.hears("⚙️ Admin Panel", (ctx) => {
    if (!isAdmin(ctx.from.id)) return;
    ctx.replyWithHTML("👑 <b>Admin Panel</b>\n\nQuyidagi amallardan birini tanlang:", 
        Markup.keyboard([["📊 Kunlik hisobot (Bugun)", "📑 Kechagi hisobot"], ["🏠 Bosh menyu"]]).resize());
});

bot.hears("📊 Kunlik hisobot (Bugun)", async (ctx) => {
    if (!isAdmin(ctx.from.id)) return;
    const today = moment().tz("Asia/Tashkent").format("YYYY-MM-DD");
    const reports = await generateReports(today);
    for (const msg of reports) {
        ctx.replyWithHTML(msg).catch(() => {});
    }
});

bot.hears("📑 Kechagi hisobot", async (ctx) => {
    if (!isAdmin(ctx.from.id)) return;
    const yesterday = moment().tz("Asia/Tashkent").subtract(1, 'day').format("YYYY-MM-DD");
    const reports = await generateReports(yesterday);
    for (const msg of reports) {
        ctx.replyWithHTML(msg).catch(() => {});
    }
});

bot.hears("🏠 Bosh menyu", (ctx) => showMenu(ctx));


bot.hears("📍 KELDIM (Kirish)", (ctx) => { ctx.session.pendingType = 'kirish'; ctx.scene.enter('LOC_SCENE'); });
bot.hears("🚪 KETDIM (Chiqish)", (ctx) => { ctx.session.pendingType = 'chiqish'; ctx.scene.enter('LOC_SCENE'); });
bot.hears("📂 Hududlarda", (ctx) => ctx.scene.enter('HUDUD_SCENE'));
bot.hears("🌴 Ta'til / O'z hisobi", (ctx) => ctx.scene.enter('VAC_SCENE'));

bot.hears("✈️ Hizmat safari", (ctx) => {
    const now = moment().tz("Asia/Tashkent");
    db.logs.push({ uid: ctx.from.id, type: 'special', status: 'Hizmat safari', date: now.format("YYYY-MM-DD"), time: now.format("HH:mm:ss") });
    saveDb(); showMenu(ctx, false, `✈️ <b>Hizmat safari qayd etildi!</b>`);
});
bot.hears("📝 Boshliq topshirig'i", (ctx) => {
    const now = moment().tz("Asia/Tashkent");
    db.logs.push({ uid: ctx.from.id, type: 'special', status: 'Topshiriq', date: now.format("YYYY-MM-DD"), time: now.format("HH:mm:ss") });
    saveDb(); showMenu(ctx, false, `📝 <b>Boshliq topshirig'i qayd etildi!</b>`);
});
bot.hears("🤒 Kasal bo'ldim", (ctx) => {
    const now = moment().tz("Asia/Tashkent");
    db.logs.push({ uid: ctx.from.id, type: 'special', status: 'Betob', date: now.format("YYYY-MM-DD"), time: now.format("HH:mm:ss") });
    saveDb(); showMenu(ctx, false, `🤒 <b>Shifo tilaymiz! Betoblik qayd etildi.</b>`);
});

bot.hears("📊 Statistika", (ctx) => {
    const today = moment().tz("Asia/Tashkent").format("YYYY-MM-DD");
    const logs = db.logs.filter(l => l.uid == ctx.from.id && l.date === today);
    let msg = `📊 <b>Bugungi holatingiz:</b>\n\n`;
    logs.forEach(l => msg += `🔹 ${l.type.toUpperCase()}: ${l.time}\n`);
    if (logs.length === 0) msg += "Hali ma'lumot yo'q.";
    ctx.replyWithHTML(msg);
});

bot.on('message', (ctx) => {
    const u = db.users[ctx.from.id];
    if (!u || !u.staff_id) return ctx.scene.enter('REG_SCENE');
    showMenu(ctx);
});

function getDistance(lat1, lon1, lat2, lon2) {
    const R = 6371e3;
    const p1 = lat1 * Math.PI / 180, p2 = lat2 * Math.PI / 180;
    const dp = (lat2-lat1) * Math.PI / 180, dl = (lon2-lon1) * Math.PI / 180;
    const a = Math.sin(dp/2)**2 + Math.cos(p1)*Math.cos(p2)*Math.sin(dl/2)**2;
    return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

async function generateReports(date) {
    const logs = db.logs.filter(l => l.date === date);
    const depts = [...new Set(STAFF_LIST.map(s => s.dept))];
    const reportParts = [];

    // Overall summary statistics
    let totalPresent = 0;
    let totalAbsent = 0;
    let totalVacation = 0;
    let totalSpecial = 0;
    let totalStaff = STAFF_LIST.length;

    for (const dept of depts) {
        let msg = `🏢 <b>${dept} bo'limi:</b>\n\n`;
        const deptStaff = STAFF_LIST.filter(s => s.dept === dept);
        
        deptStaff.forEach(s => {
            const uidString = Object.keys(db.users).find(u => db.users[u].staff_id === s.id);
            const uid = uidString ? Number(uidString) : null;
            const u = db.users[uidString];
            const v = (u && u.vacations) ? u.vacations.find(vac => {
                const startStr = vac.start.split('-')[0].trim();
                const endStr = vac.end.split('-')[0].trim();
                return moment(date).isBetween(moment(startStr, "DD.MM.YYYY"), moment(endStr, "DD.MM.YYYY"), 'day', '[]');
            }) : null;

            if (v) {
                msg += `🌴 <b>${s.name}</b>: ${v.type}\n`;
                totalVacation++;
            } else {
                const l = logs.find(log => log.uid == uid && log.type === 'kirish');
                const o = logs.find(log => log.uid == uid && log.type === 'chiqish');
                const sp = logs.find(log => log.uid == uid && log.type === 'special');
                if (sp) {
                    msg += `📂 <b>${s.name}</b>: ${sp.status}\n`;
                    totalSpecial++;
                } else if (l) {
                    const status = l.time > WORK_START ? "⚠️ Kechikdi" : "✅ Vaqtida";
                    msg += `👤 <b>${s.name}</b>: ${status} (${l.time})\n`;
                    if (o) msg += `   ⤷ 🚪 Ketdi: ${o.time} ${o.time < WORK_END ? '🔴 (Erta)' : ''}\n`;
                    else msg += `   ⤷ 🏢 Hozir ishda\n`;
                    totalPresent++;
                } else {
                    msg += `❌ <b>${s.name}</b>: Kelmadi\n`;
                    totalAbsent++;
                }
            }
        });
        reportParts.push(msg);
    }

    const summaryHeader = `📊 <b>KUNLIK HISOBOT (${date})</b>\n\n` +
        `👥 Jami xodimlar: <b>${totalStaff}</b>\n` +
        `✅ Ishda: <b>${totalPresent}</b>\n` +
        `❌ Kelmagan: <b>${totalAbsent}</b>\n` +
        `🌴 Ta'tilda: <b>${totalVacation}</b>\n` +
        `📂 Maxsus (Safari/Kasal): <b>${totalSpecial}</b>\n\n` +
        `--------------------------\n\n`;

    // Prepend summary to the first part
    if (reportParts.length > 0) reportParts[0] = summaryHeader + reportParts[0];

    return reportParts;
}

cron.schedule('0 10 * * *', async () => {
    const today = moment().tz("Asia/Tashkent").format("YYYY-MM-DD");
    if (isHoliday(moment())) return;
    const reports = await generateReports(today);
    for (const msg of reports) {
        ADMINS.forEach(id => bot.telegram.sendMessage(id, `⏰ <b>10:00 Monitoring:</b>\n\n` + msg, { parse_mode: 'HTML' }).catch(() => {}));
    }
}, { timezone: "Asia/Tashkent" });

cron.schedule('0 20 * * *', async () => {
    const today = moment().tz("Asia/Tashkent").format("YYYY-MM-DD");
    if (isHoliday(moment())) return;
    const reports = await generateReports(today);
    for (const msg of reports) {
        ADMINS.forEach(id => bot.telegram.sendMessage(id, `🌙 <b>20:00 Yakuniy:</b>\n\n` + msg, { parse_mode: 'HTML' }).catch(() => {}));
    }
}, { timezone: "Asia/Tashkent" });

cron.schedule('0 7 * * *', async () => {
    const yesterday = moment().tz("Asia/Tashkent").subtract(1, 'day').format("YYYY-MM-DD");
    const reports = await generateReports(yesterday);
    for (const msg of reports) {
        ADMINS.forEach(id => bot.telegram.sendMessage(id, `📑 <b>07:00 Kechagi hisobot:</b>\n\n` + msg, { parse_mode: 'HTML' }).catch(() => {}));
    }
}, { timezone: "Asia/Tashkent" });

bot.launch({ dropPendingUpdates: true })
    .then(() => console.log("HR BOT IS ONLINE ON RENDER"))
    .catch((err) => console.error("Bot launch error:", err));

// Enable graceful stop
process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => {
    bot.stop('SIGTERM');
    server.close();
});
