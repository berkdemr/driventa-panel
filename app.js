// Driventa İşletme – Panel arayüzü
import { api } from "./api.js?v=0.8";
import { readXlsx, readCsv, writeXlsx, download } from "./xlsx.js?v=0.8";

const $ = s => document.querySelector(s);
const root = $("#root");
const esc = s => String(s ?? "").replace(/[&<>"']/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));

const I = {
  home: '<svg viewBox="0 0 24 24"><path d="M3 10.5 12 3l9 7.5V20a1 1 0 0 1-1 1h-5v-6h-6v6H4a1 1 0 0 1-1-1z"/></svg>',
  car: '<svg viewBox="0 0 24 24"><path d="M5 17h14M3 13l2-5.5A2 2 0 0 1 6.9 6h10.2a2 2 0 0 1 1.9 1.5L21 13v4a1 1 0 0 1-1 1h-1.5a1 1 0 0 1-1-1v-1h-11v1a1 1 0 0 1-1 1H4a1 1 0 0 1-1-1z"/><path d="M3 13h18"/></svg>',
  check: '<svg viewBox="0 0 24 24"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V3h6v1M9 13l2 2 4-4"/></svg>',
  users: '<svg viewBox="0 0 24 24"><circle cx="9" cy="8" r="3.5"/><path d="M2.5 20a6.5 6.5 0 0 1 13 0M16 4.5a3.5 3.5 0 0 1 0 7M21.5 20a6.5 6.5 0 0 0-4-6"/></svg>',
  wrench: '<svg viewBox="0 0 24 24"><path d="M14.5 6.5a4 4 0 0 1 5.4-3.7l-2.6 2.6.8 2 2 .8 2.6-2.6a4 4 0 0 1-5.4 5.4L8.5 19.8a2 2 0 0 1-2.8-2.8l8.8-8.8a4 4 0 0 1 0-1.7z"/></svg>',
  cog: '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.7 1.7 0 0 0 .3 1.8l.1.1a2 2 0 1 1-2.8 2.8l-.1-.1a1.7 1.7 0 0 0-1.8-.3 1.7 1.7 0 0 0-1 1.5V21a2 2 0 1 1-4 0v-.1a1.7 1.7 0 0 0-1.1-1.5 1.7 1.7 0 0 0-1.8.3l-.1.1a2 2 0 1 1-2.8-2.8l.1-.1a1.7 1.7 0 0 0 .3-1.8 1.7 1.7 0 0 0-1.5-1H3a2 2 0 1 1 0-4h.1a1.7 1.7 0 0 0 1.5-1.1 1.7 1.7 0 0 0-.3-1.8l-.1-.1a2 2 0 1 1 2.8-2.8l.1.1a1.7 1.7 0 0 0 1.8.3H9a1.7 1.7 0 0 0 1-1.5V3a2 2 0 1 1 4 0v.1a1.7 1.7 0 0 0 1 1.5 1.7 1.7 0 0 0 1.8-.3l.1-.1a2 2 0 1 1 2.8 2.8l-.1.1a1.7 1.7 0 0 0-.3 1.8V9a1.7 1.7 0 0 0 1.5 1H21a2 2 0 1 1 0 4h-.1a1.7 1.7 0 0 0-1.5 1z"/></svg>',
  out: '<svg viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4M16 17l5-5-5-5M21 12H9"/></svg>',
  menu: '<svg viewBox="0 0 24 24"><path d="M4 7h16M4 12h16M4 17h16"/></svg>',
  ok: '<svg viewBox="0 0 24 24"><path d="m5 12 5 5 9-10"/></svg>',
  copy: '<svg viewBox="0 0 24 24"><rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>',
  plus: '<svg viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>',
  up: '<svg viewBox="0 0 24 24"><path d="M12 16V4M7 9l5-5 5 5M5 20h14"/></svg>',
  down: '<svg viewBox="0 0 24 24"><path d="M12 4v12M7 11l5 5 5-5M5 20h14"/></svg>',
  edit: '<svg viewBox="0 0 24 24"><path d="M4 20h4L19 9a2.8 2.8 0 0 0-4-4L4 16z"/></svg>',
  trash: '<svg viewBox="0 0 24 24"><path d="M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3"/></svg>',
  x: '<svg viewBox="0 0 24 24"><path d="M6 6l12 12M18 6 6 18"/></svg>',
  search: '<svg viewBox="0 0 24 24"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
  receipt: '<svg viewBox="0 0 24 24"><path d="M6 3h12v18l-3-2-3 2-3-2-3 2zM9 8h6M9 12h6M9 16h4"/></svg>',
  chart: '<svg viewBox="0 0 24 24"><path d="M4 20V10M10 20V4M16 20v-7M22 20H2"/></svg>',
  refresh: '<svg viewBox="0 0 24 24"><path d="M21 12a9 9 0 1 1-3-6.7L21 8M21 3v5h-5"/></svg>'
};
const G_LOGO = '<svg viewBox="0 0 48 48" style="stroke:none"><path fill="#FFC107" d="M43.6 20.1H42V20H24v8h11.3C33.7 32.7 29.2 36 24 36c-6.6 0-12-5.4-12-12s5.4-12 12-12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 13 4 4 13 4 24s9 20 20 20 20-9 20-20c0-1.3-.1-2.6-.4-3.9z"/><path fill="#FF3D00" d="m6.3 14.7 6.6 4.8C14.7 15.1 19 12 24 12c3.1 0 5.8 1.2 7.9 3.1l5.7-5.7C34 6.1 29.3 4 24 4 16.3 4 9.7 8.3 6.3 14.7z"/><path fill="#4CAF50" d="M24 44c5.2 0 9.9-2 13.4-5.2l-6.2-5.2C29.2 35.1 26.7 36 24 36c-5.2 0-9.6-3.3-11.3-8l-6.5 5C9.5 39.6 16.2 44 24 44z"/><path fill="#1976D2" d="M43.6 20.1H42V20H24v8h11.3c-.8 2.2-2.2 4.2-4.1 5.6l6.2 5.2C37 39.2 44 34 44 24c0-1.3-.1-2.6-.4-3.9z"/></svg>';

/* ---------- Hata mesajları (Türkçe) ---------- */
function errMsg(e) {
  const c = (e && (e.code || e.message)) || "";
  const map = {
    "auth/invalid-credential": "E-posta veya şifre hatalı.",
    "auth/wrong-password": "E-posta veya şifre hatalı.",
    "auth/user-not-found": "Bu e-postayla kayıtlı bir hesap bulunamadı.",
    "auth/invalid-email": "Geçerli bir e-posta adresi girin.",
    "auth/email-already-in-use": "Bu e-posta adresiyle zaten bir hesap var. Giriş yapmayı deneyin.",
    "auth/weak-password": "Şifre en az 6 karakter olmalı.",
    "auth/too-many-requests": "Çok fazla deneme yapıldı. Lütfen birkaç dakika sonra tekrar deneyin.",
    "auth/network-request-failed": "İnternet bağlantısı kurulamadı.",
    "auth/popup-closed-by-user": "Google penceresi kapatıldı.",
    "auth/popup-blocked": "Tarayıcı açılır pencereyi engelledi. Lütfen izin verip tekrar deneyin.",
    "auth/unauthorized-domain": "Bu adresten giriş yapılmasına izin verilmiyor (Firebase yetkili alan adları).",
    "permission-denied": "Bu işlem için yetkin yok.",
    "unavailable": "Sunucuya ulaşılamıyor. İnternet bağlantını kontrol et."
  };
  for (const k in map) if (c.includes(k)) return map[k];
  console.error(e);
  return "Beklenmeyen bir hata oluştu. Lütfen tekrar deneyin.";
}
let toastT;
function toast(msg, err) {
  const t = $("#toast"); t.textContent = msg; t.classList.toggle("err", !!err); t.classList.add("show");
  clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove("show"), 2600);
}
async function busy(btn, fn) {
  const txt = btn.innerHTML; btn.disabled = true; btn.innerHTML = '<span class="spinner" style="width:18px;height:18px;border-width:2px;margin:0"></span>';
  try { return await fn(); } finally { btn.disabled = false; btn.innerHTML = txt; }
}

/* ---------- Durum ---------- */
const S = { user: null, profile: null, company: null, member: null, page: "overview", members: null, vehicles: null, records: null, pending: 0 };

/* ---------- Giriş ekranı ---------- */
function renderAuth(mode = "login") {
  root.innerHTML = `
  <div class="auth">
    <div class="auth-side">
      <div class="brand"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
      <div>
        <h2>Filonu <span>tek panelden</span> yönet.</h2>
        <p class="muted">Şoförlerin masrafı telefondan girer, sen onaylar ve tüm filonun maliyetini anlık görürsün.</p>
        <ul>
          <li>${I.ok}Şoför kayıtları yönetici onayıyla</li>
          <li>${I.ok}Araç bazlı maliyet ve km başı gider</li>
          <li>${I.ok}Bakım, sigorta ve muayene takvimi</li>
          <li>${I.ok}Veriler Avrupa'daki sunucularda</li>
        </ul>
      </div>
      <p class="muted" style="font-size:13px">© ${new Date().getFullYear()} BeSey · <a class="link" href="https://driventa.com.tr/gizlilik.html" target="_blank" rel="noopener">Gizlilik</a> · <a class="link" href="https://driventa.com.tr/veri-isleme-sozlesmesi.html" target="_blank" rel="noopener">Veri İşleme Sözleşmesi</a></p>
    </div>
    <div class="auth-main"><div class="auth-card">
      <div class="brand" style="margin-bottom:22px"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
      <h1>${mode === "reset" ? "Şifreni sıfırla" : "Hoş geldin"}</h1>
      <p class="muted">${mode === "reset" ? "E-postana bir sıfırlama bağlantısı gönderelim." : "Yönetici paneline giriş yap veya şirketin için hesap oluştur."}</p>
      ${mode === "reset" ? "" : `<div class="tabs"><button data-mode="login" class="${mode === "login" ? "on" : ""}">Giriş yap</button><button data-mode="register" class="${mode === "register" ? "on" : ""}">Hesap oluştur</button></div>`}
      <div id="authMsg"></div>
      <form id="authForm" novalidate style="${mode === "reset" ? "margin-top:22px" : ""}">
        ${mode === "register" ? `<div class="field"><label>Ad soyad</label><input class="input" name="name" autocomplete="name" maxlength="60"></div>` : ""}
        <div class="field"><label>E-posta</label><input class="input" name="email" type="email" autocomplete="email"></div>
        ${mode === "reset" ? "" : `<div class="field"><label>Şifre</label><input class="input" name="pw" type="password" autocomplete="${mode === "register" ? "new-password" : "current-password"}"></div>`}
        <button class="btn btn-primary btn-block" type="submit">${mode === "login" ? "Giriş yap" : mode === "register" ? "Hesap oluştur" : "Bağlantı gönder"}</button>
      </form>
      ${mode === "reset" ? `<p style="margin-top:16px;text-align:center"><a class="link" href="#" data-mode="login">← Girişe dön</a></p>` : `
      <div class="or">veya</div>
      <button class="gbtn" id="gBtn">${G_LOGO} Google ile devam et</button>
      ${mode === "login" ? `<p style="margin-top:16px;text-align:center"><a class="link" href="#" data-mode="reset">Şifremi unuttum</a></p>` : ""}`}
      <p class="muted" style="font-size:12.5px;margin-top:22px;text-align:center">Şoför müsün? Driventa uygulamasını telefonuna kur ve yöneticinin verdiği katılım koduyla şirkete katıl.</p>
    </div></div>
  </div>`;
  root.querySelectorAll("[data-mode]").forEach(b => b.addEventListener("click", e => { e.preventDefault(); renderAuth(b.dataset.mode); }));
  const msg = (t, ok) => { $("#authMsg").innerHTML = `<div class="alert ${ok ? "ok" : "err"}">${esc(t)}</div>`; };
  const f = $("#authForm");
  f.addEventListener("submit", e => {
    e.preventDefault();
    const email = f.email.value.trim(), pw = f.pw ? f.pw.value : "", name = f.name ? f.name.value.trim() : "";
    if (mode === "register" && name.length < 2) return msg("Ad soyad girin.");
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return msg("Geçerli bir e-posta adresi girin.");
    if (mode !== "reset" && pw.length < 6) return msg("Şifre en az 6 karakter olmalı.");
    busy(f.querySelector("button[type=submit]"), async () => {
      try {
        if (mode === "login") await api.signIn(email, pw);
        else if (mode === "register") await api.signUp(name, email, pw);
        else { await api.resetPassword(email); msg("Sıfırlama bağlantısı e-postana gönderildi. Gelen kutunu (ve spam klasörünü) kontrol et.", true); }
      } catch (err) { msg(errMsg(err)); }
    });
  });
  const g = $("#gBtn");
  if (g) g.addEventListener("click", () => busy(g, async () => { try { await api.signInGoogle(); } catch (err) { msg(errMsg(err)); } }));
}

/* ---------- Şirket oluşturma ---------- */
function renderOnboarding() {
  root.innerHTML = `
  <div class="center"><div class="card">
    <div class="brand" style="margin-bottom:20px"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
    <h1 style="font-size:24px">Şirketini oluştur</h1>
    <p class="muted" style="margin:6px 0 20px">Bu bilgilerle filonun yönetim paneli oluşturulacak. Daha sonra Ayarlar'dan değiştirebilirsin.</p>
    <div id="obMsg"></div>
    <form id="obForm" novalidate>
      <div class="field"><label>Şirket adı *</label><input class="input" name="name" maxlength="80" placeholder="Örn. Demir Lojistik"></div>
      <div class="field"><label>Telefon</label><input class="input" name="phone" maxlength="20" inputmode="tel" placeholder="05xx xxx xx xx"></div>
      <div class="field"><label>Araç sayısı</label>
        <select class="input" name="size"><option value="">Seçin</option><option>1–5</option><option>6–20</option><option>21–50</option><option>51–200</option><option>200+</option></select></div>
      <label class="ck-row"><input type="checkbox" name="kabul"><span>Şirketi temsil etmeye yetkili olduğumu beyan eder, <a class="link" href="https://driventa.com.tr/kullanim-kosullari.html" target="_blank" rel="noopener">Kullanım Koşulları</a>'nı ve <a class="link" href="https://driventa.com.tr/veri-isleme-sozlesmesi.html" target="_blank" rel="noopener">Veri İşleme Sözleşmesi</a>'ni kabul ederim.</span></label>
      <button class="btn btn-primary btn-block" type="submit">Şirketi oluştur</button>
    </form>
    <p style="margin-top:16px;text-align:center"><a class="link" href="#" id="obOut">Farklı hesapla giriş yap</a></p>
  </div></div>`;
  $("#obOut").addEventListener("click", e => { e.preventDefault(); api.signOut(); });
  const f = $("#obForm");
  f.addEventListener("submit", e => {
    e.preventDefault();
    const name = f.name.value.trim();
    if (name.length < 2) { $("#obMsg").innerHTML = '<div class="alert err">Şirket adını girin.</div>'; return; }
    if (!f.kabul.checked) { $("#obMsg").innerHTML = '<div class="alert err">Devam etmek için Kullanım Koşulları ve Veri İşleme Sözleşmesi\'ni kabul etmelisin.</div>'; return; }
    busy(f.querySelector("button"), async () => {
      try {
        const cid = await api.createCompany(S.user, { name, phone: f.phone.value.trim(), size: f.size.value });
        await loadCompany(cid);
        toast("Şirketin oluşturuldu. Hoş geldin!");
      } catch (err) { $("#obMsg").innerHTML = `<div class="alert err">${esc(errMsg(err))}</div>`; }
    });
  });
}

/* ---------- Şoför hesabı panele girerse ---------- */
function renderDriverInfo() {
  root.innerHTML = `
  <div class="center"><div class="card" style="text-align:center">
    <div class="empty" style="padding:10px 0 0">
      <div class="ic">${I.car}</div>
      <h3>Bu hesap bir şoför hesabı</h3>
      <p>Şoförler kayıtlarını Driventa mobil uygulamasından girer. Yönetim paneline yalnızca şirket sahibi ve yöneticiler erişebilir.</p>
    </div>
    <button class="btn btn-ghost" style="margin-top:20px" id="dOut">${I.out} Çıkış yap</button>
  </div></div>`;
  $("#dOut").addEventListener("click", () => api.signOut());
}

/* ---------- Panel iskeleti ---------- */
const PAGES = [
  { id: "overview", ad: "Özet", ic: "home" },
  { id: "vehicles", ad: "Araçlar", ic: "car" },
  { id: "approvals", ad: "Kayıt Onayları", ic: "check" },
  { id: "reports", ad: "Raporlar", ic: "chart" },
  { id: "drivers", ad: "Şoförler", ic: "users" },
  { id: "maintenance", ad: "Bakım Takvimi", ic: "wrench" },
  { id: "settings", ad: "Ayarlar", ic: "cog" }
];

function renderShell() {
  const u = S.user, c = S.company, m = S.member;
  const nm = m.name || u.displayName || u.email;
  root.innerHTML = `
  <div class="shell">
    <aside class="side" id="side">
      <div class="brand"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
      ${PAGES.map(p => `<button class="nav-item ${S.page === p.id ? "on" : ""}" data-page="${p.id}">${I[p.ic]}${p.ad}${p.soon ? '<span class="soon">Yakında</span>' : ""}${p.id === "approvals" && S.pending ? `<span class="cnt">${S.pending}</span>` : ""}</button>`).join("")}
      <div class="side-foot">
        <div class="me"><div class="avatar">${esc((nm[0] || "?").toLocaleUpperCase("tr-TR"))}</div>
          <div><b>${esc(nm)}</b><small>${m.role === "owner" ? "Şirket sahibi" : "Yönetici"}</small></div></div>
        <button class="nav-item" id="outBtn">${I.out}Çıkış yap</button>
      </div>
    </aside>
    <main class="main">
      <div class="top">
        <div style="display:flex;gap:12px;align-items:center">
          <button class="menu-btn" id="menuBtn" aria-label="Menü">${I.menu}</button>
          <div><h1 id="pgTitle"></h1><p>${esc(c.name)}</p></div>
        </div>
      </div>
      <div id="page"></div>
    </main>
  </div>`;
  root.querySelectorAll("[data-page]").forEach(b => b.addEventListener("click", () => { S.page = b.dataset.page; $("#side").classList.remove("open"); renderShell(); }));
  $("#outBtn").addEventListener("click", () => api.signOut());
  $("#menuBtn").addEventListener("click", e => { e.stopPropagation(); $("#side").classList.toggle("open"); });
  document.addEventListener("click", e => { const s = $("#side"); if (s && s.classList.contains("open") && !s.contains(e.target)) s.classList.remove("open"); }, { once: true });
  const page = PAGES.find(p => p.id === S.page);
  $("#pgTitle").textContent = page.ad;
  (Pages[S.page] || Pages.soon)(page);
}

const emptyBox = (ic, t, d) => `<div class="card empty"><div class="ic">${I[ic]}</div><h3>${t}</h3><p>${d}</p></div>`;

const Pages = {
  overview() {
    $("#page").innerHTML = `
      <div class="kpis">
        <div class="kpi"><small>Araç</small><b id="kVehicles">–</b></div>
        <div class="kpi"><small>Şoför</small><b id="kDrivers">–</b></div>
        <a class="kpi" href="#" data-goto="reports"><small>Bu ay onaylanan masraf</small><b id="kMonth">–</b></a>
        <a class="kpi" href="#" data-goto="approvals"><small>Onay bekleyen</small><b id="kPending">–</b></a>
      </div>
      <div id="mtAlert"></div>
      <div class="card">
        <h3 style="font-size:17px;margin-bottom:4px">Başlarken</h3>
        <p class="muted">Filonu Driventa'ya taşımak için şu adımları izle:</p>
        <ol class="steps">
          <li>Araçlarını ekle: <a class="link" href="#" data-goto="vehicles">Araçlar</a> sayfasından tek tek veya Excel'den toplu olarak.</li>
          <li>Şoförlerine katılım kodunu ver: <b style="color:var(--accent);letter-spacing:2px">${esc(S.company.joinCode)}</b></li>
          <li>Şoförler Driventa uygulamasından bu kodla şirkete katılsın.</li>
          <li>Her şoföre aracını ata; şoförlerin girdiği kayıtları onayla.</li>
        </ol>
      </div>`;
    $("#page").querySelectorAll("[data-goto]").forEach(a => a.addEventListener("click", e => { e.preventDefault(); S.page = a.dataset.goto; renderShell(); }));
    api.listVehicles(S.company.id).then(vs => {
      S.vehicles = vs; const el = $("#kVehicles"); if (el) el.textContent = vs.length;
      const it = trackItems(vs); const late = it.filter(x => x.st === "late"), soon = it.filter(x => x.st === "soon"); const box = $("#mtAlert");
      if (box && (late.length || soon.length)) {
        box.innerHTML = `<a href="#" class="card mt-alert ${late.length ? "late" : ""}" data-goto="maintenance"><span class="ic">${I.wrench}</span>
          <span><b>Bakım ve belge takibi</b><small>${[late.length ? `${late.length} işin süresi geçti` : "", soon.length ? `${soon.length} iş 30 gün içinde` : ""].filter(Boolean).join(" · ")}${(late[0] || soon[0]) ? ` · İlk: ${esc((late[0] || soon[0]).v.plaka)} ${(late[0] || soon[0]).ad}` : ""}</small></span><span class="go">Takvimi aç →</span></a>`;
        box.querySelector("[data-goto]").addEventListener("click", e => { e.preventDefault(); S.page = "maintenance"; renderShell(); });
      }
    }).catch(() => {});
    api.listMembers(S.company.id).then(ms => { S.members = ms; const el = $("#kDrivers"); if (el) el.textContent = ms.filter(x => x.role === "driver").length; })
      .catch(() => { const el = $("#kDrivers"); if (el) el.textContent = "–"; });
    loadRecords().then(rs => {
      const d = new Date(); const pre = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
      const m = rs.filter(r => r.status === "approved" && String(r.tarih).startsWith(pre)).reduce((s, r) => s + (+r.tutar || 0), 0);
      const a = $("#kMonth"), b = $("#kPending"); if (a) a.textContent = TL(m); if (b) b.textContent = S.pending;
    }).catch(() => {});
  },

  drivers() {
    const c = S.company;
    $("#page").innerHTML = `
      <div class="grid2">
        <div class="card">
          <h3 style="font-size:17px">Katılım kodu</h3>
          <p class="muted" style="margin:4px 0 16px">Şoförler Driventa uygulamasında <b>Profil → İşletme hesabı</b> bölümüne bu kodu girerek şirketine katılır.</p>
          <div class="code" id="jc">${esc(c.joinCode)}</div>
          <div style="display:flex;gap:10px;margin-top:16px;flex-wrap:wrap">
            <button class="btn btn-ghost" id="cpBtn">${I.copy} Kopyala</button>
            <button class="btn btn-ghost" id="rgBtn">${I.refresh} Yeni kod oluştur</button>
          </div>
          <p class="muted" style="font-size:12.5px;margin-top:12px">Kod başkasının eline geçtiyse yeni kod oluştur. Eski kod hemen geçersiz olur; katılmış şoförler etkilenmez.</p>
        </div>
        <div class="card">
          <h3 style="font-size:17px">Nasıl çalışır?</h3>
          <ol class="steps">
            <li>Şoför Driventa'yı telefonuna kurar ve e-postasıyla hesap açar.</li>
            <li>Katılım kodunu girer; listede "araç atanmadı" olarak görünür.</li>
            <li>Sen <b>Araçlar</b> sayfasından şoföre aracını atarsın.</li>
            <li>Şoförün girdiği kayıtlar senin onayına düşer.</li>
          </ol>
        </div>
      </div>
      <div class="card" style="margin-top:16px">
        <h3 style="font-size:17px;margin-bottom:10px">Ekip</h3>
        <div id="mList"><div class="spinner"></div></div>
      </div>`;
    $("#cpBtn").addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(S.company.joinCode); toast("Kod kopyalandı."); } catch (e) { toast("Kopyalanamadı, kodu elle seçebilirsin.", true); }
    });
    $("#rgBtn").addEventListener("click", e => {
      if (!confirmInline(e.currentTarget, "Emin misin? Tekrar bas")) return;
      busy(e.currentTarget, async () => {
        try { const code = await api.regenerateJoinCode(c.id, c.joinCode); S.company.joinCode = code; $("#jc").textContent = code; toast("Yeni katılım kodu oluşturuldu."); }
        catch (err) { toast(errMsg(err), true); }
      });
    });
    Promise.all([api.listMembers(c.id), api.listVehicles(c.id).catch(() => [])]).then(([ms, vs]) => {
      S.members = ms; S.vehicles = vs;
      const roleAd = { owner: "Sahip", admin: "Yönetici", driver: "Şoför" };
      const vPl = id => { const v = (S.vehicles || []).find(x => x.id === id); return v ? esc(v.plaka) : ""; };
      $("#mList").innerHTML = ms.length ? `<div style="overflow-x:auto"><table class="tbl"><thead><tr><th>Ad</th><th class="hm">E-posta</th><th>Rol</th><th>Araç</th><th>Durum</th><th></th></tr></thead><tbody>
        ${ms.sort((a, b) => (a.role === "owner" ? -1 : b.role === "owner" ? 1 : 0)).map(m => `<tr class="${m.active === false ? "dim" : ""}"><td>${esc(m.name || "–")}<div class="muted sm-only" style="font-size:12px">${esc(m.email || "")}</div></td><td class="hm">${esc(m.email || "–")}</td>
          <td><span class="role ${m.role}">${roleAd[m.role] || m.role}</span></td>
          <td class="muted">${m.role === "driver" ? (m.vehicleIds && m.vehicleIds.length ? m.vehicleIds.map(vPl).filter(Boolean).join(", ") || m.vehicleIds.length + " araç" : "Atanmadı") : "–"}</td>
          <td>${m.active === false ? '<span class="st-b late">Pasif</span>' : '<span class="st-b ok">Aktif</span>'}</td>
          <td class="acts">${m.role === "driver" ? `<button class="btn btn-sm btn-ghost" data-act="${m.active === false ? "on" : "off"}" data-uid="${m.uid}">${m.active === false ? "Aktif yap" : "Pasif yap"}</button><button class="btn btn-sm btn-danger" data-rm="${m.uid}">${I.trash} Çıkar</button>` : ""}</td></tr>`).join("")}
        </tbody></table></div>
        <p class="muted" style="font-size:12.5px;margin-top:12px"><b>Pasif yap:</b> şoför giriş yapabilir ama kayıt giremez ve şirket verilerini göremez. <b>Çıkar:</b> şoför şirketten tamamen çıkarılır, araç atamaları kaldırılır; girdiği kayıtlar şirket kayıtlarında kalır.</p>` : emptyBox("users", "Henüz ekip üyesi yok", "Katılım kodunu şoförlerinle paylaş.");
      $("#mList").querySelectorAll("[data-act]").forEach(b => b.addEventListener("click", () => {
        const m = S.members.find(x => x.uid === b.dataset.uid); const on = b.dataset.act === "on";
        busy(b, async () => {
          try { await api.setMemberActive(c.id, m.uid, on); m.active = on; toast(on ? `${m.name || m.email} tekrar aktif.` : `${m.name || m.email} pasif yapıldı.`); Pages.drivers(); }
          catch (err) { toast(errMsg(err), true); }
        });
      }));
      $("#mList").querySelectorAll("[data-rm]").forEach(b => b.addEventListener("click", () => removeMemberModal(S.members.find(x => x.uid === b.dataset.rm))));
    }).catch(err => { $("#mList").innerHTML = `<div class="alert err">${esc(errMsg(err))}</div>`; });
  },

  settings() {
    const c = S.company;
    $("#page").innerHTML = `
      <div class="grid2">
        <div class="card">
          <h3 style="font-size:17px;margin-bottom:14px">Şirket bilgileri</h3>
          <form id="stForm" novalidate>
            <div class="field"><label>Şirket adı *</label><input class="input" name="name" maxlength="80" value="${esc(c.name)}"></div>
            <div class="field"><label>Telefon</label><input class="input" name="phone" maxlength="20" value="${esc(c.phone || "")}"></div>
            <div class="field"><label>Araç sayısı</label><select class="input" name="size">
              ${["", "1–5", "6–20", "21–50", "51–200", "200+"].map(o => `<option value="${o}" ${o === (c.size || "") ? "selected" : ""}>${o || "Seçin"}</option>`).join("")}</select></div>
            <button class="btn btn-primary" type="submit">Kaydet</button>
          </form>
        </div>
        <div class="card">
          <h3 style="font-size:17px;margin-bottom:10px">Hesap</h3>
          <p class="muted">Giriş yapılan hesap</p><p style="margin-bottom:14px">${esc(S.user.email)}</p>
          <p class="muted">Paket</p><p style="margin-bottom:14px">Deneme (erken erişim)</p>
          <p class="muted" style="font-size:13px">Verilerin Avrupa'daki (eur3) sunucularda saklanır. <a class="link" href="https://driventa.com.tr/gizlilik.html" target="_blank" rel="noopener">Gizlilik politikası</a></p>
        </div>
      </div>`;
    const f = $("#stForm");
    f.addEventListener("submit", e => {
      e.preventDefault();
      const name = f.name.value.trim();
      if (name.length < 2) return toast("Şirket adını girin.", true);
      busy(f.querySelector("button"), async () => {
        try {
          const data = { name, phone: f.phone.value.trim(), size: f.size.value };
          await api.updateCompany(c.id, data); Object.assign(S.company, data); toast("Kaydedildi."); renderShell();
        } catch (err) { toast(errMsg(err), true); }
      });
    });
  },

  soon(page) {
    const d = {
      vehicles: "Araç ekleme, düzenleme ve Excel'den toplu aktarım bir sonraki aşamada (F2) geliyor.",
      approvals: "Şoförlerin girdiği kayıtları onaylama ve reddetme ekranı F3 aşamasında geliyor.",
      maintenance: "Tüm filonun bakım, sigorta ve muayene takvimi F5 aşamasında geliyor."
    };
    $("#page").innerHTML = emptyBox(page.ic, `${page.ad} yakında`, d[page.id] || "Bu bölüm hazırlanıyor.");
  }
};


/* =====================================================================
   ARAÇLAR (F2)
   ===================================================================== */
const VEHICLE_TYPES = [{ id: "otomobil", ad: "Otomobil" }, { id: "motosiklet", ad: "Motosiklet" }, { id: "ticari", ad: "Ticari" }];
const FUELS = ["Benzin", "Dizel", "LPG", "Benzin + LPG", "Hibrit", "Elektrik"];
const NUM = n => new Intl.NumberFormat("tr-TR").format(n);
const trUp = s => String(s).toLocaleUpperCase("tr-TR");
const trLow = s => String(s ?? "").toLocaleLowerCase("tr-TR");
const typeAd = t => (VEHICLE_TYPES.find(x => x.id === t) || VEHICLE_TYPES[0]).ad;

/* ---------- Bakım / belge takibi alanları (F5) ---------- */
const TRACK_DATES = [["muayeneBitis", "Muayene bitiş"], ["sigortaBitis", "Trafik sigortası bitiş"], ["kaskoBitis", "Kasko bitiş"], ["sonBakimTarih", "Son periyodik bakım"]];
const BAKIM_KM_VARSAYILAN = 15000, BAKIM_AY_VARSAYILAN = 12;
/* Excel hücresindeki tarihi YYYY-AA-GG'ye çevirir. Boşsa "", okunamazsa null */
function parseDateCell(x) {
  if (x === "" || x == null) return "";
  if (typeof x === "number") { if (x > 20000 && x < 80000) { const d = new Date(Date.UTC(1899, 11, 30) + Math.round(x) * 864e5); return d.toISOString().slice(0, 10); } return null; }
  const t = String(x).trim(); if (!t) return "";
  let m = t.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/); if (m) return okDate(+m[1], +m[2], +m[3]);
  m = t.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/); if (m) return okDate(+m[3], +m[2], +m[1]);
  if (/^\d+(\.\d+)?$/.test(t)) return parseDateCell(+t);
  return null;
}
function okDate(y, mo, d) { const dt = new Date(Date.UTC(y, mo - 1, d)); return (y >= 1990 && y <= 2100 && dt.getUTCMonth() === mo - 1 && dt.getUTCDate() === d) ? dt.toISOString().slice(0, 10) : null; }
function validateTrack(f) {
  const e = {};
  TRACK_DATES.forEach(([k]) => { if (f[k] === null) e[k] = "Tarih okunamadı. Örnek: 15.03.2027"; });
  const num = x => String(x ?? "").replace(/[.\s]/g, "");
  const sk = num(f.sonBakimKm), ak = num(f.bakimAralikKm), ay = num(f.bakimAralikAy), km = +num(f.km);
  if (sk && !/^\d+$/.test(sk)) e.sonBakimKm = "Sadece rakam girin."; else if (sk && km && +sk > km) e.sonBakimKm = "Son bakım km'si aracın güncel km'sinden büyük olamaz.";
  if (ak && !(/^\d+$/.test(ak) && +ak >= 1000 && +ak <= 100000)) e.bakimAralikKm = "1.000 ile 100.000 km arasında olmalı.";
  if (ay && !(/^\d+$/.test(ay) && +ay >= 1 && +ay <= 60)) e.bakimAralikAy = "1 ile 60 ay arasında olmalı.";
  if (f.sonBakimTarih && f.sonBakimTarih > isoD(new Date())) e.sonBakimTarih = "Son bakım tarihi ileri bir tarih olamaz.";
  return e;
}
const toTrack = f => {
  const n = x => { const t = String(x ?? "").replace(/[.\s]/g, ""); return t ? +t : null; };
  return { muayeneBitis: f.muayeneBitis || null, sigortaBitis: f.sigortaBitis || null, kaskoBitis: f.kaskoBitis || null,
    sonBakimTarih: f.sonBakimTarih || null, sonBakimKm: n(f.sonBakimKm), bakimAralikKm: n(f.bakimAralikKm), bakimAralikAy: n(f.bakimAralikAy) };
};

function normalizePlate(raw) {
  const s = trUp(raw).replace(/[\s\-.]/g, "");
  const m = s.match(/^(\d{2})([A-Z]{1,3})(\d{2,5})$/);
  return m ? `${m[1]} ${m[2]} ${m[3]}` : null;
}
/* Uygulamadaki kurallarla aynı kontrol */
function validateVehicle(f, editId) {
  const e = {}; const yearNow = new Date().getFullYear();
  if (!f.marka) e.marka = "Marka boş bırakılamaz."; else if (f.marka.length > 30) e.marka = "Marka en fazla 30 karakter olabilir.";
  if (!f.model) e.model = "Model boş bırakılamaz."; else if (f.model.length > 40) e.model = "Model en fazla 40 karakter olabilir.";
  const yil = parseInt(f.yil, 10);
  if (!f.yil && f.yil !== 0) e.yil = "Model yılı girin."; else if (!(yil >= 1950 && yil <= yearNow + 1)) e.yil = `1950 ile ${yearNow + 1} arasında bir yıl girin.`;
  if (!f.plaka) e.plaka = "Plaka boş bırakılamaz.";
  else {
    const raw = trUp(f.plaka).replace(/[\s\-.]/g, "");
    if (/[ÇĞİÖŞÜ]/.test(raw)) e.plaka = "Plakada Türkçe karakter kullanılmaz.";
    else {
      const p = normalizePlate(f.plaka), il = p ? +p.slice(0, 2) : 0;
      if (!p) e.plaka = "Plaka biçimi hatalı. Örnek: 34 ABC 123";
      else if (il < 1 || il > 81) e.plaka = "İl kodu 01 ile 81 arasında olmalı.";
      else if ((S.vehicles || []).some(v => v.plaka === p && v.id !== editId)) e.plaka = "Bu plaka zaten kayıtlı.";
    }
  }
  const km = String(f.km ?? "").replace(/[.\s]/g, "");
  if (km === "") e.km = "Kilometre girin."; else if (!/^\d+$/.test(km)) e.km = "Kilometre yalnızca rakam olmalı."; else if (+km > 3000000) e.km = "Kilometre çok yüksek görünüyor.";
  if (!FUELS.includes(f.yakit)) e.yakit = "Yakıt türü seçin.";
  if (!VEHICLE_TYPES.some(t => t.id === f.tur)) e.tur = "Araç türü seçin.";
  const motor = String(f.motor ?? "").trim();
  if (motor && f.yakit !== "Elektrik" && !(/^\d+$/.test(motor) && +motor >= 49 && +motor <= 9000)) e.motor = "Motor hacmi 49–9000 cc olmalı.";
  return e;
}
const toVehicle = f => ({
  ...toTrack(f),
  tur: f.tur, marka: String(f.marka).trim(), model: String(f.model).trim(), yil: parseInt(f.yil, 10),
  plaka: normalizePlate(f.plaka), km: +String(f.km).replace(/[.\s]/g, ""), yakit: f.yakit,
  motor: f.yakit === "Elektrik" || !String(f.motor ?? "").trim() ? null : +f.motor
});
const drivers = () => (S.members || []).filter(m => m.role === "driver" && m.active);
const driverName = uid => { const m = (S.members || []).find(x => x.uid === uid); return m ? (m.name || m.email) : ""; };

function removeMemberModal(m) {
  if (!m) return;
  const vs = (S.vehicles || []).filter(v => v.assignedUid === m.uid);
  const w = openModal(`
    <div class="m-head"><h3>Şoför şirketten çıkarılsın mı?</h3><button class="icon" data-close>${I.x}</button></div>
    <p class="muted"><b style="color:var(--text)">${esc(m.name || m.email)}</b> şirketten çıkarılacak.${vs.length ? ` Atanmış araçları (${vs.map(v => esc(v.plaka)).join(", ")}) boşa alınacak.` : ""} Girdiği kayıtlar şirket kayıtlarında kalır. Şoför tekrar katılmak isterse katılım kodunu girmesi gerekir; kodu başkalarıyla da paylaştıysa yenilemeni öneririz.</p>
    <div class="m-foot"><button class="btn btn-ghost" data-close>Vazgeç</button><button class="btn btn-danger" id="rmOk">${I.trash} Çıkar</button></div>`);
  w.querySelector("#rmOk").addEventListener("click", e => busy(e.currentTarget, async () => {
    try {
      await api.removeMember(S.company.id, m.uid, vs.map(v => v.id));
      vs.forEach(v => v.assignedUid = null); S.members = S.members.filter(x => x.uid !== m.uid);
      closeModal(); toast("Şoför şirketten çıkarıldı."); Pages.drivers();
    } catch (err) { toast(errMsg(err), true); }
  }));
}

/* ---------- Pencere (modal) ---------- */
function openModal(html, wide) {
  closeModal();
  const w = document.createElement("div"); w.className = "modal-wrap"; w.id = "modal";
  w.innerHTML = `<div class="modal ${wide ? "wide" : ""}" role="dialog">${html}</div>`;
  document.body.appendChild(w);
  w.querySelectorAll("[data-close]").forEach(b => b.addEventListener("click", closeModal));
  return w;
}
function closeModal() { const m = $("#modal"); if (m) m.remove(); }
document.addEventListener("keydown", e => { if (e.key === "Escape") closeModal(); });

Pages.vehicles = function () {
  $("#page").innerHTML = `
    <div class="toolbar">
      <div class="search">${I.search}<input class="input" id="vq" placeholder="Plaka, marka veya şoför ara"></div>
      <div class="tb-actions">
        <button class="btn btn-ghost" id="vExp">${I.down} Excel'e indir</button>
        <button class="btn btn-ghost" id="vImp">${I.up} Excel'den aktar</button>
        <button class="btn btn-primary" id="vAdd">${I.plus} Araç ekle</button>
      </div>
    </div>
    <div class="card" style="padding:0;overflow:hidden"><div id="vList"><div style="padding:30px"><div class="spinner"></div></div></div></div>`;
  $("#vAdd").addEventListener("click", () => vehicleModal());
  $("#vImp").addEventListener("click", importModal);
  $("#vExp").addEventListener("click", exportVehicles);
  $("#vq").addEventListener("input", renderVehicleList);
  Promise.all([api.listVehicles(S.company.id), api.listMembers(S.company.id)])
    .then(([vs, ms]) => { S.vehicles = vs; S.members = ms; renderVehicleList(); })
    .catch(err => { $("#vList").innerHTML = `<div class="alert err" style="margin:16px">${esc(errMsg(err))}</div>`; });
};

function renderVehicleList() {
  const box = $("#vList"); if (!box) return;
  const q = trLow($("#vq") ? $("#vq").value.trim() : "");
  const all = (S.vehicles || []).slice().sort((a, b) => a.plaka.localeCompare(b.plaka, "tr"));
  const list = q ? all.filter(v => trLow(`${v.plaka} ${v.plaka.replace(/\s/g, "")} ${v.marka} ${v.model} ${driverName(v.assignedUid)}`).includes(q)) : all;
  if (!all.length) {
    box.innerHTML = `<div class="empty"><div class="ic">${I.car}</div><h3>Henüz araç yok</h3>
      <p>Araçlarını tek tek ekleyebilir veya Excel'den toplu olarak aktarabilirsin.</p>
      <div style="display:flex;gap:10px;justify-content:center;margin-top:18px;flex-wrap:wrap">
        <button class="btn btn-ghost" data-act="imp">${I.up} Excel'den aktar</button><button class="btn btn-primary" data-act="add">${I.plus} Araç ekle</button></div></div>`;
    box.querySelector('[data-act="imp"]').addEventListener("click", importModal);
    box.querySelector('[data-act="add"]').addEventListener("click", () => vehicleModal());
    return;
  }
  const ds = drivers();
  box.innerHTML = `<div style="overflow-x:auto"><table class="tbl vt">
    <thead><tr><th>Plaka</th><th>Araç</th><th class="hm">Tür</th><th class="hm">Yakıt</th><th class="num">Km</th><th>Şoför</th><th></th></tr></thead>
    <tbody>${list.map(v => `<tr>
      <td><span class="plate"><i>TR</i><b>${esc(v.plaka)}</b></span></td>
      <td><b style="font-weight:600">${esc(v.marka)} ${esc(v.model)}</b><div class="muted" style="font-size:12.5px">${esc(v.yil)}</div></td>
      <td class="hm">${typeAd(v.tur)}</td><td class="hm">${esc(v.yakit)}</td>
      <td class="num">${NUM(v.km)}</td>
      <td><select class="input sel" data-assign="${v.id}">
        <option value="">— Atanmadı —</option>
        ${ds.map(d => `<option value="${d.uid}" ${v.assignedUid === d.uid ? "selected" : ""}>${esc(d.name || d.email)}</option>`).join("")}
        ${v.assignedUid && !ds.some(d => d.uid === v.assignedUid) ? `<option value="${v.assignedUid}" selected>(Ayrılmış şoför)</option>` : ""}
      </select></td>
      <td class="acts"><button class="icon" title="Düzenle" data-edit="${v.id}">${I.edit}</button><button class="icon danger" title="Sil" data-del="${v.id}">${I.trash}</button></td>
    </tr>`).join("")}</tbody></table></div>
    <div class="tfoot">${list.length === all.length ? `${all.length} araç` : `${list.length} / ${all.length} araç`}${ds.length ? "" : ` · <span class="muted">Şoför atamak için önce şoförlerin katılım koduyla katılmalı.</span>`}</div>`;
  box.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => vehicleModal(S.vehicles.find(v => v.id === b.dataset.edit))));
  box.querySelectorAll("[data-del]").forEach(b => b.addEventListener("click", () => deleteVehicleModal(S.vehicles.find(v => v.id === b.dataset.del))));
  box.querySelectorAll("[data-assign]").forEach(sel => sel.addEventListener("change", async () => {
    const v = S.vehicles.find(x => x.id === sel.dataset.assign); const newUid = sel.value || null, oldUid = v.assignedUid || null;
    sel.disabled = true;
    try {
      await api.assignVehicle(S.company.id, v.id, newUid, oldUid);
      v.assignedUid = newUid;
      (S.members || []).forEach(m => {
        m.vehicleIds = (m.vehicleIds || []).filter(id => id !== v.id);
        if (m.uid === newUid) m.vehicleIds.push(v.id);
      });
      toast(newUid ? `${v.plaka} → ${driverName(newUid)} atandı.` : `${v.plaka} için atama kaldırıldı.`);
    } catch (err) { sel.value = oldUid || ""; toast(errMsg(err), true); }
    finally { sel.disabled = false; }
  }));
}

function vehicleModal(v, onDone) {
  const d = v || { tur: "otomobil", marka: "", model: "", yil: "", plaka: "", km: "", yakit: "", motor: "" };
  const fx = (k, lbl, input) => `<div class="field" data-f="${k}"><label>${lbl}</label>${input}<div class="msg"></div></div>`;
  const dt = k => `<input class="input" type="date" name="${k}" value="${esc(d[k] || "")}">`;
  const nm = (k, ph) => `<input class="input" name="${k}" inputmode="numeric" placeholder="${ph || ""}" value="${d[k] != null && d[k] !== "" ? NUM(d[k]) : ""}">`;
  const w = openModal(`
    <div class="m-head"><h3>${v ? "Aracı düzenle" : "Yeni araç"}</h3><button class="icon" data-close>${I.x}</button></div>
    <form id="vf" novalidate>
      <div class="g2">
        <div class="field" data-f="tur"><label>Araç türü *</label><select class="input" name="tur">${VEHICLE_TYPES.map(t => `<option value="${t.id}" ${d.tur === t.id ? "selected" : ""}>${t.ad}</option>`).join("")}</select><div class="msg"></div></div>
        <div class="field" data-f="plaka"><label>Plaka *</label><input class="input" name="plaka" maxlength="12" placeholder="34 ABC 123" style="text-transform:uppercase" value="${esc(d.plaka)}"><div class="msg"></div></div>
        <div class="field" data-f="marka"><label>Marka *</label><input class="input" name="marka" maxlength="30" value="${esc(d.marka)}"><div class="msg"></div></div>
        <div class="field" data-f="model"><label>Model *</label><input class="input" name="model" maxlength="40" value="${esc(d.model)}"><div class="msg"></div></div>
        <div class="field" data-f="yil"><label>Model yılı *</label><input class="input" name="yil" inputmode="numeric" maxlength="4" value="${esc(d.yil)}"><div class="msg"></div></div>
        <div class="field" data-f="km"><label>Güncel kilometre *</label><input class="input" name="km" inputmode="numeric" value="${d.km !== "" ? NUM(d.km) : ""}"><div class="msg"></div></div>
        <div class="field" data-f="yakit"><label>Yakıt türü *</label><select class="input" name="yakit"><option value="">Seçin</option>${FUELS.map(f => `<option ${d.yakit === f ? "selected" : ""}>${f}</option>`).join("")}</select><div class="msg"></div></div>
        <div class="field" data-f="motor"><label>Motor hacmi (cc)</label><input class="input" name="motor" inputmode="numeric" maxlength="4" value="${d.motor ? esc(d.motor) : ""}"><div class="msg"></div></div>
      </div>
      <h4 class="m-sec">Bakım ve belge takibi <span class="muted">· isteğe bağlı</span></h4>
      <p class="muted" style="font-size:12.5px;margin:-4px 0 12px">Bakım Takvimi bu bilgilerle hazırlanır. Şoförün girdiği sigorta, muayene ve bakım kayıtları onaylanınca bu alanlar kendiliğinden güncellenir.</p>
      <div class="g2">
        ${fx("muayeneBitis", "Muayene bitiş tarihi", dt("muayeneBitis"))}
        ${fx("sigortaBitis", "Trafik sigortası bitiş", dt("sigortaBitis"))}
        ${fx("kaskoBitis", "Kasko bitiş", dt("kaskoBitis"))}
        ${fx("sonBakimTarih", "Son periyodik bakım tarihi", dt("sonBakimTarih"))}
        ${fx("sonBakimKm", "Son bakımdaki km", nm("sonBakimKm", "Örn. 75.000"))}
        <div class="g2 in">
          ${fx("bakimAralikKm", "Bakım aralığı (km)", nm("bakimAralikKm", NUM(BAKIM_KM_VARSAYILAN)))}
          ${fx("bakimAralikAy", "Aralık (ay)", nm("bakimAralikAy", String(BAKIM_AY_VARSAYILAN)))}
        </div>
      </div>
      <div class="m-foot"><button type="button" class="btn btn-ghost" data-close>Vazgeç</button><button type="submit" class="btn btn-primary">${v ? "Kaydet" : "Aracı ekle"}</button></div>
    </form>`);
  const f = w.querySelector("#vf");
  ["km", "sonBakimKm", "bakimAralikKm"].forEach(k => f[k].addEventListener("input", () => { const x = f[k].value.replace(/\D/g, "").slice(0, 7); f[k].value = x ? NUM(+x) : ""; }));
  f.bakimAralikAy.addEventListener("input", () => { f.bakimAralikAy.value = f.bakimAralikAy.value.replace(/\D/g, "").slice(0, 2); });
  f.plaka.addEventListener("blur", () => { const p = normalizePlate(f.plaka.value); if (p) f.plaka.value = p; });
  f.addEventListener("input", e => { const el = f.querySelector(`[data-f="${e.target.name}"]`); if (el) el.classList.remove("err"); });
  setTimeout(() => f.plaka.focus(), 50);
  f.addEventListener("submit", e => {
    e.preventDefault();
    const val = { tur: f.tur.value, marka: f.marka.value.trim(), model: f.model.value.trim(), yil: f.yil.value.trim(), plaka: f.plaka.value.trim(), km: f.km.value.trim(), yakit: f.yakit.value, motor: f.motor.value.trim(),
      muayeneBitis: f.muayeneBitis.value, sigortaBitis: f.sigortaBitis.value, kaskoBitis: f.kaskoBitis.value, sonBakimTarih: f.sonBakimTarih.value,
      sonBakimKm: f.sonBakimKm.value.trim(), bakimAralikKm: f.bakimAralikKm.value.trim(), bakimAralikAy: f.bakimAralikAy.value.trim() };
    const errs = { ...validateVehicle(val, v ? v.id : null), ...validateTrack(val) };
    f.querySelectorAll(".field").forEach(el => { const k = el.dataset.f; el.classList.toggle("err", !!errs[k]); el.querySelector(".msg").textContent = errs[k] || ""; });
    if (Object.keys(errs).length) return;
    busy(f.querySelector('button[type=submit]'), async () => {
      try {
        const data = toVehicle(val);
        if (v) { await api.updateVehicle(S.company.id, v.id, data); Object.assign(v, data); toast("Araç güncellendi."); }
        else { const id = await api.addVehicle(S.company.id, data); S.vehicles.push({ id, ...data, assignedUid: null }); toast("Araç eklendi."); }
        closeModal(); (onDone || renderVehicleList)();
      } catch (err) { toast(errMsg(err), true); }
    });
  });
}

function deleteVehicleModal(v) {
  const w = openModal(`
    <div class="m-head"><h3>Araç silinsin mi?</h3><button class="icon" data-close>${I.x}</button></div>
    <p class="muted"><b style="color:var(--text)">${esc(v.plaka)} – ${esc(v.marka)} ${esc(v.model)}</b> silinecek.${v.assignedUid ? ` Şoför ataması (${esc(driverName(v.assignedUid))}) da kaldırılacak.` : ""} Bu işlem geri alınamaz.</p>
    <div class="m-foot"><button class="btn btn-ghost" data-close>Vazgeç</button><button class="btn btn-danger" id="delOk">${I.trash} Sil</button></div>`);
  w.querySelector("#delOk").addEventListener("click", e => busy(e.currentTarget, async () => {
    try {
      await api.deleteVehicle(S.company.id, v.id, v.assignedUid || null);
      S.vehicles = S.vehicles.filter(x => x.id !== v.id);
      (S.members || []).forEach(m => { m.vehicleIds = (m.vehicleIds || []).filter(id => id !== v.id); });
      closeModal(); renderVehicleList(); toast("Araç silindi.");
    } catch (err) { toast(errMsg(err), true); }
  }));
}

/* ---------- Excel ---------- */
const XL_HEAD = ["Plaka", "Marka", "Model", "Model Yılı", "Kilometre", "Yakıt Türü", "Araç Türü", "Motor Hacmi (cc)",
  "Muayene Bitiş", "Trafik Sigortası Bitiş", "Kasko Bitiş", "Son Bakım Tarihi", "Son Bakım Km", "Bakım Aralığı (km)", "Bakım Aralığı (ay)"];
const XL_W = [14, 16, 18, 12, 13, 15, 13, 17, 15, 21, 13, 17, 14, 18, 18];
const trDate = iso => iso ? String(iso).split("-").reverse().join(".") : "";
function templateDownload() {
  const blob = writeXlsx([
    { name: "Araçlar", widths: XL_W, rows: [XL_HEAD, ["34 ABC 123", "Renault", "Clio", 2019, 85000, "Benzin", "Otomobil", 999, "15.03.2027", "02.08.2027", "02.08.2027", "10.05.2026", 75000, 15000, 12], ["06 BSY 41", "Ford", "Transit Custom", 2021, 64200, "Dizel", "Ticari", "", "", "", "", "", "", "", ""]] },
    { name: "Açıklama", widths: [20, 70], rows: [["Sütun", "Nasıl doldurulur?"],
      ["Plaka *", "Örn. 34 ABC 123 (boşluklar önemli değil)"], ["Marka *", "Örn. Renault"], ["Model *", "Örn. Clio"],
      ["Model Yılı *", "Örn. 2019"], ["Kilometre *", "Sadece rakam. Örn. 85000"],
      ["Yakıt Türü *", FUELS.join(", ")], ["Araç Türü", "Otomobil, Motosiklet veya Ticari (boşsa Otomobil)"],
      ["Motor Hacmi (cc)", "İsteğe bağlı. Örn. 1461"],
      ["Muayene Bitiş", "İsteğe bağlı. Gün.Ay.Yıl biçiminde. Örn. 15.03.2027"], ["Trafik Sigortası Bitiş", "İsteğe bağlı. Örn. 02.08.2027"], ["Kasko Bitiş", "İsteğe bağlı. Örn. 02.08.2027"],
      ["Son Bakım Tarihi", "İsteğe bağlı. Son periyodik bakımın tarihi. Örn. 10.05.2026"], ["Son Bakım Km", "İsteğe bağlı. Son bakımdaki kilometre. Örn. 75000"],
      ["Bakım Aralığı (km)", "İsteğe bağlı. Boşsa 15000"], ["Bakım Aralığı (ay)", "İsteğe bağlı. Boşsa 12"],
      ["Not", "Örnek satırları silip kendi araçlarını yaz. * işaretli sütunlar zorunlu."]] }
  ]);
  download(blob, "driventa-arac-sablonu.xlsx");
}
function mapFuel(x) {
  const k = trLow(x).replace(/[\s+\-_/]/g, "");
  if (!k) return "";
  if (k === "benzinlpg" || k === "lpgbenzin") return "Benzin + LPG";
  if (k.startsWith("benzin")) return "Benzin";
  if (["dizel", "diesel", "motorin", "mazot"].includes(k)) return "Dizel";
  if (k === "lpg") return "LPG";
  if (k.startsWith("hibrit") || k.startsWith("hybrid")) return "Hibrit";
  if (k.startsWith("elektrik") || k === "ev") return "Elektrik";
  return String(x).trim();
}
function mapType(x) {
  const k = trLow(x).trim();
  if (!k) return "otomobil";
  if (["otomobil", "binek", "araba", "oto"].includes(k)) return "otomobil";
  if (["motosiklet", "motor", "moto", "motorsiklet"].includes(k)) return "motosiklet";
  if (["ticari", "kamyonet", "panelvan", "minibüs", "minibus", "kamyon", "hafif ticari", "pikap"].includes(k)) return "ticari";
  return "?";
}
const HEAD_KEYS = { plaka: "plaka", marka: "marka", model: "model", modelyılı: "yil", modelyili: "yil", yıl: "yil", yil: "yil",
  kilometre: "km", km: "km", güncelkilometre: "km", yakıttürü: "yakit", yakitturu: "yakit", yakıt: "yakit", yakit: "yakit",
  araçtürü: "tur", aracturu: "tur", tür: "tur", tur: "tur", tip: "tur", motorhacmi: "motor", motorhacmicc: "motor", motor: "motor", cc: "motor",
  muayenebitiş: "muayeneBitis", muayenebitis: "muayeneBitis", muayene: "muayeneBitis", muayenetarihi: "muayeneBitis",
  trafiksigortasıbitiş: "sigortaBitis", trafiksigortasibitis: "sigortaBitis", trafiksigortası: "sigortaBitis", sigortabitiş: "sigortaBitis", sigorta: "sigortaBitis",
  kaskobitiş: "kaskoBitis", kaskobitis: "kaskoBitis", kasko: "kaskoBitis",
  sonbakımtarihi: "sonBakimTarih", sonbakimtarihi: "sonBakimTarih", sonbakım: "sonBakimTarih",
  sonbakımkm: "sonBakimKm", sonbakimkm: "sonBakimKm", sonbakımkilometresi: "sonBakimKm",
  bakımaralığıkm: "bakimAralikKm", bakimaraligikm: "bakimAralikKm", bakımaralığıay: "bakimAralikAy", bakimaraligiay: "bakimAralikAy" };

function parseRows(rows) {
  const headIdx = rows.findIndex(r => r && r.some(c => HEAD_KEYS[trLow(c).replace(/[^a-zçğıöşü]/g, "")] === "plaka"));
  if (headIdx < 0) return { error: "Dosyada \"Plaka\" sütunu bulunamadı. Lütfen şablonu kullan." };
  const map = rows[headIdx].map(c => HEAD_KEYS[trLow(c).replace(/[^a-zçğıöşü]/g, "")] || null);
  const need = ["plaka", "marka", "model", "yil", "km", "yakit"];
  const missing = need.filter(k => !map.includes(k));
  if (missing.length) return { error: `Eksik sütun: ${missing.map(k => ({ plaka: "Plaka", marka: "Marka", model: "Model", yil: "Model Yılı", km: "Kilometre", yakit: "Yakıt Türü" })[k]).join(", ")}` };
  const out = []; const seen = new Set();
  rows.slice(headIdx + 1).forEach((r, i) => {
    if (!r || r.every(c => String(c ?? "").trim() === "")) return;
    const o = {}; map.forEach((k, ci) => { if (k) o[k] = r[ci] ?? ""; });
    const f = { plaka: String(o.plaka ?? "").trim(), marka: String(o.marka ?? "").trim(), model: String(o.model ?? "").trim(),
      yil: String(o.yil ?? "").trim(), km: String(o.km ?? "").trim(), yakit: mapFuel(o.yakit), tur: mapType(o.tur ?? ""), motor: String(o.motor ?? "").trim(),
      muayeneBitis: parseDateCell(o.muayeneBitis), sigortaBitis: parseDateCell(o.sigortaBitis), kaskoBitis: parseDateCell(o.kaskoBitis), sonBakimTarih: parseDateCell(o.sonBakimTarih),
      sonBakimKm: String(o.sonBakimKm ?? "").trim(), bakimAralikKm: String(o.bakimAralikKm ?? "").trim(), bakimAralikAy: String(o.bakimAralikAy ?? "").trim() };
    const errs = { ...validateVehicle(f, null), ...validateTrack(f) };
    if (f.tur === "?") errs.tur = "Araç türü Otomobil, Motosiklet veya Ticari olmalı.";
    if (errs.yakit && f.yakit) errs.yakit = `"${f.yakit}" tanınmadı. ${FUELS.join(", ")} olmalı.`;
    const p = normalizePlate(f.plaka);
    if (p && seen.has(p) && !errs.plaka) errs.plaka = "Bu plaka dosyada birden fazla kez geçiyor.";
    if (p) seen.add(p);
    out.push({ line: headIdx + i + 2, f, errs, ok: !Object.keys(errs).length });
  });
  return { rows: out };
}

function importModal() {
  const w = openModal(`
    <div class="m-head"><h3>Excel'den araç aktar</h3><button class="icon" data-close>${I.x}</button></div>
    <ol class="steps" style="margin:0 0 16px">
      <li><span><a class="link" href="#" id="tpl">Şablonu indir</a> ve araçlarını doldur.</span></li>
      <li><span>Doldurduğun dosyayı (.xlsx veya .csv) aşağıdan seç.</span></li>
      <li><span>Önizlemeyi kontrol et ve aktar.</span></li>
    </ol>
    <label class="drop" id="drop"><input type="file" id="xf" accept=".xlsx,.csv" hidden>${I.up}<b>Dosya seç</b><span class="muted">veya buraya sürükle bırak</span></label>
    <div id="impRes"></div>`, true);
  w.querySelector("#tpl").addEventListener("click", e => { e.preventDefault(); templateDownload(); });
  const input = w.querySelector("#xf"), drop = w.querySelector("#drop");
  input.addEventListener("change", () => input.files[0] && handleFile(input.files[0]));
  drop.addEventListener("dragover", e => { e.preventDefault(); drop.classList.add("over"); });
  drop.addEventListener("dragleave", () => drop.classList.remove("over"));
  drop.addEventListener("drop", e => { e.preventDefault(); drop.classList.remove("over"); const f = e.dataTransfer.files[0]; if (f) handleFile(f); });
}

async function handleFile(file) {
  const res = $("#impRes");
  if (file.size > 5 * 1024 * 1024) { res.innerHTML = '<div class="alert err">Dosya çok büyük (en fazla 5 MB).</div>'; return; }
  res.innerHTML = '<div style="padding:20px"><div class="spinner"></div></div>';
  let rows;
  try {
    if (/\.csv$/i.test(file.name)) rows = readCsv(await file.text());
    else if (/\.xlsx$/i.test(file.name)) rows = await readXlsx(await file.arrayBuffer());
    else { res.innerHTML = '<div class="alert err">Lütfen .xlsx veya .csv dosyası seç. (Eski .xls dosyasını Excel\'de "Farklı kaydet → .xlsx" ile dönüştürebilirsin.)</div>'; return; }
  } catch (e) { console.error(e); res.innerHTML = '<div class="alert err">Dosya okunamadı. Dosyanın bozuk olmadığından ve .xlsx biçiminde kaydedildiğinden emin ol.</div>'; return; }
  const parsed = parseRows(rows);
  if (parsed.error) { res.innerHTML = `<div class="alert err">${esc(parsed.error)}</div>`; return; }
  const ok = parsed.rows.filter(r => r.ok), bad = parsed.rows.filter(r => !r.ok);
  if (!parsed.rows.length) { res.innerHTML = '<div class="alert err">Dosyada araç satırı bulunamadı.</div>'; return; }
  const LBL = { plaka: "Plaka", marka: "Marka", model: "Model", yil: "Yıl", km: "Km", yakit: "Yakıt", tur: "Tür", motor: "Motor",
    muayeneBitis: "Muayene", sigortaBitis: "Trafik sigortası", kaskoBitis: "Kasko", sonBakimTarih: "Son bakım tarihi", sonBakimKm: "Son bakım km", bakimAralikKm: "Bakım aralığı (km)", bakimAralikAy: "Bakım aralığı (ay)" };
  res.innerHTML = `
    <div class="sum"><span class="pill ok">${ok.length} araç aktarılacak</span>${bad.length ? `<span class="pill bad">${bad.length} satırda hata</span>` : ""}</div>
    <div class="prev"><table class="tbl"><thead><tr><th>Satır</th><th>Plaka</th><th>Araç</th><th>Km</th><th>Yakıt</th><th>Durum</th></tr></thead><tbody>
      ${parsed.rows.map(r => `<tr class="${r.ok ? "" : "bad"}"><td class="muted">${r.line}</td><td style="white-space:nowrap">${esc(normalizePlate(r.f.plaka) || r.f.plaka || "–")}</td>
        <td>${esc(r.f.marka)} ${esc(r.f.model)} <span class="muted">${esc(r.f.yil)}</span></td><td>${esc(r.f.km)}</td><td>${esc(r.f.yakit)}</td>
        <td>${r.ok ? `<span class="okc">${I.ok}</span>` : `<span class="errc">${Object.entries(r.errs).map(([k, m]) => `<b>${LBL[k] || k}:</b> ${esc(m)}`).join("<br>")}</span>`}</td></tr>`).join("")}
    </tbody></table></div>
    ${bad.length ? '<p class="muted" style="font-size:13px;margin-top:10px">Hatalı satırlar atlanır. İstersen dosyayı düzeltip tekrar yükleyebilirsin.</p>' : ""}
    <div class="m-foot"><button class="btn btn-ghost" data-close2>Vazgeç</button><button class="btn btn-primary" id="doImp" ${ok.length ? "" : "disabled"}>${I.up} ${ok.length} aracı aktar</button></div>`;
  res.querySelector("[data-close2]").addEventListener("click", closeModal);
  const btn = res.querySelector("#doImp");
  btn.addEventListener("click", () => busy(btn, async () => {
    try {
      const list = ok.map(r => toVehicle(r.f));
      await api.bulkAddVehicles(S.company.id, list);
      S.vehicles = await api.listVehicles(S.company.id);
      closeModal(); renderVehicleList(); toast(`${list.length} araç aktarıldı.`);
    } catch (err) { toast(errMsg(err), true); }
  }));
}

function exportVehicles() {
  const vs = (S.vehicles || []).slice().sort((a, b) => a.plaka.localeCompare(b.plaka, "tr"));
  if (!vs.length) return toast("İndirilecek araç yok.", true);
  const rows = [[...XL_HEAD, "Atanan Şoför"], ...vs.map(v => [v.plaka, v.marka, v.model, v.yil, v.km, v.yakit, typeAd(v.tur), v.motor ?? "",
    trDate(v.muayeneBitis), trDate(v.sigortaBitis), trDate(v.kaskoBitis), trDate(v.sonBakimTarih), v.sonBakimKm ?? "", v.bakimAralikKm ?? "", v.bakimAralikAy ?? "", driverName(v.assignedUid) || ""])];
  const d = new Date(); const stamp = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  download(writeXlsx([{ name: "Araçlar", widths: [...XL_W, 22], rows }]), `driventa-araclar-${stamp}.xlsx`);
  toast("Excel dosyası indirildi.");
}

/* İki aşamalı onay: ilk basışta uyarır, ikinci basışta işlemi yapar */
function confirmInline(btn, label) {
  if (btn.dataset.armed) { delete btn.dataset.armed; btn.innerHTML = btn.dataset.orig; return true; }
  btn.dataset.armed = "1"; btn.dataset.orig = btn.innerHTML; btn.innerHTML = label;
  setTimeout(() => { if (btn.dataset.armed) { delete btn.dataset.armed; btn.innerHTML = btn.dataset.orig; } }, 3500);
  return false;
}

/* =====================================================================
   KAYIT ONAYLARI (F3)
   Şoförlerin mobil uygulamadan girdiği kayıtlar burada onaylanır/reddedilir.
   ===================================================================== */
const REC_TYPES = { yakit: "Yakıt", bakim: "Bakım", sigorta: "Sigorta", mtv: "MTV", muayene: "Muayene", lastik: "Lastik", otopark: "Otopark", hgs: "HGS/OGS", diger: "Diğer" };
const TL = n => { n = Math.round((+n || 0) * 100) / 100; const i = Number.isInteger(n); return new Intl.NumberFormat("tr-TR", { style: "currency", currency: "TRY", minimumFractionDigits: i ? 0 : 2, maximumFractionDigits: 2 }).format(n); };
const fmtD = iso => { if (!iso) return "–"; const [y, m, d] = String(iso).split("-").map(Number); return new Date(y, m - 1, d).toLocaleDateString("tr-TR", { day: "numeric", month: "short", year: "numeric" }); };
const fmtTs = ms => ms ? new Date(ms).toLocaleDateString("tr-TR", { day: "numeric", month: "short", year: "numeric" }) : "–";
const recDriver = r => r.createdByName || driverName(r.createdBy) || "–";
const recVehicle = r => (S.vehicles || []).find(v => v.id === r.aracId);
const REJECT_REASONS = ["Fiş / fatura eksik", "Tutar hatalı", "Kilometre hatalı", "Yanlış araç", "Mükerrer kayıt"];
let apTab = "pending";
const apSel = new Set();

async function loadRecords() {
  S.records = await api.listRecords(S.company.id);
  S.pending = S.records.filter(r => r.status === "pending").length;
  updateNavBadge();
  return S.records;
}
function updateNavBadge() {
  const b = document.querySelector('.side [data-page="approvals"]'); if (!b) return;
  let s = b.querySelector(".cnt");
  if (!S.pending) { if (s) s.remove(); return; }
  if (!s) { s = document.createElement("span"); s.className = "cnt"; b.appendChild(s); }
  s.textContent = S.pending;
}

Pages.approvals = function () {
  apSel.clear();
  $("#page").innerHTML = `
    <div class="tabs ap-tabs" id="apTabs"></div>
    <div class="toolbar">
      <div class="search">${I.search}<input class="input" id="aq" placeholder="Plaka, şoför veya açıklama ara"></div>
      <div class="tb-actions" id="apActs"></div>
    </div>
    <div class="card" style="padding:0;overflow:hidden"><div id="apList"><div style="padding:30px"><div class="spinner"></div></div></div></div>`;
  $("#aq").addEventListener("input", renderApprovals);
  const cid = S.company.id;
  Promise.all([loadRecords(), api.listVehicles(cid), api.listMembers(cid).catch(() => S.members || [])])
    .then(([, vs, ms]) => { S.vehicles = vs; S.members = ms; renderApprovals(); })
    .catch(err => { $("#apList").innerHTML = `<div class="alert err" style="margin:16px">${esc(errMsg(err))}</div>`; });
};

function recDetail(r) {
  const e = r.ekstra || {}; const parts = [];
  if (r.tur === "yakit" && e.litre) { const ev = recVehicle(r)?.yakit === "Elektrik"; parts.push(`${NUM(e.litre)} ${ev ? "kWh" : "L"}${e.birimFiyat ? " × " + TL(e.birimFiyat) : ""}${e.tamDepo === false ? " · kısmi" : ""}`); }
  if (e.altTur) parts.push(e.altTur);
  if (e.bitis) parts.push("Bitiş: " + fmtD(e.bitis));
  return parts.map(esc).join(" · ");
}

function renderApprovals() {
  const box = $("#apList"); if (!box) return;
  const all = S.records || [];
  const cnt = k => all.filter(r => r.status === k).length;
  $("#apTabs").innerHTML = [["pending", "Bekleyen"], ["approved", "Onaylanan"], ["rejected", "Reddedilen"]]
    .map(([k, l]) => `<button class="${apTab === k ? "on" : ""}" data-tab="${k}">${l} (${cnt(k)})</button>`).join("");
  $("#apTabs").querySelectorAll("[data-tab]").forEach(b => b.addEventListener("click", () => { apTab = b.dataset.tab; apSel.clear(); renderApprovals(); }));

  const q = trLow($("#aq").value.trim());
  let list = all.filter(r => r.status === apTab);
  if (q) list = list.filter(r => trLow(`${r.plaka} ${String(r.plaka || "").replace(/\s/g, "")} ${recDriver(r)} ${r.aciklama} ${REC_TYPES[r.tur]} ${(r.ekstra || {}).altTur || ""}`).includes(q));
  list.sort((a, b) => String(b.tarih).localeCompare(String(a.tarih)) || (b.createdAt || 0) - (a.createdAt || 0));
  [...apSel].forEach(id => { if (!list.some(r => r.id === id)) apSel.delete(id); });

  const P = apTab === "pending";
  $("#apActs").innerHTML = P && list.length ? `<button class="btn btn-primary" id="apBulk" ${apSel.size ? "" : "disabled"}>${I.ok} Seçilenleri onayla${apSel.size ? ` (${apSel.size})` : ""}</button>` : "";
  if ($("#apBulk")) $("#apBulk").addEventListener("click", e => approve([...apSel], e.currentTarget));

  if (!list.length) {
    const msg = q ? ["Sonuç yok", "Aramanla eşleşen kayıt bulunamadı."]
      : P ? ["Bekleyen kayıt yok", "Şoförlerin mobil uygulamadan girdiği kayıtlar burada onayına düşer."]
      : apTab === "approved" ? ["Henüz onaylanan kayıt yok", "Onayladığın kayıtlar burada listelenir."] : ["Reddedilen kayıt yok", "Reddettiğin kayıtlar ve nedenleri burada listelenir."];
    box.innerHTML = `<div class="empty"><div class="ic">${I.check}</div><h3>${msg[0]}</h3><p>${msg[1]}</p></div>`;
    return;
  }
  const total = list.reduce((s, r) => s + (+r.tutar || 0), 0);
  const allOn = P && list.every(r => apSel.has(r.id));
  box.innerHTML = `<div style="overflow-x:auto"><table class="tbl vt ap">
    <thead><tr>${P ? `<th class="ck"><input type="checkbox" id="apAll" ${allOn ? "checked" : ""} aria-label="Tümünü seç"></th>` : ""}
      <th>Tarih</th><th>Araç</th><th>Şoför</th><th>Tür</th><th class="num">Km</th><th class="num">Tutar</th>
      <th class="hm">${apTab === "rejected" ? "Red nedeni" : "Açıklama"}</th>${P ? "<th></th>" : `<th class="hm">${apTab === "approved" ? "Onay" : "Red"} tarihi</th>`}</tr></thead>
    <tbody>${list.map(r => {
      const v = recVehicle(r); const det = recDetail(r);
      const kmUp = P && r.km != null && v && r.km > (v.km || 0);
      const dUp = P && v && (r.ekstra || {}).bitis && ["muayene", "sigorta"].includes(r.tur);
      return `<tr class="${apSel.has(r.id) ? "sel" : ""}">
      ${P ? `<td class="ck"><input type="checkbox" data-ck="${r.id}" ${apSel.has(r.id) ? "checked" : ""} aria-label="Seç"></td>` : ""}
      <td style="white-space:nowrap">${fmtD(r.tarih)}</td>
      <td><span class="plate"><i>TR</i><b>${esc(r.plaka || v?.plaka || "?")}</b></span>${v ? "" : '<div class="errc">Araç silinmiş</div>'}</td>
      <td style="white-space:nowrap">${esc(recDriver(r))}</td>
      <td><b style="font-weight:600">${REC_TYPES[r.tur] || esc(r.tur)}</b>${det ? `<div class="muted" style="font-size:12.5px">${det}</div>` : ""}${dUp ? '<div class="kmup">↻ takvim güncellenir</div>' : ""}${(r.ekstra || {}).fis ? `<button class="fis-link" data-fis="${r.id}">${I.receipt} Fişi gör</button>` : ""}</td>
      <td class="num">${r.km != null ? NUM(r.km) : "–"}${kmUp ? `<div class="kmup" title="Onaylanınca aracın kilometresi ${NUM(r.km)} olur">↑ araç km</div>` : ""}</td>
      <td class="num"><b>${TL(r.tutar)}</b></td>
      <td class="hm muted" style="max-width:240px">${esc(apTab === "rejected" ? (r.redNedeni || "–") : (r.aciklama || "–"))}</td>
      ${P ? `<td class="acts"><button class="btn btn-sm btn-ok" data-ok="${r.id}">${I.ok} Onayla</button><button class="btn btn-sm btn-danger" data-rej="${r.id}">${I.x} Reddet</button></td>`
          : `<td class="hm muted" style="white-space:nowrap">${fmtTs(r.reviewedAt)}</td>`}
    </tr>`; }).join("")}</tbody></table></div>
    <div class="tfoot">${list.length} kayıt · Toplam <b style="color:var(--text)">${TL(total)}</b>${P ? ` · <span class="muted">Onaylanan kayıtlar şoför tarafından değiştirilemez.</span>` : ""}</div>`;

  box.querySelectorAll("[data-fis]").forEach(b => b.addEventListener("click", () => receiptModal(all.find(r => r.id === b.dataset.fis))));
  if (P) {
    $("#apAll").addEventListener("change", e => { list.forEach(r => e.target.checked ? apSel.add(r.id) : apSel.delete(r.id)); renderApprovals(); });
    box.querySelectorAll("[data-ck]").forEach(c => c.addEventListener("change", () => { c.checked ? apSel.add(c.dataset.ck) : apSel.delete(c.dataset.ck); renderApprovals(); }));
    box.querySelectorAll("[data-ok]").forEach(b => b.addEventListener("click", () => approve([b.dataset.ok], b)));
    box.querySelectorAll("[data-rej]").forEach(b => b.addEventListener("click", () => rejectModal(all.find(r => r.id === b.dataset.rej))));
  }
}

/* Onaylanan kayıtlardan araç bilgisi güncellemeleri: km, sigorta/kasko/muayene bitişi, son periyodik bakım */
const isPeriodic = r => { const e = r.ekstra || {}; const ops = Array.isArray(e.islemler) ? e.islemler : String(e.altTur || "").split(/,\s*/); return ops.some(x => /periyodik bakım|yağ değişimi/i.test(x)); };
function vehicleUpdatesFor(recs) {
  const up = {};
  const cur = (v, k) => (up[v.id] && up[v.id][k] !== undefined) ? up[v.id][k] : v[k];
  const set = (v, k, val) => { (up[v.id] = up[v.id] || {})[k] = val; };
  recs.slice().sort((a, b) => String(a.tarih).localeCompare(String(b.tarih))).forEach(r => {
    const v = recVehicle(r); if (!v) return; const e = r.ekstra || {};
    if (r.km != null && r.km > (cur(v, "km") || 0)) set(v, "km", r.km);
    if (e.bitis) {
      const k = r.tur === "muayene" ? "muayeneBitis" : r.tur === "sigorta" ? (e.altTur === "Kasko" ? "kaskoBitis" : e.altTur === "Trafik sigortası" ? "sigortaBitis" : null) : null;
      if (k && e.bitis > (cur(v, k) || "")) set(v, k, e.bitis);
    }
    if (r.tur === "bakim" && isPeriodic(r) && r.tarih >= (cur(v, "sonBakimTarih") || "")) {
      set(v, "sonBakimTarih", r.tarih); if (r.km != null) set(v, "sonBakimKm", r.km);
    }
  });
  return up;
}

function approve(ids, btn) {
  if (!ids.length) return;
  const recs = ids.map(id => S.records.find(r => r.id === id)).filter(Boolean);
  const up = vehicleUpdatesFor(recs);
  busy(btn, async () => {
    try {
      await api.approveRecords(S.company.id, ids, S.user.uid, up);
      const now = Date.now();
      recs.forEach(r => { r.status = "approved"; r.reviewedAt = now; r.redNedeni = null; apSel.delete(r.id); });
      Object.entries(up).forEach(([vid, d]) => { const v = S.vehicles.find(x => x.id === vid); if (v) Object.assign(v, d); });
      S.pending = S.records.filter(r => r.status === "pending").length; updateNavBadge();
      const n = Object.keys(up).length;
      toast(`${ids.length} kayıt onaylandı.${n ? ` ${n} aracın bilgileri güncellendi.` : ""}`);
      renderApprovals();
    } catch (err) {
      toast(errMsg(err), true);
      try { await loadRecords(); renderApprovals(); } catch (e) { /* liste yenilenemedi */ }
    }
  });
}

function receiptModal(r) {
  if (!r) return;
  const w = openModal(`
    <div class="m-head"><h3>Fiş · ${esc(r.plaka || "")}</h3><button class="icon" data-close aria-label="Kapat">${I.x}</button></div>
    <p class="muted" style="margin-bottom:12px">${REC_TYPES[r.tur] || ""} · ${TL(r.tutar)} · ${fmtD(r.tarih)} · ${esc(recDriver(r))}${(r.ekstra || {}).litre ? ` · ${NUM((r.ekstra || {}).litre)} L` : ""}</p>
    <div id="fisBox" class="fis-box"><div class="spinner"></div></div>
    <div class="m-foot"><button class="btn btn-ghost" data-close>Kapat</button></div>`);
  api.getReceipt(S.company.id, r.id).then(img => {
    const box = w.querySelector("#fisBox"); if (!box) return;
    box.innerHTML = img ? `<a href="data:image/jpeg;base64,${img}" download="fis-${esc(String(r.plaka || "").replace(/\s/g, ""))}-${r.tarih}.jpg" title="İndir"><img src="data:image/jpeg;base64,${img}" alt="Fiş fotoğrafı"></a>`
      : `<p class="muted">Fiş fotoğrafı bulunamadı. Şoför fotoğrafı henüz göndermemiş olabilir (internet bağlantısı bekleniyor).</p>`;
  }).catch(err => { const box = w.querySelector("#fisBox"); if (box) box.innerHTML = `<div class="alert err">${esc(errMsg(err))}</div>`; });
}

function rejectModal(r) {
  if (!r) return;
  const w = openModal(`
    <div class="m-head"><h3>Kaydı reddet</h3><button class="icon" data-close aria-label="Kapat">${I.x}</button></div>
    <p class="muted" style="margin-bottom:14px">${esc(r.plaka || "")} · ${REC_TYPES[r.tur] || ""} · ${TL(r.tutar)} · ${fmtD(r.tarih)} · ${esc(recDriver(r))}</p>
    <div class="field"><label>Red nedeni</label>
      <div class="chips" style="display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px">${REJECT_REASONS.map(x => `<button type="button" class="chip" data-rs="${esc(x)}">${esc(x)}</button>`).join("")}</div>
      <textarea class="input" id="rjText" maxlength="200" rows="3" placeholder="Şoför bu açıklamayı uygulamada görecek ve kaydı düzeltip yeniden gönderebilecek."></textarea>
    </div>
    <div class="m-foot"><button class="btn btn-ghost" data-close>Vazgeç</button><button class="btn btn-danger" id="rjOk">${I.x} Reddet</button></div>`);
  const ta = w.querySelector("#rjText");
  w.querySelectorAll("[data-rs]").forEach(b => b.addEventListener("click", () => { ta.value = b.dataset.rs; w.querySelectorAll("[data-rs]").forEach(x => x.classList.toggle("on", x === b)); ta.focus(); }));
  w.querySelector("#rjOk").addEventListener("click", e => {
    const reason = ta.value.trim();
    if (!reason) { ta.classList.add("err-b"); return toast("Şoförün düzeltebilmesi için bir neden yaz veya seç.", true); }
    busy(e.currentTarget, async () => {
      try {
        await api.rejectRecord(S.company.id, r.id, reason, S.user.uid);
        r.status = "rejected"; r.redNedeni = reason; r.reviewedAt = Date.now(); apSel.delete(r.id);
        S.pending = S.records.filter(x => x.status === "pending").length; updateNavBadge();
        closeModal(); toast("Kayıt reddedildi. Şoför nedeni uygulamada görecek."); renderApprovals();
      } catch (err) { toast(errMsg(err), true); }
    });
  });
}

/* =====================================================================
   RAPORLAR (F4) – sadece ONAYLANAN kayıtlar hesaba katılır
   ===================================================================== */
const pad2 = n => String(n).padStart(2, "0");
const isoD = d => `${d.getFullYear()}-${pad2(d.getMonth() + 1)}-${pad2(d.getDate())}`;
const parseIso = s => { const [y, m, d] = String(s).split("-").map(Number); return new Date(y, m - 1, d); };
const TLc = n => new Intl.NumberFormat("tr-TR", { notation: "compact", maximumFractionDigits: 1 }).format(n || 0);
const AYLAR = ["Oca", "Şub", "Mar", "Nis", "May", "Haz", "Tem", "Ağu", "Eyl", "Eki", "Kas", "Ara"];
const RP_PRESETS = [["month", "Bu ay"], ["last", "Geçen ay"], ["3m", "Son 3 ay"], ["year", "Bu yıl"], ["custom", "Özel"]];
const RP = { preset: "month", from: "", to: "", vid: "" };
const typeGroup = t => t === "yakit" ? "yakit" : (t === "bakim" || t === "lastik") ? "bakim" : "diger";

function rpRange() {
  const t = new Date(); const y = t.getFullYear(), m = t.getMonth();
  switch (RP.preset) {
    case "last": return [isoD(new Date(y, m - 1, 1)), isoD(new Date(y, m, 0))];
    case "3m": return [isoD(new Date(y, m - 2, 1)), isoD(t)];
    case "year": return [isoD(new Date(y, 0, 1)), isoD(t)];
    case "custom": return [RP.from || isoD(new Date(y, m, 1)), RP.to || isoD(t)];
    default: return [isoD(new Date(y, m, 1)), isoD(t)];
  }
}

Pages.reports = function () {
  $("#page").innerHTML = `<div class="card"><div class="spinner"></div></div>`;
  const cid = S.company.id;
  Promise.all([loadRecords(), api.listVehicles(cid), api.listMembers(cid).catch(() => S.members || [])])
    .then(([, vs, ms]) => { S.vehicles = vs; S.members = ms; renderReports(); })
    .catch(err => { $("#page").innerHTML = `<div class="alert err">${esc(errMsg(err))}</div>`; });
};

/* Seçili aralık ve araç için hesaplar */
function rpCompute() {
  const [from, to] = rpRange();
  const recs = (S.records || []).filter(r => r.status === "approved" && r.tarih >= from && r.tarih <= to && (!RP.vid || r.aracId === RP.vid));
  const sum = a => a.reduce((s, r) => s + (+r.tutar || 0), 0);
  const total = sum(recs);
  const byGroup = { yakit: 0, bakim: 0, diger: 0 };
  recs.forEach(r => byGroup[typeGroup(r.tur)] += +r.tutar || 0);
  // Türlere göre
  const byType = {};
  recs.forEach(r => { byType[r.tur] = (byType[r.tur] || 0) + (+r.tutar || 0); });
  const types = Object.entries(byType).sort((a, b) => b[1] - a[1]);
  // Araçlara göre
  const vmap = {};
  const vs = (S.vehicles || []).filter(v => !RP.vid || v.id === RP.vid);
  vs.forEach(v => vmap[v.id] = { v, plaka: v.plaka, recs: [] });
  recs.forEach(r => { (vmap[r.aracId] = vmap[r.aracId] || { v: null, plaka: r.plaka || "?", recs: [] }).recs.push(r); });
  const vehicles = Object.values(vmap).map(x => {
    const kms = x.recs.map(r => r.km).filter(k => k != null && k > 0);
    const dist = kms.length > 1 ? Math.max(...kms) - Math.min(...kms) : 0;
    const t = sum(x.recs);
    const g = { yakit: 0, bakim: 0, diger: 0 }; x.recs.forEach(r => g[typeGroup(r.tur)] += +r.tutar || 0);
    const litre = x.recs.filter(r => r.tur === "yakit").reduce((s, r) => s + (+(r.ekstra || {}).litre || 0), 0);
    return { ...x, total: t, g, dist, perKm: dist > 0 ? t / dist : null, litre, n: x.recs.length };
  }).sort((a, b) => b.total - a.total || String(a.plaka).localeCompare(String(b.plaka), "tr"));
  const distAll = vehicles.reduce((s, x) => s + x.dist, 0);
  const costWithDist = vehicles.filter(x => x.dist > 0).reduce((s, x) => s + x.total, 0);
  // Şoförlere göre
  const dmap = {};
  recs.forEach(r => { const k = r.createdBy || "?"; (dmap[k] = dmap[k] || { name: recDriver(r), recs: [] }).recs.push(r); });
  const drivers = Object.values(dmap).map(d => ({ name: d.name, n: d.recs.length, total: sum(d.recs),
    yakit: d.recs.filter(r => r.tur === "yakit").reduce((s, r) => s + (+r.tutar || 0), 0),
    last: d.recs.map(r => r.tarih).sort().pop() })).sort((a, b) => b.total - a.total);
  // Zaman serisi: 62 günden kısa aralıkta günlük, uzunsa aylık
  const days = Math.round((parseIso(to) - parseIso(from)) / 864e5) + 1;
  const daily = days <= 62; const buckets = [];
  if (daily) { for (let d = parseIso(from); d <= parseIso(to); d.setDate(d.getDate() + 1)) buckets.push({ key: isoD(d), label: `${d.getDate()} ${AYLAR[d.getMonth()]}`, v: 0 }); }
  else { const a = parseIso(from), b = parseIso(to); for (let d = new Date(a.getFullYear(), a.getMonth(), 1); d <= b; d.setMonth(d.getMonth() + 1)) buckets.push({ key: `${d.getFullYear()}-${pad2(d.getMonth() + 1)}`, label: `${AYLAR[d.getMonth()]} ${String(d.getFullYear()).slice(2)}`, v: 0 }); }
  const bi = {}; buckets.forEach((b, i) => bi[b.key] = i);
  recs.forEach(r => { const k = daily ? r.tarih : r.tarih.slice(0, 7); if (bi[k] != null) buckets[bi[k]].v += +r.tutar || 0; });
  return { from, to, recs, total, byGroup, types, vehicles, drivers, buckets, daily, perKm: distAll > 0 ? costWithDist / distAll : null, distAll };
}

function renderReports() {
  const D = rpCompute();
  const vs = (S.vehicles || []).slice().sort((a, b) => String(a.plaka).localeCompare(String(b.plaka), "tr"));
  const pend = (S.records || []).filter(r => r.status === "pending").length;
  $("#page").innerHTML = `
    <div class="rp-bar">
      <div class="seg" id="rpSeg">${RP_PRESETS.map(([k, l]) => `<button class="${RP.preset === k ? "on" : ""}" data-p="${k}">${l}</button>`).join("")}</div>
      ${RP.preset === "custom" ? `<div class="rp-dates"><input type="date" class="input" id="rpFrom" value="${D.from}" max="${D.to}"><span class="muted">–</span><input type="date" class="input" id="rpTo" value="${D.to}" min="${D.from}"></div>` : ""}
      <select class="input rp-sel" id="rpVeh"><option value="">Tüm araçlar</option>${vs.map(v => `<option value="${v.id}" ${RP.vid === v.id ? "selected" : ""}>${esc(v.plaka)} · ${esc(v.marka)} ${esc(v.model)}</option>`).join("")}</select>
      <button class="btn btn-primary" id="rpXls" ${D.recs.length ? "" : "disabled"}>${I.down} Excel'e aktar</button>
    </div>
    <p class="muted rp-note">${fmtD(D.from)} – ${fmtD(D.to)} · Sadece onaylanan kayıtlar${pend ? ` · <a class="link" href="#" data-goto="approvals">${pend} kayıt onay bekliyor</a>` : ""}</p>
    <div class="kpis">
      <div class="kpi"><small>Toplam masraf</small><b>${TL(D.total)}</b><span class="kpi-sub">${D.recs.length} kayıt</span></div>
      <div class="kpi"><small>Yakıt</small><b>${TL(D.byGroup.yakit)}</b><span class="kpi-sub">Toplamın %${D.total ? Math.round(D.byGroup.yakit / D.total * 100) : 0}</span></div>
      <div class="kpi"><small>Bakım ve lastik</small><b>${TL(D.byGroup.bakim)}</b><span class="kpi-sub">Toplamın %${D.total ? Math.round(D.byGroup.bakim / D.total * 100) : 0}</span></div>
      <div class="kpi"><small>Km başına maliyet</small><b>${D.perKm != null ? TL(Math.round(D.perKm * 100) / 100) : "–"}</b><span class="kpi-sub">${D.distAll ? NUM(D.distAll) + " km yol" : "Km girilen kayıt yetersiz"}</span></div>
    </div>
    ${D.recs.length ? `
    <div class="rp-grid">
      <div class="card"><h3 class="rp-h">${D.daily ? "Günlük" : "Aylık"} masraf</h3><div class="chart" id="rpChart"></div></div>
      <div class="card"><h3 class="rp-h">Masraf türleri</h3>${D.types.map(([t, v]) => `
        <div class="hb"><div class="hb-top"><span>${REC_TYPES[t] || esc(t)}</span><span><b>${TL(v)}</b> <span class="muted">· %${D.total ? (v / D.total * 100).toLocaleString("tr-TR", { maximumFractionDigits: 1 }) : 0}</span></span></div>
        <div class="hb-track"><div class="hb-fill" style="width:${D.types[0][1] ? Math.max(1.5, v / D.types[0][1] * 100) : 0}%"></div></div></div>`).join("")}
      </div>
    </div>` : `<div class="card empty"><div class="ic">${I.chart}</div><h3>Bu aralıkta onaylanan kayıt yok</h3><p>Başka bir tarih aralığı seç ya da <a class="link" href="#" data-goto="approvals">Kayıt Onayları</a> sayfasından bekleyen kayıtları onayla.</p></div>`}
    <div class="card" style="padding:0;overflow:hidden;margin-top:16px">
      <h3 class="rp-h" style="padding:16px 18px 0">Araçlara göre</h3>
      <div style="overflow-x:auto"><table class="tbl vt rp-t"><thead><tr><th>Araç</th><th class="hm">Şoför</th><th class="num">Yakıt</th><th class="num">Bakım/Lastik</th><th class="num hm">Diğer</th><th class="num">Toplam</th><th class="num">Yol (km)</th><th class="num">₺ / km</th><th class="num hm">Kayıt</th></tr></thead>
      <tbody>${D.vehicles.map(x => `<tr class="${x.n ? "" : "dim"}">
        <td><span class="plate"><i>TR</i><b>${esc(x.plaka)}</b></span><div class="muted" style="font-size:12.5px">${x.v ? esc(x.v.marka + " " + x.v.model) : "Silinmiş araç"}</div></td>
        <td class="hm">${x.v ? esc(driverName(x.v.assignedUid) || "–") : "–"}</td>
        <td class="num">${TL(x.g.yakit)}</td><td class="num">${TL(x.g.bakim)}</td><td class="num hm">${TL(x.g.diger)}</td>
        <td class="num"><b>${TL(x.total)}</b></td><td class="num">${x.dist ? NUM(x.dist) : "–"}</td>
        <td class="num">${x.perKm != null ? TL(Math.round(x.perKm * 100) / 100) : "–"}</td><td class="num hm">${x.n}</td></tr>`).join("") || `<tr><td colspan="9" class="muted" style="text-align:center;padding:24px">Araç yok</td></tr>`}</tbody></table></div>
      <div class="tfoot">Yol, aracın bu aralıktaki kayıtlarında girilen en yüksek ve en düşük km farkıdır. En az iki kayıtta km girilmesi gerekir.</div>
    </div>
    <div class="card" style="padding:0;overflow:hidden;margin-top:16px">
      <h3 class="rp-h" style="padding:16px 18px 0">Şoförlere göre</h3>
      ${D.drivers.length ? `<div style="overflow-x:auto"><table class="tbl vt rp-t"><thead><tr><th>Şoför</th><th class="num">Kayıt</th><th class="num">Yakıt</th><th class="num">Toplam</th><th class="num hm">Son kayıt</th></tr></thead>
      <tbody>${D.drivers.map(d => `<tr><td>${esc(d.name)}</td><td class="num">${d.n}</td><td class="num">${TL(d.yakit)}</td><td class="num"><b>${TL(d.total)}</b></td><td class="num hm">${fmtD(d.last)}</td></tr>`).join("")}</tbody></table></div>`
      : `<p class="muted" style="padding:12px 18px 18px">Bu aralıkta şoför kaydı yok.</p>`}
    </div>`;

  $("#rpSeg").querySelectorAll("[data-p]").forEach(b => b.addEventListener("click", () => {
    if (b.dataset.p === "custom" && RP.preset !== "custom") { const [f, t] = rpRange(); RP.from = f; RP.to = t; }
    RP.preset = b.dataset.p; renderReports();
  }));
  if ($("#rpFrom")) {
    const ch = () => { const f = $("#rpFrom").value, t = $("#rpTo").value; if (!f || !t) return; if (f > t) return toast("Başlangıç tarihi bitişten sonra olamaz.", true); RP.from = f; RP.to = t; renderReports(); };
    $("#rpFrom").addEventListener("change", ch); $("#rpTo").addEventListener("change", ch);
  }
  $("#rpVeh").addEventListener("change", e => { RP.vid = e.target.value; renderReports(); });
  $("#rpXls").addEventListener("click", () => exportReport(D));
  $("#page").querySelectorAll("[data-goto]").forEach(a => a.addEventListener("click", e => { e.preventDefault(); S.page = a.dataset.goto; renderShell(); }));
  const ch = $("#rpChart");
  if (ch) { ch.innerHTML = barChart(D.buckets, ch.clientWidth); bindChartHover(); }
}

/* Tek seriden oluşan çubuk grafik (SVG) – üzerine gelince tutar görünür */
function barChart(b, width) {
  const W = Math.max(300, Math.round(width || 640)), H = W < 500 ? 200 : 240, L = 44, R = 6, T = 12, B = 26;
  const max = Math.max(...b.map(x => x.v), 0);
  const step = niceStep(max / 4 || 1); const top = Math.max(step * 4, step);
  const cw = (W - L - R) / b.length; const bw = Math.max(2, Math.min(36, cw * 0.62));
  const y = v => T + (H - T - B) * (1 - v / top);
  const every = Math.ceil(b.length / Math.max(3, Math.floor((W - L) / 64)));
  let g = "";
  for (let i = 0; i <= 4; i++) { const v = step * i; if (v > top) break; g += `<line x1="${L}" x2="${W - R}" y1="${y(v)}" y2="${y(v)}" class="gl"/><text x="${L - 8}" y="${y(v) + 4}" class="ax" text-anchor="end">${TLc(v)}</text>`; }
  const bars = b.map((x, i) => {
    const cx = L + cw * i + cw / 2; const h = (H - T - B) - (y(x.v) - T);
    const r = Math.min(4, bw / 2, h);
    const yy = y(x.v), x0 = cx - bw / 2, base = H - B;
    const path = x.v > 0 ? `M${x0},${base}V${yy + r}Q${x0},${yy} ${x0 + r},${yy}H${x0 + bw - r}Q${x0 + bw},${yy} ${x0 + bw},${yy + r}V${base}Z` : "";
    return `<g class="bar" data-l="${esc(x.label)}" data-v="${x.v}">
      <rect x="${L + cw * i}" y="${T}" width="${cw}" height="${H - T - B}" class="hit"/>
      ${path ? `<path d="${path}" class="bf"/>` : ""}
      ${i % every === 0 ? `<text x="${cx}" y="${H - 9}" class="ax" text-anchor="middle">${esc(x.label)}</text>` : ""}</g>`;
  }).join("");
  return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Masraf grafiği">${g}<line x1="${L}" x2="${W - R}" y1="${H - B}" y2="${H - B}" class="bl"/>${bars}</svg><div class="tip" id="rpTip"></div>`;
}
let rpResizeT;
window.addEventListener("resize", () => { clearTimeout(rpResizeT); rpResizeT = setTimeout(() => { if (S.page === "reports" && $("#rpChart")) renderReports(); }, 250); });
function niceStep(raw) { const p = Math.pow(10, Math.floor(Math.log10(raw))); const n = raw / p; return (n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10) * p; }
function bindChartHover() {
  const c = $("#rpChart"); if (!c) return; const tip = $("#rpTip");
  c.querySelectorAll(".bar").forEach(g => {
    const show = e => {
      c.querySelectorAll(".bar.on").forEach(x => x.classList.remove("on")); g.classList.add("on");
      tip.innerHTML = `<small>${g.dataset.l}</small><b>${TL(+g.dataset.v)}</b>`;
      const rc = c.getBoundingClientRect(), rb = g.querySelector(".hit").getBoundingClientRect();
      tip.style.left = Math.min(Math.max(rb.left - rc.left + rb.width / 2, 50), rc.width - 50) + "px"; tip.classList.add("show");
    };
    g.addEventListener("mouseenter", show); g.addEventListener("click", show);
  });
  c.addEventListener("mouseleave", () => { tip.classList.remove("show"); c.querySelectorAll(".bar.on").forEach(x => x.classList.remove("on")); });
}

function exportReport(D) {
  const fmt = s => s ? s.split("-").reverse().join(".") : "";
  const veh = RP.vid ? (S.vehicles.find(v => v.id === RP.vid)?.plaka || "") : "Tüm araçlar";
  const r2 = n => Math.round(n * 100) / 100;
  const ozet = [["Driventa İşletme – Masraf Raporu", ""], ["Şirket", S.company.name], ["Tarih aralığı", `${fmt(D.from)} – ${fmt(D.to)}`], ["Araç", veh],
    ["Not", "Sadece onaylanan kayıtlar"], ["", ""], ["Masraf türü", "Tutar (₺)"],
    ...D.types.map(([t, v]) => [REC_TYPES[t] || t, r2(v)]), ["TOPLAM", r2(D.total)], ["", ""],
    ["Toplam yol (km)", D.distAll || ""], ["Km başına maliyet (₺)", D.perKm != null ? r2(D.perKm) : ""]];
  const araclar = [["Plaka", "Marka", "Model", "Atanan Şoför", "Yakıt (₺)", "Bakım/Lastik (₺)", "Diğer (₺)", "Toplam (₺)", "Yakıt (L/kWh)", "Yol (km)", "₺/km", "Kayıt sayısı"],
    ...D.vehicles.map(x => [x.plaka, x.v?.marka || "", x.v?.model || "", x.v ? (driverName(x.v.assignedUid) || "") : "", r2(x.g.yakit), r2(x.g.bakim), r2(x.g.diger), r2(x.total), r2(x.litre) || "", x.dist || "", x.perKm != null ? r2(x.perKm) : "", x.n])];
  const soforler = [["Şoför", "Kayıt sayısı", "Yakıt (₺)", "Toplam (₺)", "Son kayıt"], ...D.drivers.map(d => [d.name, d.n, r2(d.yakit), r2(d.total), fmt(d.last)])];
  const kayitlar = [["Tarih", "Plaka", "Şoför", "Tür", "Detay", "Tutar (₺)", "Km", "Litre/kWh", "Birim fiyat (₺)", "Açıklama", "Onay tarihi", "Fiş"],
    ...D.recs.slice().sort((a, b) => a.tarih.localeCompare(b.tarih)).map(r => { const e = r.ekstra || {};
      return [fmt(r.tarih), r.plaka || recVehicle(r)?.plaka || "", recDriver(r), REC_TYPES[r.tur] || r.tur, e.altTur || "", r2(+r.tutar || 0), r.km ?? "", e.litre ?? "", e.birimFiyat ?? "", r.aciklama || "", r.reviewedAt ? fmt(isoD(new Date(r.reviewedAt))) : "", e.fis ? "Var" : ""]; })];
  const blob = writeXlsx([
    { name: "Özet", widths: [26, 34], rows: ozet },
    { name: "Araçlar", widths: [14, 14, 16, 20, 12, 15, 12, 13, 13, 11, 10, 12], rows: araclar },
    { name: "Şoförler", widths: [24, 12, 13, 13, 13], rows: soforler },
    { name: "Kayıtlar", widths: [12, 14, 20, 12, 26, 12, 11, 11, 14, 28, 13, 8], rows: kayitlar }
  ]);
  download(blob, `driventa-rapor-${D.from}_${D.to}.xlsx`);
  toast("Rapor Excel dosyası indirildi.");
}

/* =====================================================================
   BAKIM TAKVİMİ (F5)
   ===================================================================== */
const MT_KINDS = [["muayene", "Muayene"], ["sigorta", "Trafik sigortası"], ["kasko", "Kasko"], ["bakim", "Periyodik bakım"]];
const MT = { view: "all", kind: "", q: "" };
const dayDiff = (a, b) => Math.round((parseIso(b) - parseIso(a)) / 864e5);
const addMonths = (iso, n) => { const d = parseIso(iso); const day = d.getDate(); d.setDate(1); d.setMonth(d.getMonth() + n); const last = new Date(d.getFullYear(), d.getMonth() + 1, 0).getDate(); d.setDate(Math.min(day, last)); return isoD(d); };

/* Araçlardan takip listesini çıkarır */
function trackItems(vs) {
  const today = isoD(new Date()); const out = [];
  (vs || []).forEach(v => {
    [["muayeneBitis", "muayene", "Muayene"], ["sigortaBitis", "sigorta", "Trafik sigortası"], ["kaskoBitis", "kasko", "Kasko"]].forEach(([k, t, ad]) => {
      if (v[k]) out.push({ v, t, ad, date: v[k], days: dayDiff(today, v[k]) });
    });
    if (v.sonBakimTarih || v.sonBakimKm != null) {
      const ak = v.bakimAralikKm || BAKIM_KM_VARSAYILAN, aa = v.bakimAralikAy || BAKIM_AY_VARSAYILAN;
      const date = v.sonBakimTarih ? addMonths(v.sonBakimTarih, aa) : null;
      const dueKm = v.sonBakimKm != null ? v.sonBakimKm + ak : null;
      out.push({ v, t: "bakim", ad: "Periyodik bakım", date, days: date ? dayDiff(today, date) : null, dueKm, remKm: dueKm != null ? dueKm - (v.km || 0) : null });
    }
  });
  out.forEach(x => {
    const late = (x.days != null && x.days < 0) || (x.remKm != null && x.remKm <= 0);
    const soon = (x.days != null && x.days <= 30) || (x.remKm != null && x.remKm <= 1000);
    x.st = late ? "late" : soon ? "soon" : "ok";
    // sıralama anahtarı: km bazlı yaklaşanları da tahmini güne çevir (günde ~100 km)
    x.sortKey = Math.min(x.days != null ? x.days : 1e6, x.remKm != null ? Math.floor(x.remKm / 100) : 1e6);
  });
  const rank = { late: 0, soon: 1, ok: 2 };
  return out.sort((a, b) => rank[a.st] - rank[b.st] || a.sortKey - b.sortKey);
}
function missingInfo(v) {
  const m = [];
  if (!v.muayeneBitis) m.push("Muayene");
  if (!v.sigortaBitis) m.push("Trafik sigortası");
  if (!v.sonBakimTarih && v.sonBakimKm == null) m.push("Son bakım");
  return m;
}
function remainText(x) {
  const parts = [];
  if (x.days != null) parts.push(x.days < 0 ? `${NUM(-x.days)} gün geçti` : x.days === 0 ? "Bugün" : `${NUM(x.days)} gün kaldı`);
  if (x.remKm != null) parts.push(x.remKm <= 0 ? `${NUM(-x.remKm)} km geçti` : `${NUM(x.remKm)} km kaldı`);
  return parts.join(" · ");
}
const ST_LBL = { late: ["late", "Süresi geçti"], soon: ["soon", "Yaklaşıyor"], ok: ["ok", "Zamanı var"] };

Pages.maintenance = function () {
  $("#page").innerHTML = `<div class="card"><div class="spinner"></div></div>`;
  Promise.all([api.listVehicles(S.company.id), api.listMembers(S.company.id).catch(() => S.members || [])])
    .then(([vs, ms]) => { S.vehicles = vs; S.members = ms; renderMaintenance(); })
    .catch(err => { $("#page").innerHTML = `<div class="alert err">${esc(errMsg(err))}</div>`; });
};

function renderMaintenance() {
  const vs = S.vehicles || [];
  const all = trackItems(vs);
  const late = all.filter(x => x.st === "late").length, soon = all.filter(x => x.st === "soon").length;
  const miss = vs.filter(v => missingInfo(v).length);
  if (!vs.length) {
    $("#page").innerHTML = emptyBox("car", "Henüz araç yok", `Bakım takvimi için önce <a class="link" href="#" data-goto="vehicles">Araçlar</a> sayfasından araç ekle.`);
    $("#page").querySelectorAll("[data-goto]").forEach(a => a.addEventListener("click", e => { e.preventDefault(); S.page = a.dataset.goto; renderShell(); }));
    return;
  }
  const q = trLow(MT.q);
  const views = [["all", "Tümü", all.length], ["late", "Süresi geçen", late], ["soon", "30 gün içinde", soon], ["miss", "Bilgisi eksik", miss.length]];
  $("#page").innerHTML = `
    <div class="kpis">
      <button class="kpi kbtn ${late ? "k-late" : ""}" data-v="late"><small>Süresi geçen</small><b>${late}</b><span class="kpi-sub">iş</span></button>
      <button class="kpi kbtn ${soon ? "k-soon" : ""}" data-v="soon"><small>30 gün içinde</small><b>${soon}</b><span class="kpi-sub">iş</span></button>
      <button class="kpi kbtn" data-v="all"><small>Takip edilen</small><b>${all.length}</b><span class="kpi-sub">iş · ${vs.length} araç</span></button>
      <button class="kpi kbtn" data-v="miss"><small>Bilgisi eksik</small><b>${miss.length}</b><span class="kpi-sub">araç</span></button>
    </div>
    <div class="rp-bar">
      <div class="seg" id="mtSeg">${views.map(([k, l, n]) => `<button class="${MT.view === k ? "on" : ""}" data-v="${k}">${l} (${n})</button>`).join("")}</div>
      ${MT.view !== "miss" ? `<select class="input rp-sel" id="mtKind"><option value="">Tüm işler</option>${MT_KINDS.map(([k, l]) => `<option value="${k}" ${MT.kind === k ? "selected" : ""}>${l}</option>`).join("")}</select>` : ""}
      <div class="search" style="flex:1;min-width:180px;max-width:300px">${I.search}<input class="input" id="mtQ" placeholder="Plaka veya şoför ara" value="${esc(MT.q)}"></div>
      <button class="btn btn-ghost" id="mtXls">${I.down} Excel'e indir</button>
    </div>
    <div class="card" style="padding:0;overflow:hidden;margin-top:8px"><div id="mtList"></div></div>
    <p class="muted" style="font-size:12.5px;margin-top:12px">Periyodik bakım; son bakım tarihine bakım aralığı (varsayılan 12 ay) ve son bakım km'sine km aralığı (varsayılan 15.000 km) eklenerek hesaplanır, hangisi önce gelirse. 1.000 km'den az kaldıysa da "yaklaşıyor" sayılır.</p>`;

  const box = $("#mtList");
  const match = v => !q || trLow(`${v.plaka} ${String(v.plaka).replace(/\s/g, "")} ${v.marka} ${v.model} ${driverName(v.assignedUid)}`).includes(q);
  if (MT.view === "miss") {
    const list = miss.filter(match);
    box.innerHTML = list.length ? `<div style="overflow-x:auto"><table class="tbl vt"><thead><tr><th>Araç</th><th class="hm">Şoför</th><th>Eksik bilgi</th><th></th></tr></thead><tbody>
      ${list.map(v => `<tr><td>${plateCell(v)}</td><td class="hm">${esc(driverName(v.assignedUid) || "–")}</td>
        <td>${missingInfo(v).map(m => `<span class="st-b miss">${m}</span>`).join(" ")}</td>
        <td class="acts"><button class="btn btn-sm btn-ghost" data-edit="${v.id}">${I.edit} Bilgileri gir</button></td></tr>`).join("")}</tbody></table></div>`
      : `<div class="empty"><div class="ic">${I.ok}</div><h3>${q ? "Sonuç yok" : "Eksik bilgi yok"}</h3><p>${q ? "Aramanla eşleşen araç bulunamadı." : "Tüm araçların muayene, sigorta ve bakım bilgileri girilmiş."}</p></div>`;
  } else {
    let list = all.filter(x => (MT.view === "all" || x.st === MT.view) && (!MT.kind || x.t === MT.kind) && match(x.v));
    box.innerHTML = list.length ? `<div style="overflow-x:auto"><table class="tbl vt mt-t"><thead><tr><th>Durum</th><th>Araç</th><th>İş</th><th>Son tarih</th><th>Kalan</th><th class="hm">Şoför</th><th></th></tr></thead><tbody>
      ${list.map(x => { const [c, l] = ST_LBL[x.st]; return `<tr>
        <td><span class="st-b ${c}">${l}</span></td>
        <td>${plateCell(x.v)}</td>
        <td><b style="font-weight:600">${x.ad}</b>${x.t === "bakim" && x.dueKm != null ? `<div class="muted" style="font-size:12.5px">Hedef: ${NUM(x.dueKm)} km${x.date ? " veya tarih" : ""}, hangisi önce gelirse</div>` : ""}</td>
        <td style="white-space:nowrap">${x.date ? fmtD(x.date) : "–"}</td>
        <td class="rem ${c}">${remainText(x)}</td>
        <td class="hm">${esc(driverName(x.v.assignedUid) || "–")}</td>
        <td class="acts"><button class="btn btn-sm btn-ghost" data-edit="${x.v.id}" title="Tarihleri düzenle">${I.edit}</button></td></tr>`; }).join("")}</tbody></table></div>
      <div class="tfoot">${list.length} iş</div>`
      : `<div class="empty"><div class="ic">${I.wrench}</div><h3>${all.length ? "Bu filtrede iş yok" : "Henüz takip bilgisi yok"}</h3><p>${all.length ? "Başka bir filtre seç." : 'Araçların muayene, sigorta ve bakım bilgilerini gir. <a class="link" href="#" data-mv="miss">Bilgisi eksik araçları göster</a>'}</p></div>`;
  }

  const go = v => { MT.view = v; renderMaintenance(); };
  $("#page").querySelectorAll(".kbtn, #mtSeg [data-v]").forEach(b => b.addEventListener("click", () => go(b.dataset.v)));
  $("#page").querySelectorAll("[data-mv]").forEach(a => a.addEventListener("click", e => { e.preventDefault(); go(a.dataset.mv); }));
  if ($("#mtKind")) $("#mtKind").addEventListener("change", e => { MT.kind = e.target.value; renderMaintenance(); });
  const qi = $("#mtQ"); qi.addEventListener("input", () => { MT.q = qi.value; const pos = qi.selectionStart; renderMaintenance(); const n = $("#mtQ"); n.focus(); n.setSelectionRange(pos, pos); });
  box.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => vehicleModal(S.vehicles.find(v => v.id === b.dataset.edit), renderMaintenance)));
  $("#mtXls").addEventListener("click", () => exportMaintenance(all, vs));
}
const plateCell = v => `<span class="plate"><i>TR</i><b>${esc(v.plaka)}</b></span><div class="muted" style="font-size:12.5px">${esc(v.marka)} ${esc(v.model)} · ${NUM(v.km || 0)} km</div>`;

function exportMaintenance(all, vs) {
  if (!vs.length) return toast("İndirilecek araç yok.", true);
  const rows = [["Durum", "Plaka", "Araç", "İş", "Son tarih", "Bakım km'si", "Kalan gün", "Kalan km", "Güncel km", "Şoför"],
    ...all.map(x => [ST_LBL[x.st][1], x.v.plaka, `${x.v.marka} ${x.v.model}`, x.ad, trDate(x.date), x.dueKm ?? "", x.days ?? "", x.remKm ?? "", x.v.km ?? "", driverName(x.v.assignedUid) || ""])];
  const eksik = [["Plaka", "Araç", "Eksik bilgi", "Şoför"], ...vs.filter(v => missingInfo(v).length).map(v => [v.plaka, `${v.marka} ${v.model}`, missingInfo(v).join(", "), driverName(v.assignedUid) || ""])];
  download(writeXlsx([{ name: "Bakım Takvimi", widths: [14, 14, 22, 18, 13, 13, 11, 11, 12, 20], rows }, { name: "Bilgisi Eksik", widths: [14, 22, 36, 20], rows: eksik }]),
    `driventa-bakim-takvimi-${isoD(new Date())}.xlsx`);
  toast("Bakım takvimi Excel dosyası indirildi.");
}

/* ---------- Yükleme akışı ---------- */
async function loadCompany(cid) {
  const [company, member] = await Promise.all([api.getCompany(cid), api.getMember(cid, S.user.uid)]);
  if (!company || !member || !member.active) { renderOnboarding(); return; }
  S.company = company; S.member = member;
  if (member.role === "driver") return renderDriverInfo();
  renderShell();
  loadRecords().catch(() => {}); // menüdeki "onay bekleyen" sayısı için
}

api.onUser(async user => {
  S.user = user; S.company = null; S.member = null; S.page = "overview"; S.records = null; S.pending = 0;
  if (!user) return renderAuth("login");
  root.innerHTML = '<div class="center"><div class="spinner"></div></div>';
  try {
    S.profile = await api.getProfile(user.uid);
    if (S.profile && S.profile.companyId) await loadCompany(S.profile.companyId);
    else renderOnboarding();
  } catch (err) {
    root.innerHTML = `<div class="center"><div class="card"><div class="alert err">${esc(errMsg(err))}</div>
      <button class="btn btn-ghost" id="retry">Tekrar dene</button> <button class="btn btn-ghost" id="lo">Çıkış yap</button></div></div>`;
    $("#retry").addEventListener("click", () => location.reload());
    $("#lo").addEventListener("click", () => api.signOut());
  }
});
