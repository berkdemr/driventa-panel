// Driventa İşletme – Panel arayüzü
import { api } from "./api.js";

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
const S = { user: null, profile: null, company: null, member: null, page: "overview", members: null };

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
      <p class="muted" style="font-size:13px">© ${new Date().getFullYear()} BeSey · <a class="link" href="https://driventa.com.tr/gizlilik.html">Gizlilik</a></p>
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
  { id: "vehicles", ad: "Araçlar", ic: "car", soon: true },
  { id: "approvals", ad: "Kayıt Onayları", ic: "check", soon: true },
  { id: "drivers", ad: "Şoförler", ic: "users" },
  { id: "maintenance", ad: "Bakım Takvimi", ic: "wrench", soon: true },
  { id: "settings", ad: "Ayarlar", ic: "cog" }
];

function renderShell() {
  const u = S.user, c = S.company, m = S.member;
  const nm = m.name || u.displayName || u.email;
  root.innerHTML = `
  <div class="shell">
    <aside class="side" id="side">
      <div class="brand"><img src="assets/logo.png" alt=""><span>Driventa<small>İŞLETME</small></span></div>
      ${PAGES.map(p => `<button class="nav-item ${S.page === p.id ? "on" : ""}" data-page="${p.id}">${I[p.ic]}${p.ad}${p.soon ? '<span class="soon">Yakında</span>' : ""}</button>`).join("")}
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
        <div class="kpi"><small>Araç</small><b>0</b></div>
        <div class="kpi"><small>Şoför</small><b id="kDrivers">–</b></div>
        <div class="kpi"><small>Bu ay masraf</small><b>₺0</b></div>
        <div class="kpi"><small>Onay bekleyen</small><b>0</b></div>
      </div>
      <div class="card">
        <h3 style="font-size:17px;margin-bottom:4px">Başlarken</h3>
        <p class="muted">Filonu Driventa'ya taşımak için şu adımları izle:</p>
        <ol class="steps">
          <li>Araçlarını ekle (Araçlar sayfası – yakında, Excel'den toplu aktarım dahil).</li>
          <li>Şoförlerine katılım kodunu ver: <b style="color:var(--accent);letter-spacing:2px">${esc(S.company.joinCode)}</b></li>
          <li>Şoförler Driventa uygulamasından bu kodla şirkete katılsın.</li>
          <li>Her şoföre aracını ata; şoförlerin girdiği kayıtları onayla.</li>
        </ol>
      </div>`;
    api.listMembers(S.company.id).then(ms => { S.members = ms; const el = $("#kDrivers"); if (el) el.textContent = ms.filter(x => x.role === "driver").length; })
      .catch(() => { const el = $("#kDrivers"); if (el) el.textContent = "–"; });
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
            <li>Sen şoföre aracını atarsın (Araçlar sayfası ile birlikte gelecek).</li>
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
    api.listMembers(c.id).then(ms => {
      S.members = ms;
      const roleAd = { owner: "Sahip", admin: "Yönetici", driver: "Şoför" };
      $("#mList").innerHTML = ms.length ? `<div style="overflow-x:auto"><table class="tbl"><thead><tr><th>Ad</th><th>E-posta</th><th>Rol</th><th>Araç</th></tr></thead><tbody>
        ${ms.sort((a, b) => (a.role === "owner" ? -1 : 1)).map(m => `<tr><td>${esc(m.name || "–")}</td><td>${esc(m.email || "–")}</td>
          <td><span class="role ${m.role}">${roleAd[m.role] || m.role}</span></td>
          <td class="muted">${m.role === "driver" ? (m.vehicleIds && m.vehicleIds.length ? m.vehicleIds.length + " araç" : "Atanmadı") : "–"}</td></tr>`).join("")}
        </tbody></table></div>` : emptyBox("users", "Henüz ekip üyesi yok", "Katılım kodunu şoförlerinle paylaş.");
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

/* İki aşamalı onay: ilk basışta uyarır, ikinci basışta işlemi yapar */
function confirmInline(btn, label) {
  if (btn.dataset.armed) { delete btn.dataset.armed; btn.innerHTML = btn.dataset.orig; return true; }
  btn.dataset.armed = "1"; btn.dataset.orig = btn.innerHTML; btn.innerHTML = label;
  setTimeout(() => { if (btn.dataset.armed) { delete btn.dataset.armed; btn.innerHTML = btn.dataset.orig; } }, 3500);
  return false;
}

/* ---------- Yükleme akışı ---------- */
async function loadCompany(cid) {
  const [company, member] = await Promise.all([api.getCompany(cid), api.getMember(cid, S.user.uid)]);
  if (!company || !member || !member.active) { renderOnboarding(); return; }
  S.company = company; S.member = member;
  if (member.role === "driver") return renderDriverInfo();
  renderShell();
}

api.onUser(async user => {
  S.user = user; S.company = null; S.member = null; S.page = "overview";
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
