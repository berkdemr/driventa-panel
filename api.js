// Driventa İşletme – Firebase veri katmanı
// Panelin geri kalanı Firebase'i doğrudan değil, bu dosyadaki fonksiyonları kullanır.
import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import {
  getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword,
  signInWithPopup, GoogleAuthProvider, signOut as fbSignOut, sendPasswordResetEmail, updateProfile
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import {
  getFirestore, doc, getDoc, setDoc, updateDoc, deleteDoc, collection, getDocs, serverTimestamp
} from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import { firebaseConfig } from "./firebase-config.js";

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
auth.languageCode = "tr";
const db = getFirestore(app);

const CODE_CHARS = "ABCDEFGHJKMNPQRSTUVWXYZ23456789"; // karışan harfler (I, L, O, 0, 1) yok
const randomCode = () => Array.from(crypto.getRandomValues(new Uint32Array(6)), n => CODE_CHARS[n % CODE_CHARS.length]).join("");

async function freeJoinCode() {
  for (let i = 0; i < 8; i++) {
    const code = randomCode();
    const snap = await getDoc(doc(db, "joinCodes", code));
    if (!snap.exists()) return code;
  }
  throw new Error("code-generation-failed");
}

export const api = {
  onUser(cb) { return onAuthStateChanged(auth, cb); },

  async signIn(email, pw) { await signInWithEmailAndPassword(auth, email, pw); },
  async signUp(name, email, pw) {
    const cred = await createUserWithEmailAndPassword(auth, email, pw);
    await updateProfile(cred.user, { displayName: name });
    await setDoc(doc(db, "users", cred.user.uid), { name, email, createdAt: serverTimestamp() }, { merge: true });
  },
  async signInGoogle() {
    const cred = await signInWithPopup(auth, new GoogleAuthProvider());
    const u = cred.user;
    await setDoc(doc(db, "users", u.uid), { name: u.displayName || "", email: u.email }, { merge: true });
  },
  async resetPassword(email) { await sendPasswordResetEmail(auth, email); },
  async signOut() { await fbSignOut(auth); },

  async getProfile(uid) {
    const s = await getDoc(doc(db, "users", uid));
    return s.exists() ? s.data() : null;
  },

  /* Şirket oluşturma sırası güvenlik kurallarına göre ayarlandı:
     1) şirket  2) sahibin üyelik kaydı  3) katılım kodu  4) kullanıcı profili */
  async createCompany(user, { name, phone, size }) {
    const cid = doc(collection(db, "companies")).id;
    const joinCode = await freeJoinCode();
    await setDoc(doc(db, "companies", cid), {
      name, phone: phone || "", size: size || "", ownerUid: user.uid, joinCode,
      plan: "trial", createdAt: serverTimestamp()
    });
    await setDoc(doc(db, "companies", cid, "members", user.uid), {
      role: "owner", name: user.displayName || "", email: user.email || "",
      vehicleIds: [], active: true, joinedAt: serverTimestamp()
    });
    await setDoc(doc(db, "joinCodes", joinCode), { companyId: cid });
    await setDoc(doc(db, "users", user.uid), { companyId: cid, role: "owner" }, { merge: true });
    return cid;
  },

  async getCompany(cid) {
    const s = await getDoc(doc(db, "companies", cid));
    return s.exists() ? { id: cid, ...s.data() } : null;
  },
  async getMember(cid, uid) {
    const s = await getDoc(doc(db, "companies", cid, "members", uid));
    return s.exists() ? { uid, ...s.data() } : null;
  },
  async updateCompany(cid, data) { await updateDoc(doc(db, "companies", cid), data); },

  async regenerateJoinCode(cid, oldCode) {
    const code = await freeJoinCode();
    await setDoc(doc(db, "joinCodes", code), { companyId: cid });
    await updateDoc(doc(db, "companies", cid), { joinCode: code });
    if (oldCode) { try { await deleteDoc(doc(db, "joinCodes", oldCode)); } catch (e) { /* eski kod zaten yoksa sorun değil */ } }
    return code;
  },

  async listMembers(cid) {
    const s = await getDocs(collection(db, "companies", cid, "members"));
    return s.docs.map(d => ({ uid: d.id, ...d.data() }));
  }
};
