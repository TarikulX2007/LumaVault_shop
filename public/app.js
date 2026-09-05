/* ================================
   🔥 এখানে তোমার নিজের Firebase কনফিগ বসাও (Firebase Console > Project Settings > General > Your apps)
   এই একটা অবজেক্ট রিপ্লেস করলেই বাকি পুরো কোড এমনিতেই কাজ করবে।
   ================================ */
const firebaseConfig = {
  apiKey: "AIzaSyA4teIWdY5lN7DK9sqItsSxMki0n-CTXlY",
  authDomain: "tgz-selling-website.firebaseapp.com",
  projectId: "tgz-selling-website",
  storageBucket: "tgz-selling-website.firebasestorage.app",
  messagingSenderId: "1070967687214",
  appId: "1:1070967687214:web:bddb6ec4d35390678437bf"
};

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js";
import {
  getAuth, browserLocalPersistence, setPersistence,
  createUserWithEmailAndPassword, signInWithEmailAndPassword,
  sendEmailVerification, onAuthStateChanged, signOut, updateProfile
} from "https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js";
import {
  getFirestore, doc, getDoc, setDoc, updateDoc, deleteDoc, addDoc,
  collection, query, where, getDocs, onSnapshot, orderBy,
  serverTimestamp, runTransaction, increment, getCountFromServer
} from "https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore.js";
import {
  getStorage, ref, uploadBytes, getDownloadURL
} from "https://www.gstatic.com/firebasejs/10.14.1/firebase-storage.js";
import * as XLSX from "https://cdn.sheetjs.com/xlsx-0.20.3/package/xlsx.mjs";

let app, auth, db, storage;
try {
  app = initializeApp(firebaseConfig);
  auth = getAuth(app);
  db = getFirestore(app);
  storage = getStorage(app);
  setPersistence(auth, browserLocalPersistence);
} catch (err) {
  console.error("Firebase failed to initialize:", err);
  const banner = document.createElement("div");
  banner.style.cssText = "position:fixed;top:0;left:0;right:0;z-index:9999;background:#F0576B;color:#fff;padding:12px 16px;font-family:sans-serif;font-size:13.5px;text-align:center;";
  banner.textContent = "Firebase failed to initialize (" + (err?.message || err) + "). Check your firebaseConfig, internet connection, and that this page is served over http(s):// rather than opened as a local file.";
  document.body.prepend(banner);
}

/* ---------- helpers ---------- */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const money = (n) => "৳" + Number(n || 0).toLocaleString();
const fmtDate = (ts) => { try { return ts?.toDate ? ts.toDate().toLocaleString() : "—"; } catch(e){ return "—"; } };
const Icon = {
  wallet: `<svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><rect x="2" y="6" width="20" height="14" rx="2"/><path d="M2 10h20"/><circle cx="17" cy="15" r="1.4"/></svg>`,
  box: `<svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8M12 13v8"/></svg>`,
  users: `<svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="9" cy="7" r="4"/><path d="M2 21c0-4 3-6 7-6s7 2 7 6M17 11a4 4 0 100-8M22 21c0-3.5-2.5-5.5-5-6"/></svg>`,
  receipt: `<svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M6 2h12v20l-3-2-3 2-3-2-3 2V2z"/><path d="M9 7h6M9 11h6"/></svg>`,
  clock: `<svg width="20" height="20" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 3"/></svg>`,
  cart: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="9" cy="21" r="1"/><circle cx="19" cy="21" r="1"/><path d="M2 3h2l2.6 12.6a2 2 0 002 1.4H19a2 2 0 002-1.6L22 8H6"/></svg>`,
  check: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M20 6L9 17l-5-5"/></svg>`,
  x: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M18 6L6 18M6 6l12 12"/></svg>`,
  trash: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M3 6h18M8 6V4a2 2 0 012-2h4a2 2 0 012 2v2m3 0l-1 14a2 2 0 01-2 2H7a2 2 0 01-2-2L4 6"/></svg>`,
  camera: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M4 8h3l2-3h6l2 3h3v11a1 1 0 01-1 1H5a1 1 0 01-1-1V8z"/><circle cx="12" cy="13" r="3.5"/></svg>`,
  logout: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M9 21H5a2 2 0 01-2-2V5a2 2 0 012-2h4"/><path d="M16 17l5-5-5-5M21 12H9"/></svg>`,
  ban: `<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M5.5 5.5l13 13"/></svg>`,
  coins: `<svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><ellipse cx="9" cy="6" rx="6" ry="3"/><path d="M3 6v6c0 1.66 2.69 3 6 3s6-1.34 6-3V6"/><path d="M3 12v6c0 1.66 2.69 3 6 3 1.13 0 2.19-.19 3-.52M15 9.5c3.31 0 6 1.34 6 3S18.31 15 15 15"/><path d="M15 15v3c0 1.66-2.69 3-6 3"/></svg>`,
  plus: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M12 5v14M5 12h14"/></svg>`,
  emptyBox: `<svg width="36" height="36" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M21 8l-9-5-9 5 9 5 9-5z"/><path d="M3 8v8l9 5 9-5V8M12 13v8"/></svg>`,
  download: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M12 3v12m0 0l-4-4m4 4l4-4M4 21h16"/></svg>`,
  user: `<svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="12" cy="8" r="4"/><path d="M4 21c0-4 4-6 8-6s8 2 8 6"/></svg>`,
  at: `<svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M16 12v1.5a2.5 2.5 0 005 0V12a9 9 0 10-4 7.5"/></svg>`,
  lock: `<svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><rect x="4" y="10" width="16" height="10" rx="2"/><path d="M8 10V7a4 4 0 018 0v3"/></svg>`,
  moon: `<svg width="19" height="19" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z"/></svg>`,
  sun: `<svg width="19" height="19" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.93 4.93l1.41 1.41M17.66 17.66l1.41 1.41M2 12h2M20 12h2M6.34 17.66l-1.41 1.41M19.07 4.93l-1.41 1.41"/></svg>`,
  more: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="5" cy="12" r="1"/><circle cx="12" cy="12" r="1"/><circle cx="19" cy="12" r="1"/></svg>`,
  info: `<svg width="17" height="17" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/></svg>`,
  send: `<svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" viewBox="0 0 24 24"><path d="M22 2L11 13"/><path d="M22 2l-7 20-4-9-9-4 20-7z"/></svg>`
};
function showMsg(elId, text, type){ const el = $(elId); el.textContent = text; el.className = "form-msg " + type; }
function hideMsg(elId){ $(elId).className = "form-msg"; }
function withTimeout(promise, ms, label){
  return Promise.race([
    promise,
    new Promise((_, reject) => setTimeout(() => reject({ code: "custom/timeout", message: (label||"Request") + " timed out after " + (ms/1000) + "s. Check your internet connection and Firebase setup." }), ms))
  ]);
}

/* ---------- theme ---------- */
const ThemeMod = {
  init(){
    const saved = localStorage.getItem("vb-theme") || "dark";
    document.documentElement.setAttribute("data-theme", saved);
    this.syncIcon(saved);
    $("themeToggle").addEventListener("click", () => this.toggle());
    $("themeToggleLanding").addEventListener("click", () => this.toggle());
  },
  toggle(){
    const cur = document.documentElement.getAttribute("data-theme");
    const next = cur === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("vb-theme", next);
    this.syncIcon(next);
  },
  syncIcon(theme){
    $("themeIconMoon").classList.toggle("hidden", theme === "light");
    $("themeIconSun").classList.toggle("hidden", theme === "dark");
    document.querySelector("#themeToggleLanding .tl-moon").classList.toggle("hidden", theme === "light");
    document.querySelector("#themeToggleLanding .tl-sun").classList.toggle("hidden", theme === "dark");
    $("themeToggleLandingLabel").textContent = theme === "dark" ? "Light mode" : "Dark mode";
  }
};
window.ThemeMod = ThemeMod;

/* ---------- state ---------- */
const State = {
  user: null, profile: null, products: [], paymentMethods: [],
  siteLive: true, notifications: [], notificationListener: null,
  siteListener: null
};

const UI = {
  toast(message, type="ok"){
    let host = $("toast-root");
    if(!host){ host = document.createElement("div"); host.id = "toast-root"; host.style.cssText = "position:fixed;right:18px;bottom:22px;z-index:1200;display:flex;flex-direction:column;gap:10px;max-width:min(360px,calc(100vw - 36px));"; document.body.appendChild(host); }
    const el = document.createElement("div");
    el.style.cssText = "padding:13px 15px;border:1px solid var(--border);border-radius:12px;background:var(--surface);color:var(--text);box-shadow:var(--shadow);font-size:13.5px;line-height:1.4;animation:modalPop .2s ease both;";
    if(type === "err") el.style.borderColor = "rgba(240,87,107,.45)";
    if(type === "ok") el.style.borderColor = "rgba(34,211,184,.35)";
    el.textContent = message; host.appendChild(el);
    setTimeout(() => el.remove(), 3200);
  },
  async ask(title, label, value="1", kind="number"){
    return new Promise(resolve => {
      const root = $("modal-root");
      root.innerHTML = `<div class="modal-overlay" id="askOverlay"><div class="modal">
        <div class="modal-head"><h3>${esc(title)}</h3><button class="modal-close" id="askClose" aria-label="Close">✕</button></div>
        <div class="field"><label>${esc(label)}</label><input id="askInput" type="${kind}" value="${esc(value)}" min="1" autofocus></div>
        <div style="display:flex;gap:10px;"><button class="btn btn-ghost btn-block" id="askCancel">Cancel</button><button class="btn btn-primary btn-block" id="askOk">Continue</button></div>
      </div></div>`;
      const finish = result => { root.innerHTML = ""; resolve(result); };
      $("askOverlay").addEventListener("click", e => { if(e.target.id === "askOverlay") finish(null); });
      $("askClose").onclick = () => finish(null); $("askCancel").onclick = () => finish(null);
      $("askOk").onclick = () => finish($("askInput").value);
      $("askInput").onkeydown = e => { if(e.key === "Enter") $("askOk").click(); if(e.key === "Escape") finish(null); };
      $("askInput").focus();
    });
  },
  async confirm(title, message, confirmLabel="Confirm", danger=false){
    return new Promise(resolve => {
      const root = $("modal-root");
      root.innerHTML = `<div class="modal-overlay" id="confirmOverlay"><div class="modal">
        <div class="modal-head"><h3>${esc(title)}</h3><button class="modal-close" id="confirmClose" aria-label="Close">✕</button></div>
        <p style="color:var(--text-muted);font-size:14px;line-height:1.55;margin-bottom:20px;">${esc(message)}</p>
        <div style="display:flex;gap:10px;"><button class="btn btn-ghost btn-block" id="confirmCancel">Cancel</button><button class="btn ${danger?"btn-danger":"btn-primary"} btn-block" id="confirmOk">${esc(confirmLabel)}</button></div>
      </div></div>`;
      const finish = result => { root.innerHTML = ""; resolve(result); };
      $("confirmOverlay").addEventListener("click", e => { if(e.target.id === "confirmOverlay") finish(false); });
      $("confirmClose").onclick = () => finish(false); $("confirmCancel").onclick = () => finish(false); $("confirmOk").onclick = () => finish(true);
    });
  }
};
window.UI = UI;

const SiteStatus = {
  render(){
    document.querySelectorAll(".live-pill").forEach(el => {
      el.classList.remove("hidden");
      el.classList.toggle("offline", !State.siteLive);
      const label = el.querySelector(".live-label");
      if(label) label.textContent = State.siteLive ? "Live" : "Offline";
    });
    document.querySelectorAll('[data-role="site-offline-notice"]').forEach(el => el.classList.toggle("hidden", State.siteLive));
  },
  start(){
    if(State.siteListener) return;
    const saved = localStorage.getItem("lumavault-site-live");
    if(saved !== null){ State.siteLive = saved !== "false"; this.render(); }
    State.siteListener = onSnapshot(doc(db,"settings","site"), snap => {
      const savedValue = localStorage.getItem("lumavault-site-live");
      State.siteLive = snap.exists() ? snap.data().isLive !== false : (savedValue === null ? true : savedValue !== "false");
      localStorage.setItem("lumavault-site-live", String(State.siteLive));
      this.render();
    }, err => {
      console.warn("Site status listener:", err);
      this.render();
    });
  },
  async setLive(isLive){
    State.siteLive = isLive;
    localStorage.setItem("lumavault-site-live", String(isLive));
    this.render();
    try{
      await setDoc(doc(db,"settings","site"), { isLive, updatedAt: serverTimestamp() }, { merge:true });
      return true;
    }catch(err){
      console.warn("Could not sync site status to Firebase:", err);
      return false;
    }
  }
};

const Notifications = {
  key(){ return "lumavault-read-" + (State.user?.uid || "guest"); },
  deletedKey(){ return "lumavault-deleted-notifications-" + (State.user?.uid || "guest"); },
  cacheKey(){ return "lumavault-notifications"; },
  date(n){ return fmtDate(n.createdAt) !== "—" ? fmtDate(n.createdAt) : (n.createdAtText || "Just now"); },
  deleted(){
    try{ return JSON.parse(localStorage.getItem(this.deletedKey()) || "[]"); }catch(e){ return []; }
  },
  visibleItems(){
    const removed = new Set(this.deleted());
    return State.notifications.filter(n => !removed.has(n.id));
  },
  unread(){
    const read = JSON.parse(localStorage.getItem(this.key()) || "[]");
    return this.visibleItems().some(n => !read.includes(n.id));
  },
  syncDot(){ document.querySelectorAll("#userNotificationDot,[data-notification-dot]").forEach(el => el.classList.toggle("hidden", !this.unread())); },
  markRead(){
    localStorage.setItem(this.key(), JSON.stringify(State.notifications.map(n => n.id)));
    this.syncDot();
  },
  open(){
    const root = $("modal-root");
    const items = this.visibleItems();
    root.innerHTML = `<div class="modal-overlay" id="notificationOverlay"><div class="modal modal-wide">
      <div class="modal-head"><h3>Notifications</h3><button class="modal-close" onclick="Modals.close()">✕</button></div>
      <div>${items.length ? items.map(n => `<div class="notification-item"><h4>${esc(n.title || "Message from admin")}</h4><p>${esc(n.message || "")}</p><time>${esc(this.date(n))}</time></div>`).join("") : `<div class="empty-state">${Icon.info}<div>No notifications yet.</div></div>`}</div>
      <div class="notification-actions">
        <button type="button" class="btn btn-danger btn-block" aria-label="Delete all notifications" onclick="window.Notifications.deleteAll()">${Icon.trash}Delete all</button>
        <button type="button" class="btn btn-ghost btn-block" aria-label="Enable phone notifications" onclick="window.Notifications.enablePhoneNotifications()">${Icon.check}Enable alerts</button>
      </div>
    </div></div>`;
    $("notificationOverlay").addEventListener("click", e => { if(e.target.id === "notificationOverlay") Modals.close(); });
    this.markRead();
  },
  async deleteAll(){
    const items = this.visibleItems();
    if(!items.length){ UI.toast("There are no notifications to delete."); return; }
    const ok = await UI.confirm("Delete all notifications","This will remove all notifications from your user panel.","Delete all",true);
    if(!ok) return;
    const removed = new Set(this.deleted());
    items.forEach(n => removed.add(n.id));
    localStorage.setItem(this.deletedKey(), JSON.stringify([...removed]));
    this.syncDot();
    this.open();
    UI.toast("All notifications deleted from your panel.");
  },
  start(){
    if(State.notificationListener || !State.user) return;
    try{ State.notifications = JSON.parse(localStorage.getItem(this.cacheKey()) || "[]"); }catch(e){ State.notifications = []; }
    this.syncDot();
    State.notificationListener = onSnapshot(collection(db,"notifications"), snap => {
      const previous = State.notifications;
      const remote = snap.docs.map(d => ({id:d.id,...d.data()}));
      const local = State.notifications.filter(n => String(n.id).startsWith("local-"));
      State.notifications = [...remote, ...local].sort((a,b) => (b.createdAt?.seconds || 0) - (a.createdAt?.seconds || 0));
      localStorage.setItem(this.cacheKey(), JSON.stringify(State.notifications.slice(0,50)));
      this.syncDot();
      const latest = State.notifications[0];
      if(latest && previous.length && latest.id !== previous[0]?.id && "Notification" in window && Notification.permission === "granted"){
        new Notification(latest.title || "LumaVault Shop", { body: latest.message || "You have a new message." });
      }
    }, err => {
      console.warn("Notification listener:", err);
      this.syncDot();
    });
  },
  async enablePhoneNotifications(){
    if(!("Notification" in window)){ UI.toast("This browser does not support phone notifications.","err"); return; }
    const permission = await Notification.requestPermission();
    UI.toast(permission === "granted" ? "Phone notifications enabled." : "Notification permission was not granted.", permission === "granted" ? "ok" : "err");
  }
};
window.Notifications = Notifications;

/* ---------- router ---------- */
const ALL_VIEWS = ["landing","signup","login","verify","admin-login","app","admin"];
/* Remembers which page the user was last on (per view), so a refresh
   lands back on the same page instead of always resetting to dashboard. */
const RouteMemory = {
  key: "lv-last-route",
  save(view, page){
    try{ localStorage.setItem(this.key, JSON.stringify({ view, page: page || null })); }catch(e){}
  },
  load(){
    try{ return JSON.parse(localStorage.getItem(this.key) || "null"); }catch(e){ return null; }
  },
  clear(){ try{ localStorage.removeItem(this.key); }catch(e){} }
};
window.RouteMemory = RouteMemory;

const Router = {
  go(view){
    ALL_VIEWS.forEach(v => $("view-" + v).classList.add("hidden"));
    const el = $("view-" + view);
    el.classList.remove("hidden");
    $("themeToggle").classList.toggle("hidden", view === "app" || view === "admin");
    el.classList.remove("anim-in"); void el.offsetWidth; el.classList.add("anim-in");
  },
  userPage(page){
    document.querySelectorAll("#view-app .nav-item[data-page]").forEach(n => n.classList.toggle("active", n.dataset.page === page));
    Pages.renderUser(page);
    const root = $("user-pages"); root.classList.remove("anim-in"); void root.offsetWidth; root.classList.add("anim-in");
    RouteMemory.save("app", page);
  },
  adminPage(page){
    document.querySelectorAll("#view-admin .nav-item[data-page]").forEach(n => n.classList.toggle("active", n.dataset.page === page));
    Pages.renderAdmin(page);
    const root = $("admin-pages"); root.classList.remove("anim-in"); void root.offsetWidth; root.classList.add("anim-in");
    RouteMemory.save("admin", page);
  }
};
window.Router = Router;

/* ---------- auth actions ---------- */
const Auth = {
  async signup(e){
    e.preventDefault();
    hideMsg("signupMsg");
    const firstName = $("suFirstName").value.trim();
    const lastName = $("suLastName").value.trim();
    const username = $("suUsername").value.trim();
    const email = $("suEmail").value.trim();
    const password = $("suPassword").value;
    const confirm = $("suConfirm").value;
    if(password !== confirm){ showMsg("signupMsg","Passwords do not match.","err"); return; }
    const usernameLower = username.toLowerCase();
    $("signupBtn").disabled = true; $("signupBtn").textContent = "Creating...";
    try{
      await withTimeout((async () => {
        const unameRef = doc(db, "usernames", usernameLower);
        const unameSnap = await getDoc(unameRef);
        if(unameSnap.exists()){ throw { code: "custom/username-taken" }; }
        const cred = await createUserWithEmailAndPassword(auth, email, password);
        await setDoc(unameRef, { uid: cred.user.uid });
        await setDoc(doc(db, "pendingProfiles", cred.user.uid), {
          firstName, lastName, username, email, createdAt: serverTimestamp()
        });
        await sendEmailVerification(cred.user);
      })(), 15000, "Signup");
      $("verifyEmailText").textContent = "We've sent a verification link to " + email + ". Verify it, then log in.";
      Router.go("verify");
    }catch(err){
      showMsg("signupMsg", Auth.friendlyError(err), "err");
      console.error("Signup error:", err);
    }finally{
      $("signupBtn").disabled = false; $("signupBtn").textContent = "Create account";
    }
  },
  async login(e){
    e.preventDefault();
    hideMsg("loginMsg");
    const email = $("liEmail").value.trim();
    const password = $("liPassword").value;
    $("loginBtn").disabled = true; $("loginBtn").textContent = "Logging in...";
    try{
      const cred = await withTimeout(signInWithEmailAndPassword(auth, email, password), 15000, "Login");
      if(!cred.user.emailVerified){
        $("verifyEmailText").textContent = "Please verify your email, then log in again.";
        Router.go("verify");
        return;
      }
      await withTimeout(Auth.ensureProfile(cred.user), 15000, "Login");
    }catch(err){
      showMsg("loginMsg", Auth.friendlyError(err), "err");
      console.error("Login error:", err);
    }finally{
      $("loginBtn").disabled = false; $("loginBtn").textContent = "Log in";
    }
  },
  async ensureProfile(user){
    const uref = doc(db, "users", user.uid);
    const usnap = await getDoc(uref);
    if(!usnap.exists()){
      const pendSnap = await getDoc(doc(db, "pendingProfiles", user.uid));
      const pend = pendSnap.exists() ? pendSnap.data() : {};
      await setDoc(uref, {
        uid: user.uid,
        firstName: pend.firstName || "",
        lastName: pend.lastName || "",
        username: pend.username || "",
        email: user.email,
        balance: 0,
        blocked: false,
        role: "user",
        photoURL: "",
        createdAt: serverTimestamp()
      });
    }
  },
  async adminLogin(e){
    e.preventDefault();
    hideMsg("adminLoginMsg");
    const email = $("alEmail").value.trim();
    const password = $("alPassword").value;
    $("adminLoginBtn").disabled = true; $("adminLoginBtn").textContent = "Signing in...";
    try{
      const cred = await signInWithEmailAndPassword(auth, email, password);
      const usnap = await getDoc(doc(db, "users", cred.user.uid));
      if(!usnap.exists() || usnap.data().role !== "admin"){
        await signOut(auth);
        throw { code: "custom/not-admin" };
      }
    }catch(err){
      showMsg("adminLoginMsg", Auth.friendlyError(err), "err");
    }finally{
      $("adminLoginBtn").disabled = false; $("adminLoginBtn").textContent = "Sign in";
    }
  },
  async resend(){
    hideMsg("verifyMsg");
    try{
      if(auth.currentUser){
        await sendEmailVerification(auth.currentUser);
        showMsg("verifyMsg","Verification email sent again.","ok");
      }else{
        showMsg("verifyMsg","Please log in again first.","err");
      }
    }catch(err){ showMsg("verifyMsg", Auth.friendlyError(err), "err"); }
  },
  async logout(){ await signOut(auth); RouteMemory.clear(); Router.go("login"); },
  friendlyError(err){
    if(err?.code === "custom/timeout") return err.message;
    const map = {
      "auth/email-already-in-use": "This email is already registered.",
      "auth/invalid-email": "That email address looks invalid.",
      "auth/weak-password": "Password should be at least 6 characters.",
      "auth/user-not-found": "No account found with that email.",
      "auth/wrong-password": "Incorrect password.",
      "auth/invalid-credential": "Incorrect email or password.",
      "auth/too-many-requests": "Too many attempts. Try again later.",
      "custom/username-taken": "That username is already taken.",
      "custom/not-admin": "This account does not have admin access."
    };
    return map[err?.code] || "Something went wrong. Please try again.";
  }
};
window.Auth = Auth;
$("signupForm").addEventListener("submit", Auth.signup);
$("loginForm").addEventListener("submit", Auth.login);
$("adminLoginForm").addEventListener("submit", (e) => Auth.adminLogin(e));
$("resendBtn").addEventListener("click", Auth.resend);
$("verifyLogoutBtn").addEventListener("click", Auth.logout);

/* ---------- auth state observer ---------- */
let unsubProfile = null, unsubProducts = null;
onAuthStateChanged(auth, async (user) => {
  if(unsubProfile) { unsubProfile(); unsubProfile = null; }
  if(!user){
    State.user = null; State.profile = null; State.notifications = [];
    if(State.notificationListener){ State.notificationListener(); State.notificationListener = null; }
    RouteMemory.clear();
    Router.go("landing"); return;
  }
  if(!user.emailVerified){ return; } // stays on verify screen if navigated there
  await Auth.ensureProfile(user);

  const uref = doc(db, "users", user.uid);
  const usnap = await getDoc(uref);
  const role = usnap.exists() ? usnap.data().role : "user";
  State.user = user;
  if(usnap.exists()) State.profile = usnap.data();
  SiteStatus.start();
  Notifications.start();
  unsubProfile = onSnapshot(uref, (snap) => {
    if(!snap.exists()) return;
    State.profile = snap.data();
    if(State.profile.role === "admin" && !document.getElementById("view-admin").classList.contains("hidden")){
      $("adminName").textContent = State.profile.firstName + " " + State.profile.lastName;
    }
    if(!document.getElementById("view-app").classList.contains("hidden")){
      SidebarUI.update();
    }
  });

  // If the admin-login form is mid-submit, let Auth.adminLogin's wrapper below
  // decide the route once it has confirmed the role — don't race it here.
  if(!$("view-admin-login").classList.contains("hidden")) return;

  // Route by the account's actual role (from Firestore), not by which view
  // happens to be visible right now — this is what makes a refresh land the
  // admin back in the admin panel instead of the user panel.
  const saved = RouteMemory.load();
  if(role === "admin"){
    $("adminName").textContent = (State.profile?.firstName || "Admin") + " " + (State.profile?.lastName || "");
    Router.go("admin");
    startAdminListeners();
    Router.adminPage(saved && saved.view === "admin" && saved.page ? saved.page : "dashboard");
  }else{
    Router.go("app");
    startProductListener();
    Router.userPage(saved && saved.view === "app" && saved.page ? saved.page : "dashboard");
  }
});

// after successful adminLogin sign-in, move to admin shell
const _origAdminLogin = Auth.adminLogin.bind(Auth);
Auth.adminLogin = async function(e){
  await _origAdminLogin(e);
  if(auth.currentUser){
    const usnap = await getDoc(doc(db,"users",auth.currentUser.uid));
    if(usnap.exists() && usnap.data().role === "admin"){
      State.profile = usnap.data();
      $("adminName").textContent = (State.profile.firstName || "Admin") + " " + (State.profile.lastName || "");
      Router.go("admin");
      startAdminListeners();
      Router.adminPage("dashboard");
    }
  }
};

function startProductListener(){
  if(unsubProducts) return;
  unsubProducts = onSnapshot(collection(db,"products"), (snap) => {
    State.products = snap.docs.map(d => ({ id: d.id, ...d.data() }));
    if(document.querySelector('.nav-item.active[data-page="products"]') || document.querySelector('#user-pages [data-role="products-grid"]')){
      Pages.renderUser("products");
    }
    if(document.querySelector('#user-pages [data-role="dashboard-root"]')){ Pages.renderUser("dashboard"); }
  });
  onSnapshot(collection(db,"paymentMethods"), (snap) => {
    State.paymentMethods = snap.docs.map(d => ({ id: d.id, ...d.data() }));
  });
}

/* ---------- Sidebar UI ---------- */
const SidebarUI = {
  update(){
    if(!State.profile) return;
    const name = (State.profile.firstName || "") + " " + (State.profile.lastName || "");
    $("sideName").textContent = name.trim() || State.profile.username || "User";
    $("sideBalance").textContent = money(State.profile.balance);
    $("sideAvatar").innerHTML = State.profile.photoURL
      ? `<img src="${esc(State.profile.photoURL)}" style="width:100%;height:100%;border-radius:50%;object-fit:cover;">`
      : esc((State.profile.firstName || "U")[0].toUpperCase());
  }
};

/* ================= USER PAGES ================= */
const Pages = {
  renderUser(page){
    const root = $("user-pages");
    if(page === "dashboard") return this.userDashboard(root);
    if(page === "products") return this.userProducts(root);
    if(page === "orders") return this.userOrders(root);
    if(page === "history") return this.userHistory(root);
    if(page === "profile") return this.userProfile(root);
    if(page === "deposit") return this.userDeposit(root);
  },

  userDashboard(root){
    root.innerHTML = `
      <div class="page" data-role="dashboard-root">
        <div class="page-head"><div><h1>Dashboard</h1><p>Your account at a glance.</p></div></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="ico">${Icon.wallet}</div><div class="label">Balance</div><div class="value accent">${money(State.profile?.balance)}</div></div>
          <div class="stat-card"><div class="ico">${Icon.box}</div><div class="label">Products available</div><div class="value">${State.products.length}</div></div>
          <div class="stat-card"><div class="ico">${Icon.receipt}</div><div class="label">Recent orders</div><div class="value" id="db-order-count">—</div></div>
          <div class="stat-card"><div class="ico">${Icon.clock}</div><div class="label">Recent deposits</div><div class="value" id="db-deposit-count">—</div></div>
        </div>
        <div class="card">
          <h3>Recent activity</h3>
          <div id="db-recent-orders"><div class="empty-state">Loading…</div></div>
        </div>
      </div>`;
    getDocs(query(collection(db,"orders"), where("uid","==",State.user.uid))).then(snap => {
      const orders = snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(b.createdAt?.seconds||0)-(a.createdAt?.seconds||0));
      $("db-order-count").textContent = orders.length;
      const recent = orders.slice(0,5);
      const statusIco = (s) => s === "completed" ? Icon.check : s === "rejected" ? Icon.x : Icon.clock;
      $("db-recent-orders").innerHTML = recent.length ? `<div class="status-list">
        ${recent.map(o=>`
          <div class="status-item">
            <div class="status-ico ${esc(o.status)}">${statusIco(o.status)}</div>
            <div class="status-body">
              <div class="status-title">${esc(o.productName)} <span class="status-qty">×${o.qty}</span></div>
              <div class="status-sub">${money(o.totalPrice)} · ${fmtDate(o.createdAt)}</div>
            </div>
            <span class="badge ${esc(o.status)}">${esc(o.status)}</span>
          </div>`).join("")}
      </div>` : `<div class="empty-state">${Icon.emptyBox}<div>No orders yet.</div></div>`;
    });
    getDocs(query(collection(db,"deposits"), where("uid","==",State.user.uid))).then(snap => {
      $("db-deposit-count").textContent = snap.size;
    });
  },

  userProducts(root){
    const cats = ["Proxy","VPN","Other"];
    root.innerHTML = `
      <div class="page">
        <div class="page-head"><div><h1>Products</h1><p>Buy instantly from your balance.</p></div></div>
        <div class="tabs" id="prod-tabs">${cats.map((c,i)=>`<div class="tab ${i===0?'active':''}" data-cat="${c}">${c}</div>`).join("")}</div>
        <div class="product-grid" data-role="products-grid" id="prod-grid"></div>
      </div>`;
    let activeCat = cats[0];
    const renderGrid = () => {
      const items = State.products.filter(p => p.category === activeCat);
      $("prod-grid").innerHTML = items.length ? items.map(p => `
        <div class="product-card">
          <span class="cat-tag">${esc(p.category)}</span>
          <h3>${esc(p.name)}</h3>
          <div class="desc">${esc(p.description || "")}</div>
          ${["Proxy","VPN"].includes(p.category) ? `<div class="stock-note">${(p.stock||[]).length} in stock</div>` : ""}
          <div class="price-row">
            <div class="price">${money(p.price)}</div>
            <button class="btn btn-primary btn-sm" onclick="ProductActions.buy('${p.id}')">${Icon.cart}Buy</button>
          </div>
        </div>`).join("") : `<div class="empty-state">${Icon.emptyBox}<div>No products in this category yet.</div></div>`;
    };
    document.querySelectorAll("#prod-tabs .tab").forEach(t => t.addEventListener("click", () => {
      document.querySelectorAll("#prod-tabs .tab").forEach(x=>x.classList.remove("active"));
      t.classList.add("active"); activeCat = t.dataset.cat; renderGrid();
    }));
    renderGrid();
  },

  userOrders(root){
    root.innerHTML = `
      <div class="page">
        <div class="page-head"><div><h1>Orders</h1><p>Everything you've purchased.</p></div></div>
        <div class="card"><div class="table-wrap" id="orders-table"><div class="empty-state">Loading…</div></div></div>
      </div>`;
    getDocs(query(collection(db,"orders"), where("uid","==",State.user.uid))).then(snap => {
      const orders = snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(b.createdAt?.seconds||0)-(a.createdAt?.seconds||0));
      $("orders-table").innerHTML = orders.length ? `<table><thead><tr><th>Product</th><th>Category</th><th>Qty</th><th>Total</th><th>Status</th><th>Date</th><th></th></tr></thead><tbody>
        ${orders.map(o=>`<tr><td>${esc(o.productName)}</td><td>${esc(o.category)}</td><td>${o.qty}</td><td>${money(o.totalPrice)}</td><td><span class="badge ${esc(o.status)}">${esc(o.status)}</span></td><td>${fmtDate(o.createdAt)}</td>
        <td>${o.deliveredContent ? `<button class="btn btn-ghost btn-sm" onclick='ProductActions.viewDelivery(${JSON.stringify(JSON.stringify(o.deliveredContent))}, ${JSON.stringify(o.productName)}, ${o.qty}, ${JSON.stringify(o.category)})'>View</button>` : ""}</td></tr>`).join("")}
      </tbody></table>` : `<div class="empty-state">${Icon.emptyBox}<div>No orders yet. Head to Products to get started.</div></div>`;
    });
  },

  userHistory(root){
    root.innerHTML = `
      <div class="page">
        <div class="page-head"><div><h1>History</h1><p>Purchases and deposits together.</p></div></div>
        <div class="card"><div class="table-wrap" id="history-table"><div class="empty-state">Loading…</div></div></div>
      </div>`;
    Promise.all([
      getDocs(query(collection(db,"orders"), where("uid","==",State.user.uid))),
      getDocs(query(collection(db,"deposits"), where("uid","==",State.user.uid)))
    ]).then(([oSnap,dSnap]) => {
      const items = [
        ...oSnap.docs.map(d=>({type:"Purchase",label:d.data().productName,amount:-d.data().totalPrice,status:d.data().status,createdAt:d.data().createdAt})),
        ...dSnap.docs.map(d=>({type:"Deposit",label:d.data().method,amount:d.data().amount,status:d.data().status,createdAt:d.data().createdAt}))
      ].sort((a,b)=>(b.createdAt?.seconds||0)-(a.createdAt?.seconds||0));
      $("history-table").innerHTML = items.length ? `<table><thead><tr><th>Type</th><th>Detail</th><th>Amount</th><th>Status</th><th>Date</th></tr></thead><tbody>
        ${items.map(i=>{
          const label = (i.type === "Deposit" && i.status === "approved") ? "Success" : i.status;
          return `<tr><td>${i.type}</td><td>${esc(i.label)}</td><td>${i.amount<0?'-':'+'}${money(Math.abs(i.amount))}</td><td><span class="badge ${esc(i.status)}">${esc(label)}</span></td><td>${fmtDate(i.createdAt)}</td></tr>`;
        }).join("")}
      </tbody></table>` : `<div class="empty-state">${Icon.emptyBox}<div>Nothing here yet.</div></div>`;
    });
  },

  userDeposit(root){
    const methods = State.paymentMethods;
    root.innerHTML = `
      <div class="page">
        <div class="page-head"><div><h1>Deposit funds</h1><p>Add balance securely from your preferred payment method.</p></div></div>
        <div class="notice-card ${State.siteLive ? "hidden" : ""}" data-role="site-offline-notice">${Icon.info}<div><strong>Deposits are temporarily paused.</strong><br>The shop is currently offline. Please try again when it is live.</div></div>
        <div class="card">
          <h3>Submit a deposit request</h3>
          <div class="form-msg" id="depositMsg"></div>
          <div class="field"><label>Amount (৳)</label><input type="number" id="depAmount" min="1" placeholder="Enter amount"></div>
          <div class="field"><label>Payment method</label>
            <select id="depMethod" ${methods.length && State.siteLive ? "" : "disabled"}>
              ${methods.length ? methods.map(m=>`<option value="${esc(m.id)}">${esc(m.name)}</option>`).join("") : `<option value="">No payment methods configured</option>`}
            </select>
          </div>
          <div class="deposit-method-card" id="selectedMethodCard">
            <div><div class="method-label">Send money to this number</div><div class="method-number" id="selectedMethodNumber">${methods.length ? esc(methods[0].details || "—") : "—"}</div></div>
            <button class="btn btn-ghost btn-sm copy-number" id="copyMethodNumber" ${methods.length ? "" : "disabled"}>${Icon.check}Copy</button>
          </div>
          <div class="field"><label>Transaction ID</label><input type="text" id="depTransactionId" placeholder="Paste your transaction ID" ${State.siteLive ? "" : "disabled"}></div>
          <div class="field"><label>Payment note (optional)</label><input type="text" id="depNote" placeholder="Any additional information" ${State.siteLive ? "" : "disabled"}></div>
          <button class="btn btn-primary" id="depositSubmit" ${(!methods.length || !State.siteLive) ? "disabled" : ""}>${Icon.check}Submit deposit</button>
        </div>
      </div>`;
    const methodById = id => State.paymentMethods.find(m => m.id === id);
    $("depMethod")?.addEventListener("change", e => {
      const method = methodById(e.target.value);
      $("selectedMethodNumber").textContent = method?.details || "—";
    });
    $("copyMethodNumber")?.addEventListener("click", async () => {
      const number = $("selectedMethodNumber").textContent;
      if(!number || number === "—") return;
      try { await navigator.clipboard.writeText(number); UI.toast("Payment number copied."); }
      catch(e){ UI.toast("Could not copy the payment number.","err"); }
    });
    $("depositSubmit")?.addEventListener("click", async () => {
      hideMsg("depositMsg");
      if(!State.siteLive){ showMsg("depositMsg","Deposits are paused while the website is offline.","err"); return; }
      const amount = parseFloat($("depAmount").value);
      const method = methodById($("depMethod").value);
      const transactionId = $("depTransactionId").value.trim();
      if(!amount || amount <= 0){ showMsg("depositMsg","Enter a valid amount.","err"); return; }
      if(!method){ showMsg("depositMsg","Select a payment method.","err"); return; }
      if(!transactionId){ showMsg("depositMsg","Enter the transaction ID.","err"); return; }
      const btn = $("depositSubmit"); btn.disabled = true; btn.textContent = "Submitting...";
      try{
        await addDoc(collection(db,"deposits"), {
          uid: State.user.uid, username: State.profile?.username || "",
          amount, method: method.name, methodDetails: method.details || "",
          transactionId, note: $("depNote").value.trim(), status: "pending", createdAt: serverTimestamp()
        });
        showMsg("depositMsg","Deposit submitted. Admin will review it shortly.","ok");
        $("depAmount").value = ""; $("depTransactionId").value = ""; $("depNote").value = "";
      }catch(err){ showMsg("depositMsg","Could not submit the deposit request.","err"); }
      finally{ btn.disabled = false; btn.innerHTML = Icon.check + "Submit deposit"; }
    });
  },

  userProfile(root){
    const p = State.profile || {};
    root.innerHTML = `
      <div class="page">
        <div class="page-head"><div><h1>Profile</h1><p>Manage your account.</p></div></div>
        <div class="card">
          <h3>Profile photo</h3>
          <div class="profile-photo-row">
            <div class="avatar" id="profAvatar">${p.photoURL ? `<img src="${esc(p.photoURL)}" style="width:100%;height:100%;border-radius:50%;object-fit:cover;">` : esc((p.firstName||"U")[0]?.toUpperCase())}</div>
            <div>
              <button class="btn btn-ghost btn-sm" id="changePhotoBtn">${Icon.camera}Change photo</button>
              <input type="file" accept="image/*" class="file-hidden" id="photoInput">
            </div>
          </div>
        </div>
        <div class="card">
          <h3>Personal information</h3>
          <div class="form-msg" id="profileMsg"></div>
          <div class="row2">
            <div class="field"><label>${Icon.user}First name</label><input type="text" id="pfFirstName" value="${esc(p.firstName)}"></div>
            <div class="field"><label>${Icon.user}Last name</label><input type="text" id="pfLastName" value="${esc(p.lastName)}"></div>
          </div>
          <div class="row2">
            <div class="field"><label>${Icon.at}Username</label><input type="text" value="${esc(p.username)}" disabled></div>
            <div class="field"><label>${Icon.lock}Email</label><input type="email" value="${esc(p.email)}" disabled></div>
          </div>
          <button class="btn btn-primary" id="saveProfileBtn">${Icon.check}Save changes</button>
        </div>
        <div class="card">
          <h3>Appearance</h3>
          <div class="appearance-row">
            <div class="info">
              <div class="ico" id="profThemeIco">${document.documentElement.getAttribute("data-theme")==="dark" ? Icon.moon : Icon.sun}</div>
              <div class="txt">
                <div class="t1" id="profThemeLabel">${document.documentElement.getAttribute("data-theme")==="dark" ? "Dark mode" : "Light mode"}</div>
                <div class="t2">Switch between light and dark.</div>
              </div>
            </div>
            <button class="btn btn-ghost btn-sm" id="profThemeToggle">Switch</button>
          </div>
        </div>
        <div class="card">
          <h3>Wallet</h3>
          <p style="color:var(--text-muted);font-size:13.5px;line-height:1.5;margin-bottom:14px;">Current balance: <strong style="color:var(--accent-2);">${money(p.balance)}</strong></p>
          <button class="btn btn-primary" onclick="Router.userPage('deposit')">${Icon.wallet}Deposit funds</button>
        </div>
        <div class="card"><button class="btn btn-danger btn-block" onclick="Auth.logout()">${Icon.logout}Log out</button></div>
      </div>`;
    $("profThemeToggle").addEventListener("click", () => {
      ThemeMod.toggle();
      const dark = document.documentElement.getAttribute("data-theme") === "dark";
      $("profThemeIco").innerHTML = dark ? Icon.moon : Icon.sun;
      $("profThemeLabel").textContent = dark ? "Dark mode" : "Light mode";
    });
    $("changePhotoBtn").addEventListener("click", () => $("photoInput").click());
    $("photoInput").addEventListener("change", async (e) => {
      const file = e.target.files[0]; if(!file) return;
      hideMsg("profileMsg");
      try{
        const r = ref(storage, "profilePhotos/" + State.user.uid);
        await uploadBytes(r, file);
        const url = await getDownloadURL(r);
        await updateDoc(doc(db,"users",State.user.uid), { photoURL: url });
        showMsg("profileMsg","Photo updated.","ok");
      }catch(err){ showMsg("profileMsg","Could not upload photo. Storage may need to be enabled on a Blaze plan.","err"); }
    });
    $("saveProfileBtn").addEventListener("click", async () => {
      hideMsg("profileMsg");
      try{
        await updateDoc(doc(db,"users",State.user.uid), {
          firstName: $("pfFirstName").value.trim(),
          lastName: $("pfLastName").value.trim()
        });
        showMsg("profileMsg","Profile updated.","ok");
      }catch(err){ showMsg("profileMsg","Could not save changes.","err"); }
    });
  }
};

/* ---------- product purchase logic ---------- */
window.ProductActions = {
  async buy(productId){
    if(!State.siteLive){ UI.toast("The shop is currently offline. Purchases are paused.","err"); return; }
    const product = State.products.find(p => p.id === productId);
    if(!product) return;
    const isDigital = ["Proxy","VPN"].includes(product.category);
    let qty = 1;
    if(isDigital){
      const input = await UI.ask("Choose quantity", `How many "${product.name}" would you like to buy?`, "1", "number");
      if(input === null) return;
      qty = parseInt(input, 10);
      if(!qty || qty < 1){ UI.toast("Enter a valid quantity.","err"); return; }
    }
    const total = product.price * qty;
    if((State.profile?.balance || 0) < total){ UI.toast("Insufficient balance. Please deposit first.","err"); return; }
    try{
      const result = await runTransaction(db, async (tx) => {
        const prodRef = doc(db, "products", productId);
        const prodSnap = await tx.get(prodRef);
        const userRef = doc(db, "users", State.user.uid);
        const userSnap = await tx.get(userRef);
        if(!prodSnap.exists()) throw new Error("Product no longer exists.");
        const pData = prodSnap.data();
        if((userSnap.data().balance||0) < total) throw new Error("Insufficient balance.");
        let delivered = null;
        if(isDigital){
          const stock = pData.stock || [];
          if(stock.length < qty) throw new Error("Not enough stock available.");
          delivered = stock.slice(0, qty);
          tx.update(prodRef, { stock: stock.slice(qty) });
        }
        tx.update(userRef, { balance: increment(-total) });
        const orderRef = doc(collection(db,"orders"));
        tx.set(orderRef, {
          uid: State.user.uid, productId, productName: pData.name, category: pData.category,
          qty, totalPrice: total, deliveredContent: delivered, status: "completed", createdAt: serverTimestamp()
        });
        return delivered;
      });
      if(isDigital && result){
        this.viewDelivery(JSON.stringify(result), product.name, qty, product.category);
      }else{
        UI.toast("Purchase complete.");
      }
      Router.userPage("products");
    }catch(err){
      UI.toast(err.message || "Purchase failed.","err");
    }
  },
  viewDelivery(jsonContent, productName, qty, category){
    let items;
    try{ items = JSON.parse(jsonContent); }catch(e){ items = []; }
    if(!Array.isArray(items)) items = [items];
    Modals.showDelivery(productName, items, category);
  }
};

/* ---------- modals ---------- */
const Modals = {
  openDeposit(){
    Router.userPage("deposit");
    return;
    /*
    const methods = State.paymentMethods;
    const root = $("modal-root");
    root.innerHTML = `
      <div class="modal-overlay" id="depositOverlay">
        <div class="modal">
          <div class="modal-head"><h3>Deposit funds</h3><div class="modal-close" onclick="Modals.close()">✕</div></div>
          <div class="form-msg" id="depositMsg"></div>
          <div class="field"><label>Amount (৳)</label><input type="number" id="depAmount" min="1"></div>
          <div class="field"><label>Payment method</label>
            <select id="depMethod">
              ${methods.length ? methods.map(m=>`<option value="${esc(m.name)}">${esc(m.name)} — ${esc(m.details||"")}</option>`).join("") : `<option value="">No methods configured</option>`}
            </select>
          </div>
          <button class="btn btn-primary btn-block" id="depositSubmit">${Icon.check}Submit request</button>
        </div>
      </div>`;
    $("depositOverlay").addEventListener("click", (e) => { if(e.target.id === "depositOverlay") Modals.close(); });
    $("depositSubmit").addEventListener("click", async () => {
      hideMsg("depositMsg");
      const amount = parseFloat($("depAmount").value);
      const method = $("depMethod").value;
      if(!amount || amount <= 0){ showMsg("depositMsg","Enter a valid amount.","err"); return; }
      if(!method){ showMsg("depositMsg","Select a payment method.","err"); return; }
      try{
        await addDoc(collection(db,"deposits"), { uid: State.user.uid, amount, method, status: "pending", createdAt: serverTimestamp() });
        showMsg("depositMsg","Deposit request submitted. Await admin approval.","ok");
        setTimeout(() => Modals.close(), 1400);
      }catch(err){ showMsg("depositMsg","Could not submit request.","err"); }
    });
    */
  },
  showDelivery(productName, items, category){
    const root = $("modal-root");
    const listHtml = items.map((it,i)=>`
      <div class="delivery-item">
        <div class="delivery-num">${i+1}.</div>
        <div class="delivery-text">${esc(it)}</div>
      </div>`).join("");
    root.innerHTML = `
      <div class="modal-overlay" id="deliveryOverlay">
        <div class="modal">
          <div class="modal-head"><h3>${esc(productName)}</h3><div class="modal-close" onclick="Modals.close()">✕</div></div>
          <div class="delivery-list">${listHtml || `<div class="empty-state">${Icon.emptyBox}<div>No content.</div></div>`}</div>
          <div style="display:flex; gap:10px;">
            <button class="btn btn-ghost btn-block" id="deliveryCopyBtn">${Icon.check}Copy</button>
            <button class="btn btn-primary btn-block" id="deliveryDownloadBtn">${Icon.download}Download</button>
          </div>
        </div>
      </div>`;
    $("deliveryOverlay").addEventListener("click", (e) => { if(e.target.id === "deliveryOverlay") Modals.close(); });
    $("deliveryCopyBtn").addEventListener("click", async () => {
      try{
        await navigator.clipboard.writeText(items.join("\n"));
        const btn = $("deliveryCopyBtn");
        if(btn){ btn.innerHTML = Icon.check + "Copied!"; setTimeout(()=>{ if($("deliveryCopyBtn")) $("deliveryCopyBtn").innerHTML = Icon.check + "Copy"; }, 1500); }
      }catch(e){ UI.toast("Could not copy. Your browser may be blocking clipboard access.","err"); }
    });
    $("deliveryDownloadBtn").addEventListener("click", () => {
      try{
        const catLabel = category === "VPN" ? "VPN" : category === "Proxy" ? "Proxy" : (category || "Item");
        const fileName = `TGZ-${catLabel}.xlsx`;
        const rows = [["#","Content"], ...items.map((it,i)=>[i+1, it])];
        const ws = XLSX.utils.aoa_to_sheet(rows);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, catLabel);
        XLSX.writeFile(wb, fileName);
      }catch(e){ UI.toast("Could not build the download file.","err"); console.error(e); }
    });
  },
  close(){ $("modal-root").innerHTML = ""; }
};
window.Modals = Modals;
window.Router = Router;
window.Auth = Auth;

/* ================= ADMIN PAGES ================= */
let adminUsersUnsub, adminDepositsUnsub;
function startAdminListeners(){
  if(!adminUsersUnsub){
    adminUsersUnsub = onSnapshot(collection(db,"users"), () => {
      if($("view-admin").classList.contains("hidden")) return;
      if(document.querySelector('[data-role="users-table"]')) Pages.adminUsers($("admin-pages"));
      if(document.querySelector('[data-role="admin-dash"]')) Pages.adminDashboard($("admin-pages"));
    });
  }
  if(!adminDepositsUnsub){
    adminDepositsUnsub = onSnapshot(collection(db,"deposits"), () => {
      if($("view-admin").classList.contains("hidden")) return;
      if(document.querySelector('[data-role="deposits-table"]')) Pages.adminDeposits($("admin-pages"));
      if(document.querySelector('[data-role="admin-dash"]')) Pages.adminDashboard($("admin-pages"));
    });
  }
  onSnapshot(collection(db,"products"), (snap) => {
    State.products = snap.docs.map(d=>({id:d.id,...d.data()}));
    if(document.querySelector('[data-role="admin-products-table"]')) Pages.adminProducts($("admin-pages"));
  });
  onSnapshot(collection(db,"paymentMethods"), (snap) => {
    State.paymentMethods = snap.docs.map(d=>({id:d.id,...d.data()}));
    if(document.querySelector('[data-role="admin-payments-table"]')) Pages.adminPayments($("admin-pages"));
  });
}

Object.assign(Pages, {
  renderAdmin(page){
    const root = $("admin-pages");
    if(page === "dashboard") return this.adminDashboard(root);
    if(page === "users") return this.adminUsers(root);
    if(page === "products") return this.adminProducts(root);
    if(page === "payments") return this.adminPayments(root);
    if(page === "deposits") return this.adminDeposits(root);
  },

  async adminDashboard(root){
    root.innerHTML = `
      <div class="page" data-role="admin-dash">
        <div class="page-head"><div><h1>Dashboard</h1><p>Store overview.</p></div></div>
        <div class="stat-grid">
          <div class="stat-card"><div class="ico">${Icon.users}</div><div class="label">Total users</div><div class="value" id="stat-users">—</div></div>
          <div class="stat-card"><div class="ico">${Icon.receipt}</div><div class="label">Total orders</div><div class="value" id="stat-orders">—</div></div>
          <div class="stat-card"><div class="ico">${Icon.wallet}</div><div class="label">Approved deposits</div><div class="value accent" id="stat-deposits">—</div></div>
          <div class="stat-card"><div class="ico">${Icon.clock}</div><div class="label">Pending deposits</div><div class="value" id="stat-pending">—</div></div>
        </div>
      </div>`;
    const [uCount, oCount, depSnap] = await Promise.all([
      getCountFromServer(collection(db,"users")),
      getCountFromServer(collection(db,"orders")),
      getDocs(collection(db,"deposits"))
    ]);
    $("stat-users").textContent = uCount.data().count;
    $("stat-orders").textContent = oCount.data().count;
    const deps = depSnap.docs.map(d=>d.data());
    const approvedTotal = deps.filter(d=>d.status==="approved").reduce((s,d)=>s+Number(d.amount||0),0);
    const pendingCount = deps.filter(d=>d.status==="pending").length;
    $("stat-deposits").textContent = money(approvedTotal);
    $("stat-pending").textContent = pendingCount;
  },

  adminUsers(root){
    root.innerHTML = `
      <div class="page">
        <div class="page-head"><div><h1>Users</h1><p>Manage balances and access.</p></div></div>
        <div class="card"><div class="table-wrap" data-role="users-table" id="users-table"><div class="empty-state">Loading…</div></div></div>
      </div>`;
    getDocs(collection(db,"users")).then(snap => {
      const users = snap.docs.map(d=>({id:d.id,...d.data()}));
      $("users-table").innerHTML = users.length ? `<table><thead><tr><th>Name</th><th>Username</th><th>Email</th><th>Balance</th><th>Status</th><th>Actions</th></tr></thead><tbody>
        ${users.map(u=>`<tr>
          <td>${esc(u.firstName)} ${esc(u.lastName)}</td>
          <td>${esc(u.username)}</td>
          <td>${esc(u.email)}</td>
          <td>${money(u.balance)}</td>
          <td><span class="badge ${u.blocked?'blocked':'active'}">${u.blocked?'Blocked':'Active'}</span></td>
          <td style="display:flex; gap:6px; flex-wrap:wrap;">
            <button class="btn btn-ghost btn-sm" onclick="AdminActions.toggleBlock('${u.id}', ${u.blocked ? 'false' : 'true'})">${u.blocked?Icon.check+'Unblock':Icon.ban+'Block'}</button>
            <button class="btn btn-ghost btn-sm" onclick="AdminActions.adjustBalance('${u.id}')">${Icon.coins}Adjust balance</button>
          </td>
        </tr>`).join("")}
      </tbody></table>` : `<div class="empty-state">${Icon.emptyBox}<div>No users yet.</div></div>`;
    });
  },

  adminProducts(root){
    root.innerHTML = `
      <div class="page" data-role="admin-products-table">
        <div class="page-head"><div><h1>Products</h1><p>Add and manage store items.</p></div></div>
        <div class="card">
          <h3>Add product</h3>
          <div class="form-msg" id="addProductMsg"></div>
          <div class="row2">
            <div class="field"><label>Name</label><input type="text" id="npName"></div>
            <div class="field"><label>Category</label><select id="npCategory"><option>Proxy</option><option>VPN</option><option>Other</option></select></div>
          </div>
          <div class="row2">
            <div class="field"><label>Price (৳)</label><input type="number" id="npPrice" min="0"></div>
            <div class="field"><label>Description</label><input type="text" id="npDesc"></div>
          </div>
          <div class="field"><label>Stock lines (one item per line — used for Proxy/VPN auto-delivery)</label><textarea id="npStock" rows="4" placeholder="1.2.3.4:8080:user:pass"></textarea></div>
          <button class="btn btn-primary" id="addProductBtn">${Icon.plus}Add product</button>
        </div>
        <div class="card">
          <h3>Existing products</h3>
          <div class="table-wrap" id="admin-products-table"><div class="empty-state">Loading…</div></div>
        </div>
      </div>`;
    const renderList = () => {
      $("admin-products-table").innerHTML = State.products.length ? `<table><thead><tr><th>Name</th><th>Category</th><th>Price</th><th>Stock</th><th></th></tr></thead><tbody>
        ${State.products.map(p=>`<tr><td>${esc(p.name)}</td><td>${esc(p.category)}</td><td>${money(p.price)}</td><td>${["Proxy","VPN"].includes(p.category) ? (p.stock||[]).length : "—"}</td>
        <td><button class="btn btn-danger btn-sm" onclick="AdminActions.deleteProduct('${p.id}')">${Icon.trash}Delete</button></td></tr>`).join("")}
      </tbody></table>` : `<div class="empty-state">${Icon.emptyBox}<div>No products yet.</div></div>`;
    };
    renderList();
    $("addProductBtn").addEventListener("click", async () => {
      hideMsg("addProductMsg");
      const name = $("npName").value.trim();
      const category = $("npCategory").value;
      const price = parseFloat($("npPrice").value);
      const description = $("npDesc").value.trim();
      const stockLines = $("npStock").value.split("\n").map(s=>s.trim()).filter(Boolean);
      if(!name || !price){ showMsg("addProductMsg","Name and price are required.","err"); return; }
      try{
        await addDoc(collection(db,"products"), { name, category, price, description, stock: stockLines, createdAt: serverTimestamp() });
        showMsg("addProductMsg","Product added.","ok");
        $("npName").value=""; $("npPrice").value=""; $("npDesc").value=""; $("npStock").value="";
      }catch(err){ showMsg("addProductMsg","Could not add product.","err"); }
    });
  },

  adminPayments(root){
    root.innerHTML = `
      <div class="page" data-role="admin-payments-table">
        <div class="page-head"><div><h1>Payment methods</h1><p>Options shown to users during deposit.</p></div></div>
        <div class="card">
          <h3>Add method</h3>
          <div class="form-msg" id="addMethodMsg"></div>
          <div class="row2">
            <div class="field"><label>Name (e.g. bKash)</label><input type="text" id="pmName"></div>
            <div class="field"><label>Number / details</label><input type="text" id="pmDetails"></div>
          </div>
          <div class="field"><label>Instructions (optional)</label><input type="text" id="pmInstructions"></div>
          <button class="btn btn-primary" id="addMethodBtn">${Icon.plus}Add method</button>
        </div>
        <div class="card"><h3>Existing methods</h3><div class="table-wrap" id="admin-payments-table"><div class="empty-state">Loading…</div></div></div>
      </div>`;
    const renderList = () => {
      $("admin-payments-table").innerHTML = State.paymentMethods.length ? `<table><thead><tr><th>Name</th><th>Details</th><th>Instructions</th><th></th></tr></thead><tbody>
        ${State.paymentMethods.map(m=>`<tr><td>${esc(m.name)}</td><td>${esc(m.details)}</td><td>${esc(m.instructions||"")}</td><td><button class="btn btn-danger btn-sm" onclick="AdminActions.deleteMethod('${m.id}')">${Icon.trash}Delete</button></td></tr>`).join("")}
      </tbody></table>` : `<div class="empty-state">${Icon.emptyBox}<div>No payment methods yet.</div></div>`;
    };
    renderList();
    $("addMethodBtn").addEventListener("click", async () => {
      hideMsg("addMethodMsg");
      const name = $("pmName").value.trim();
      const details = $("pmDetails").value.trim();
      const instructions = $("pmInstructions").value.trim();
      if(!name || !details){ showMsg("addMethodMsg","Name and details are required.","err"); return; }
      try{
        await addDoc(collection(db,"paymentMethods"), { name, details, instructions, createdAt: serverTimestamp() });
        showMsg("addMethodMsg","Method added.","ok");
        $("pmName").value=""; $("pmDetails").value=""; $("pmInstructions").value="";
      }catch(err){ showMsg("addMethodMsg","Could not add method.","err"); }
    });
  },

  adminDeposits(root){
    root.innerHTML = `
      <div class="page">
        <div class="page-head"><div><h1>Deposits</h1><p>Approve or reject requests.</p></div></div>
        <div class="card"><div class="table-wrap" data-role="deposits-table" id="deposits-table"><div class="empty-state">Loading…</div></div></div>
      </div>`;
    getDocs(collection(db,"deposits")).then(async (snap) => {
      const deposits = snap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(b.createdAt?.seconds||0)-(a.createdAt?.seconds||0));
      const userIds = [...new Set(deposits.map(d=>d.uid))];
      const userMap = {};
      await Promise.all(userIds.map(async uid => {
        const s = await getDoc(doc(db,"users",uid));
        if(s.exists()) userMap[uid] = s.data();
      }));
       $("deposits-table").innerHTML = deposits.length ? `<table><thead><tr><th>User</th><th>Amount</th><th>Method</th><th>Transaction ID</th><th>Status</th><th>Date</th><th>Actions</th></tr></thead><tbody>
        ${deposits.map(d=>`<tr>
           <td><strong>${esc(d.username || userMap[d.uid]?.username || d.uid)}</strong><br><span style="font-size:11px;color:var(--text-muted);">${esc(userMap[d.uid]?.email || "")}</span></td>
          <td>${money(d.amount)}</td>
           <td>${esc(d.method)}<br><span style="font-size:11px;color:var(--text-muted);">${esc(d.methodDetails || "")}</span></td>
           <td><span class="content-box" style="display:inline-block;margin:0;padding:7px 9px;max-width:180px;">${esc(d.transactionId || "—")}</span></td>
          <td><span class="badge ${esc(d.status)}">${esc(d.status)}</span></td>
          <td>${fmtDate(d.createdAt)}</td>
          <td style="display:flex; gap:6px;">
            ${d.status==="pending" ? `<button class="btn btn-primary btn-sm" onclick="AdminActions.approveDeposit('${d.id}')">${Icon.check}Approve</button><button class="btn btn-danger btn-sm" onclick="AdminActions.rejectDeposit('${d.id}')">${Icon.x}Reject</button>` : ""}
          </td>
        </tr>`).join("")}
      </tbody></table>` : `<div class="empty-state">${Icon.emptyBox}<div>No deposit requests yet.</div></div>`;
    });
  }
});

const AdminUI = {
  menu(){
    return `<button type="button" onclick="window.AdminUI.toggleSite()">${State.siteLive ? Icon.x : Icon.check}${State.siteLive ? "Turn website off" : "Turn website on"}</button>
      <button type="button" onclick="window.AdminUI.openNotification()">${Icon.send}Send notification</button>
      <button type="button" onclick="window.ThemeMod.toggle()">${document.documentElement.getAttribute("data-theme")==="dark" ? Icon.sun : Icon.moon}${document.documentElement.getAttribute("data-theme")==="dark" ? "Light mode" : "Dark mode"}</button>
      <button type="button" class="danger-item" onclick="window.Auth.logout()">${Icon.logout}Log out</button>`;
  },
  toggleMore(location="top"){
    let menu = location === "sidebar" ? $("adminMoreMenuSidebar") : $("adminMoreMenu");
    if(!menu){
      menu = document.createElement("div");
      menu.id = location === "sidebar" ? "adminMoreMenuSidebar" : "adminMoreMenu"; menu.className = "more-menu";
      menu.style.cssText = "position:fixed;top:70px;right:18px;";
      document.body.appendChild(menu);
    }
    const wasHidden = menu.classList.contains("hidden");
    menu.innerHTML = this.menu();
    menu.classList.toggle("hidden", !wasHidden);
    if(wasHidden){
      setTimeout(() => {
        const close = e => { if(!menu.contains(e.target) && !e.target.closest('[onclick*="toggleMore"]')){ menu.classList.add("hidden"); document.removeEventListener("click", close); } };
        document.addEventListener("click", close);
      }, 0);
    }
  },
  async toggleSite(){
    try{
      const next = !State.siteLive;
      const synced = await SiteStatus.setLive(next);
      UI.toast(next ? "Website is live now." : "Website is offline. Purchases and deposits are paused.", synced ? "ok" : "err");
      ["adminMoreMenu","adminMoreMenuSidebar"].forEach(id => { const menu = $(id); if(menu){ menu.innerHTML = this.menu(); menu.classList.remove("hidden"); }});
    }catch(err){ UI.toast("Could not update website status.","err"); }
  },
  openNotification(){
    const root = $("modal-root");
    root.innerHTML = `<div class="modal-overlay" id="sendNotificationOverlay"><div class="modal modal-wide">
      <div class="modal-head"><h3>Send notification</h3><button class="modal-close" onclick="Modals.close()">✕</button></div>
      <p style="color:var(--text-muted);font-size:13px;line-height:1.5;margin-bottom:18px;">This message will appear in every user's notification center. Users who allowed browser notifications will also receive it on their phone/browser.</p>
      <div class="form-msg" id="notificationMsg"></div>
      <div class="field"><label>Title</label><input id="notificationTitle" maxlength="80" placeholder="Important update"></div>
      <div class="field"><label>Message</label><textarea id="notificationMessage" rows="4" maxlength="500" placeholder="Write your message..."></textarea></div>
      <button class="btn btn-primary btn-block" id="sendNotificationBtn">${Icon.send}Send notification</button>
    </div></div>`;
    $("sendNotificationOverlay").addEventListener("click", e => { if(e.target.id === "sendNotificationOverlay") Modals.close(); });
    $("sendNotificationBtn").onclick = async () => {
      const title = $("notificationTitle").value.trim();
      const message = $("notificationMessage").value.trim();
      if(!title || !message){ showMsg("notificationMsg","Title and message are required.","err"); return; }
      const btn = $("sendNotificationBtn"); btn.disabled = true; btn.textContent = "Sending...";
      try{
        const localNotice = { id:"local-" + Date.now(), title, message, createdBy: State.user?.uid || "admin", createdAtText:new Date().toLocaleString() };
        let synced = true;
        try{
          const ref = await addDoc(collection(db,"notifications"), { title, message, createdBy: State.user.uid, createdAt: serverTimestamp() });
          localNotice.id = ref.id;
        }catch(syncErr){
          synced = false;
          console.warn("Could not sync notification to Firebase:", syncErr);
        }
        let cached = [];
        try{ cached = JSON.parse(localStorage.getItem(Notifications.cacheKey()) || "[]"); }catch(e){}
        localStorage.setItem(Notifications.cacheKey(), JSON.stringify([localNotice, ...cached.filter(n => n.id !== localNotice.id)].slice(0,50)));
        Modals.close(); UI.toast("Notification sent to users.");
        if(!synced) UI.toast("Saved on this device; Firebase notification access is not enabled.","err");
      }catch(err){ showMsg("notificationMsg","Could not send notification.","err"); btn.disabled = false; btn.innerHTML = Icon.send + "Send notification"; }
    };
  }
};
window.AdminUI = AdminUI;

window.AdminActions = {
  async toggleBlock(uid, blockedVal){
    await updateDoc(doc(db,"users",uid), { blocked: blockedVal === "true" });
    Pages.adminUsers($("admin-pages"));
  },
  async adjustBalance(uid){
    const input = await UI.ask("Adjust balance", "Amount to add (use a negative number to remove)", "0", "number");
    if(input === null) return;
    const amt = parseFloat(input);
    if(isNaN(amt) || amt === 0){ UI.toast("Enter a non-zero amount.","err"); return; }
    await updateDoc(doc(db,"users",uid), { balance: increment(amt) });
    Pages.adminUsers($("admin-pages"));
  },
  async deleteProduct(id){
    if(!await UI.confirm("Delete product","This product and its stock will be removed.","Delete",true)) return;
    await deleteDoc(doc(db,"products",id));
  },
  async deleteMethod(id){
    if(!await UI.confirm("Delete payment method","Users will no longer see this method on the deposit page.","Delete",true)) return;
    await deleteDoc(doc(db,"paymentMethods",id));
  },
  async approveDeposit(id){
    try{
      await runTransaction(db, async (tx) => {
        const dRef = doc(db,"deposits",id);
        const dSnap = await tx.get(dRef);
        if(!dSnap.exists() || dSnap.data().status !== "pending") throw new Error("Already processed.");
        const dep = dSnap.data();
        tx.update(dRef, { status: "approved" });
        tx.update(doc(db,"users",dep.uid), { balance: increment(Number(dep.amount)) });
      });
      Pages.adminDeposits($("admin-pages"));
    }catch(err){ UI.toast(err.message || "Could not approve.","err"); }
  },
  async rejectDeposit(id){
    await updateDoc(doc(db,"deposits",id), { status: "rejected" });
    Pages.adminDeposits($("admin-pages"));
  }
};

/* ---------- init ---------- */
ThemeMod.init();
Router.go("landing");
