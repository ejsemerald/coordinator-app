import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-app.js';
import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, signOut, updatePassword, reauthenticateWithCredential, EmailAuthProvider } from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-auth.js';
import { initializeFirestore, getFirestore, persistentLocalCache, persistentMultipleTabManager, doc, getDoc, setDoc, deleteDoc, collection, getDocs, query, where, writeBatch, serverTimestamp } from 'https://www.gstatic.com/firebasejs/10.14.1/firebase-firestore.js';

const app = initializeApp(window.FBCONFIG);
const auth = getAuth(app);
let db;
try { db = initializeFirestore(app, { localCache: persistentLocalCache({ tabManager: persistentMultipleTabManager() }) }); }
catch (e) { db = getFirestore(app); }

const DOMAIN = window.APP_CFG.emailDomain;
const emailOf = u => String(u).trim().toLowerCase() + '@' + DOMAIN;
const userOf = e => String(e || '').split('@')[0];

const FB = {
  emailOf, userOf,
  validUser: u => /^[a-z0-9._-]{3,30}$/.test(String(u).trim().toLowerCase()),
  onAuth: cb => onAuthStateChanged(auth, cb),
  current: () => auth.currentUser,
  signIn: (u, p) => signInWithEmailAndPassword(auth, emailOf(u), p),
  async signUp(u, p, name) {
    const c = await createUserWithEmailAndPassword(auth, emailOf(u), p);
    await FB.createPending(c.user, name);
    return c.user;
  },
  createPending: (user, name) => setDoc(doc(db, 'users', user.uid), { name: String(name || '').trim(), username: userOf(user.email), role: 'pending', isAdmin: false, createdAt: serverTimestamp() }),
  signOut: () => signOut(auth),
  async getProfile(uid) { const s = await getDoc(doc(db, 'users', uid)); return s.exists() ? s.data() : null; },
  async changePassword(oldp, newp) {
    const u = auth.currentUser;
    await reauthenticateWithCredential(u, EmailAuthProvider.credential(u.email, oldp));
    await updatePassword(u, newp);
  },
  async listUsers() { const s = await getDocs(collection(db, 'users')); return s.docs.map(d => Object.assign({ uid: d.id }, d.data())); },
  async listRoster() { const s = await getDocs(collection(db, 'roster')); return s.docs.map(d => Object.assign({ id: d.id }, d.data())); },
  async importRoster(people) {
    for (let i = 0; i < people.length; i += 200) {
      const b = writeBatch(db);
      people.slice(i, i + 200).forEach(p => b.set(doc(db, 'roster', p.id), p, { merge: true }));
      await b.commit();
    }
  },
  /* userPatch: بيتكتب على users/{uid}. rosterId/rosterPatch: بيتكتب على roster/{id} (اختياري) */
  async approve(uid, userPatch, rosterId, rosterPatch) {
    const b = writeBatch(db);
    b.set(doc(db, 'users', uid), userPatch, { merge: true });
    if (rosterId) b.set(doc(db, 'roster', rosterId), rosterPatch, { merge: true });
    await b.commit();
  },
  removeUserDoc: uid => deleteDoc(doc(db, 'users', uid)),
  /* مزامنة الزيارات والتقارير الشهرية */
  putDoc: (c, id, data) => setDoc(doc(db, c, id), data),
  delDoc: (c, id) => deleteDoc(doc(db, c, id)),
  async queryDocs(c, filters) {
    const q = (filters && filters.length) ? query(collection(db, c), ...filters.map(f => where(f[0], f[1], f[2]))) : collection(db, c);
    const s = await getDocs(q);
    return s.docs.map(d => Object.assign({ id: d.id }, d.data()));
  }
};
window.FB = FB;
window.__fbResolve && window.__fbResolve(FB);
