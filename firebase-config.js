// firebase-config.js
// Smart 5G Dashboard — Firebase Configuration
// Deploy as: ES Module loaded via <script type="module"> in index.html

import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged }
  from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore, doc, setDoc, getDoc, updateDoc, collection, query, where, getDocs }
  from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

const firebaseConfig = {
  apiKey:            "AIzaSyDFlb0sO0S6PMzku8LIrMopFcHA5ohiBsg",
  authDomain:        "smartdashboard-ae001.firebaseapp.com",
  projectId:         "smartdashboard-ae001",
  storageBucket:     "smartdashboard-ae001.firebasestorage.app",
  messagingSenderId: "828088207222",
  appId:             "1:828088207222:web:ed27dd97badcaacc8ddc57"
};

const _app  = initializeApp(firebaseConfig);
const _auth = getAuth(_app);
const _db   = getFirestore(_app);
const _googleProvider = new GoogleAuthProvider();

// Expose to global scope so non-module app.js can access
window.firebaseAuth     = _auth;
window.firebaseDb       = _db;
window.googleProvider   = _googleProvider;
window.fbSignInWithPopup    = signInWithPopup;
window.fbSignOut            = signOut;
window.fbOnAuthStateChanged = onAuthStateChanged;
window.fbDoc        = doc;
window.fbSetDoc     = setDoc;
window.fbGetDoc     = getDoc;
window.fbUpdateDoc  = updateDoc;
window.fbCollection = collection;
window.fbQuery      = query;
window.fbWhere      = where;
window.fbGetDocs    = getDocs;
