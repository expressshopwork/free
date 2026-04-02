// ============================================================
// Firebase Configuration — Placeholder
// ============================================================
// To enable Firebase Authentication and Firestore, replace the
// placeholder values below with your actual Firebase project
// credentials from the Firebase Console:
//   https://console.firebase.google.com/
//
// Steps:
//   1. Create a Firebase project
//   2. Enable Authentication (Google Sign-In provider)
//   3. Enable Firestore Database
//   4. Copy your web app's config from Project Settings
//   5. Replace the values below
//   6. Uncomment the initializeApp and export lines
//
// IMPORTANT: Do NOT commit real credentials to source control.
// Use environment variables or a secrets manager in production.
// ============================================================

const FIREBASE_CONFIG = {
  apiKey:            "YOUR_API_KEY",
  authDomain:        "YOUR_PROJECT_ID.firebaseapp.com",
  projectId:         "YOUR_PROJECT_ID",
  storageBucket:     "YOUR_PROJECT_ID.appspot.com",
  messagingSenderId: "YOUR_MESSAGING_SENDER_ID",
  appId:             "YOUR_APP_ID"
};

// Uncomment the lines below once you have real Firebase credentials:
//
// import { initializeApp } from "https://www.gstatic.com/firebasejs/10.x.x/firebase-app.js";
// import { getAuth, GoogleAuthProvider, signInWithPopup, signOut, onAuthStateChanged }
//   from "https://www.gstatic.com/firebasejs/10.x.x/firebase-auth.js";
// import { getFirestore, collection, addDoc, getDocs, query, where, orderBy }
//   from "https://www.gstatic.com/firebasejs/10.x.x/firebase-firestore.js";
//
// const app  = initializeApp(FIREBASE_CONFIG);
// const auth = getAuth(app);
// const db   = getFirestore(app);
// const googleProvider = new GoogleAuthProvider();
//
// Firestore Security Rules (deploy via Firebase CLI):
// rules_version = '2';
// service cloud.firestore {
//   match /databases/{database}/documents {
//     match /transactions/{doc} {
//       allow read, write: if request.auth != null
//         && request.auth.uid == resource.data.userId;   // agents: own records only
//     }
//     match /transactions/{doc} {
//       allow read: if request.auth != null
//         && get(/databases/$(database)/documents/users/$(request.auth.uid)).data.role == 'supervisor'
//         && resource.data.branch == get(/databases/$(database)/documents/users/$(request.auth.uid)).data.branch;
//     }
//     match /kpiSettings/{doc} {
//       allow read:  if request.auth != null;
//       allow write: if request.auth != null
//         && get(/databases/$(database)/documents/users/$(request.auth.uid)).data.role in ['supervisor','admin'];
//     }
//   }
// }
