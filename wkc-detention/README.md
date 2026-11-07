# WKC detention desk

Responsive teacher dashboard and student check-in portal for detention sessions.

## Run locally

```bash
npm install
npm run dev
```

Without Firebase settings the app runs in a clearly labelled local demo mode with a seeded roster. Add these Vite variables to use the remote adapters:

```bash
VITE_FIREBASE_API_KEY=...
VITE_FIREBASE_PROJECT_ID=...
VITE_FIREBASE_DATABASE_ID=(default)
VITE_FIREBASE_COLLECTION=detention
VITE_DETENTION_WEBHOOK_URL=https://script.google.com/macros/s/.../exec
```

The Firestore REST adapter is the source-of-truth boundary for roster, session, and check-in writes when configured. It signs in through Firebase Anonymous Auth before reading or writing Firestore. `VITE_DETENTION_WEBHOOK_URL` is an optional Google Apps Script/webhook boundary for forwarding check-in events. If remote auth or Firestore is unavailable, the UI keeps operating with the local seeded fallback.
