import { initializeApp } from "firebase/app";
import { getAnalytics } from "firebase/analytics";
import { getAuth, GoogleAuthProvider } from "firebase/auth"; // 新增載入 Auth 模組

const firebaseConfig = {
  apiKey: "AIzaSyBZWTL8zHzyXThABjgtcSsE-YIz-C1ojYo",
  authDomain: "wkc-system-discipline.firebaseapp.com",
  projectId: "wkc-system-discipline",
  storageBucket: "wkc-system-discipline.firebasestorage.app",
  messagingSenderId: "917375802225",
  appId: "1:917375802225:web:84ba777aea965c491634d0",
  measurementId: "G-BJYEPHSTLS"
};

const app = initializeApp(firebaseConfig);
const analytics = getAnalytics(app);

// 初始化 Auth 與 Google Provider
const auth = getAuth(app);
const googleProvider = new GoogleAuthProvider();

// 將 auth 與 googleProvider 一併匯出
export { app, analytics, auth, googleProvider };