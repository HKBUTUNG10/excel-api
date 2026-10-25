import React, { useState } from 'react';
import { Shield } from 'lucide-react';
import { signInWithPopup } from 'firebase/auth';
import { auth, googleProvider } from './firebase';

export default function Login({ onLogin }) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [errorMsg, setErrorMsg] = useState('');
  const [loading, setLoading] = useState(false);

  // List your admin Google emails here
  const ADMIN_EMAILS = [
    'admin1@wkc.edu.hk',
    'director@wkc.edu.hk'
  ];

  const handleLocalLogin = (e) => {
    e.preventDefault();
    setErrorMsg('');
    if (username === 'admin' && password === 'admin123') {
      onLogin('admin');
    } else if (username === 'user' && password === 'user123') {
      onLogin('user');
    } else {
      setErrorMsg('帳號或密碼錯誤');
    }
  };

  const handleGoogleLogin = async () => {
    setLoading(true);
    setErrorMsg('');
    try {
      const result = await signInWithPopup(auth, googleProvider);
      const email = result.user.email;

      if (!email.endsWith('@wkc.edu.hk')) {
        await auth.signOut();
        setErrorMsg('請使用 @wkc.edu.hk 學校 Google 帳號登入！');
        setLoading(false);
        return;
      }

      if (ADMIN_EMAILS.includes(email)) {
        onLogin('admin');
      } else {
        onLogin('user');
      }
    } catch (error) {
      console.error('Login error:', error);
      if (error.code === 'auth/popup-blocked') {
        setErrorMsg('瀏覽器阻擋了登入視窗，請允許彈出視窗。');
      } else {
        setErrorMsg(`登入失敗，請稍後再試。(${error.code || error.message})`);
      }
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-[#0f172a] flex flex-col items-center justify-center p-4">
      <div className="bg-white rounded-[2rem] p-10 max-w-sm w-full shadow-2xl flex flex-col items-center">
        
        <div className="bg-blue-50 p-4 rounded-full mb-5">
          <Shield className="w-8 h-8 text-blue-600" fill="currentColor" strokeWidth={1} />
        </div>

        <h1 className="text-2xl font-bold text-gray-900 mb-1">系統登入</h1>
        <p className="text-gray-500 text-sm mb-8 font-medium">WKC Data Analytics</p>

        {errorMsg && (
          <div className="mb-6 w-full p-3 bg-red-50 border border-red-200 rounded-lg text-red-600 text-sm text-center">
            {errorMsg}
          </div>
        )}

        <form onSubmit={handleLocalLogin} className="w-full space-y-4 mb-6">
          <div className="text-left">
            <label className="block text-sm font-medium text-slate-600 mb-1">帳號</label>
            <input 
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 bg-gray-50"
              placeholder="輸入 admin 或 user"
            />
          </div>
          <div className="text-left">
            <label className="block text-sm font-medium text-slate-600 mb-1">密碼</label>
            <input 
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              className="w-full px-4 py-3 border border-slate-200 rounded-xl focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 bg-gray-50"
              placeholder="輸入密碼"
            />
          </div>
          
          <button 
            type="submit" 
            className="w-full flex items-center justify-center py-3 px-4 bg-blue-600 hover:bg-blue-700 text-white font-semibold rounded-xl transition-all shadow-sm mt-4"
          >
            登入
          </button>
        </form>

        <div className="w-full flex items-center mb-6">
          <div className="flex-1 border-t border-gray-200"></div>
          <span className="px-3 text-sm text-gray-400">或</span>
          <div className="flex-1 border-t border-gray-200"></div>
        </div>

        <button
          onClick={handleGoogleLogin}
          disabled={loading}
          type="button"
          className="w-full flex items-center justify-center gap-3 py-3 px-4 bg-white border border-gray-300 hover:bg-gray-50 text-gray-700 font-semibold rounded-xl transition-all shadow-sm disabled:opacity-50"
        >
          <img
            src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg"
            alt="Google logo"
            className="w-5 h-5"
          />
          {loading ? '登入中...' : '使用 Google 帳號登入'}
        </button>
      </div>
      
      <p className="text-slate-500 text-xs font-semibold tracking-wider mt-8 uppercase">
        Restricted to wkc.edu.hk accounts only
      </p>
    </div>
  );
}