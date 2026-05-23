// sidebar.js — shared across all pages

const API_BASE = "https://aman-ai-rl6v.onrender.com";

const LANG = localStorage.getItem("amanLang") || "en";
const TOKEN = localStorage.getItem("amanToken");
const USER  = JSON.parse(localStorage.getItem("amanUser") || "null");

const T = {
  en: {
    newChat: "New Chat", home: "Home", settings: "Settings", search: "Search",
    general: "General", recentChats: "Recent Chats", login: "Login",
    logout: "Logout", loggedInAs: "Logged in as",
    welcomeTitle: "Welcome to Aman.ai",
    welcomeDesc: "Upload your SDLC phase documents to receive a professional security analysis report for each phase.",
    hct: "Higher Colleges of Technology",
    tagline: "AI-Driven Software Security Analyst Assistant",
    startChat: "Start a new conversation and ask Aman.ai about your software security, vulnerabilities, or best practices.",
    analyze: "Analyze", downloadReport: "Download Report",
    generateFinal: "Generate Final Comprehensive Security Report",
    downloadFinal: "Download Final Report",
    analyzing: "Analyzing", generating: "Generating...", downloading: "Downloading...",
    conversation: "Conversation",
    convDesc: "Ask Aman.ai anything about software security.",
    msgPlaceholder: "Message Aman.ai...",
    helloMsg: "Hello! I'm Aman.ai. Ask me anything about software security, or attach a file for analysis.",
    settingsTitle: "Settings", settingsDesc: "Manage your preferences for Aman.ai.",
    theme: "Theme", dark: "Dark (Default)", light: "Light",
    language: "Language", english: "English", arabic: "Arabic",
    clearHistory: "Clear Chat History", clearBtn: "Clear All History",
    saveSettings: "Save Settings", backHome: "← Back to Home",
    saved: "✓ Settings saved successfully.",
    searchTitle: "Search", searchDesc: "Search through your previous Aman.ai conversations.",
    searchLabel: "Search query", searchPlaceholder: "Search conversations...",
    searchBtn: "Search",
  },
  ar: {
    newChat: "محادثة جديدة", home: "الرئيسية", settings: "الإعدادات", search: "بحث",
    general: "عام", recentChats: "المحادثات الأخيرة", login: "تسجيل الدخول",
    logout: "تسجيل الخروج", loggedInAs: "مسجل دخول كـ",
    welcomeTitle: "مرحباً بك في Aman.ai",
    welcomeDesc: "ارفع وثائق مراحل دورة حياة تطوير البرمجيات للحصول على تقرير أمني احترافي لكل مرحلة.",
    hct: "كليات التقنية العليا",
    tagline: "مساعد ذكاء اصطناعي لتحليل أمن البرمجيات",
    startChat: "ابدأ محادثة جديدة واسأل Aman.ai عن أمن البرمجيات.",
    analyze: "تحليل", downloadReport: "تحميل التقرير",
    generateFinal: "إنشاء التقرير الأمني الشامل النهائي",
    downloadFinal: "تحميل التقرير النهائي",
    analyzing: "جارٍ التحليل", generating: "جارٍ الإنشاء...", downloading: "جارٍ التحميل...",
    conversation: "المحادثة",
    convDesc: "اسأل Aman.ai عن أمن البرمجيات.",
    msgPlaceholder: "اكتب رسالتك...",
    helloMsg: "مرحباً! أنا Aman.ai. اسألني أي شيء عن أمن البرمجيات.",
    settingsTitle: "الإعدادات", settingsDesc: "إدارة تفضيلاتك في Aman.ai.",
    theme: "المظهر", dark: "داكن (افتراضي)", light: "فاتح",
    language: "اللغة", english: "الإنجليزية", arabic: "العربية",
    clearHistory: "مسح سجل المحادثات", clearBtn: "مسح الكل",
    saveSettings: "حفظ الإعدادات", backHome: "→ العودة للرئيسية",
    saved: "✓ تم حفظ الإعدادات بنجاح.",
    searchTitle: "بحث", searchDesc: "ابحث في محادثاتك السابقة مع Aman.ai.",
    searchLabel: "نص البحث", searchPlaceholder: "ابحث في المحادثات...",
    searchBtn: "بحث",
  }
};

const t = (key) => (T[LANG] || T.en)[key] || key;

function applyLang() {
  if (LANG === "ar") {
    document.documentElement.setAttribute("dir", "rtl");
    document.documentElement.setAttribute("lang", "ar");
    document.body.style.fontFamily = "'Segoe UI', Tahoma, Arial, sans-serif";
  } else {
    document.documentElement.setAttribute("dir", "ltr");
    document.documentElement.setAttribute("lang", "en");
  }
}

function applyTheme() {
  if (localStorage.getItem("amanTheme") === "light") {
    document.body.classList.add("theme-light");
  } else {
    document.body.classList.remove("theme-light");
  }
}

async function logout() {
  if (TOKEN) {
    try {
      await fetch(`${API_BASE}/api/logout`, {
        method: "POST",
        headers: { Authorization: "Bearer " + TOKEN }
      });
    } catch (e) {
      console.warn("Logout request failed:", e);
    }
  }

  localStorage.removeItem("amanToken");
  localStorage.removeItem("amanUser");
  window.location.href = "login.html";
}

async function loadSidebarSessions(containerId, activeSessionId) {
  if (!TOKEN) return;

  try {
    const res = await fetch(`${API_BASE}/api/chat/sessions`, {
      headers: { Authorization: "Bearer " + TOKEN }
    });

    const data = await res.json();
    const container = document.getElementById(containerId);
    if (!container) return;

    container.querySelectorAll(".session-btn").forEach(b => b.remove());

    (data.sessions || []).forEach(s => {
      const btn = document.createElement("button");
      btn.className = "sidebar-btn session-btn" + (s.id === activeSessionId ? " active-session" : "");
      btn.style.cssText = "font-size:12px; padding:8px 10px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;";
      btn.title = s.title;
      btn.textContent = s.title;
      btn.onclick = () => {
        window.location.href = `newchat.html?session=${s.id}`;
      };
      container.appendChild(btn);
    });
  } catch (e) {
    console.warn("Could not load chat sessions:", e);
  }
}

function buildSidebar(options = {}) {
  const activePage = options.activePage || "";
  const sessionsContainerId = options.sessionsContainerId || "sessionsList";

  const authEl = document.getElementById("sidebarAuth");
  if (authEl) {
    if (USER) {
      authEl.innerHTML = `
        <span style="font-size:12px;color:#8b92b8;">${t("loggedInAs")}<br>
        <strong style="color:#e5e7ff;">${USER.email}</strong></span>
        <button class="sidebar-btn" style="margin-top:8px;" onclick="logout()">
          <span class="icon">🚪</span> ${t("logout")}
        </button>`;
    } else {
      authEl.innerHTML = `<a class="login-bottom" href="login.html">${t("login")}</a>`;
    }
  }

  document.querySelectorAll("[data-tkey]").forEach(el => {
    el.textContent = t(el.getAttribute("data-tkey"));
  });

  loadSidebarSessions(sessionsContainerId, options.activeSessionId || null);
}

document.addEventListener("DOMContentLoaded", () => {
  applyLang();
  applyTheme();
});