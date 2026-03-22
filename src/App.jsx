import { useState, useEffect } from 'react'
import { auth, db, googleProvider } from './firebase.js'
import { signInWithPopup, signOut, onAuthStateChanged } from 'firebase/auth'
import {
  collection, addDoc, getDocs, getDoc, doc,
  query, where, serverTimestamp
} from 'firebase/firestore'
import * as XLSX from 'xlsx'

// ─── Admin emails ─────────────────────────────────────────────────────────────
const ADMIN_EMAILS = ['issac@sjps.kh.edu.tw']
const isAdmin = (user) => user && ADMIN_EMAILS.includes(user.email)

// ─── Hash Router ──────────────────────────────────────────────────────────────
function getHash() { return window.location.hash.replace('#', '') || '/' }
function useHash() {
  const [hash, setHash] = useState(getHash)
  useEffect(() => {
    const h = () => setHash(getHash())
    window.addEventListener('hashchange', h)
    return () => window.removeEventListener('hashchange', h)
  }, [])
  return hash
}
function navigate(path) { window.location.hash = path }

// ─── CSS ──────────────────────────────────────────────────────────────────────
const css = `
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+TC:wght@300;400;500;700&family=DM+Mono:wght@400;500&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
:root{
  --bg:#f0ede8;--surface:#faf9f7;--border:#e2ddd8;
  --ink:#1a1714;--ink2:#6b6560;--accent:#2d6a4f;--accent2:#52b788;
  --danger:#c1440e;--warn:#d4a017;--admin:#5b4fcf;
  --radius:12px;--shadow:0 2px 12px rgba(0,0,0,.08)
}
body{font-family:'Noto Sans TC',sans-serif;background:var(--bg);color:var(--ink)}
.app{min-height:100vh}
.login-page{min-height:100vh;display:flex;align-items:center;justify-content:center;
  background:linear-gradient(135deg,#1a1714 0%,#2d4a3e 50%,#1a1714 100%);position:relative}
.login-page::before{content:'';position:absolute;inset:0;
  background:radial-gradient(ellipse at 30% 50%,rgba(82,183,136,.15) 0%,transparent 60%),
             radial-gradient(ellipse at 70% 20%,rgba(212,160,23,.1) 0%,transparent 50%)}
.login-card{background:rgba(250,249,247,.97);border-radius:20px;padding:48px;width:420px;
  position:relative;box-shadow:0 24px 80px rgba(0,0,0,.4)}
.login-logo{font-size:13px;font-weight:500;color:var(--ink2);letter-spacing:.2em;text-transform:uppercase;margin-bottom:32px}
.login-title{font-size:28px;font-weight:700;margin-bottom:8px;line-height:1.2}
.login-sub{font-size:14px;color:var(--ink2);margin-bottom:36px}
.google-btn{width:100%;display:flex;align-items:center;justify-content:center;gap:12px;
  padding:14px 20px;background:white;border:1.5px solid var(--border);border-radius:10px;
  font-size:15px;font-weight:500;cursor:pointer;transition:all .2s;font-family:'Noto Sans TC',sans-serif;color:var(--ink)}
.google-btn:hover{border-color:var(--accent);box-shadow:0 4px 16px rgba(45,106,79,.15);transform:translateY(-1px)}
.google-btn:disabled{opacity:.6;cursor:not-allowed;transform:none}
.login-features{margin-top:32px;display:flex;flex-direction:column;gap:10px}
.feat-item{display:flex;align-items:center;gap:10px;font-size:13px;color:var(--ink2)}
.feat-dot{width:6px;height:6px;border-radius:50%;background:var(--accent2);flex-shrink:0}
.layout{display:flex;min-height:100vh}
.sidebar{width:240px;background:#1a1714;flex-shrink:0;display:flex;flex-direction:column;position:fixed;height:100vh;z-index:10}
.sidebar-brand{padding:24px 20px;border-bottom:1px solid rgba(255,255,255,.08)}
.brand-name{font-size:16px;font-weight:700;color:white}
.brand-sub{font-size:11px;color:rgba(255,255,255,.4);margin-top:2px;letter-spacing:.1em;text-transform:uppercase}
.admin-badge{display:inline-block;background:var(--admin);color:white;font-size:10px;font-weight:700;
  padding:2px 7px;border-radius:10px;letter-spacing:.05em;margin-top:4px}
.sidebar-nav{flex:1;padding:16px 12px;display:flex;flex-direction:column;gap:4px}
.nav-item{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:8px;
  font-size:14px;color:rgba(255,255,255,.6);cursor:pointer;transition:all .15s}
.nav-item:hover{background:rgba(255,255,255,.06);color:white}
.nav-item.active{background:var(--accent);color:white}
.nav-item.admin-nav.active{background:var(--admin)}
.nav-icon{font-size:16px;width:20px;text-align:center}
.sidebar-user{padding:16px 20px;border-top:1px solid rgba(255,255,255,.08);display:flex;align-items:center;gap:10px}
.user-avatar{width:32px;height:32px;border-radius:50%;background:var(--accent);
  display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:white;flex-shrink:0}
.user-avatar.admin-avatar{background:var(--admin)}
.user-name{font-size:13px;color:rgba(255,255,255,.8);font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:110px}
.logout-btn{margin-left:auto;font-size:12px;color:rgba(255,255,255,.3);cursor:pointer;padding:4px;border-radius:4px;flex-shrink:0}
.logout-btn:hover{color:rgba(255,255,255,.7)}
.main{flex:1;margin-left:240px;padding:32px;max-width:calc(100% - 240px)}
.page-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:28px}
.page-title{font-size:24px;font-weight:700}
.page-sub{font-size:14px;color:var(--ink2);margin-top:4px}
.btn{display:inline-flex;align-items:center;gap:6px;padding:10px 18px;border-radius:8px;
  font-size:14px;font-weight:500;cursor:pointer;border:none;transition:all .15s;font-family:'Noto Sans TC',sans-serif}
.btn:disabled{opacity:.5;cursor:not-allowed}
.btn-primary{background:var(--accent);color:white}
.btn-primary:hover:not(:disabled){background:#235c42;transform:translateY(-1px);box-shadow:0 4px 12px rgba(45,106,79,.3)}
.btn-secondary{background:white;color:var(--ink);border:1.5px solid var(--border)}
.btn-secondary:hover:not(:disabled){border-color:var(--accent);color:var(--accent)}
.btn-admin{background:var(--admin);color:white}
.btn-admin:hover:not(:disabled){background:#4a3fb5;transform:translateY(-1px)}
.btn-excel{background:#1d6f42;color:white}
.btn-excel:hover:not(:disabled){background:#155232;transform:translateY(-1px);box-shadow:0 4px 12px rgba(29,111,66,.3)}
.btn-sm{padding:6px 12px;font-size:13px}
.card{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:24px}
.tabs{display:flex;gap:4px;background:#f0ede8;border-radius:10px;padding:4px;margin-bottom:20px;width:fit-content}
.tab{padding:8px 18px;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer;transition:all .15s;color:var(--ink2)}
.tab.active{background:white;color:var(--ink);box-shadow:0 1px 4px rgba(0,0,0,.1)}
.form-label{font-size:13px;font-weight:600;color:var(--ink);margin-bottom:8px;display:block}
.form-input{width:100%;padding:10px 14px;border-radius:8px;border:1.5px solid var(--border);
  font-size:14px;font-family:'Noto Sans TC',sans-serif;color:var(--ink);background:white;transition:border-color .15s}
.form-input:focus{outline:none;border-color:var(--accent)}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:16px}
table{width:100%;border-collapse:collapse;font-size:14px}
th{text-align:left;padding:10px 14px;font-size:12px;font-weight:600;color:var(--ink2);
  text-transform:uppercase;letter-spacing:.05em;border-bottom:2px solid var(--border);background:#f5f3ef}
td{padding:12px 14px;border-bottom:1px solid var(--border);color:var(--ink);vertical-align:middle}
tr:last-child td{border-bottom:none}
tr:hover td{background:#f9f8f6}
.score-badge{display:inline-block;padding:3px 10px;border-radius:20px;font-size:13px;font-weight:700}
.score-high{background:#e8f5ee;color:var(--accent)}
.score-mid{background:#fff8e1;color:#a67c00}
.score-low{background:#ffeee8;color:var(--danger)}
.tag{display:inline-block;padding:2px 8px;border-radius:4px;font-size:11px;font-weight:600}
.tag-class{background:#e8eeff;color:#3b5bdb}
.stats-row{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:24px}
.stat-box{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:20px}
.stat-label{font-size:12px;color:var(--ink2);font-weight:500;letter-spacing:.05em;text-transform:uppercase;margin-bottom:8px}
.stat-value{font-size:28px;font-weight:700}
.stat-sub{font-size:12px;color:var(--ink2);margin-top:4px}
.quiz-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:16px}
.quiz-card{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:20px;transition:all .2s}
.quiz-card:hover{border-color:var(--accent2);box-shadow:0 4px 20px rgba(45,106,79,.12);transform:translateY(-2px)}
.quiz-tag-pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;background:#e8f5ee;color:var(--accent);margin-bottom:10px}
.quiz-name{font-size:16px;font-weight:700;margin-bottom:6px}
.quiz-meta{font-size:12px;color:var(--ink2);display:flex;gap:12px;flex-wrap:wrap}
.quiz-url{margin-top:12px;background:#f5f3ef;border-radius:6px;padding:8px 10px;
  font-family:'DM Mono',monospace;font-size:11px;color:var(--ink2);display:flex;align-items:center;justify-content:space-between;gap:8px}
.copy-btn{font-size:11px;color:var(--accent);cursor:pointer;font-weight:600;flex-shrink:0}
.new-quiz-card{background:transparent;border-radius:var(--radius);border:2px dashed var(--border);
  padding:20px;cursor:pointer;transition:all .2s;display:flex;align-items:center;justify-content:center;
  gap:8px;color:var(--ink2);font-size:14px;font-weight:500;min-height:140px}
.new-quiz-card:hover{border-color:var(--accent);color:var(--accent);background:#f0f9f4}
.q-editor{border:1.5px solid var(--border);border-radius:10px;padding:18px;margin-bottom:12px;background:white}
.q-editor-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
.q-num-label{font-size:11px;font-weight:700;color:var(--ink2);text-transform:uppercase;letter-spacing:.1em}
.options-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px;margin-bottom:12px}
.option-row{display:flex;align-items:center;gap:8px}
.opt-label{font-size:12px;font-weight:700;color:var(--ink2);width:18px;flex-shrink:0}
.hint-toggle{font-size:12px;padding:5px 12px;border-radius:6px;border:1.5px solid var(--border);
  background:white;cursor:pointer;color:var(--ink2);font-family:'Noto Sans TC',sans-serif;transition:all .15s;margin-right:6px}
.hint-toggle:hover{border-color:var(--warn);color:var(--warn)}
.hint-area{background:#fffbf0;border:1.5px solid #f0e0a0;border-radius:8px;padding:12px;margin-top:10px}
.hint-area-title{font-size:11px;font-weight:700;color:#a67c00;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px}
.expl-area{background:#f0f9f4;border:1.5px solid #b7e4c7;border-radius:8px;padding:12px;margin-top:8px}
.expl-area-title{font-size:11px;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px}
.upload-zone{border:2px dashed var(--border);border-radius:10px;padding:32px;text-align:center;cursor:pointer;transition:all .2s;background:white}
.upload-zone:hover{border-color:var(--accent);background:#f0f9f4}
.paste-zone{width:100%;min-height:160px;padding:14px;border-radius:10px;border:1.5px solid var(--border);font-size:13px;font-family:'DM Mono',monospace;color:var(--ink);background:white;resize:vertical;transition:border-color .15s;line-height:1.5}
.paste-zone:focus{outline:none;border-color:var(--accent)}
.paste-zone::placeholder{color:var(--ink2);font-family:'Noto Sans TC',sans-serif;font-size:13px}
.parse-preview{background:#f0f9f4;border:1.5px solid #b7e4c7;border-radius:10px;padding:14px;margin-top:12px}
.parse-preview-title{font-size:12px;font-weight:700;color:var(--accent);margin-bottom:10px}
.preview-row{font-size:12px;color:var(--ink);padding:6px 0;border-bottom:1px solid #d8f0e4;display:flex;gap:8px;align-items:flex-start}
.preview-row:last-child{border-bottom:none}
.preview-qnum{font-weight:700;color:var(--accent);flex-shrink:0;width:40px}
.preview-qtext{flex:1}
.q-analytics{display:flex;flex-direction:column;gap:12px}
.q-row{background:#f9f8f6;border-radius:10px;padding:14px 16px}
.q-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px}
.q-text-sm{font-size:13px;font-weight:600;flex:1}
.q-rate{font-size:13px;font-weight:700}
.q-rate.good{color:var(--accent)}
.q-rate.bad{color:var(--danger)}
.q-bar-bg{height:8px;background:var(--border);border-radius:4px;overflow:hidden}
.q-bar{height:100%;border-radius:4px;transition:width .6s}
.bar-green{background:var(--accent2)}
.bar-red{background:#ff7c5c}
.q-opts-row{display:flex;gap:6px;margin-top:8px;flex-wrap:wrap}
.q-opt-chip{font-size:11px;padding:3px 8px;border-radius:4px;background:white;border:1px solid var(--border);color:var(--ink2)}
.q-opt-chip.correct{background:#e8f5ee;border-color:var(--accent2);color:var(--accent);font-weight:600}
.answer-dots{display:flex;gap:4px}
.dot{width:16px;height:16px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:white}
.dot-correct{background:var(--accent2)}
.dot-wrong{background:var(--danger)}
.student-page{min-height:100vh;background:var(--bg)}
.student-topbar{background:#1a1714;padding:14px 20px;display:flex;align-items:center;justify-content:space-between}
.student-topbar-title{color:white;font-size:15px;font-weight:700}
.student-topbar-sub{color:rgba(255,255,255,.5);font-size:12px}
.student-body{max-width:640px;margin:0 auto;padding:24px 16px}
.info-card{background:white;border-radius:var(--radius);padding:24px;box-shadow:var(--shadow);margin-bottom:20px}
.info-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:12px;margin-bottom:20px}
.sq-card{background:white;border-radius:var(--radius);padding:20px;box-shadow:var(--shadow);margin-bottom:16px}
.sq-card.correct-card{border-left:4px solid var(--accent2)}
.sq-card.wrong-card{border-left:4px solid var(--danger)}
.sq-num{font-size:11px;font-weight:700;color:var(--ink2);letter-spacing:.1em;text-transform:uppercase;margin-bottom:8px;display:flex;align-items:center;justify-content:space-between}
.sq-text{font-size:15px;font-weight:600;margin-bottom:14px;line-height:1.6}
.sq-opts{display:flex;flex-direction:column;gap:8px}
.sq-opt{display:flex;align-items:center;gap:12px;padding:11px 14px;border-radius:8px;border:2px solid var(--border);cursor:pointer;transition:all .15s}
.sq-opt:hover:not(.revealed){border-color:var(--accent2);background:#f0f9f4}
.sq-opt.selected{border-color:var(--accent);background:#e8f5ee}
.sq-opt.correct-reveal{border-color:var(--accent2);background:#e8f5ee}
.sq-opt.wrong-reveal{border-color:var(--danger);background:#ffeee8;opacity:.75}
.opt-circle{width:26px;height:26px;border-radius:50%;border:2px solid var(--border);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;flex-shrink:0;transition:all .15s;color:var(--ink2)}
.sq-opt.selected .opt-circle{border-color:var(--accent);background:var(--accent);color:white}
.sq-opt.correct-reveal .opt-circle{border-color:var(--accent2);background:var(--accent2);color:white}
.sq-opt.wrong-reveal .opt-circle{border-color:var(--danger);background:var(--danger);color:white}
.opt-text{font-size:14px;color:var(--ink);flex:1}
.hint-btn{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border-radius:6px;
  border:1.5px solid #f0e0a0;background:#fffbf0;color:#a67c00;font-size:12px;font-weight:600;cursor:pointer;font-family:'Noto Sans TC',sans-serif}
.hint-bubble{background:#fffbf0;border:1.5px solid #f0e0a0;border-radius:8px;padding:10px 14px;font-size:13px;color:#7a5c00;margin-top:10px;line-height:1.6}
.hint-bubble-title{font-size:11px;font-weight:700;color:#a67c00;margin-bottom:4px}
.expl-bubble{background:#f0f9f4;border:1.5px solid #b7e4c7;border-radius:8px;padding:12px 14px;margin-top:12px}
.expl-bubble-title{font-size:11px;font-weight:700;color:var(--accent);margin-bottom:6px}
.expl-bubble-text{font-size:13px;color:#1f4e39;line-height:1.7}
.wrong-expl-bubble{background:#ffeee8;border:1.5px solid #fcd5c5;border-radius:8px;padding:12px 14px;margin-top:12px}
.wrong-expl-title{font-size:11px;font-weight:700;color:var(--danger);margin-bottom:6px}
.wrong-expl-text{font-size:13px;color:#7a2010;line-height:1.7}
.result-card{background:white;border-radius:20px;padding:36px;box-shadow:var(--shadow);text-align:center;margin-bottom:24px}
.result-score{font-size:72px;font-weight:700;color:var(--accent);line-height:1;margin:16px 0 4px}
.result-label{font-size:14px;color:var(--ink2)}
.result-breakdown{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-top:20px}
.rb-item{background:var(--bg);border-radius:10px;padding:12px}
.rb-val{font-size:20px;font-weight:700}
.rb-lab{font-size:11px;color:var(--ink2);margin-top:2px}
.result-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700}
.result-badge.correct{background:#e8f5ee;color:var(--accent)}
.result-badge.wrong{background:#ffeee8;color:var(--danger)}
.toast{position:fixed;bottom:24px;right:24px;background:#1a1714;color:white;padding:12px 20px;
  border-radius:8px;font-size:14px;z-index:1000;box-shadow:0 4px 20px rgba(0,0,0,.3);animation:slideUp .3s ease}
@keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:translateY(0);opacity:1}}
.loading{display:flex;align-items:center;justify-content:center;min-height:200px;font-size:14px;color:var(--ink2);gap:10px}
.spinner{width:20px;height:20px;border:2px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.empty-state{text-align:center;padding:60px 20px;color:var(--ink2)}
.empty-icon{font-size:48px;margin-bottom:16px}
.progress-bar-wrap{background:white;border-radius:10px;padding:12px 16px;margin-bottom:16px;box-shadow:var(--shadow);display:flex;align-items:center;gap:12px}
.progress-track{flex:1;height:6px;background:var(--border);border-radius:3px;overflow:hidden}
.progress-fill{height:100%;background:var(--accent2);border-radius:3px;transition:width .3s}
.admin-header{background:linear-gradient(135deg,#3730a3,#5b4fcf);border-radius:var(--radius);padding:24px;margin-bottom:24px;color:white}
.admin-header-title{font-size:20px;font-weight:700;margin-bottom:4px}
.admin-header-sub{font-size:13px;opacity:.75}
.teacher-row{display:flex;align-items:center;justify-content:space-between;padding:14px 0;border-bottom:1px solid var(--border)}
.teacher-row:last-child{border-bottom:none}
.teacher-info{display:flex;align-items:center;gap:12px}
.teacher-avatar{width:36px;height:36px;border-radius:50%;background:var(--accent);display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:700;color:white;flex-shrink:0}
.teacher-name{font-size:14px;font-weight:600}
.teacher-email{font-size:12px;color:var(--ink2)}
.teacher-stats{display:flex;gap:16px}
.t-stat{text-align:center}
.t-stat-val{font-size:18px;font-weight:700;color:var(--accent)}
.t-stat-lab{font-size:11px;color:var(--ink2)}
@media(max-width:768px){
  .sidebar{display:none}
  .main{margin-left:0;max-width:100%;padding:16px}
  .stats-row{grid-template-columns:1fr 1fr}
  .form-row{grid-template-columns:1fr}
  .options-grid{grid-template-columns:1fr}
  .info-grid{grid-template-columns:1fr}
}
`

// ─── Helpers ──────────────────────────────────────────────────────────────────
function formatDate(ts) {
  if (!ts) return ''
  const d = ts.toDate ? ts.toDate() : new Date(ts)
  return d.toLocaleDateString('zh-TW')
}
function formatTime(ts) {
  if (!ts) return ''
  const d = ts.toDate ? ts.toDate() : new Date(ts)
  return d.toLocaleTimeString('zh-TW', { hour: '2-digit', minute: '2-digit' })
}
function scoreBadgeClass(s) { return s >= 80 ? 'score-high' : s >= 60 ? 'score-mid' : 'score-low' }

function Toast({ msg }) {
  return <div className="toast">✓ {msg}</div>
}

// ─── Parse rows helper (shared by Excel & Paste) ─────────────────────────────
// Forces every cell to string so "3/8" stays "3/8" and never becomes 0.375
function parseQuestionRows(rows) {
  return rows.slice(1).filter(r => r[0]).map(r => {
    const correctLetter = String(r[5] || 'A').trim().toUpperCase()
    const correctIdx = ['A','B','C','D'].indexOf(correctLetter)
    return {
      text:        String(r[0] || '').trim(),
      options:     [String(r[1]||''), String(r[2]||''), String(r[3]||''), String(r[4]||'')],
      correct:     correctIdx >= 0 ? correctIdx : 0,
      points:      parseInt(String(r[6]).replace(/[^0-9]/g,'')) || 10,
      hint:        String(r[7] || '').trim(),
      explanation: String(r[8] || '').trim(),
      showHint:    !!r[7],
      showExpl:    !!r[8],
    }
  })
}

// Parse tab-separated text (from AI output or copy-paste)
function parsePasteText(text) {
  const lines = text.trim().split('\n').filter(l => l.trim())
  if (lines.length < 2) return null
  // Support both tab and multiple-spaces as delimiter
  const rows = lines.map(l => l.split('\t'))
  return parseQuestionRows(rows)
}
function exportToExcel(quiz, responses) {
  if (!responses.length) { alert('目前沒有作答紀錄'); return }

  // Sheet 1: 成績總表
  const scoreData = responses.map((r, i) => {
    const row = {
      '序號': i + 1,
      '班級': r.class,
      '座號': r.seat,
      '姓名': r.name,
      '總分': r.score,
      '作答時間': formatTime(r.submittedAt),
    }
    quiz.questions.forEach((q, qi) => {
      const ans = r.answers?.[qi]
      row[`第${qi + 1}題`] = ans !== undefined && ans >= 0 ? ['A','B','C','D'][ans] : '-'
      row[`第${qi + 1}題是否正確`] = ans === q.correct ? '✓' : '✗'
    })
    return row
  })

  // Sheet 2: 題目分析
  const analysisData = quiz.questions.map((q, qi) => {
    const total = responses.length
    const correct = responses.filter(r => r.answers?.[qi] === q.correct).length
    const optCounts = [0,1,2,3].map(oi => responses.filter(r => r.answers?.[qi] === oi).length)
    return {
      '題號': `第${qi + 1}題`,
      '題目': q.text,
      '正確答案': ['A','B','C','D'][q.correct],
      '配分': q.points,
      '答對人數': correct,
      '答對率': total ? `${Math.round(correct / total * 100)}%` : '-',
      '選A人數': optCounts[0],
      '選B人數': optCounts[1],
      '選C人數': optCounts[2],
      '選D人數': optCounts[3],
    }
  })

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(scoreData), '學生成績')
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysisData), '題目分析')

  const filename = `${quiz.title}_成績_${new Date().toLocaleDateString('zh-TW').replace(/\//g, '-')}.xlsx`
  XLSX.writeFile(wb, filename)
}

// ─── Login ────────────────────────────────────────────────────────────────────
function LoginPage() {
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const handleLogin = async () => {
    setLoading(true); setError('')
    try { await signInWithPopup(auth, googleProvider) }
    catch { setError('登入失敗，請再試一次'); setLoading(false) }
  }
  return (
    <div className="login-page">
      <div className="login-card">
        <div className="login-logo">📋 QuizFlow</div>
        <div className="login-title">線上測驗<br/>管理平台</div>
        <div className="login-sub">為老師設計的出題與成績分析工具</div>
        <button className="google-btn" onClick={handleLogin} disabled={loading}>
          {loading ? <span className="spinner"/> : (
            <svg width="18" height="18" viewBox="0 0 18 18">
              <path fill="#4285F4" d="M16.51 8H8.98v3h4.3c-.18 1-.74 1.48-1.6 2.04v2.01h2.6a7.8 7.8 0 0 0 2.38-5.88c0-.57-.05-.66-.15-1.18z"/>
              <path fill="#34A853" d="M8.98 17c2.16 0 3.97-.72 5.3-1.94l-2.6-2a4.8 4.8 0 0 1-7.18-2.54H1.83v2.07A8 8 0 0 0 8.98 17z"/>
              <path fill="#FBBC05" d="M4.5 10.52a4.8 4.8 0 0 1 0-3.04V5.41H1.83a8 8 0 0 0 0 7.18l2.67-2.07z"/>
              <path fill="#EA4335" d="M8.98 4.18c1.17 0 2.23.4 3.06 1.2l2.3-2.3A8 8 0 0 0 1.83 5.4L4.5 7.49a4.77 4.77 0 0 1 4.48-3.3z"/>
            </svg>
          )}
          {loading ? '登入中...' : '使用 Google 帳號登入'}
        </button>
        {error && <div style={{marginTop:12,fontSize:13,color:'var(--danger)',textAlign:'center'}}>{error}</div>}
        <div className="login-features">
          {['每位老師獨立空間，資料安全分離','支援 Excel 批量匯入與成績匯出','學生無需登入，輸入班級座號即可作答'].map(f => (
            <div key={f} className="feat-item"><div className="feat-dot"/>{f}</div>
          ))}
        </div>
      </div>
    </div>
  )
}

// ─── Sidebar ──────────────────────────────────────────────────────────────────
function Sidebar({ hash, user, onLogout }) {
  const admin = isAdmin(user)
  const initial = user?.displayName?.[0] || user?.email?.[0] || '?'
  const teacherNav = [
    { path:'/',         icon:'◉', label:'我的測驗' },
    { path:'/create',   icon:'＋', label:'新增測驗' },
  ]
  const adminNav = [
    { path:'/admin',    icon:'⊞', label:'管理後台', adminClass:true },
  ]
  return (
    <div className="sidebar">
      <div className="sidebar-brand">
        <div className="brand-name">📋 QuizFlow</div>
        <div className="brand-sub">Teacher Dashboard</div>
        {admin && <div className="admin-badge">管理員</div>}
      </div>
      <div className="sidebar-nav">
        {teacherNav.map(n => (
          <div key={n.path} className={`nav-item ${hash===n.path?'active':''}`} onClick={() => navigate(n.path)}>
            <span className="nav-icon">{n.icon}</span>{n.label}
          </div>
        ))}
        {admin && (
          <>
            <div style={{height:1,background:'rgba(255,255,255,.08)',margin:'8px 4px'}}/>
            {adminNav.map(n => (
              <div key={n.path} className={`nav-item ${n.adminClass?'admin-nav':''} ${hash===n.path?'active':''}`}
                onClick={() => navigate(n.path)}>
                <span className="nav-icon">{n.icon}</span>{n.label}
              </div>
            ))}
          </>
        )}
      </div>
      <div className="sidebar-user">
        <div className={`user-avatar ${admin?'admin-avatar':''}`}>{initial.toUpperCase()}</div>
        <div style={{overflow:'hidden'}}>
          <div className="user-name">{user?.displayName || user?.email}</div>
        </div>
        <div className="logout-btn" onClick={onLogout}>登出</div>
      </div>
    </div>
  )
}

// ─── Admin Panel ──────────────────────────────────────────────────────────────
function AdminPanel() {
  const [teachers, setTeachers] = useState([])
  const [quizzes, setQuizzes] = useState([])
  const [responses, setResponses] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    async function load() {
      const [qSnap, rSnap] = await Promise.all([
        getDocs(collection(db, 'quizzes')),
        getDocs(collection(db, 'responses')),
      ])
      const allQuizzes = qSnap.docs.map(d => ({ id: d.id, ...d.data() }))
      const allResponses = rSnap.docs.map(d => ({ id: d.id, ...d.data() }))

      // 彙整老師資料
      const teacherMap = {}
      allQuizzes.forEach(q => {
        const id = q.teacherId
        if (!teacherMap[id]) teacherMap[id] = { id, name: q.teacherName || '未知', email: '', quizCount: 0, responseCount: 0 }
        teacherMap[id].quizCount++
        teacherMap[id].responseCount += allResponses.filter(r => r.quizId === q.id).length
      })
      setTeachers(Object.values(teacherMap).sort((a, b) => b.quizCount - a.quizCount))
      setQuizzes(allQuizzes)
      setResponses(allResponses)
      setLoading(false)
    }
    load()
  }, [])

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>

  const totalResponses = responses.length
  const avgResponsesPerQuiz = quizzes.length ? Math.round(totalResponses / quizzes.length) : 0

  return (
    <div>
      <div className="admin-header">
        <div className="admin-header-title">⊞ 系統管理後台</div>
        <div className="admin-header-sub">查看所有老師的使用狀況與全站統計</div>
      </div>

      <div className="stats-row">
        {[
          { label:'老師帳號數', value: teachers.length, sub:'位老師', color:'var(--admin)' },
          { label:'測驗總數',   value: quizzes.length,  sub:'份測驗' },
          { label:'作答總次數', value: totalResponses,   sub:'筆紀錄', color:'var(--accent)' },
          { label:'平均每份',   value: avgResponsesPerQuiz, sub:'人作答' },
        ].map(s => (
          <div key={s.label} className="stat-box">
            <div className="stat-label">{s.label}</div>
            <div className="stat-value" style={{color: s.color || 'var(--ink)'}}>{s.value}</div>
            <div className="stat-sub">{s.sub}</div>
          </div>
        ))}
      </div>

      <div className="card" style={{marginBottom:20}}>
        <div style={{fontSize:16,fontWeight:700,marginBottom:20}}>👩‍🏫 老師帳號統計</div>
        {teachers.length === 0 ? (
          <div className="empty-state" style={{padding:40}}>
            <div className="empty-icon">📭</div><div>目前還沒有老師登入使用</div>
          </div>
        ) : teachers.map((t, i) => (
          <div key={t.id} className="teacher-row">
            <div className="teacher-info">
              <div className="teacher-avatar">{(t.name?.[0] || '?').toUpperCase()}</div>
              <div>
                <div className="teacher-name">{t.name}</div>
                <div className="teacher-email">{t.id}</div>
              </div>
            </div>
            <div className="teacher-stats">
              <div className="t-stat">
                <div className="t-stat-val">{t.quizCount}</div>
                <div className="t-stat-lab">測驗數</div>
              </div>
              <div className="t-stat">
                <div className="t-stat-val">{t.responseCount}</div>
                <div className="t-stat-lab">作答數</div>
              </div>
            </div>
          </div>
        ))}
      </div>

      <div className="card">
        <div style={{fontSize:16,fontWeight:700,marginBottom:16}}>📋 所有測驗列表</div>
        {quizzes.length === 0 ? (
          <div className="empty-state" style={{padding:40}}><div>目前沒有測驗</div></div>
        ) : (
          <div style={{overflowX:'auto'}}>
            <table>
              <thead><tr><th>測驗名稱</th><th>老師</th><th>科目</th><th>題數</th><th>作答人數</th><th>建立日期</th></tr></thead>
              <tbody>
                {quizzes.sort((a,b) => (b.createdAt?.seconds||0) - (a.createdAt?.seconds||0)).map(q => (
                  <tr key={q.id}>
                    <td style={{fontWeight:600}}>{q.title}</td>
                    <td style={{fontSize:13,color:'var(--ink2)'}}>{q.teacherName || '-'}</td>
                    <td><span className="quiz-tag-pill" style={{marginBottom:0}}>{q.subject || '-'}</span></td>
                    <td>{q.questions?.length || 0} 題</td>
                    <td>
                      <span className="score-badge score-high">
                        {responses.filter(r => r.quizId === q.id).length} 人
                      </span>
                    </td>
                    <td style={{fontSize:13,color:'var(--ink2)'}}>{formatDate(q.createdAt)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  )
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({ user }) {
  const [quizzes, setQuizzes] = useState([])
  const [responseCounts, setResponseCounts] = useState({})
  const [loading, setLoading] = useState(true)
  const [toast, setToast] = useState('')

  useEffect(() => {
    async function load() {
      try {
        const q = query(collection(db,'quizzes'), where('teacherId','==',user.uid))
        const snap = await getDocs(q)
        const list = snap.docs.map(d => ({ id:d.id, ...d.data() }))
          .sort((a,b) => (b.createdAt?.seconds||0) - (a.createdAt?.seconds||0))
        setQuizzes(list)
        // load response counts
        const counts = {}
        await Promise.all(list.map(async quiz => {
          const rs = await getDocs(query(collection(db,'responses'), where('quizId','==',quiz.id)))
          counts[quiz.id] = rs.size
        }))
        setResponseCounts(counts)
      } catch(e) { console.error(e) }
      setLoading(false)
    }
    load()
  }, [user.uid])

  const copyLink = (id) => {
    const url = `${window.location.origin}${window.location.pathname}#/s/${id}`
    navigator.clipboard.writeText(url)
    setToast('連結已複製！'); setTimeout(() => setToast(''), 2000)
  }

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>

  return (
    <div>
      <div className="page-header">
        <div><div className="page-title">我的測驗</div><div className="page-sub">管理你建立的所有測驗卷</div></div>
        <button className="btn btn-primary" onClick={() => navigate('/create')}>＋ 新增測驗</button>
      </div>
      {quizzes.length === 0 ? (
        <div className="empty-state">
          <div className="empty-icon">📋</div>
          <div style={{fontSize:16,fontWeight:600,marginBottom:8}}>還沒有測驗</div>
          <button className="btn btn-primary" onClick={() => navigate('/create')}>＋ 新增測驗</button>
        </div>
      ) : (
        <div className="quiz-grid">
          {quizzes.map(q => (
            <div key={q.id} className="quiz-card">
              <div className="quiz-tag-pill">{q.subject || '未分類'}</div>
              <div className="quiz-name">{q.title}</div>
              <div className="quiz-meta">
                <span>📝 {q.questions?.length || 0} 題</span>
                <span>👥 {responseCounts[q.id] || 0} 人作答</span>
                <span>🗓 {formatDate(q.createdAt)}</span>
              </div>
              <div className="quiz-url">
                <span style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>...#/s/{q.id}</span>
                <span className="copy-btn" onClick={() => copyLink(q.id)}>複製連結</span>
              </div>
              <div style={{display:'flex',gap:8,marginTop:12}}>
                <button className="btn btn-secondary btn-sm" style={{flex:1}} onClick={() => navigate(`/results/${q.id}`)}>📊 成績</button>
                <button className="btn btn-secondary btn-sm" style={{flex:1}} onClick={() => navigate(`/analytics/${q.id}`)}>📈 分析</button>
              </div>
            </div>
          ))}
          <div className="new-quiz-card" onClick={() => navigate('/create')}>
            <span style={{fontSize:20}}>＋</span>建立新測驗
          </div>
        </div>
      )}
      {toast && <Toast msg={toast}/>}
    </div>
  )
}

// ─── Create Quiz ──────────────────────────────────────────────────────────────
const emptyQ = () => ({ text:'', options:['','','',''], correct:0, points:10, hint:'', explanation:'', showHint:false, showExpl:false })

function CreateQuiz({ user }) {
  const [mode, setMode] = useState('manual')
  const [title, setTitle] = useState('')
  const [subject, setSubject] = useState('')
  const [allowHint, setAllowHint] = useState(true)
  const [showExplAfter, setShowExplAfter] = useState(true)
  const [showCorrect, setShowCorrect] = useState(true)
  const [questions, setQuestions] = useState([emptyQ()])
  const [saving, setSaving] = useState(false)
  const [createdId, setCreatedId] = useState(null)
  const [toast, setToast] = useState('')

  const [pasteText, setPasteText] = useState('')
  const [pastePreview, setPastePreview] = useState(null)

  const handlePasteChange = (text) => {
    setPasteText(text)
    if (text.trim()) {
      const parsed = parsePasteText(text)
      setPastePreview(parsed)
    } else {
      setPastePreview(null)
    }
  }

  const handlePasteImport = () => {
    if (!pastePreview || pastePreview.length === 0) { alert('無法解析，請確認格式'); return }
    setQuestions(pastePreview)
    setPasteText(''); setPastePreview(null); setMode('manual')
    setToast(`成功匯入 ${pastePreview.length} 道題目`); setTimeout(()=>setToast(''),2000)
  }

  const addQ = () => setQuestions(p => [...p, emptyQ()])
  const removeQ = i => setQuestions(p => p.filter((_,idx) => idx!==i))
  const updateQ = (i,f,v) => setQuestions(p => { const q=[...p]; q[i]={...q[i],[f]:v}; return q })
  const updateOpt = (qi,oi,v) => setQuestions(p => { const q=[...p]; q[qi].options[oi]=v; return q })
  const toggleField = (i,f) => setQuestions(p => { const q=[...p]; q[i]={...q[i],[f]:!q[i][f]}; return q })

  const handleExcel = (e) => {
    const file = e.target.files[0]; if (!file) return
    const reader = new FileReader()
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, { type:'binary' })
        const ws = wb.Sheets[wb.SheetNames[0]]
        // raw:false → forces all cells to formatted string, prevents 3/8 → 0.375
        const rows = XLSX.utils.sheet_to_json(ws, { header:1, raw:false, defval:'' })
        const parsed = parseQuestionRows(rows)
        if (parsed.length > 0) {
          setQuestions(parsed); setMode('manual')
          setToast(`成功匯入 ${parsed.length} 道題目`); setTimeout(()=>setToast(''),2000)
        }
      } catch { alert('Excel 格式錯誤，請確認欄位格式') }
    }
    reader.readAsBinaryString(file)
  }

  const handleSave = async () => {
    if (!title.trim()) { alert('請輸入測驗名稱'); return }
    if (questions.some(q=>!q.text.trim())) { alert('請填寫所有題目'); return }
    setSaving(true)
    try {
      const clean = questions.map(({ showHint, showExpl, ...q }) => q)
      const docRef = await addDoc(collection(db,'quizzes'), {
        teacherId: user.uid, teacherName: user.displayName || user.email,
        title: title.trim(), subject: subject.trim()||'未分類',
        questions: clean, settings: { allowHint, showExplAfter, showCorrect },
        createdAt: serverTimestamp(),
      })
      setCreatedId(docRef.id)
    } catch(e) { alert('儲存失敗，請確認 Firebase 設定'); console.error(e) }
    setSaving(false)
  }

  if (createdId) {
    const url = `${window.location.origin}${window.location.pathname}#/s/${createdId}`
    return (
      <div style={{maxWidth:480,margin:'60px auto',textAlign:'center'}}>
        <div style={{fontSize:48,marginBottom:16}}>🎉</div>
        <div style={{fontSize:22,fontWeight:700,marginBottom:8}}>測驗已建立！</div>
        <div style={{fontSize:14,color:'var(--ink2)',marginBottom:24}}>將以下連結傳給學生即可開始作答</div>
        <div className="card" style={{padding:20,marginBottom:16}}>
          <div style={{fontFamily:"'DM Mono',monospace",fontSize:13,background:'#f5f3ef',padding:'10px 14px',borderRadius:8,marginBottom:12,wordBreak:'break-all'}}>{url}</div>
          <button className="btn btn-primary" style={{width:'100%'}} onClick={() => navigator.clipboard.writeText(url)}>複製學生作答連結</button>
        </div>
        <div style={{display:'flex',gap:10,justifyContent:'center'}}>
          <button className="btn btn-secondary" onClick={() => navigate('/')}>回到首頁</button>
          <button className="btn btn-secondary" onClick={() => { setCreatedId(null);setTitle('');setSubject('');setQuestions([emptyQ()]) }}>再建一份</button>
        </div>
      </div>
    )
  }

  return (
    <div style={{maxWidth:700}}>
      <div className="page-header">
        <div><div className="page-title">新增測驗</div><div className="page-sub">手動輸入或上傳 Excel，可為每題設定提示與解析</div></div>
      </div>
      <div className="tabs">
        <div className={`tab ${mode==='manual'?'active':''}`} onClick={()=>setMode('manual')}>✎ 手動輸入</div>
        <div className={`tab ${mode==='excel'?'active':''}`} onClick={()=>setMode('excel')}>📊 上傳 Excel</div>
        <div className={`tab ${mode==='paste'?'active':''}`} onClick={()=>setMode('paste')}>📋 貼上文字</div>
      </div>
      {mode==='excel' && (
        <div className="card" style={{marginBottom:20}}>
          <label className="upload-zone" style={{display:'block'}}>
            <div style={{fontSize:36,marginBottom:12}}>📊</div>
            <div style={{fontSize:14,fontWeight:600,marginBottom:6}}>點擊選擇 Excel 檔案</div>
            <div style={{fontSize:12,color:'var(--ink2)'}}>支援 .xlsx / .xls · 分數如 3/8 會正確保留</div>
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleExcel}/>
          </label>
          <div style={{marginTop:16}}>
            <div style={{fontSize:12,fontWeight:600,marginBottom:8,color:'var(--ink2)'}}>Excel 欄位格式（第一列為標題）：</div>
            <div style={{overflowX:'auto'}}>
              <table style={{fontSize:11}}>
                <thead><tr>
                  {['題目','選A','選B','選C','選D','正確(A/B/C/D)','配分','提示(選填)','解析(選填)'].map(h=>(
                    <th key={h} style={{padding:'5px 8px',background:'#f0ede8',border:'1px solid var(--border)',whiteSpace:'nowrap'}}>{h}</th>
                  ))}
                </tr></thead>
                <tbody><tr>
                  {['算算看3/8+...','1/4','3/8','1/2','5/8','B','10','分母相同時...','分母不變分子相加'].map((c,i)=>(
                    <td key={i} style={{padding:'5px 8px',border:'1px solid var(--border)',color:'var(--ink2)',whiteSpace:'nowrap'}}>{c}</td>
                  ))}
                </tr></tbody>
              </table>
            </div>
          </div>
        </div>
      )}
      {mode==='paste' && (
        <div className="card" style={{marginBottom:20}}>
          <div style={{fontSize:14,fontWeight:700,marginBottom:4}}>📋 貼上 AI 生成的題目文字</div>
          <div style={{fontSize:12,color:'var(--ink2)',marginBottom:12}}>
            請確認格式為 <b>Tab 分隔</b>，第一列是標題列，欄位順序：<br/>
            <span style={{fontFamily:"'DM Mono',monospace",fontSize:11,background:'#f5f3ef',padding:'2px 6px',borderRadius:4,display:'inline-block',marginTop:4}}>
              題目 + 選A + 選B + 選C + 選D + 正確答案(A/B/C/D) + 配分 + 提示 + 解析
            </span>
          </div>
          <textarea
            className="paste-zone"
            placeholder={"直接把 AI 給你的表格文字貼在這裡...\n\n範例（Tab 分隔，第一列標題）：\n題目\t選A\t選B\t選C\t選D\t正確(A/B/C/D)\t配分\t提示(選填)\t解析(選填)\n小明吃了3/8個披薩...\t1/8\t3/8\t5/8\t8/3\tB\t10\t分母是份數\t全部8份吃了3份"}
            value={pasteText}
            onChange={e => handlePasteChange(e.target.value)}
          />
          {pastePreview && pastePreview.length > 0 && (
            <div className="parse-preview">
              <div className="parse-preview-title">✅ 預覽：成功解析 {pastePreview.length} 道題目</div>
              {pastePreview.slice(0,3).map((q,i) => (
                <div key={i} className="preview-row">
                  <div className="preview-qnum">第{i+1}題</div>
                  <div className="preview-qtext">
                    <div style={{fontWeight:600,marginBottom:3}}>{q.text.slice(0,40)}{q.text.length>40?'...':''}</div>
                    <div style={{color:'var(--ink2)',fontSize:11}}>
                      {q.options.map((o,oi) => (
                        <span key={oi} style={{marginRight:8,color:oi===q.correct?'var(--accent)':'var(--ink2)',fontWeight:oi===q.correct?700:400}}>
                          {['A','B','C','D'][oi]}.{o}
                        </span>
                      ))}
                    </div>
                  </div>
                </div>
              ))}
              {pastePreview.length > 3 && <div style={{fontSize:12,color:'var(--ink2)',paddingTop:8}}>...還有 {pastePreview.length-3} 道題目</div>}
              <button className="btn btn-primary" style={{marginTop:12,width:'100%'}} onClick={handlePasteImport}>
                匯入這 {pastePreview.length} 道題目 →
              </button>
            </div>
          )}
          {pasteText && (!pastePreview || pastePreview.length === 0) && (
            <div style={{marginTop:12,padding:'10px 14px',background:'#ffeee8',borderRadius:8,fontSize:13,color:'var(--danger)'}}>
              ⚠️ 無法解析，請確認第一列是標題列，且欄位之間用 Tab 分隔
            </div>
          )}
        </div>
      )}
      {/* Main form card */}
      <div className="card">
        <div style={{marginBottom:20}}>
          <div className="form-row">
            <div><label className="form-label">測驗名稱 *</label>
              <input className="form-input" placeholder="例：第三章自然科測驗" value={title} onChange={e=>setTitle(e.target.value)}/></div>
            <div><label className="form-label">科目</label>
              <input className="form-input" placeholder="例：自然科學" value={subject} onChange={e=>setSubject(e.target.value)}/></div>
          </div>
        </div>
        <div style={{background:'#f5f3ef',borderRadius:8,padding:'10px 14px',marginBottom:20,display:'flex',gap:20,flexWrap:'wrap'}}>
          <span style={{fontSize:13,fontWeight:600}}>設定</span>
          {[[allowHint,setAllowHint,'允許學生查看提示'],[showExplAfter,setShowExplAfter,'提交後顯示解析'],[showCorrect,setShowCorrect,'顯示正確答案']].map(([v,s,l])=>(
            <label key={l} style={{display:'flex',alignItems:'center',gap:6,fontSize:13,cursor:'pointer'}}>
              <input type="checkbox" checked={v} onChange={e=>s(e.target.checked)} style={{accentColor:'var(--accent)'}}/>{l}
            </label>
          ))}
        </div>
        <div style={{borderTop:'1px solid var(--border)',paddingTop:20}}>
          <div style={{fontSize:14,fontWeight:700,marginBottom:16}}>📝 題目設定</div>
          {questions.map((q,qi)=>(
            <div key={qi} className="q-editor">
              <div className="q-editor-header">
                <span className="q-num-label">第 {qi+1} 題</span>
                <div style={{display:'flex',alignItems:'center',gap:12}}>
                  <span style={{fontSize:12,color:'var(--ink2)'}}>配分</span>
                  <input className="form-input" type="number" min="1" style={{width:60,padding:'4px 8px',fontSize:13}}
                    value={q.points} onChange={e=>updateQ(qi,'points',parseInt(e.target.value)||1)}/>
                  {questions.length>1 && <span style={{fontSize:12,color:'var(--danger)',cursor:'pointer'}} onClick={()=>removeQ(qi)}>✕ 刪除</span>}
                </div>
              </div>
              <input className="form-input" placeholder="輸入題目..." value={q.text}
                onChange={e=>updateQ(qi,'text',e.target.value)} style={{marginBottom:10}}/>
              <div className="options-grid">
                {['A','B','C','D'].map((lbl,oi)=>(
                  <div key={oi} className="option-row">
                    <span className="opt-label">{lbl}.</span>
                    <input className="form-input" style={{fontSize:13,padding:'7px 10px'}} placeholder={`選項 ${lbl}`}
                      value={q.options[oi]} onChange={e=>updateOpt(qi,oi,e.target.value)}/>
                    <input type="radio" style={{accentColor:'var(--accent)',width:16,height:16,cursor:'pointer'}}
                      name={`c_${qi}`} checked={q.correct===oi} onChange={()=>updateQ(qi,'correct',oi)} id={`r_${qi}_${oi}`}/>
                    <label style={{fontSize:11,color:'var(--ink2)',cursor:'pointer'}} htmlFor={`r_${qi}_${oi}`}>正確</label>
                  </div>
                ))}
              </div>
              <div style={{display:'flex',flexWrap:'wrap',gap:6}}>
                <button className="hint-toggle" onClick={()=>toggleField(qi,'showHint')}>{q.showHint?'▼':'▶'} 💡 提示</button>
                <button className="hint-toggle" style={{borderColor:'#b7e4c7',color:'var(--accent)'}} onClick={()=>toggleField(qi,'showExpl')}>{q.showExpl?'▼':'▶'} 📖 解析</button>
              </div>
              {q.showHint && <div className="hint-area" style={{marginTop:10}}>
                <div className="hint-area-title">💡 提示內容（學生作答中可主動查看）</div>
                <textarea className="form-input" style={{resize:'vertical',minHeight:56,fontSize:13}}
                  placeholder="例：想想含有葉綠素的細胞器是哪一個？" value={q.hint} onChange={e=>updateQ(qi,'hint',e.target.value)}/>
              </div>}
              {q.showExpl && <div className="expl-area" style={{marginTop:8}}>
                <div className="expl-area-title">📖 解析內容（提交後才顯示給學生）</div>
                <textarea className="form-input" style={{resize:'vertical',minHeight:72,fontSize:13}}
                  placeholder="例：葉綠體含有葉綠素，是光合作用的場所..." value={q.explanation} onChange={e=>updateQ(qi,'explanation',e.target.value)}/>
              </div>}
            </div>
          ))}
          <button className="btn btn-secondary btn-sm" onClick={addQ}>＋ 新增題目</button>
        </div>
        <div style={{marginTop:24,display:'flex',justifyContent:'flex-end',gap:10}}>
          <button className="btn btn-secondary" onClick={()=>navigate('/')}>取消</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>{saving?'儲存中...':'建立測驗並產生連結 →'}</button>
        </div>
      </div>
      {toast && <Toast msg={toast}/>}
    </div>
  )
}

// ─── Results ──────────────────────────────────────────────────────────────────
function Results({ quizId }) {
  const [quiz, setQuiz] = useState(null)
  const [responses, setResponses] = useState([])
  const [loading, setLoading] = useState(true)
  const [filter, setFilter] = useState('all')

  useEffect(() => {
    async function load() {
      const qSnap = await getDoc(doc(db,'quizzes',quizId))
      if (!qSnap.exists()) { setLoading(false); return }
      setQuiz({ id:qSnap.id, ...qSnap.data() })
      const rSnap = await getDocs(query(collection(db,'responses'),where('quizId','==',quizId)))
      setResponses(rSnap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(b.submittedAt?.seconds||0)-(a.submittedAt?.seconds||0)))
      setLoading(false)
    }
    load()
  }, [quizId])

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>
  if (!quiz) return <div className="empty-state"><div>找不到這份測驗</div></div>

  const classes = ['all',...new Set(responses.map(r=>r.class).filter(Boolean))]
  const filtered = filter==='all' ? responses : responses.filter(r=>r.class===filter)
  const avg = filtered.length ? Math.round(filtered.reduce((s,r)=>s+r.score,0)/filtered.length) : 0

  return (
    <div>
      <div className="page-header">
        <div>
          <div style={{fontSize:13,color:'var(--ink2)',cursor:'pointer',marginBottom:6}} onClick={()=>navigate('/')}>← 返回</div>
          <div className="page-title">{quiz.title}</div>
          <div className="page-sub">成績查詢</div>
        </div>
        <button className="btn btn-excel" onClick={()=>exportToExcel(quiz, responses)}>
          📥 匯出 Excel
        </button>
      </div>
      <div className="stats-row">
        {[
          {label:'作答人數',value:responses.length,sub:'位學生'},
          {label:'平均分數',value:avg,sub:'分',color:'var(--accent)'},
          {label:'最高分',value:responses.length?Math.max(...responses.map(r=>r.score)):'-',sub:'分'},
          {label:'及格率',value:responses.length?Math.round(responses.filter(r=>r.score>=60).length/responses.length*100)+'%':'-',sub:'≥60分'},
        ].map(s=>(
          <div key={s.label} className="stat-box">
            <div className="stat-label">{s.label}</div>
            <div className="stat-value" style={{color:s.color||'var(--ink)'}}>{s.value}</div>
            <div className="stat-sub">{s.sub}</div>
          </div>
        ))}
      </div>
      <div className="card">
        <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:16}}>
          <span style={{fontSize:13,fontWeight:600}}>篩選班級：</span>
          {classes.map(c=>(
            <button key={c} className={`btn btn-sm ${filter===c?'btn-primary':'btn-secondary'}`} onClick={()=>setFilter(c)}>
              {c==='all'?'全部':`${c}班`}
            </button>
          ))}
        </div>
        {filtered.length===0 ? (
          <div className="empty-state" style={{padding:40}}><div className="empty-icon">📭</div><div>目前沒有作答紀錄</div></div>
        ) : (
          <div style={{overflowX:'auto'}}>
            <table>
              <thead><tr><th>班級</th><th>座號</th><th>姓名</th><th>分數</th><th>答題狀況</th><th>作答時間</th></tr></thead>
              <tbody>
                {filtered.map(r=>(
                  <tr key={r.id}>
                    <td><span className="tag tag-class">{r.class}班</span></td>
                    <td style={{fontFamily:"'DM Mono',monospace"}}>{r.seat}</td>
                    <td style={{fontWeight:500}}>{r.name}</td>
                    <td><span className={`score-badge ${scoreBadgeClass(r.score)}`}>{r.score}分</span></td>
                    <td><div className="answer-dots">
                      {quiz.questions.map((q,i)=>(
                        <div key={i} className={`dot ${r.answers?.[i]===q.correct?'dot-correct':'dot-wrong'}`} title={`第${i+1}題`}>
                          {r.answers?.[i]===q.correct?'✓':'✗'}
                        </div>
                      ))}
                    </div></td>
                    <td style={{color:'var(--ink2)',fontSize:13}}>{formatTime(r.submittedAt)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  )
}

// ─── Analytics ────────────────────────────────────────────────────────────────
function Analytics({ quizId }) {
  const [quiz, setQuiz] = useState(null)
  const [responses, setResponses] = useState([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    async function load() {
      const qSnap = await getDoc(doc(db,'quizzes',quizId))
      if (!qSnap.exists()) { setLoading(false); return }
      setQuiz({ id:qSnap.id, ...qSnap.data() })
      const rSnap = await getDocs(query(collection(db,'responses'),where('quizId','==',quizId)))
      setResponses(rSnap.docs.map(d=>({id:d.id,...d.data()})))
      setLoading(false)
    }
    load()
  }, [quizId])

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>
  if (!quiz) return <div className="empty-state"><div>找不到這份測驗</div></div>

  const n = responses.length
  const stats = quiz.questions.map((q,qi) => {
    const correct = responses.filter(r=>r.answers?.[qi]===q.correct).length
    const rate = n ? Math.round(correct/n*100) : 0
    const optCounts = [0,1,2,3].map(oi=>responses.filter(r=>r.answers?.[qi]===oi).length)
    return { q, rate, correct, optCounts }
  })

  return (
    <div>
      <div className="page-header">
        <div>
          <div style={{fontSize:13,color:'var(--ink2)',cursor:'pointer',marginBottom:6}} onClick={()=>navigate('/')}>← 返回</div>
          <div className="page-title">{quiz.title}</div>
          <div className="page-sub">答題分析 — {n} 人作答</div>
        </div>
        <button className="btn btn-excel" onClick={()=>exportToExcel(quiz, responses)}>📥 匯出 Excel</button>
      </div>
      <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16,marginBottom:24}}>
        <div className="card">
          <div style={{fontSize:14,fontWeight:700,marginBottom:14}}>各題答對率</div>
          <div className="q-analytics">
            {stats.map((s,i)=>(
              <div key={i} className="q-row">
                <div className="q-head">
                  <div className="q-text-sm">第{i+1}題：{s.q.text.length>20?s.q.text.slice(0,20)+'...':s.q.text}</div>
                  <div className={`q-rate ${s.rate>=70?'good':'bad'}`}>{s.rate}%</div>
                </div>
                <div className="q-bar-bg"><div className={`q-bar ${s.rate>=70?'bar-green':'bar-red'}`} style={{width:`${s.rate}%`}}/></div>
              </div>
            ))}
          </div>
        </div>
        <div className="card">
          <div style={{fontSize:14,fontWeight:700,marginBottom:14}}>選項分佈</div>
          <div className="q-analytics">
            {stats.map((s,i)=>(
              <div key={i} className="q-row">
                <div className="q-head" style={{marginBottom:8}}>
                  <div className="q-text-sm">第{i+1}題</div>
                  <div style={{fontSize:12,color:'var(--ink2)'}}>{s.correct}/{n} 答對</div>
                </div>
                <div className="q-opts-row">
                  {s.q.options.map((opt,oi)=>(
                    <div key={oi} className={`q-opt-chip ${oi===s.q.correct?'correct':''}`}>
                      {['A','B','C','D'][oi]}. {opt.slice(0,5)}({s.optCounts[oi]}人)
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
      <div className="card">
        <div style={{fontSize:14,fontWeight:700,marginBottom:14}}>⚠️ 最難題目排名</div>
        <table>
          <thead><tr><th>排名</th><th>題目</th><th>答對率</th><th>最多人答錯的選項</th></tr></thead>
          <tbody>
            {[...stats].sort((a,b)=>a.rate-b.rate).map((s,i)=>{
              const wrongIdx = s.optCounts.reduce((mi,v,oi)=>oi!==s.q.correct&&v>s.optCounts[mi]?oi:mi, s.q.correct===0?1:0)
              return (
                <tr key={i}>
                  <td style={{fontWeight:700,color:i<2?'var(--danger)':'var(--ink2)'}}>#{i+1}</td>
                  <td>{s.q.text}</td>
                  <td><span className={`score-badge ${scoreBadgeClass(s.rate)}`}>{s.rate}%</span></td>
                  <td style={{color:'var(--danger)',fontSize:13}}>
                    {wrongIdx!==s.q.correct?`${['A','B','C','D'][wrongIdx]}. ${s.q.options[wrongIdx]}（${s.optCounts[wrongIdx]}人）`:'-'}
                  </td>
                </tr>
              )
            })}
          </tbody>
        </table>
      </div>
    </div>
  )
}

// ─── Student Quiz ─────────────────────────────────────────────────────────────
function StudentQuiz({ quizId }) {
  const [step, setStep] = useState('info')
  const [info, setInfo] = useState({ name:'', class:'', seat:'' })
  const [quiz, setQuiz] = useState(null)
  const [loading, setLoading] = useState(true)
  const [answers, setAnswers] = useState({})
  const [shownHints, setShownHints] = useState({})
  const [score, setScore] = useState(0)
  const [saving, setSaving] = useState(false)

  useEffect(() => {
    async function load() {
      try {
        const snap = await getDoc(doc(db,'quizzes',quizId))
        if (snap.exists()) setQuiz({ id:snap.id, ...snap.data() })
      } catch(e) { console.error(e) }
      setLoading(false)
    }
    load()
  }, [quizId])

  const toggleHint = i => setShownHints(h=>({...h,[i]:!h[i]}))
  const totalPoints = quiz?.questions?.reduce((s,q)=>s+q.points,0)||0
  const allAnswered = quiz?.questions?.every((_,i)=>answers[i]!==undefined)

  const handleSubmit = async () => {
    setSaving(true)
    let s=0; quiz.questions.forEach((q,i)=>{ if(answers[i]===q.correct) s+=q.points })
    setScore(s)
    try {
      await addDoc(collection(db,'responses'),{
        quizId, name:info.name, class:info.class, seat:info.seat,
        score:s, answers:quiz.questions.map((_,i)=>answers[i]??-1),
        submittedAt: serverTimestamp(),
      })
    } catch(e) { console.error('儲存失敗',e) }
    setSaving(false); setStep('result')
  }

  if (loading) return (
    <div className="student-page">
      <div className="student-topbar"><div className="student-topbar-title">📋 QuizFlow</div></div>
      <div className="loading"><div className="spinner"/>載入題目中...</div>
    </div>
  )
  if (!quiz) return (
    <div className="student-page">
      <div className="student-topbar"><div className="student-topbar-title">📋 QuizFlow</div></div>
      <div className="empty-state"><div className="empty-icon">❌</div><div>找不到這份測驗，請確認連結是否正確</div></div>
    </div>
  )

  const s = quiz.settings||{}

  if (step==='result') {
    const correctCount = quiz.questions.filter((_,i)=>answers[i]===quiz.questions[i].correct).length
    return (
      <div className="student-page">
        <div className="student-topbar"><div><div className="student-topbar-title">📋 {quiz.title}</div></div></div>
        <div className="student-body">
          <div className="result-card">
            <div style={{fontSize:16,fontWeight:700}}>✅ 作答完成！</div>
            <div style={{fontSize:14,color:'var(--ink2)',marginTop:6}}>{info.name} 同學，你的成績如下</div>
            <div className="result-score">{score}</div>
            <div className="result-label">滿分 {totalPoints} 分</div>
            <div className="result-breakdown">
              <div className="rb-item"><div className="rb-val" style={{color:'var(--accent)'}}>{correctCount}</div><div className="rb-lab">答對</div></div>
              <div className="rb-item"><div className="rb-val" style={{color:'var(--danger)'}}>{quiz.questions.length-correctCount}</div><div className="rb-lab">答錯</div></div>
              <div className="rb-item"><div className="rb-val">{totalPoints?Math.round(score/totalPoints*100):0}%</div><div className="rb-lab">正確率</div></div>
            </div>
          </div>
          {(s.showCorrect!==false||s.showExplAfter!==false) && (
            <div>
              <div style={{fontSize:16,fontWeight:700,marginBottom:16}}>📋 詳細解析</div>
              {quiz.questions.map((q,qi)=>{
                const isCorrect = answers[qi]===q.correct
                return (
                  <div key={qi} className={`sq-card ${isCorrect?'correct-card':'wrong-card'}`}>
                    <div className="sq-num">
                      <span>第 {qi+1} 題 · {q.points}分</span>
                      <span className={`result-badge ${isCorrect?'correct':'wrong'}`}>{isCorrect?'✓ 答對':'✗ 答錯'}</span>
                    </div>
                    <div className="sq-text">{q.text}</div>
                    <div className="sq-opts" style={{pointerEvents:'none'}}>
                      {q.options.map((opt,oi)=>{
                        let cls='revealed'
                        if(s.showCorrect!==false&&oi===q.correct) cls+=' correct-reveal'
                        else if(oi===answers[qi]) cls+=' wrong-reveal'
                        return (
                          <div key={oi} className={`sq-opt ${cls}`}>
                            <div className="opt-circle">{['A','B','C','D'][oi]}</div>
                            <div className="opt-text">{opt}
                              {s.showCorrect!==false&&oi===q.correct&&<span style={{fontSize:11,color:'var(--accent)',marginLeft:8,fontWeight:700}}>← 正確答案</span>}
                              {oi===answers[qi]&&oi!==q.correct&&<span style={{fontSize:11,color:'var(--danger)',marginLeft:8}}>← 你的選擇</span>}
                            </div>
                          </div>
                        )
                      })}
                    </div>
                    {s.showExplAfter!==false&&q.explanation&&(
                      isCorrect
                        ? <div className="expl-bubble"><div className="expl-bubble-title">📖 解析</div><div className="expl-bubble-text">{q.explanation}</div></div>
                        : <div className="wrong-expl-bubble"><div className="wrong-expl-title">📖 看看哪裡答錯了</div><div className="wrong-expl-text">{q.explanation}</div></div>
                    )}
                  </div>
                )
              })}
            </div>
          )}
        </div>
      </div>
    )
  }

  return (
    <div className="student-page">
      <div className="student-topbar">
        <div>
          <div className="student-topbar-title">{quiz.title}</div>
          <div className="student-topbar-sub">{quiz.questions.length} 題 · 共 {totalPoints} 分</div>
        </div>
      </div>
      <div className="student-body">
        {step==='info' && (
          <div className="info-card">
            <div style={{fontSize:16,fontWeight:700,marginBottom:16}}>請先填寫基本資料</div>
            <div className="info-grid">
              {[['班級','例：701','class'],['座號','例：01','seat'],['姓名','你的名字','name']].map(([label,ph,field])=>(
                <div key={field}>
                  <label className="form-label" style={{fontSize:12}}>{label}</label>
                  <input className="form-input" placeholder={ph} value={info[field]} onChange={e=>setInfo({...info,[field]:e.target.value})}/>
                </div>
              ))}
            </div>
            <button onClick={()=>setStep('quiz')} disabled={!info.name||!info.class||!info.seat}
              style={{width:'100%',background:(!info.name||!info.class||!info.seat)?'#ccc':'var(--accent)',color:'white',
                padding:'13px 20px',borderRadius:8,border:'none',fontSize:15,fontWeight:700,
                cursor:(!info.name||!info.class||!info.seat)?'not-allowed':'pointer',fontFamily:"'Noto Sans TC',sans-serif"}}>
              開始作答 →
            </button>
          </div>
        )}
        {step==='quiz' && (
          <>
            <div className="progress-bar-wrap">
              <span style={{fontSize:13,color:'var(--ink2)',flexShrink:0}}>已作答 {Object.keys(answers).length}/{quiz.questions.length}</span>
              <div className="progress-track"><div className="progress-fill" style={{width:`${Object.keys(answers).length/quiz.questions.length*100}%`}}/></div>
            </div>
            {quiz.questions.map((q,qi)=>(
              <div key={qi} className="sq-card">
                <div className="sq-num">
                  <span>第 {qi+1} 題 · {q.points}分</span>
                  {s.allowHint!==false&&q.hint&&(
                    <button className="hint-btn" onClick={()=>toggleHint(qi)}>💡 {shownHints[qi]?'收起提示':'查看提示'}</button>
                  )}
                </div>
                <div className="sq-text">{q.text}</div>
                {q.hint&&shownHints[qi]&&<div className="hint-bubble"><div className="hint-bubble-title">💡 提示</div>{q.hint}</div>}
                <div className="sq-opts" style={{marginTop:q.hint&&shownHints[qi]?12:0}}>
                  {q.options.map((opt,oi)=>(
                    <div key={oi} className={`sq-opt ${answers[qi]===oi?'selected':''}`} onClick={()=>setAnswers({...answers,[qi]:oi})}>
                      <div className="opt-circle">{['A','B','C','D'][oi]}</div>
                      <div className="opt-text">{opt}</div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
            <button onClick={handleSubmit} disabled={!allAnswered||saving}
              style={{width:'100%',background:allAnswered&&!saving?'var(--accent)':'#ccc',color:'white',
                padding:'15px 20px',borderRadius:10,border:'none',fontSize:16,fontWeight:700,
                cursor:allAnswered&&!saving?'pointer':'not-allowed',fontFamily:"'Noto Sans TC',sans-serif",marginBottom:32}}>
              {saving?'儲存中...':allAnswered?'提交答案 →':`還有 ${quiz.questions.length-Object.keys(answers).length} 題未作答`}
            </button>
          </>
        )}
      </div>
    </div>
  )
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(undefined)
  const hash = useHash()

  useEffect(() => { return onAuthStateChanged(auth, u => setUser(u||null)) }, [])

  const isStudentRoute = hash.startsWith('/s/')

  if (isStudentRoute) return (
    <><style>{css}</style><StudentQuiz quizId={hash.replace('/s/','')}/></>
  )

  if (user === undefined) return (
    <><style>{css}</style><div className="loading" style={{minHeight:'100vh'}}><div className="spinner"/>載入中...</div></>
  )

  if (!user) return (
    <><style>{css}</style><LoginPage/></>
  )

  const handleLogout = () => signOut(auth)

  let content
  if (hash.startsWith('/results/'))   content = <Results quizId={hash.replace('/results/','')}/>
  else if (hash.startsWith('/analytics/')) content = <Analytics quizId={hash.replace('/analytics/','')}/>
  else if (hash === '/create')         content = <CreateQuiz user={user}/>
  else if (hash === '/admin' && isAdmin(user)) content = <AdminPanel/>
  else                                 content = <Dashboard user={user}/>

  return (
    <><style>{css}</style>
      <div className="app">
        <div className="layout">
          <Sidebar hash={hash} user={user} onLogout={handleLogout}/>
          <div className="main">{content}</div>
        </div>
      </div>
    </>
  )
}
