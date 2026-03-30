import { useState, useEffect, useMemo } from 'react'
import { auth, db, googleProvider } from './firebase.js'
import { signInWithPopup, signOut, onAuthStateChanged } from 'firebase/auth'
import { collection, addDoc, getDocs, getDoc, doc, deleteDoc, setDoc, query, where, serverTimestamp } from 'firebase/firestore'
import * as XLSX from 'xlsx'
import katex from 'katex'
import 'katex/dist/katex.min.css'
import QRCode from 'qrcode'

// ─── Constants ────────────────────────────────────────────────────────────────
const ADMIN_EMAILS = ['issac@sjps.kh.edu.tw']
const isAdmin = (u) => u && ADMIN_EMAILS.includes(u.email)
const ABCD = ['A','B','C','D']
const padSeat = (s) => String(parseInt(String(s).trim()) || 0).padStart(2, '0')
const emptyQ = () => ({text:'',options:['','','',''],correct:0,points:10,hint:'',explanation:'',showHint:false,showExpl:false})

// ─── Hash Router ──────────────────────────────────────────────────────────────
function getHash() { return window.location.hash.replace('#','') || '/' }
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

// ─── Math Rendering ───────────────────────────────────────────────────────────
function renderMath(text) {
  if (!text || typeof text !== 'string') return text || ''
  let r = text
  r = r.replace(/\$\$([^$]+?)\$\$/g, (m, math) => {
    try { return katex.renderToString(math.trim(), { displayMode:true, throwOnError:false }) }
    catch { return m }
  })
  r = r.replace(/\$([^$\n]+?)\$/g, (m, math) => {
    try { return katex.renderToString(math.trim(), { throwOnError:false }) }
    catch { return m }
  })
  return r
}
function MathText({ text, className, style, tag='span' }) {
  const html = renderMath(text || '')
  const Tag = tag
  if (html === (text||'')) return <Tag className={className} style={style}>{text}</Tag>
  return <Tag className={className} style={style} dangerouslySetInnerHTML={{__html:html}}/>
}

// ─── Helpers ──────────────────────────────────────────────────────────────────
const fmtDate = (ts) => { if(!ts) return ''; const d=ts.toDate?ts.toDate():new Date(ts); return d.toLocaleDateString('zh-TW') }
const fmtTime = (ts) => { if(!ts) return ''; const d=ts.toDate?ts.toDate():new Date(ts); return d.toLocaleTimeString('zh-TW',{hour:'2-digit',minute:'2-digit'}) }
const scoreClass = (s) => s>=80?'score-high':s>=60?'score-mid':'score-low'
const errorColor = (rate) => rate>=80?'err-red':rate>=50?'err-orange':rate>=30?'err-yellow':'err-green'

// ─── Quiz Status ──────────────────────────────────────────────────────────────
// status: 'draft' | 'active' | 'ended'
const STATUS = {
  draft:  { label:'草稿',   emoji:'📝', color:'#6b6560', bg:'#f0ede8',  border:'#e2ddd8' },
  active: { label:'派發中', emoji:'🟢', color:'#2d6a4f', bg:'#e8f5ee',  border:'#52b788' },
  ended:  { label:'已結束', emoji:'🔴', color:'#c1440e', bg:'#ffeee8',  border:'#fcd5c5' },
}
function getStatus(q) { return q.status || 'draft' }

function parseQuestionRows(rows) {
  return rows.slice(1).filter(r=>r[0]).map(r => ({
    text: String(r[0]||'').trim(),
    options: [String(r[1]||''),String(r[2]||''),String(r[3]||''),String(r[4]||'')],
    correct: Math.max(0, ABCD.indexOf(String(r[5]||'A').trim().toUpperCase())),
    points: parseInt(String(r[6]).replace(/[^0-9]/g,''))||10,
    hint: String(r[7]||'').trim(),
    explanation: String(r[8]||'').trim(),
    showHint: !!r[7], showExpl: !!r[8],
  }))
}
function parsePasteText(text) {
  if (!text || !text.trim()) return null
  const lines = text.trim().split('\n').filter(l => l.trim())
  if (lines.length < 2) return null

  // Detect format
  const firstData = lines[0]

  // Format 1: Markdown pipe table  | 題目 | 選A | 選B | ...
  if (firstData.includes('|')) {
    const dataLines = lines.filter(l => !l.match(/^\s*\|[\s\-|]+\|\s*$/)) // remove separator rows like |---|---|
    if (dataLines.length < 2) return null
    const rows = dataLines.map(l =>
      l.replace(/^\||\|$/g, '').split('|').map(c => c.trim())
    )
    return parseQuestionRows(rows)
  }

  // Format 2: Tab separated (from Excel copy or AI table)
  if (firstData.includes('\t')) {
    return parseQuestionRows(lines.map(l => l.split('\t')))
  }

  // Format 3: Multiple spaces as separator (2+ spaces)
  // Only if consistent — try splitting on 2+ spaces
  const bySpaces = lines.map(l => l.split(/\s{2,}/))
  if (bySpaces[0].length >= 6) {
    return parseQuestionRows(bySpaces)
  }

  return null
}

function exportToExcel(quiz, responses) {
  if (!responses.length) { alert('目前沒有作答紀錄'); return }
  const groups = {}
  responses.forEach(r => {
    const key = `${r.class}-${padSeat(r.seat)}`
    if (!groups[key]) groups[key] = { class:r.class, seat:padSeat(r.seat), name:r.name, attempts:[] }
    groups[key].attempts.push(r)
  })
  const sorted = Object.values(groups)
    .sort((a,b) => a.class.localeCompare(b.class)||a.seat.localeCompare(b.seat))
    .map(g => {
      g.attempts.sort((a,b) => (a.submittedAt?.seconds||0)-(b.submittedAt?.seconds||0))
      return g
    })

  const scoreData = sorted.map((g,i) => {
    const first = g.attempts[0], best = g.attempts.reduce((b,a)=>a.score>b.score?a:b, g.attempts[0])
    const last = g.attempts[g.attempts.length-1]
    const row = { '序號':i+1,'班級':g.class,'座號':g.seat,'姓名':g.name,
      '作答次數':g.attempts.length,'第一次成績':first?.score??'-','最高成績':best?.score??'-','最新成績':last?.score??'-' }
    quiz.questions.forEach((q,qi) => {
      const ans = last?.answers?.[qi]
      row[`第${qi+1}題`] = ans!=null&&ans>=0?ABCD[ans]:'-'
      row[`第${qi+1}題正確`] = ans===q.correct?'✓':'✗'
    })
    return row
  })
  const analysis = quiz.questions.map((q,qi) => {
    const n=responses.length, correct=responses.filter(r=>r.answers?.[qi]===q.correct).length
    const opts = [0,1,2,3].map(oi=>responses.filter(r=>r.answers?.[qi]===oi).length)
    return { '題號':`第${qi+1}題`,'題目':q.text,'正確答案':ABCD[q.correct],'配分':q.points,
      '答對人數':correct,'答對率':n?`${Math.round(correct/n*100)}%`:'-',
      '選A人數':opts[0],'選B人數':opts[1],'選C人數':opts[2],'選D人數':opts[3] }
  })
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(scoreData), '學生成績')
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(analysis), '題目分析')
  XLSX.writeFile(wb, `${quiz.title}_成績_${new Date().toLocaleDateString('zh-TW').replace(/\//g,'-')}.xlsx`)
}

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
/* Login */
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
.google-btn{width:100%;display:flex;align-items:center;justify-content:center;gap:12px;padding:14px 20px;
  background:white;border:1.5px solid var(--border);border-radius:10px;font-size:15px;font-weight:500;
  cursor:pointer;transition:all .2s;font-family:'Noto Sans TC',sans-serif;color:var(--ink)}
.google-btn:hover{border-color:var(--accent);transform:translateY(-1px);box-shadow:0 4px 16px rgba(45,106,79,.15)}
.google-btn:disabled{opacity:.6;cursor:not-allowed;transform:none}
.login-features{margin-top:32px;display:flex;flex-direction:column;gap:10px}
.feat-item{display:flex;align-items:center;gap:10px;font-size:13px;color:var(--ink2)}
.feat-dot{width:6px;height:6px;border-radius:50%;background:var(--accent2);flex-shrink:0}
/* Layout */
.layout{display:flex;min-height:100vh}
.sidebar{width:240px;background:#1a1714;flex-shrink:0;display:flex;flex-direction:column;position:fixed;height:100vh;z-index:10}
.sidebar-brand{padding:24px 20px;border-bottom:1px solid rgba(255,255,255,.08)}
.brand-name{font-size:16px;font-weight:700;color:white}
.brand-sub{font-size:11px;color:rgba(255,255,255,.4);margin-top:2px;letter-spacing:.1em;text-transform:uppercase}
.admin-badge{display:inline-block;background:var(--admin);color:white;font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px;margin-top:4px}
.sidebar-nav{flex:1;padding:16px 12px;display:flex;flex-direction:column;gap:4px;overflow-y:auto}
.nav-item{display:flex;align-items:center;gap:10px;padding:10px 12px;border-radius:8px;font-size:14px;color:rgba(255,255,255,.6);cursor:pointer;transition:all .15s}
.nav-item:hover{background:rgba(255,255,255,.06);color:white}
.nav-item.active{background:var(--accent);color:white}
.nav-item.admin-nav.active{background:var(--admin)}
.nav-icon{font-size:16px;width:20px;text-align:center}
.sidebar-user{padding:16px 20px;border-top:1px solid rgba(255,255,255,.08);display:flex;align-items:center;gap:10px}
.user-avatar{width:32px;height:32px;border-radius:50%;background:var(--accent);display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:700;color:white;flex-shrink:0}
.user-avatar.admin-av{background:var(--admin)}
.user-name{font-size:13px;color:rgba(255,255,255,.8);font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:110px}
.logout-btn{margin-left:auto;font-size:12px;color:rgba(255,255,255,.3);cursor:pointer;padding:4px;flex-shrink:0}
.logout-btn:hover{color:rgba(255,255,255,.7)}
.main{flex:1;margin-left:240px;padding:32px;max-width:calc(100% - 240px)}
/* Common */
.page-header{display:flex;align-items:flex-start;justify-content:space-between;margin-bottom:28px}
.page-title{font-size:24px;font-weight:700}
.page-sub{font-size:14px;color:var(--ink2);margin-top:4px}
.btn{display:inline-flex;align-items:center;gap:6px;padding:10px 18px;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer;border:none;transition:all .15s;font-family:'Noto Sans TC',sans-serif}
.btn:disabled{opacity:.5;cursor:not-allowed!important;transform:none!important}
.btn-primary{background:var(--accent);color:white}
.btn-primary:hover:not(:disabled){background:#235c42;transform:translateY(-1px);box-shadow:0 4px 12px rgba(45,106,79,.3)}
.btn-secondary{background:white;color:var(--ink);border:1.5px solid var(--border)}
.btn-secondary:hover:not(:disabled){border-color:var(--accent);color:var(--accent)}
.btn-danger{background:var(--danger);color:white}
.btn-danger:hover:not(:disabled){background:#a33a0c;transform:translateY(-1px)}
.btn-excel{background:#1d6f42;color:white}
.btn-excel:hover:not(:disabled){background:#155232;transform:translateY(-1px)}
.btn-admin{background:var(--admin);color:white}
.btn-admin:hover:not(:disabled){background:#4a3fb5;transform:translateY(-1px)}
.btn-sm{padding:6px 12px;font-size:13px}
.card{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);padding:24px}
.tabs{display:flex;gap:4px;background:#f0ede8;border-radius:10px;padding:4px;margin-bottom:20px;width:fit-content}
.tab{padding:8px 18px;border-radius:8px;font-size:14px;font-weight:500;cursor:pointer;transition:all .15s;color:var(--ink2)}
.tab.active{background:white;color:var(--ink);box-shadow:0 1px 4px rgba(0,0,0,.1)}
.form-label{font-size:13px;font-weight:600;color:var(--ink);margin-bottom:8px;display:block}
.form-input{width:100%;padding:10px 14px;border-radius:8px;border:1.5px solid var(--border);font-size:14px;font-family:'Noto Sans TC',sans-serif;color:var(--ink);background:white;transition:border-color .15s}
.form-input:focus{outline:none;border-color:var(--accent)}
.form-select{width:100%;padding:10px 14px;border-radius:8px;border:1.5px solid var(--border);font-size:14px;font-family:'Noto Sans TC',sans-serif;color:var(--ink);background:white;cursor:pointer}
.form-select:focus{outline:none;border-color:var(--accent)}
.form-row{display:grid;grid-template-columns:1fr 1fr;gap:16px}
table{width:100%;border-collapse:collapse;font-size:14px}
th{text-align:left;padding:10px 14px;font-size:12px;font-weight:600;color:var(--ink2);text-transform:uppercase;letter-spacing:.05em;border-bottom:2px solid var(--border);background:#f5f3ef}
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
.quiz-tag-pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:11px;font-weight:600;background:#e8f5ee;color:var(--accent)}
.quiz-name{font-size:16px;font-weight:700;margin-bottom:6px}
.quiz-meta{font-size:12px;color:var(--ink2);display:flex;gap:12px;flex-wrap:wrap}
.quiz-url{margin-top:12px;background:#f5f3ef;border-radius:6px;padding:8px 10px;font-family:'DM Mono',monospace;font-size:11px;color:var(--ink2);display:flex;align-items:center;justify-content:space-between;gap:8px}
.copy-btn{font-size:11px;color:var(--accent);cursor:pointer;font-weight:600;flex-shrink:0}
.new-quiz-card{background:transparent;border-radius:var(--radius);border:2px dashed var(--border);padding:20px;cursor:pointer;transition:all .2s;display:flex;align-items:center;justify-content:center;gap:8px;color:var(--ink2);font-size:14px;font-weight:500;min-height:140px}
.new-quiz-card:hover{border-color:var(--accent);color:var(--accent);background:#f0f9f4}
/* Status badges */
.status-badge{display:inline-flex;align-items:center;gap:5px;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700}
.status-draft{background:#f0ede8;color:#6b6560;border:1px solid #e2ddd8}
.status-active{background:#e8f5ee;color:#2d6a4f;border:1px solid #b7e4c7}
.status-ended{background:#ffeee8;color:#c1440e;border:1px solid #fcd5c5}
.quiz-card-footer{display:flex;gap:6px;flex-wrap:wrap}
/* Edit modal */
.edit-modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.55);z-index:300;display:flex;align-items:stretch;justify-content:flex-end}
.edit-modal-panel{background:white;width:100%;max-width:680px;display:flex;flex-direction:column;overflow:hidden;animation:slideInRight .25s ease}
@keyframes slideInRight{from{transform:translateX(100%)}to{transform:translateX(0)}}
.edit-modal-header{padding:20px 24px;border-bottom:1px solid var(--border);display:flex;align-items:center;justify-content:space-between;flex-shrink:0}
.edit-modal-title{font-size:18px;font-weight:700}
.edit-modal-body{flex:1;overflow-y:auto;padding:24px}
/* Quiz editor */
.q-editor{border:1.5px solid var(--border);border-radius:10px;padding:18px;margin-bottom:12px;background:white}
.q-preview-panel{background:#f8f7ff;border:1.5px solid #c5bff0;border-radius:8px;padding:16px;margin-top:10px}
.q-preview-title{font-size:11px;font-weight:700;color:#5b4fcf;letter-spacing:.08em;text-transform:uppercase;margin-bottom:10px}
.q-preview-question{font-size:15px;font-weight:600;color:var(--ink);line-height:1.7;margin-bottom:12px}
.q-preview-opts{display:flex;flex-direction:column;gap:6px}
.q-preview-opt{display:flex;align-items:center;gap:10px;padding:7px 12px;border-radius:7px;font-size:14px}
.q-preview-opt.is-correct{background:#e8f5ee;border:1.5px solid #b7e4c7;font-weight:600;color:var(--accent)}
.q-preview-opt.not-correct{background:white;border:1.5px solid var(--border);color:var(--ink2)}
.q-preview-opt-letter{font-weight:700;width:20px;flex-shrink:0}
.preview-btn{font-size:12px;padding:4px 10px;border-radius:6px;border:1.5px solid #c5bff0;background:#f0eeff;color:#5b4fcf;cursor:pointer;font-family:'Noto Sans TC',sans-serif;font-weight:600;transition:all .15s}
.preview-btn:hover{background:#e4dfff}
.preview-btn.active{background:#5b4fcf;color:white;border-color:#5b4fcf}
.batch-paste-area{background:#f5f3ef;border-radius:10px;padding:16px;margin-top:12px}
.batch-paste-title{font-size:13px;font-weight:700;margin-bottom:10px;display:flex;align-items:center;gap:6px}
.q-editor-header{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px}
.q-num-label{font-size:11px;font-weight:700;color:var(--ink2);text-transform:uppercase;letter-spacing:.1em}
.options-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-top:10px;margin-bottom:12px}
.option-row{display:flex;align-items:center;gap:8px}
.opt-label{font-size:12px;font-weight:700;color:var(--ink2);width:18px;flex-shrink:0}
.hint-toggle{font-size:12px;padding:5px 12px;border-radius:6px;border:1.5px solid var(--border);background:white;cursor:pointer;color:var(--ink2);font-family:'Noto Sans TC',sans-serif;transition:all .15s;margin-right:6px}
.hint-toggle:hover{border-color:var(--warn);color:var(--warn)}
.hint-area{background:#fffbf0;border:1.5px solid #f0e0a0;border-radius:8px;padding:12px;margin-top:10px}
.hint-area-title{font-size:11px;font-weight:700;color:#a67c00;text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px}
.expl-area{background:#f0f9f4;border:1.5px solid #b7e4c7;border-radius:8px;padding:12px;margin-top:8px}
.expl-area-title{font-size:11px;font-weight:700;color:var(--accent);text-transform:uppercase;letter-spacing:.08em;margin-bottom:8px}
.upload-zone{border:2px dashed var(--border);border-radius:10px;padding:32px;text-align:center;cursor:pointer;transition:all .2s;background:white}
.upload-zone:hover{border-color:var(--accent);background:#f0f9f4}
.paste-zone{width:100%;min-height:160px;padding:14px;border-radius:10px;border:1.5px solid var(--border);font-size:13px;font-family:'DM Mono',monospace;color:var(--ink);background:white;resize:vertical;transition:border-color .15s;line-height:1.5}
.paste-zone:focus{outline:none;border-color:var(--accent)}
.parse-preview{background:#f0f9f4;border:1.5px solid #b7e4c7;border-radius:10px;padding:14px;margin-top:12px}
.parse-preview-title{font-size:12px;font-weight:700;color:var(--accent);margin-bottom:10px}
.preview-row{font-size:12px;padding:6px 0;border-bottom:1px solid #d8f0e4;display:flex;gap:8px}
.preview-row:last-child{border-bottom:none}
/* Analytics */
.q-analytics{display:flex;flex-direction:column;gap:10px}
.q-row{border-radius:10px;padding:14px 16px;border-left:4px solid transparent}
.err-red{background:#fff2f2;border-left-color:var(--danger)}
.err-orange{background:#fff6ee;border-left-color:var(--warn)}
.err-yellow{background:#fffce0;border-left-color:#e9c32a}
.err-green{background:#f0f9f4;border-left-color:var(--accent2)}
.q-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:8px;gap:8px}
.q-text-sm{font-size:13px;font-weight:600;flex:1}
.err-rate{font-size:12px;font-weight:700;padding:2px 8px;border-radius:4px}
.err-rate-red{background:#ffeee8;color:var(--danger)}
.err-rate-orange{background:#fff3e0;color:#c07000}
.err-rate-yellow{background:#fffbe0;color:#9a7c00}
.err-rate-green{background:#e8f5ee;color:var(--accent)}
.q-bar-bg{height:6px;background:rgba(0,0,0,.1);border-radius:3px;overflow:hidden;margin-bottom:8px}
.q-bar{height:100%;border-radius:3px;transition:width .6s}
.bar-green{background:var(--accent2)}
.bar-red{background:#ff7c5c}
.q-opts-row{display:flex;gap:6px;flex-wrap:wrap}
.q-opt-chip{font-size:11px;padding:3px 8px;border-radius:4px;background:white;border:1px solid var(--border);color:var(--ink2)}
.q-opt-chip.correct{background:#e8f5ee;border-color:var(--accent2);color:var(--accent);font-weight:600}
.wrong-seats{margin-top:8px;padding-top:8px;border-top:1px solid rgba(0,0,0,.08)}
.wrong-seat-badge{display:inline-block;padding:2px 7px;border-radius:4px;font-size:11px;background:#ffeee8;color:var(--danger);margin:2px;font-family:'DM Mono',monospace}
.answer-dots{display:flex;gap:4px}
.dot{width:16px;height:16px;border-radius:50%;display:flex;align-items:center;justify-content:center;font-size:9px;font-weight:700;color:white}
.dot-correct{background:var(--accent2)}
.dot-wrong{background:var(--danger)}
/* Student */
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
.sq-text{font-size:15px;font-weight:600;margin-bottom:14px;line-height:1.7}
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
.opt-text{font-size:14px;color:var(--ink);flex:1;line-height:1.5}
.hint-btn{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border-radius:6px;border:1.5px solid #f0e0a0;background:#fffbf0;color:#a67c00;font-size:12px;font-weight:600;cursor:pointer;font-family:'Noto Sans TC',sans-serif}
.hint-bubble{background:#fffbf0;border:1.5px solid #f0e0a0;border-radius:8px;padding:10px 14px;font-size:13px;color:#7a5c00;margin-top:10px;line-height:1.7}
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
.result-btns{display:flex;gap:12px;margin-top:20px}
.progress-bar-wrap{background:white;border-radius:10px;padding:12px 16px;margin-bottom:16px;box-shadow:var(--shadow);display:flex;align-items:center;gap:12px}
.progress-track{flex:1;height:6px;background:var(--border);border-radius:3px;overflow:hidden}
.progress-fill{height:100%;background:var(--accent2);border-radius:3px;transition:width .3s}
/* Admin */
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
/* Status badges */
.status-badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:700;border:1.5px solid transparent}
.status-draft{background:#f0ede8;color:#6b6560;border-color:#e2ddd8}
.status-active{background:#e8f5ee;color:#2d6a4f;border-color:#52b788}
.status-ended{background:#ffeee8;color:#c1440e;border-color:#fcd5c5}
.status-btn{display:inline-flex;align-items:center;gap:5px;padding:5px 12px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;border:1.5px solid;font-family:'Noto Sans TC',sans-serif;transition:all .15s}
.warn-banner{background:#fff8e1;border:1.5px solid #f0d060;border-radius:10px;padding:14px 16px;margin-bottom:20px;display:flex;gap:10px;align-items:flex-start}
.warn-banner-icon{font-size:18px;flex-shrink:0;line-height:1.4}
.warn-banner-text{font-size:13px;color:#7a5c00;line-height:1.6}
.quiz-card-footer{display:flex;gap:6px;margin-top:10px;padding-top:10px;border-top:1px solid var(--border);flex-wrap:wrap;align-items:center}
.roster-class-block{background:var(--surface);border-radius:var(--radius);border:1px solid var(--border);margin-bottom:16px;overflow:hidden}
.roster-class-header{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;background:#f5f3ef;border-bottom:1px solid var(--border)}
.roster-class-name{font-size:15px;font-weight:700}
.roster-student-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(120px,1fr));gap:8px;padding:14px 18px}
.roster-student{display:flex;align-items:center;gap:6px;padding:6px 10px;background:white;border-radius:6px;border:1px solid var(--border);font-size:13px}
.roster-seat{font-family:'DM Mono',monospace;font-size:11px;font-weight:700;color:var(--accent);flex-shrink:0}
.roster-name{overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
/* Modals */
.modal-overlay{position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:200;display:flex;align-items:center;justify-content:center;padding:16px}
.modal-box{background:white;border-radius:16px;padding:28px;width:100%;max-width:420px;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal-title{font-size:18px;font-weight:700;margin-bottom:8px}
.modal-sub{font-size:14px;color:var(--ink2);margin-bottom:24px;line-height:1.6}
.modal-actions{display:flex;gap:10px;justify-content:flex-end}
/* Toast */
.toast{position:fixed;bottom:24px;right:24px;background:#1a1714;color:white;padding:12px 20px;border-radius:8px;font-size:14px;z-index:1000;box-shadow:0 4px 20px rgba(0,0,0,.3);animation:slideUp .3s ease}
@keyframes slideUp{from{transform:translateY(20px);opacity:0}to{transform:translateY(0);opacity:1}}
.loading{display:flex;align-items:center;justify-content:center;min-height:200px;font-size:14px;color:var(--ink2);gap:10px}
.spinner{width:20px;height:20px;border:2px solid var(--border);border-top-color:var(--accent);border-radius:50%;animation:spin .7s linear infinite}
@keyframes spin{to{transform:rotate(360deg)}}
.empty-state{text-align:center;padding:60px 20px;color:var(--ink2)}
.empty-icon{font-size:48px;margin-bottom:16px}
/* Attempt history */
.attempt-row{font-size:12px;color:var(--ink2);display:flex;gap:16px;padding:6px 0;border-bottom:1px solid var(--border)}
.attempt-row:last-child{border-bottom:none}
/* QR */
.qr-img{display:block;margin:0 auto;border-radius:8px;image-rendering:pixelated}
.float-top-btn{position:fixed;bottom:28px;right:20px;width:44px;height:44px;border-radius:50%;background:var(--accent);color:white;border:none;font-size:20px;cursor:pointer;box-shadow:0 4px 16px rgba(0,0,0,.2);display:flex;align-items:center;justify-content:center;z-index:50;transition:all .2s;line-height:1}
.float-top-btn:hover{background:#235c42;transform:translateY(-2px);box-shadow:0 6px 20px rgba(0,0,0,.25)}
.unanswered-highlight{animation:pulse-border .6s ease 2;border-color:var(--danger)!important}
@keyframes pulse-border{0%,100%{box-shadow:none}50%{box-shadow:0 0 0 4px rgba(193,68,14,.3)}}
@media(max-width:768px){
  .sidebar{display:none}
  .main{margin-left:0;max-width:100%;padding:16px}
  .stats-row{grid-template-columns:1fr 1fr}
  .form-row{grid-template-columns:1fr}
  .options-grid{grid-template-columns:1fr}
  .info-grid{grid-template-columns:1fr}
}
`

// ─── Toast ────────────────────────────────────────────────────────────────────
function Toast({ msg }) { return <div className="toast">✓ {msg}</div> }

// ─── Modals ───────────────────────────────────────────────────────────────────
function ConfirmModal({ title, message, onConfirm, onCancel }) {
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-box" onClick={e=>e.stopPropagation()}>
        <div className="modal-title">🗑️ {title}</div>
        <div className="modal-sub">{message}</div>
        <div className="modal-actions">
          <button className="btn btn-secondary" onClick={onCancel}>取消</button>
          <button className="btn btn-danger" onClick={onConfirm}>確認刪除</button>
        </div>
      </div>
    </div>
  )
}

function ProfileModal({ user, displayName, onSave, onCancel }) {
  const [name, setName] = useState(displayName || user?.displayName || '')
  const [saving, setSaving] = useState(false)
  const handleSave = async () => {
    if (!name.trim()) return
    setSaving(true)
    try {
      await setDoc(doc(db,'users',user.uid), { displayName:name.trim(), email:user.email, updatedAt:serverTimestamp() }, { merge:true })
      onSave(name.trim())
    } catch(e) { console.error(e) }
    setSaving(false)
  }
  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div className="modal-box" onClick={e=>e.stopPropagation()}>
        <div className="modal-title">✏️ 修改名稱</div>
        <div style={{marginBottom:20}}>
          <label className="form-label">你的名稱</label>
          <input className="form-input" placeholder="例：王小明老師" value={name}
            onChange={e=>setName(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleSave()} autoFocus/>
        </div>
        <div className="modal-actions">
          <button className="btn btn-secondary" onClick={onCancel}>取消</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving||!name.trim()}>
            {saving?'儲存中...':'儲存'}
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── QuestionEditorCard ───────────────────────────────────────────────────────
function PreviewModal({ q, qi, onClose }) {
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box" onClick={e=>e.stopPropagation()} style={{maxWidth:520}}>
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:16}}>
          <div style={{fontSize:16,fontWeight:700}}>👁 第 {qi+1} 題預覽</div>
          <button onClick={onClose} style={{background:'none',border:'none',fontSize:20,cursor:'pointer',color:'var(--ink2)',lineHeight:1}}>✕</button>
        </div>
        <div className="q-preview-question">
          {q.text ? <MathText text={q.text}/> : <span style={{color:'var(--ink2)',fontStyle:'italic'}}>尚未輸入題目</span>}
        </div>
        <div className="q-preview-opts" style={{margin:'12px 0'}}>
          {q.options.map((opt,oi)=>(
            <div key={oi} className={`q-preview-opt ${oi===q.correct?'is-correct':'not-correct'}`}>
              <span className="q-preview-opt-letter">{ABCD[oi]}.</span>
              {opt ? <MathText text={opt}/> : <span style={{fontStyle:'italic',opacity:.5}}>選項 {ABCD[oi]}</span>}
              {oi===q.correct && <span style={{marginLeft:'auto',fontSize:11,fontWeight:700}}>✓ 正確答案</span>}
            </div>
          ))}
        </div>
        {(q.hint||q.explanation) && (
          <div style={{display:'flex',flexDirection:'column',gap:8,marginTop:8}}>
            {q.hint && <div style={{fontSize:13,color:'#a67c00',background:'#fffbf0',padding:'8px 12px',borderRadius:8,lineHeight:1.6}}>
              <strong>💡 提示：</strong><MathText text={q.hint}/>
            </div>}
            {q.explanation && <div style={{fontSize:13,color:'var(--accent)',background:'#f0f9f4',padding:'8px 12px',borderRadius:8,lineHeight:1.6}}>
              <strong>📖 解析：</strong><MathText text={q.explanation}/>
            </div>}
          </div>
        )}
        <div className="modal-actions" style={{marginTop:20}}>
          <button className="btn btn-secondary" onClick={onClose}>關閉</button>
        </div>
      </div>
    </div>
  )
}

function QuestionEditorCard({ q, qi, total, updateQ, updateOpt, toggleField, removeQ }) {
  const [showPreview, setShowPreview] = useState(false)
  return (
    <>
    <div className="q-editor">
      <div className="q-editor-header">
        <span className="q-num-label">第 {qi+1} 題</span>
        <div style={{display:'flex',alignItems:'center',gap:8}}>
          <button className="preview-btn" onClick={()=>setShowPreview(true)}>👁 預覽</button>
          <span style={{fontSize:12,color:'var(--ink2)'}}>配分</span>
          <input className="form-input" type="number" min="1" style={{width:60,padding:'4px 8px',fontSize:13}}
            value={q.points} onChange={e=>updateQ(qi,'points',parseInt(e.target.value)||1)}/>
          {total>1 && <span style={{fontSize:12,color:'var(--danger)',cursor:'pointer'}} onClick={()=>removeQ(qi)}>✕ 刪除</span>}
        </div>
      </div>
      <input className="form-input" placeholder="輸入題目（可用 $\frac{3}{8}$ 表示分數）" value={q.text}
        onChange={e=>updateQ(qi,'text',e.target.value)} style={{marginBottom:6}}/>
      <div className="options-grid">
        {ABCD.map((lbl,oi)=>(
          <div key={oi} className="option-row">
            <span className="opt-label">{lbl}.</span>
            <input className="form-input" style={{fontSize:13,padding:'7px 10px'}} placeholder={`選項 ${lbl}`}
              value={q.options[oi]} onChange={e=>updateOpt(qi,oi,e.target.value)}/>
            <input type="radio" style={{accentColor:'var(--accent)',width:16,height:16,cursor:'pointer'}}
              name={`c_${qi}`} checked={q.correct===oi} onChange={()=>updateQ(qi,'correct',oi)} id={`r_${qi}_${oi}`}/>
            <label style={{fontSize:11,color:'var(--ink2)',cursor:'pointer'}} htmlFor={`r_${qi}_${oi}`}>✓正確</label>
          </div>
        ))}
      </div>
      <div style={{display:'flex',flexWrap:'wrap',gap:6}}>
        <button className="hint-toggle" onClick={()=>toggleField(qi,'showHint')}>{q.showHint?'▼':'▶'} 💡 提示</button>
        <button className="hint-toggle" style={{borderColor:'#b7e4c7',color:'var(--accent)'}} onClick={()=>toggleField(qi,'showExpl')}>{q.showExpl?'▼':'▶'} 📖 解析</button>
      </div>
      {q.showHint && <div className="hint-area" style={{marginTop:10}}>
        <div className="hint-area-title">💡 提示（學生作答中可主動查看）</div>
        <textarea className="form-input" style={{resize:'vertical',minHeight:56,fontSize:13}}
          placeholder="例：分母是全部份數，分子是已吃的份數" value={q.hint} onChange={e=>updateQ(qi,'hint',e.target.value)}/>
      </div>}
      {q.showExpl && <div className="expl-area" style={{marginTop:8}}>
        <div className="expl-area-title">📖 解析（提交後才顯示）</div>
        <textarea className="form-input" style={{resize:'vertical',minHeight:72,fontSize:13}}
          placeholder="例：全部8份（分母），吃了3份（分子），所以是3/8" value={q.explanation} onChange={e=>updateQ(qi,'explanation',e.target.value)}/>
      </div>}
    </div>
    {showPreview && <PreviewModal q={q} qi={qi} onClose={()=>setShowPreview(false)}/>}
    </>
  )
}

// ─── BatchPasteSection ────────────────────────────────────────────────────────
function BatchPasteSection({ onImport, appendMode=false }) {
  const [pasteText, setPasteText] = useState('')
  const [preview, setPreview] = useState(null)
  const [format, setFormat] = useState('auto') // 'auto'|'tab'|'pipe'

  const handleChange = (text) => {
    setPasteText(text)
    setPreview(text.trim() ? parsePasteText(text) : null)
  }

  const handleImport = () => {
    if (!preview?.length) return
    onImport(preview)
    setPasteText(''); setPreview(null)
  }

  const formatExamples = {
    tab:  "題目\t選A\t選B\t選C\t選D\t正確(A/B/C/D)\t配分\t提示(選填)\t解析(選填)\n小明吃了幾分之幾\t1/8\t3/8\t5/8\t8/3\tB\t10\t分母是份數\t8份吃3份",
    pipe: "| 題目 | 選A | 選B | 選C | 選D | 正確 | 配分 | 提示 | 解析 |\n|---|---|---|---|---|---|---|---|---|\n| 小明吃了幾分之幾 | 1/8 | 3/8 | 5/8 | 8/3 | B | 10 | 分母是份數 | 8份吃3份 |",
  }

  return (
    <div className="batch-paste-area">
      <div className="batch-paste-title">
        📋 {appendMode ? '批次新增題目（貼上後追加到現有題目後面）' : '貼上 AI 生成的題目'}
      </div>

      {/* Format tabs */}
      <div style={{display:'flex',gap:6,marginBottom:10}}>
        {[['auto','🔍 自動偵測'],['tab','📊 Tab 分隔（Excel）'],['pipe','📝 管線格式（AI）']].map(([f,label])=>(
          <button key={f} onClick={()=>{ setFormat(f); if(f!=='auto') setPasteText(formatExamples[f]||''); }}
            style={{padding:'5px 12px',borderRadius:6,fontSize:12,fontWeight:600,cursor:'pointer',fontFamily:"'Noto Sans TC',sans-serif",
              background:format===f?'var(--accent)':'white',color:format===f?'white':'var(--ink2)',
              border:format===f?'1.5px solid var(--accent)':'1.5px solid var(--border)',transition:'all .15s'}}>
            {label}
          </button>
        ))}
      </div>

      <div style={{fontSize:12,color:'var(--ink2)',marginBottom:8}}>
        {format==='pipe' && '支援 Markdown 表格（從 AI 複製的表格格式）'}
        {format==='tab'  && '支援 Tab 分隔（從 Excel 直接複製貼上）'}
        {format==='auto' && '自動偵測 Tab 或 Markdown 管線格式'}
      </div>

      <textarea className="paste-zone"
        placeholder={format==='pipe'
          ? "| 題目 | 選A | 選B | 選C | 選D | 正確(A/B/C/D) | 配分 | 提示 | 解析 |\n|---|---|---|---|---|---|---|---|---|\n| 題目內容 | 甲 | 乙 | 丙 | 丁 | A | 10 | | |"
          : "第一列是標題列，欄位用 Tab 分隔\n\n題目\t選A\t選B\t選C\t選D\t正確(A/B/C/D)\t配分\t提示(選填)\t解析(選填)\n小明吃了3/8個披薩...\t1/8\t3/8\t5/8\t8/3\tB\t10\t分母是份數\t全部8份吃了3份"}
        value={pasteText}
        onChange={e=>handleChange(e.target.value)}
        style={{minHeight:140}}
      />

      {preview && preview.length > 0 && (
        <div className="parse-preview" style={{marginTop:10}}>
          <div className="parse-preview-title">✅ 成功解析 {preview.length} 道題目</div>
          {preview.slice(0,3).map((q,i)=>(
            <div key={i} className="preview-row">
              <div style={{fontWeight:700,color:'var(--accent)',flexShrink:0,width:40}}>第{i+1}題</div>
              <div style={{flex:1}}>
                <div style={{fontWeight:600,marginBottom:2,fontSize:13}}>{q.text.slice(0,40)}{q.text.length>40?'...':''}</div>
                <div style={{fontSize:11,color:'var(--ink2)'}}>
                  {q.options.map((o,oi)=>(
                    <span key={oi} style={{marginRight:8,color:oi===q.correct?'var(--accent)':'var(--ink2)',fontWeight:oi===q.correct?700:400}}>
                      {ABCD[oi]}.{o.slice(0,6)}
                    </span>
                  ))}
                </div>
              </div>
            </div>
          ))}
          {preview.length>3 && <div style={{fontSize:12,color:'var(--ink2)',padding:'6px 0'}}>...還有 {preview.length-3} 道題目</div>}
          <button className="btn btn-primary" style={{marginTop:10,width:'100%'}} onClick={handleImport}>
            {appendMode ? `＋ 追加這 ${preview.length} 道題目` : `匯入這 ${preview.length} 道題目 →`}
          </button>
        </div>
      )}
      {pasteText && (!preview||preview.length===0) && (
        <div style={{marginTop:10,padding:'10px 14px',background:'#ffeee8',borderRadius:8,fontSize:13,color:'var(--danger)'}}>
          ⚠️ 無法解析。請確認：<br/>
          • Tab 分隔：第一列是標題，欄位用 Tab 分隔<br/>
          • 管線格式：每列以 | 開頭和結尾
        </div>
      )}
    </div>
  )
}

// ─── Edit Quiz Modal ──────────────────────────────────────────────────────────
function EditQuizModal({ quiz, onSave, onCancel }) {
  const isActive = getStatus(quiz) === 'active'
  const [warnAcked, setWarnAcked] = useState(false)

  // Pre-fill all fields from existing quiz
  const [title, setTitle] = useState(quiz.title || '')
  const [subject, setSubject] = useState(quiz.subject || '')
  const s = quiz.settings || {}
  const [allowHint, setAllowHint] = useState(s.allowHint !== false)
  const [showExplAfter, setShowExplAfter] = useState(s.showExplAfter !== false)
  const [showCorrect, setShowCorrect] = useState(s.showCorrect !== false)
  const [allowMultipleAttempts, setAllowMultipleAttempts] = useState(s.allowMultipleAttempts !== false)
  const [selectedClasses, setSelectedClasses] = useState(new Set(s.allowedClasses||[]))
  const [useRoster, setUseRoster] = useState(s.useRoster || false)
  const [allowNameEdit, setAllowNameEdit] = useState(s.allowNameEdit !== false)
  const [startTime, setStartTime] = useState(s.startTime || '')
  const [endTime, setEndTime] = useState(s.endTime || '')
  const [rosterClasses, setRosterClasses] = useState([])
  const [questions, setQuestions] = useState(
    (quiz.questions||[]).map(q => ({...q, showHint:!!q.hint, showExpl:!!q.explanation}))
  )
  const [saving, setSaving] = useState(false)
  const [mode, setMode] = useState('questions') // 'questions' | 'settings'

  useEffect(() => {
    getDoc(doc(db,'rosters',quiz.teacherId)).then(snap => {
      if (snap.exists()) setRosterClasses(Object.keys(snap.data().classes||{}).sort())
    }).catch(()=>{})
  }, [quiz.teacherId])

  const toggleClass = (cls) => setSelectedClasses(prev => {
    const next = new Set(prev)
    if (next.has(cls)) next.delete(cls); else next.add(cls)
    return next
  })

  const addQ = () => setQuestions(p=>[...p, emptyQ()])
  const removeQ = i => setQuestions(p=>p.filter((_,idx)=>idx!==i))
  const updateQ = (i,f,v) => setQuestions(p=>{const q=[...p];q[i]={...q[i],[f]:v};return q})
  const updateOpt = (qi,oi,v) => setQuestions(p=>{const q=[...p];q[qi].options[oi]=v;return q})
  const toggleField = (i,f) => setQuestions(p=>{const q=[...p];q[i]={...q[i],[f]:!q[i][f]};return q})

  const handleSave = async () => {
    if (!title.trim()) { alert('請輸入測驗名稱'); return }
    if (questions.some(q=>!q.text.trim())) { alert('請填寫所有題目'); return }
    setSaving(true)
    try {
      const clean = questions.map(({showHint,showExpl,...q})=>q)
      const classes = useRoster ? [...selectedClasses] : []
      await setDoc(doc(db,'quizzes',quiz.id), {
        ...quiz,
        title: title.trim(),
        subject: subject.trim() || '未分類',
        questions: clean,
        settings: { allowHint, showExplAfter, showCorrect, allowMultipleAttempts, allowedClasses:classes, useRoster, allowNameEdit, startTime, endTime },
        updatedAt: serverTimestamp(),
      })
      onSave({ ...quiz, title:title.trim(), subject:subject.trim()||'未分類', questions:clean,
        settings:{ allowHint, showExplAfter, showCorrect, allowMultipleAttempts, allowedClasses:classes, useRoster, allowNameEdit, startTime, endTime } })
    } catch(e) { alert('儲存失敗'); console.error(e) }
    setSaving(false)
  }

  // If active and not yet acknowledged warning → show warning first
  if (isActive && !warnAcked) {
    return (
      <div className="modal-overlay" onClick={onCancel}>
        <div className="modal-box" onClick={e=>e.stopPropagation()} style={{maxWidth:460}}>
          <div style={{fontSize:32,marginBottom:12,textAlign:'center'}}>⚠️</div>
          <div className="modal-title" style={{textAlign:'center'}}>測驗派發中，確定要編輯？</div>
          <div className="modal-sub" style={{marginBottom:0}}>
            此測驗目前正在派發，修改題目或設定<strong>不會影響已提交的成績</strong>，
            但可能導致進行中的學生出現題目不一致的情況。<br/><br/>
            如需大幅修改，建議先將測驗「暫停」，完成後再重新開放。
          </div>
          <div className="modal-actions" style={{marginTop:24}}>
            <button className="btn btn-secondary" onClick={onCancel}>取消</button>
            <button className="btn" style={{background:'var(--warn)',color:'white'}} onClick={()=>setWarnAcked(true)}>
              了解，繼續編輯
            </button>
          </div>
        </div>
      </div>
    )
  }

  return (
    <div className="modal-overlay" onClick={onCancel}>
      <div onClick={e=>e.stopPropagation()} style={{
        background:'white', borderRadius:16, width:'100%', maxWidth:720,
        maxHeight:'90vh', overflow:'hidden', display:'flex', flexDirection:'column',
        boxShadow:'0 20px 60px rgba(0,0,0,.3)'
      }}>
        {/* Header */}
        <div style={{padding:'20px 24px', borderBottom:'1px solid var(--border)', display:'flex', alignItems:'center', justifyContent:'space-between', flexShrink:0}}>
          <div>
            <div style={{fontSize:18,fontWeight:700}}>✏️ 編輯測驗</div>
            <div style={{fontSize:13,color:'var(--ink2)',marginTop:2}}>{quiz.title}</div>
          </div>
          <div style={{display:'flex',gap:8,alignItems:'center'}}>
            <div className="tabs" style={{margin:0}}>
              <div className={`tab ${mode==='questions'?'active':''}`} onClick={()=>setMode('questions')}>題目</div>
              <div className={`tab ${mode==='settings'?'active':''}`} onClick={()=>setMode('settings')}>設定</div>
            </div>
            <button onClick={onCancel} style={{background:'none',border:'none',cursor:'pointer',fontSize:20,color:'var(--ink2)',padding:'4px 8px'}}>✕</button>
          </div>
        </div>

        {/* Scrollable body */}
        <div style={{flex:1, overflowY:'auto', padding:'20px 24px'}}>
          {isActive && (
            <div className="warn-banner" style={{marginBottom:16}}>
              <div className="warn-banner-icon">⚠️</div>
              <div className="warn-banner-text">測驗派發中 — 修改將立即生效，已提交的成績不受影響</div>
            </div>
          )}

          {mode==='questions' && (
            <>
              <div className="form-row" style={{marginBottom:16}}>
                <div><label className="form-label">測驗名稱</label>
                  <input className="form-input" value={title} onChange={e=>setTitle(e.target.value)}/></div>
                <div><label className="form-label">科目</label>
                  <input className="form-input" value={subject} onChange={e=>setSubject(e.target.value)}/></div>
              </div>
              <div style={{fontSize:13,fontWeight:700,marginBottom:12}}>📝 題目（支援 $公式$，每題右上角可預覽）</div>
              {questions.map((q,qi)=>(
                <QuestionEditorCard key={qi} q={q} qi={qi} total={questions.length}
                  updateQ={updateQ} updateOpt={updateOpt} toggleField={toggleField} removeQ={removeQ}/>
              ))}
              <button className="btn btn-secondary btn-sm" onClick={addQ}>＋ 新增題目</button>
              <BatchPasteSection appendMode={true} onImport={parsed => setQuestions(prev=>[...prev,...parsed])}/>
            </>
          )}

          {mode==='settings' && (
            <div>
              <div style={{background:'#f5f3ef',borderRadius:10,padding:16,marginBottom:16}}>
                <div style={{fontSize:13,fontWeight:700,marginBottom:12}}>⚙️ 測驗設定</div>
                <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10}}>
                  {[
                    [allowHint,setAllowHint,'💡 允許提示'],
                    [showExplAfter,setShowExplAfter,'📖 提交後顯示解析'],
                    [showCorrect,setShowCorrect,'✓ 顯示正確答案'],
                    [allowMultipleAttempts,setAllowMultipleAttempts,'🔄 允許重複作答'],
                  ].map(([v,sv,l])=>(
                    <label key={l} style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer'}}>
                      <input type="checkbox" checked={v} onChange={e=>sv(e.target.checked)} style={{accentColor:'var(--accent)',width:15,height:15}}/>{l}
                    </label>
                  ))}
                </div>

                {/* Input mode + class selection */}
                <div style={{marginTop:14}}>
                  <label className="form-label" style={{fontSize:12}}>學生資料輸入方式</label>
                  <div style={{display:'flex',gap:8,marginBottom:12}}>
                    {[['free','✎ 自由輸入'],['roster','📚 使用班級名單']].map(([val,label])=>(
                      <button key={val} onClick={()=>setUseRoster(val==='roster')}
                        style={{padding:'7px 16px',borderRadius:8,fontSize:13,fontWeight:600,cursor:'pointer',fontFamily:"'Noto Sans TC',sans-serif",
                          background:(!useRoster&&val==='free')||(useRoster&&val==='roster')?'var(--accent)':'white',
                          color:(!useRoster&&val==='free')||(useRoster&&val==='roster')?'white':'var(--ink2)',
                          border:(!useRoster&&val==='free')||(useRoster&&val==='roster')?'2px solid var(--accent)':'2px solid var(--border)',
                          transition:'all .15s'}}>
                        {label}
                      </button>
                    ))}
                  </div>

                  {useRoster && (
                    <div style={{marginBottom:12}}>
                      <label className="form-label" style={{fontSize:12}}>開放班級（留空 = 全部班級均可進入）</label>
                      {rosterClasses.length === 0 ? (
                        <div style={{fontSize:13,color:'var(--ink2)',padding:'8px 12px',background:'#f5f3ef',borderRadius:8}}>
                          尚未建立班級名單，請先到「班級名單」頁面新增
                        </div>
                      ) : (
                        <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                          {rosterClasses.map(cls => (
                            <button key={cls} onClick={()=>toggleClass(cls)}
                              style={{padding:'6px 16px',borderRadius:20,fontSize:13,fontWeight:700,cursor:'pointer',fontFamily:"'Noto Sans TC',sans-serif",transition:'all .15s',
                                background:selectedClasses.has(cls)?'var(--accent)':'white',
                                color:selectedClasses.has(cls)?'white':'var(--ink2)',
                                border:selectedClasses.has(cls)?'2px solid var(--accent)':'2px solid var(--border)'}}>
                              {selectedClasses.has(cls)?'✓ ':''}{cls}班
                            </button>
                          ))}
                        </div>
                      )}
                      <label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer',marginTop:10}}>
                        <input type="checkbox" checked={allowNameEdit} onChange={e=>setAllowNameEdit(e.target.checked)} style={{accentColor:'var(--accent)',width:15,height:15}}/>
                        允許學生修改自動帶出的姓名
                      </label>
                    </div>
                  )}

                  <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
                    <div>
                      <label className="form-label" style={{fontSize:12}}>🕐 開始時間（選填）</label>
                      <input className="form-input" type="datetime-local" value={startTime}
                        onChange={e=>setStartTime(e.target.value)} style={{fontSize:13,padding:'8px 12px'}}/>
                    </div>
                    <div>
                      <label className="form-label" style={{fontSize:12}}>🕐 截止時間（選填）</label>
                      <input className="form-input" type="datetime-local" value={endTime}
                        onChange={e=>setEndTime(e.target.value)} style={{fontSize:13,padding:'8px 12px'}}/>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Footer */}
        <div style={{padding:'16px 24px',borderTop:'1px solid var(--border)',display:'flex',justifyContent:'flex-end',gap:10,flexShrink:0,background:'white'}}>
          <button className="btn btn-secondary" onClick={onCancel}>取消</button>
          <button className="btn btn-primary" onClick={handleSave} disabled={saving}>
            {saving ? '儲存中...' : '儲存變更'}
          </button>
        </div>
      </div>
    </div>
  )
}

// ─── QR Code Modal ────────────────────────────────────────────────────────────
function QRModal({ url, quizTitle, onClose }) {
  const [qrSrc, setQrSrc] = useState('')
  useEffect(() => {
    QRCode.toDataURL(url, { width:280, margin:2, color:{dark:'#1a1714',light:'#ffffff'} })
      .then(setQrSrc).catch(()=>{})
  }, [url])
  const handleDownload = () => {
    const a = document.createElement('a')
    a.href = qrSrc; a.download = `${quizTitle || 'quiz'}-QRCode.png`; a.click()
  }
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-box" onClick={e=>e.stopPropagation()} style={{textAlign:'center',maxWidth:360}}>
        <div className="modal-title" style={{marginBottom:4}}>📱 學生掃描入口</div>
        <div style={{fontSize:13,color:'var(--ink2)',marginBottom:20}}>{quizTitle}</div>
        {qrSrc ? <img src={qrSrc} className="qr-img" width={240} height={240} alt="QR Code"/> : <div className="loading"><div className="spinner"/>產生中...</div>}
        <div style={{fontFamily:"'DM Mono',monospace",fontSize:11,background:'#f5f3ef',padding:'8px 12px',borderRadius:6,margin:'16px 0',wordBreak:'break-all',textAlign:'left'}}>{url}</div>
        <div style={{display:'flex',gap:10,justifyContent:'center'}}>
          <button className="btn btn-primary" onClick={handleDownload} disabled={!qrSrc}>⬇ 下載 QR Code</button>
          <button className="btn btn-secondary" onClick={onClose}>關閉</button>
        </div>
      </div>
    </div>
  )
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
          {loading?'登入中...':'使用 Google 帳號登入'}
        </button>
        {error && <div style={{marginTop:12,fontSize:13,color:'var(--danger)',textAlign:'center'}}>{error}</div>}
        <div className="login-features">
          {['每位老師獨立空間，資料安全分離','班級名單匯入，學生選座號自動帶出姓名','支援數學公式顯示與 QR Code 分發'].map(f => (
            <div key={f} className="feat-item"><div className="feat-dot"/>{f}</div>
          ))}
        </div>
      </div>
    </div>
  )
}

// ─── Sidebar ──────────────────────────────────────────────────────────────────
function Sidebar({ hash, user, displayName, onLogout, onEditProfile }) {
  const admin = isAdmin(user)
  const showName = displayName || user?.displayName || user?.email || ''
  const initial = showName[0] || '?'
  const teacherNav = [
    { path:'/', icon:'◉', label:'我的測驗' },
    { path:'/create', icon:'＋', label:'新增測驗' },
    { path:'/roster', icon:'📚', label:'班級名單' },
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
          <div key={n.path} className={`nav-item ${hash===n.path?'active':''}`} onClick={()=>navigate(n.path)}>
            <span className="nav-icon">{n.icon}</span>{n.label}
          </div>
        ))}
        {admin && <>
          <div style={{height:1,background:'rgba(255,255,255,.08)',margin:'8px 4px'}}/>
          <div className={`nav-item admin-nav ${hash==='/admin'?'active':''}`} onClick={()=>navigate('/admin')}>
            <span className="nav-icon">⊞</span>管理後台
          </div>
        </>}
      </div>
      <div className="sidebar-user">
        <div className={`user-avatar ${admin?'admin-av':''}`}>{initial.toUpperCase()}</div>
        <div style={{overflow:'hidden',flex:1}}>
          <div className="user-name">{showName}</div>
          <div style={{fontSize:11,color:'rgba(255,255,255,.35)',cursor:'pointer',marginTop:1}} onClick={onEditProfile}>✏️ 修改名稱</div>
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
      const [qSnap, rSnap] = await Promise.all([getDocs(collection(db,'quizzes')), getDocs(collection(db,'responses'))])
      const allQ = qSnap.docs.map(d=>({id:d.id,...d.data()}))
      const allR = rSnap.docs.map(d=>({id:d.id,...d.data()}))
      const tMap = {}
      allQ.forEach(q => {
        if (!tMap[q.teacherId]) tMap[q.teacherId] = { id:q.teacherId, name:q.teacherName||'未知', quizCount:0, responseCount:0 }
        tMap[q.teacherId].quizCount++
        tMap[q.teacherId].responseCount += allR.filter(r=>r.quizId===q.id).length
      })
      setTeachers(Object.values(tMap).sort((a,b)=>b.quizCount-a.quizCount))
      setQuizzes(allQ); setResponses(allR); setLoading(false)
    }
    load()
  }, [])

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>
  return (
    <div>
      <div className="admin-header">
        <div className="admin-header-title">⊞ 系統管理後台</div>
        <div className="admin-header-sub">查看所有老師的使用狀況與全站統計</div>
      </div>
      <div className="stats-row">
        {[{label:'老師帳號',value:teachers.length,sub:'位',color:'var(--admin)'},
          {label:'測驗總數',value:quizzes.length,sub:'份'},
          {label:'作答次數',value:responses.length,sub:'筆',color:'var(--accent)'},
          {label:'平均每份',value:quizzes.length?Math.round(responses.length/quizzes.length):0,sub:'人作答'}
        ].map(s=>(
          <div key={s.label} className="stat-box">
            <div className="stat-label">{s.label}</div>
            <div className="stat-value" style={{color:s.color||'var(--ink)'}}>{s.value}</div>
            <div className="stat-sub">{s.sub}</div>
          </div>
        ))}
      </div>
      <div className="card" style={{marginBottom:16}}>
        <div style={{fontSize:16,fontWeight:700,marginBottom:20}}>👩‍🏫 老師帳號統計</div>
        {teachers.length===0 ? <div className="empty-state" style={{padding:40}}><div>目前還沒有老師登入</div></div>
          : teachers.map(t=>(
          <div key={t.id} className="teacher-row">
            <div className="teacher-info">
              <div className="teacher-avatar">{(t.name?.[0]||'?').toUpperCase()}</div>
              <div><div className="teacher-name">{t.name}</div><div className="teacher-email">{t.id}</div></div>
            </div>
            <div className="teacher-stats">
              <div className="t-stat"><div className="t-stat-val">{t.quizCount}</div><div className="t-stat-lab">測驗數</div></div>
              <div className="t-stat"><div className="t-stat-val">{t.responseCount}</div><div className="t-stat-lab">作答數</div></div>
            </div>
          </div>
        ))}
      </div>
    </div>
  )
}

// ─── Roster Manager ───────────────────────────────────────────────────────────
function RosterManager({ user }) {
  const [classes, setClasses] = useState({})
  const [loading, setLoading] = useState(true)
  const [newClass, setNewClass] = useState('')
  const [editClass, setEditClass] = useState(null)
  const [editMode, setEditMode] = useState('single') // 'single' | 'paste'
  const [newSeat, setNewSeat] = useState('')
  const [newStudentName, setNewStudentName] = useState('')
  const [pasteRoster, setPasteRoster] = useState('')
  const [toast, setToast] = useState('')

  useEffect(() => {
    getDoc(doc(db,'rosters',user.uid)).then(snap => {
      if (snap.exists()) setClasses(snap.data().classes||{})
      setLoading(false)
    }).catch(()=>setLoading(false))
  }, [user.uid])

  const saveRoster = async (updated) => {
    await setDoc(doc(db,'rosters',user.uid), { classes:updated, updatedAt:serverTimestamp() })
    setClasses(updated)
  }

  const handleAddClass = async () => {
    const name = newClass.trim(); if (!name||classes[name]) return
    const updated = { ...classes, [name]:[] }
    await saveRoster(updated); setNewClass('')
    setToast(`已新增班級 ${name}`); setTimeout(()=>setToast(''),2000)
  }

  const handleDeleteClass = async (name) => {
    if (!confirm(`確定刪除班級 ${name} 的所有名單？`)) return
    const updated = {...classes}; delete updated[name]
    await saveRoster(updated)
    setToast(`已刪除班級 ${name}`); setTimeout(()=>setToast(''),2000)
  }

  const handleAddStudent = async (cls) => {
    const seat = padSeat(newSeat), name = newStudentName.trim()
    if (!seat||!name) return
    const updated = {...classes, [cls]:[...(classes[cls]||[])]}
    const idx = updated[cls].findIndex(s=>s.seat===seat)
    if (idx>=0) updated[cls][idx]={seat,name}
    else updated[cls].push({seat,name})
    updated[cls].sort((a,b)=>a.seat.localeCompare(b.seat))
    await saveRoster(updated); setNewSeat(''); setNewStudentName('')
    setToast('已儲存'); setTimeout(()=>setToast(''),2000)
  }

  const handlePasteRoster = async (cls) => {
    // Parse lines like "01 陳小明" or "1 陳小明" or "01\t陳小明"
    const lines = pasteRoster.trim().split('\n').filter(l=>l.trim())
    const parsed = lines.map(line => {
      const parts = line.trim().split(/[\t\s]+/)
      if (parts.length < 2) return null
      const seat = padSeat(parts[0])
      const name = parts.slice(1).join(' ').trim()
      return seat && name ? { seat, name } : null
    }).filter(Boolean)
    if (!parsed.length) { alert('無法解析，請確認格式：每行「座號 姓名」'); return }
    const updated = {...classes, [cls]:[...(classes[cls]||[])]}
    parsed.forEach(({seat,name}) => {
      const idx = updated[cls].findIndex(s=>s.seat===seat)
      if (idx>=0) updated[cls][idx]={seat,name}
      else updated[cls].push({seat,name})
    })
    updated[cls].sort((a,b)=>a.seat.localeCompare(b.seat))
    await saveRoster(updated)
    setPasteRoster('')
    setToast(`已匯入 ${parsed.length} 位學生`); setTimeout(()=>setToast(''),2500)
  }

  const handleDeleteStudent = async (cls, seat) => {
    const updated = {...classes, [cls]:classes[cls].filter(s=>s.seat!==seat)}
    await saveRoster(updated)
  }

  const handleExcelUpload = (e) => {
    const file = e.target.files[0]; if (!file) return
    e.target.value = ''
    const reader = new FileReader()
    reader.onload = async (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, {type:'binary'})
        const ws = wb.Sheets[wb.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:false, defval:''})
        const updated = {...classes}
        rows.slice(1).filter(r=>r[0]&&r[2]).forEach(r => {
          const cls = String(r[0]).trim()
          const seat = padSeat(r[1])
          const name = String(r[2]).trim()
          if (!updated[cls]) updated[cls] = []
          const idx = updated[cls].findIndex(s=>s.seat===seat)
          if (idx>=0) updated[cls][idx]={seat,name}
          else updated[cls].push({seat,name})
        })
        Object.keys(updated).forEach(c=>updated[c].sort((a,b)=>a.seat.localeCompare(b.seat)))
        await saveRoster(updated)
        setToast('名單匯入成功！'); setTimeout(()=>setToast(''),2500)
      } catch { alert('Excel 格式錯誤') }
    }
    reader.readAsBinaryString(file)
  }

  const exportRosterExcel = () => {
    const rows = []
    Object.entries(classes).forEach(([cls, students]) => {
      students.forEach(s => rows.push({'班級':cls,'座號':s.seat,'姓名':s.name}))
    })
    if (!rows.length) { alert('沒有名單資料'); return }
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), '班級名單')
    XLSX.writeFile(wb, '班級名單.xlsx')
  }

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>

  return (
    <div>
      <div className="page-header">
        <div><div className="page-title">班級名單</div>
          <div className="page-sub">匯入名單後，學生作答只需選班級座號，姓名自動帶出</div></div>
        <div style={{display:'flex',gap:8}}>
          <button className="btn btn-excel" onClick={exportRosterExcel}>📥 匯出名單</button>
          <label className="btn btn-secondary" style={{cursor:'pointer'}}>
            📊 匯入 Excel
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleExcelUpload}/>
          </label>
        </div>
      </div>

      <div className="card" style={{marginBottom:20}}>
        <div style={{fontSize:13,fontWeight:700,marginBottom:12}}>Excel 格式（第一列為標題）：</div>
        <div style={{overflowX:'auto'}}>
          <table style={{fontSize:12,width:'auto'}}>
            <thead><tr>
              {['班級','座號','姓名'].map(h=>(
                <th key={h} style={{padding:'5px 12px',background:'#f0ede8',border:'1px solid var(--border)'}}>{h}</th>
              ))}
            </tr></thead>
            <tbody>
              {[['103','01','陳小明'],['103','02','李美華'],['105','01','張大偉']].map((r,i)=>(
                <tr key={i}>{r.map((c,j)=><td key={j} style={{padding:'5px 12px',border:'1px solid var(--border)',color:'var(--ink2)'}}>{c}</td>)}</tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Add new class */}
      <div style={{display:'flex',gap:10,marginBottom:20}}>
        <input className="form-input" placeholder="新班級名稱（如 103）" value={newClass}
          onChange={e=>setNewClass(e.target.value)} onKeyDown={e=>e.key==='Enter'&&handleAddClass()} style={{maxWidth:240}}/>
        <button className="btn btn-primary" onClick={handleAddClass} disabled={!newClass.trim()}>新增班級</button>
      </div>

      {Object.keys(classes).length===0 ? (
        <div className="empty-state"><div className="empty-icon">📚</div><div>還沒有班級名單，請上傳 Excel 或新增班級</div></div>
      ) : Object.entries(classes).sort(([a],[b])=>a.localeCompare(b)).map(([cls,students])=>(
        <div key={cls} className="roster-class-block">
          <div className="roster-class-header">
            <div className="roster-class-name">📋 {cls} 班 <span style={{fontSize:13,color:'var(--ink2)',fontWeight:400}}>（{students.length} 人）</span></div>
            <div style={{display:'flex',gap:8,alignItems:'center'}}>
              <button className="btn btn-secondary btn-sm" onClick={()=>{ const next=editClass===cls?null:cls; setEditClass(next); if(!next){setPasteRoster('');setNewSeat('');setNewStudentName('')} }}>
                {editClass===cls?'收起':'✎ 新增學生'}
              </button>
              <button className="btn btn-sm" style={{background:'#ffeee8',color:'var(--danger)',border:'1px solid #fcd5c5'}}
                onClick={()=>handleDeleteClass(cls)}>刪除班級</button>
            </div>
          </div>
          {editClass===cls && (
            <div style={{background:'#f9f8f6',borderBottom:'1px solid var(--border)'}}>
              {/* Mode tabs */}
              <div style={{display:'flex',gap:0,borderBottom:'1px solid var(--border)'}}>
                {[['single','✎ 逐一新增'],['paste','📋 貼上名單']].map(([m,label])=>(
                  <button key={m} onClick={()=>setEditMode(m)}
                    style={{padding:'8px 16px',fontSize:13,fontWeight:editMode===m?700:400,
                      color:editMode===m?'var(--accent)':'var(--ink2)',background:'none',border:'none',
                      borderBottom:editMode===m?'2px solid var(--accent)':'2px solid transparent',cursor:'pointer'}}>
                    {label}
                  </button>
                ))}
              </div>
              {editMode==='single' && (
                <div style={{padding:'12px 18px',display:'flex',gap:8,flexWrap:'wrap',alignItems:'flex-end'}}>
                  <div>
                    <div style={{fontSize:11,color:'var(--ink2)',marginBottom:4}}>座號</div>
                    <input className="form-input" placeholder="如：1 或 01" value={newSeat}
                      onChange={e=>setNewSeat(e.target.value)} style={{width:120}}/>
                  </div>
                  <div>
                    <div style={{fontSize:11,color:'var(--ink2)',marginBottom:4}}>姓名</div>
                    <input className="form-input" placeholder="學生姓名" value={newStudentName}
                      onChange={e=>setNewStudentName(e.target.value)} style={{width:160}}
                      onKeyDown={e=>e.key==='Enter'&&handleAddStudent(cls)}/>
                  </div>
                  <button className="btn btn-primary btn-sm" onClick={()=>handleAddStudent(cls)}
                    disabled={!newSeat||!newStudentName.trim()}>新增 / 更新</button>
                </div>
              )}
              {editMode==='paste' && (
                <div style={{padding:'12px 18px'}}>
                  <div style={{fontSize:12,color:'var(--ink2)',marginBottom:8}}>
                    每行一位學生，格式：<code style={{background:'#eee',padding:'1px 5px',borderRadius:3}}>座號 姓名</code>
                    <span style={{marginLeft:8,color:'var(--ink2)'}}>（座號輸入 1 會自動變成 01）</span>
                  </div>
                  <textarea
                    style={{width:'100%',minHeight:120,padding:'10px 12px',borderRadius:8,border:'1.5px solid var(--border)',
                      fontSize:13,fontFamily:"'DM Mono',monospace",resize:'vertical',lineHeight:1.7}}
                    placeholder={"01 陳小明\n02 李美華\n03 張大偉\n04 王小花"}
                    value={pasteRoster}
                    onChange={e=>setPasteRoster(e.target.value)}
                  />
                  <button className="btn btn-primary btn-sm" style={{marginTop:8}}
                    onClick={()=>handlePasteRoster(cls)} disabled={!pasteRoster.trim()}>
                    匯入名單
                  </button>
                </div>
              )}
            </div>
          )}
          <div className="roster-student-grid">
            {students.map(s=>(
              <div key={s.seat} className="roster-student">
                <span className="roster-seat">{s.seat}</span>
                <span className="roster-name">{s.name}</span>
                <span style={{marginLeft:'auto',cursor:'pointer',color:'var(--danger)',fontSize:12}} onClick={()=>handleDeleteStudent(cls,s.seat)}>✕</span>
              </div>
            ))}
          </div>
        </div>
      ))}
      {toast && <Toast msg={toast}/>}
    </div>
  )
}

// ─── Dashboard ────────────────────────────────────────────────────────────────
function Dashboard({ user }) {
  const [quizzes, setQuizzes] = useState([])
  const [responseCounts, setResponseCounts] = useState({})
  const [loading, setLoading] = useState(true)
  const [toast, setToast] = useState('')
  const [deleteTarget, setDeleteTarget] = useState(null)
  const [qrTarget, setQrTarget] = useState(null)
  const [editTarget, setEditTarget] = useState(null)
  const [statusChanging, setStatusChanging] = useState(null)

  useEffect(() => {
    async function load() {
      try {
        const snap = await getDocs(query(collection(db,'quizzes'), where('teacherId','==',user.uid)))
        const list = snap.docs.map(d=>({id:d.id,...d.data()}))
          .sort((a,b)=>(b.createdAt?.seconds||0)-(a.createdAt?.seconds||0))
        setQuizzes(list)
        const counts = {}
        await Promise.all(list.map(async q => {
          const rs = await getDocs(query(collection(db,'responses'), where('quizId','==',q.id)))
          counts[q.id] = rs.size
        }))
        setResponseCounts(counts)
      } catch(e){console.error(e)}
      setLoading(false)
    }
    load()
  }, [user.uid])

  const showToast = (msg) => { setToast(msg); setTimeout(()=>setToast(''),2500) }

  const copyLink = (id) => {
    navigator.clipboard.writeText(`${window.location.origin}${window.location.pathname}#/s/${id}`)
    showToast('連結已複製！')
  }

  const handleDeleteConfirm = async () => {
    if (!deleteTarget) return
    try {
      const rSnap = await getDocs(query(collection(db,'responses'), where('quizId','==',deleteTarget.id)))
      await Promise.all(rSnap.docs.map(d=>deleteDoc(d.ref)))
      await deleteDoc(doc(db,'quizzes',deleteTarget.id))
      setQuizzes(prev=>prev.filter(q=>q.id!==deleteTarget.id))
      showToast(`「${deleteTarget.title}」已刪除`)
    } catch(e){console.error(e); alert('刪除失敗')}
    setDeleteTarget(null)
  }

  const handleStatusChange = async (quiz, newStatus) => {
    setStatusChanging(quiz.id)
    try {
      await setDoc(doc(db,'quizzes',quiz.id), {...quiz, status:newStatus, updatedAt:serverTimestamp()})
      setQuizzes(prev=>prev.map(q=>q.id===quiz.id?{...q,status:newStatus}:q))
      const labels = {draft:'草稿',active:'派發中',ended:'已結束'}
      showToast(`「${quiz.title}」已設為${labels[newStatus]}`)
    } catch(e){console.error(e)}
    setStatusChanging(null)
  }

  const handleEditSave = (updated) => {
    setQuizzes(prev=>prev.map(q=>q.id===updated.id?{...q,...updated}:q))
    setEditTarget(null)
    showToast('已儲存變更')
  }

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>

  // Group by status for better UX
  const activeQuizzes = quizzes.filter(q=>getStatus(q)==='active')
  const draftQuizzes  = quizzes.filter(q=>getStatus(q)==='draft')
  const endedQuizzes  = quizzes.filter(q=>getStatus(q)==='ended')
  const orderedQuizzes = [...activeQuizzes, ...draftQuizzes, ...endedQuizzes]

  const QuizCard = ({ q }) => {
    const st = getStatus(q)
    const stInfo = STATUS[st]
    const isEnded = st === 'ended'
    const isChanging = statusChanging === q.id

    return (
      <div className="quiz-card" style={{opacity: isEnded ? 0.8 : 1}}>
        {/* Top row: status badge + delete */}
        <div style={{display:'flex',alignItems:'center',justifyContent:'space-between',marginBottom:10}}>
          <span className={`status-badge status-${st}`}>{stInfo.emoji} {stInfo.label}</span>
          <button onClick={()=>setDeleteTarget({id:q.id,title:q.title})}
            style={{background:'none',border:'none',cursor:'pointer',fontSize:14,color:'var(--ink2)',padding:'2px 4px',lineHeight:1}}>🗑️</button>
        </div>

        {/* Title + subject */}
        <div style={{marginBottom:4}}>
          <div style={{display:'flex',alignItems:'center',gap:6,marginBottom:4}}>
            <span className="quiz-tag-pill" style={{marginBottom:0}}>{q.subject||'未分類'}</span>
            {q.settings?.allowMultipleAttempts===false
              ? <span style={{fontSize:10,color:'var(--danger)',background:'#ffeee8',padding:'1px 6px',borderRadius:10,fontWeight:700}}>🔒 限一次</span>
              : <span style={{fontSize:10,color:'var(--accent)',background:'#e8f5ee',padding:'1px 6px',borderRadius:10,fontWeight:700}}>🔄 可多次</span>
            }
          </div>
          <div className="quiz-name">{q.title}</div>
        </div>

        {/* Meta */}
        <div className="quiz-meta" style={{marginBottom:6}}>
          <span>📝 {q.questions?.length||0} 題</span>
          <span>👥 {responseCounts[q.id]||0} 人次</span>
          <span>🗓 {fmtDate(q.createdAt)}</span>
        </div>
        {/* Time info */}
        {(q.settings?.startTime||q.settings?.endTime) && (
          <div style={{fontSize:11,color:'var(--ink2)',marginBottom:6,display:'flex',gap:10,flexWrap:'wrap'}}>
            {q.settings.startTime && <span>🕐 開始 {q.settings.startTime.replace('T',' ').slice(0,16)}</span>}
            {q.settings.endTime   && <span>🕐 截止 {q.settings.endTime.replace('T',' ').slice(0,16)}</span>}
          </div>
        )}

        {/* Link (only show when active) */}
        {st === 'active' && (
          <div className="quiz-url" style={{marginBottom:10}}>
            <span style={{overflow:'hidden',textOverflow:'ellipsis',whiteSpace:'nowrap'}}>...#/s/{q.id}</span>
            <span className="copy-btn" onClick={()=>copyLink(q.id)}>複製連結</span>
          </div>
        )}

        {/* Status action buttons */}
        <div style={{display:'flex',gap:6,marginBottom:8,flexWrap:'wrap'}}>
          {st === 'draft' && (
            <button className="btn btn-primary btn-sm" style={{flex:1}} disabled={isChanging}
              onClick={()=>handleStatusChange(q,'active')}>
              {isChanging ? '...' : '▶ 開始派發'}
            </button>
          )}
          {st === 'active' && <>
            <button className="btn btn-sm" style={{flex:1,background:'#fff8e1',color:'#a67c00',border:'1.5px solid #f0d060'}}
              disabled={isChanging} onClick={()=>handleStatusChange(q,'draft')}>
              {isChanging ? '...' : '⏸ 暫停'}
            </button>
            <button className="btn btn-sm" style={{flex:1,background:'#ffeee8',color:'var(--danger)',border:'1.5px solid #fcd5c5'}}
              disabled={isChanging} onClick={()=>{ if(confirm('確定結束這份測驗？學生將無法再作答。')) handleStatusChange(q,'ended') }}>
              ⏹ 結束
            </button>
          </>}
          {st === 'ended' && (
            <button className="btn btn-secondary btn-sm" style={{flex:1}} disabled={isChanging}
              onClick={()=>handleStatusChange(q,'active')}>
              {isChanging ? '...' : '↺ 重新開放'}
            </button>
          )}
        </div>

        {/* Bottom row: view + edit + QR */}
        <div className="quiz-card-footer">
          <button className="btn btn-secondary btn-sm" style={{flex:1}} onClick={()=>navigate(`/results/${q.id}`)}>📊 成績</button>
          <button className="btn btn-secondary btn-sm" style={{flex:1}} onClick={()=>navigate(`/analytics/${q.id}`)}>📈 分析</button>
          {!isEnded && (
            <button className="btn btn-secondary btn-sm" onClick={()=>setEditTarget(q)} title="編輯">✏️</button>
          )}
          <button className="btn btn-secondary btn-sm" onClick={()=>setQrTarget(q)} title="QR Code">📱</button>
        </div>
      </div>
    )
  }

  return (
    <div>
      <div className="page-header">
        <div><div className="page-title">我的測驗</div>
          <div className="page-sub">
            {activeQuizzes.length>0 && <span style={{color:'var(--accent)',fontWeight:600}}>🟢 {activeQuizzes.length} 份派發中 </span>}
            共 {quizzes.length} 份測驗
          </div>
        </div>
        <button className="btn btn-primary" onClick={()=>navigate('/create')}>＋ 新增測驗</button>
      </div>

      {quizzes.length===0 ? (
        <div className="empty-state">
          <div className="empty-icon">📋</div>
          <div style={{fontSize:16,fontWeight:600,marginBottom:8}}>還沒有測驗</div>
          <button className="btn btn-primary" onClick={()=>navigate('/create')}>＋ 新增測驗</button>
        </div>
      ) : (
        <div className="quiz-grid">
          {orderedQuizzes.map(q=><QuizCard key={q.id} q={q}/>)}
          <div className="new-quiz-card" onClick={()=>navigate('/create')}>
            <span style={{fontSize:20}}>＋</span>建立新測驗
          </div>
        </div>
      )}

      {toast && <Toast msg={toast}/>}
      {deleteTarget && <ConfirmModal title="刪除測驗"
        message={`確定要刪除「${deleteTarget.title}」嗎？\n所有學生作答紀錄也會一併刪除，無法復原。`}
        onConfirm={handleDeleteConfirm} onCancel={()=>setDeleteTarget(null)}/>}
      {qrTarget && <QRModal
        url={`${window.location.origin}${window.location.pathname}#/s/${qrTarget.id}`}
        quizTitle={qrTarget.title} onClose={()=>setQrTarget(null)}/>}
      {editTarget && <EditQuizModal quiz={editTarget} onSave={handleEditSave} onCancel={()=>setEditTarget(null)}/>}
    </div>
  )
}

// ─── Create Quiz ──────────────────────────────────────────────────────────────
function CreateQuiz({ user }) {
  const [mode, setMode] = useState('manual')
  const [title, setTitle] = useState('')
  const [subject, setSubject] = useState('')
  const [allowHint, setAllowHint] = useState(true)
  const [showExplAfter, setShowExplAfter] = useState(true)
  const [showCorrect, setShowCorrect] = useState(true)
  const [allowMultipleAttempts, setAllowMultipleAttempts] = useState(true)
  const [selectedClasses, setSelectedClasses] = useState(new Set())
  const [useRoster, setUseRoster] = useState(false)
  const [allowNameEdit, setAllowNameEdit] = useState(true)
  const [startTime, setStartTime] = useState('')
  const [endTime, setEndTime] = useState('')
  const [rosterClasses, setRosterClasses] = useState([])
  const [questions, setQuestions] = useState([emptyQ()])
  const [saving, setSaving] = useState(false)
  const [createdId, setCreatedId] = useState(null)
  const [toast, setToast] = useState('')
  const [pasteText, setPasteText] = useState('')
  const [pastePreview, setPastePreview] = useState(null)

  const addQ = () => setQuestions(p=>[...p,emptyQ()])
  const removeQ = i => setQuestions(p=>p.filter((_,idx)=>idx!==i))
  const updateQ = (i,f,v) => setQuestions(p=>{const q=[...p];q[i]={...q[i],[f]:v};return q})
  const updateOpt = (qi,oi,v) => setQuestions(p=>{const q=[...p];q[qi].options[oi]=v;return q})
  const toggleField = (i,f) => setQuestions(p=>{const q=[...p];q[i]={...q[i],[f]:!q[i][f]};return q})

  useEffect(() => {
    if (!user?.uid) return
    getDoc(doc(db,'rosters',user.uid)).then(snap => {
      if (snap.exists()) setRosterClasses(Object.keys(snap.data().classes||{}).sort())
    }).catch(()=>{})
  }, [user?.uid])

  const toggleClass = (cls) => setSelectedClasses(prev => {
    const next = new Set(prev)
    if (next.has(cls)) next.delete(cls); else next.add(cls)
    return next
  })

  const handlePasteChange = (text) => {
    setPasteText(text)
    setPastePreview(text.trim() ? parsePasteText(text) : null)
  }
  const handlePasteImport = () => {
    if (!pastePreview?.length) return
    setQuestions(pastePreview); setPasteText(''); setPastePreview(null); setMode('manual')
    setToast(`成功匯入 ${pastePreview.length} 道題目`); setTimeout(()=>setToast(''),2000)
  }

  const handleExcel = (e) => {
    const file = e.target.files[0]; if (!file) return; e.target.value=''
    const reader = new FileReader()
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(ev.target.result, {type:'binary'})
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1,raw:false,defval:''})
        const parsed = parseQuestionRows(rows)
        if (parsed.length) { setQuestions(parsed); setMode('manual'); setToast(`匯入 ${parsed.length} 道題目`); setTimeout(()=>setToast(''),2000) }
      } catch { alert('Excel 格式錯誤') }
    }
    reader.readAsBinaryString(file)
  }

  const handleSave = async () => {
    if (!title.trim()) { alert('請輸入測驗名稱'); return }
    if (questions.some(q=>!q.text.trim())) { alert('請填寫所有題目'); return }
    setSaving(true)
    try {
      const clean = questions.map(({showHint,showExpl,...q})=>q)
      const classes = useRoster ? [...selectedClasses] : []
      const docRef = await addDoc(collection(db,'quizzes'), {
        teacherId:user.uid, teacherName:user.displayName||user.email,
        title:title.trim(), subject:subject.trim()||'未分類', questions:clean,
        status: 'draft',
        settings:{ allowHint, showExplAfter, showCorrect, allowMultipleAttempts, allowedClasses:classes, useRoster, allowNameEdit, startTime, endTime },
      })
      setCreatedId(docRef.id)
    } catch(e) { alert('儲存失敗，請確認 Firebase 設定'); console.error(e) }
    setSaving(false)
  }

  if (createdId) {
    const url = `${window.location.origin}${window.location.pathname}#/s/${createdId}`
    return (
      <div style={{maxWidth:480,margin:'60px auto',textAlign:'center'}}>
        <div style={{fontSize:48,marginBottom:16}}>✅</div>
        <div style={{fontSize:22,fontWeight:700,marginBottom:8}}>測驗已建立（草稿）</div>
        <div style={{fontSize:14,color:'var(--ink2)',marginBottom:20}}>
          目前為草稿狀態，學生尚無法作答。<br/>回到首頁後點「▶ 開始派發」即可開放。
        </div>
        <div className="card" style={{padding:20,marginBottom:16,textAlign:'left'}}>
          <div style={{fontSize:12,color:'var(--ink2)',marginBottom:6}}>學生作答連結（派發後才有效）</div>
          <div style={{fontFamily:"'DM Mono',monospace",fontSize:12,background:'#f5f3ef',padding:'10px 14px',borderRadius:8,marginBottom:12,wordBreak:'break-all'}}>{url}</div>
          <button className="btn btn-secondary" style={{width:'100%'}} onClick={()=>navigator.clipboard.writeText(url)}>複製連結</button>
        </div>
        <div style={{display:'flex',gap:10,justifyContent:'center'}}>
          <button className="btn btn-primary" onClick={()=>navigate('/')}>回到首頁開始派發 →</button>
          <button className="btn btn-secondary" onClick={()=>{setCreatedId(null);setTitle('');setSubject('');setQuestions([emptyQ()])}}>再建一份</button>
        </div>
      </div>
    )
  }

  return (
    <div style={{maxWidth:720}}>
      <div className="page-header">
        <div><div className="page-title">新增測驗</div><div className="page-sub">手動輸入、上傳 Excel 或貼上 AI 生成的題目</div></div>
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
            <div style={{fontSize:14,fontWeight:600,marginBottom:6}}>點擊上傳 Excel（支援 .xlsx/.xls）</div>
            <div style={{fontSize:12,color:'var(--ink2)'}}>分數如 3/8 會正確保留，不會變成小數</div>
            <input type="file" accept=".xlsx,.xls" style={{display:'none'}} onChange={handleExcel}/>
          </label>
          <div style={{marginTop:16,fontSize:12,fontWeight:600,color:'var(--ink2)',marginBottom:8}}>欄位順序（第一列為標題）：</div>
          <div style={{overflowX:'auto'}}>
            <table style={{fontSize:11}}>
              <thead><tr>
                {['題目','選A','選B','選C','選D','正確(A/B/C/D)','配分','提示(選填)','解析(選填)'].map(h=>(
                  <th key={h} style={{padding:'5px 8px',background:'#f0ede8',border:'1px solid var(--border)',whiteSpace:'nowrap'}}>{h}</th>
                ))}
              </tr></thead>
              <tbody><tr>
                {['算算看3/8+...','1/4','3/8','1/2','5/8','B','10','分母相同...','分母不變...'].map((c,i)=>(
                  <td key={i} style={{padding:'5px 8px',border:'1px solid var(--border)',color:'var(--ink2)',whiteSpace:'nowrap'}}>{c}</td>
                ))}
              </tr></tbody>
            </table>
          </div>
        </div>
      )}

      {mode==='paste' && (
        <div className="card" style={{marginBottom:20}}>
          <BatchPasteSection onImport={parsed => { setQuestions(parsed); setMode('manual') }}/>
        </div>
      )}

      <div className="card">
        <div style={{marginBottom:20}}>
          <div className="form-row">
            <div><label className="form-label">測驗名稱 *</label>
              <input className="form-input" placeholder="例：分數運算第一章測驗" value={title} onChange={e=>setTitle(e.target.value)}/></div>
            <div><label className="form-label">科目</label>
              <input className="form-input" placeholder="例：數學" value={subject} onChange={e=>setSubject(e.target.value)}/></div>
          </div>
        </div>

        {/* Settings */}
        <div style={{background:'#f5f3ef',borderRadius:10,padding:16,marginBottom:20}}>
          <div style={{fontSize:13,fontWeight:700,marginBottom:12}}>⚙️ 測驗設定</div>
          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:10,marginBottom:14}}>
            {[
              [allowHint,setAllowHint,'💡 允許學生查看提示'],
              [showExplAfter,setShowExplAfter,'📖 提交後顯示解析'],
              [showCorrect,setShowCorrect,'✓ 顯示正確答案'],
              [allowMultipleAttempts,setAllowMultipleAttempts,'🔄 允許重複作答'],
            ].map(([v,s,l])=>(
              <label key={l} style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer',padding:'6px 0'}}>
                <input type="checkbox" checked={v} onChange={e=>s(e.target.checked)} style={{accentColor:'var(--accent)',width:15,height:15}}/>{l}
              </label>
            ))}
          </div>

          <div style={{marginBottom:12}}>
            <label className="form-label" style={{fontSize:12}}>學生資料輸入方式</label>
            <div style={{display:'flex',gap:8}}>
              {[['free','✎ 自由輸入'],['roster','📚 使用班級名單']].map(([val,label])=>(
                <button key={val} onClick={()=>setUseRoster(val==='roster')}
                  style={{padding:'7px 16px',borderRadius:8,fontSize:13,fontWeight:600,cursor:'pointer',fontFamily:"'Noto Sans TC',sans-serif",
                    background:(!useRoster&&val==='free')||(useRoster&&val==='roster')?'var(--accent)':'white',
                    color:(!useRoster&&val==='free')||(useRoster&&val==='roster')?'white':'var(--ink2)',
                    border:(!useRoster&&val==='free')||(useRoster&&val==='roster')?'2px solid var(--accent)':'2px solid var(--border)',
                    transition:'all .15s'}}>
                  {label}
                </button>
              ))}
            </div>
          </div>

          {useRoster && (
            <div style={{marginBottom:12}}>
              <label className="form-label" style={{fontSize:12}}>開放班級（留空 = 全部班級均可進入）</label>
              {rosterClasses.length === 0 ? (
                <div style={{fontSize:13,color:'var(--ink2)',padding:'8px 12px',background:'white',borderRadius:8,border:'1px solid var(--border)'}}>
                  尚未建立班級名單，請先到「班級名單」頁面新增
                </div>
              ) : (
                <div style={{display:'flex',gap:8,flexWrap:'wrap'}}>
                  {rosterClasses.map(cls => (
                    <button key={cls} onClick={()=>toggleClass(cls)}
                      style={{padding:'6px 16px',borderRadius:20,fontSize:13,fontWeight:700,cursor:'pointer',fontFamily:"'Noto Sans TC',sans-serif",transition:'all .15s',
                        background:selectedClasses.has(cls)?'var(--accent)':'white',
                        color:selectedClasses.has(cls)?'white':'var(--ink2)',
                        border:selectedClasses.has(cls)?'2px solid var(--accent)':'2px solid var(--border)'}}>
                      {selectedClasses.has(cls)?'✓ ':''}{cls}班
                    </button>
                  ))}
                </div>
              )}
              <label style={{display:'flex',alignItems:'center',gap:8,fontSize:13,cursor:'pointer',marginTop:10}}>
                <input type="checkbox" checked={allowNameEdit} onChange={e=>setAllowNameEdit(e.target.checked)} style={{accentColor:'var(--accent)',width:15,height:15}}/>
                允許學生修改自動帶出的姓名
              </label>
            </div>
          )}

          <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:12}}>
            <div>
              <label className="form-label" style={{fontSize:12}}>🕐 開始時間（選填）</label>
              <input className="form-input" type="datetime-local" value={startTime} onChange={e=>setStartTime(e.target.value)}
                style={{fontSize:13,padding:'8px 12px'}}/>
            </div>
            <div>
              <label className="form-label" style={{fontSize:12}}>🕐 截止時間（選填）</label>
              <input className="form-input" type="datetime-local" value={endTime} onChange={e=>setEndTime(e.target.value)}
                style={{fontSize:13,padding:'8px 12px'}}/>
            </div>
          </div>
        </div>

        <div style={{borderTop:'1px solid var(--border)',paddingTop:20}}>
          <div style={{fontSize:14,fontWeight:700,marginBottom:16}}>📝 題目設定（支援 $公式$，每題右上角可預覽）</div>
          {questions.map((q,qi)=>(
            <QuestionEditorCard key={qi} q={q} qi={qi} total={questions.length}
              updateQ={updateQ} updateOpt={updateOpt} toggleField={toggleField} removeQ={removeQ}/>
          ))}
          <button className="btn btn-secondary btn-sm" onClick={addQ}>＋ 新增題目</button>
          <BatchPasteSection appendMode={true} onImport={parsed => setQuestions(prev=>[...prev,...parsed])}/>
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
      setQuiz({id:qSnap.id,...qSnap.data()})
      const rSnap = await getDocs(query(collection(db,'responses'),where('quizId','==',quizId)))
      setResponses(rSnap.docs.map(d=>({id:d.id,...d.data()})).sort((a,b)=>(a.submittedAt?.seconds||0)-(b.submittedAt?.seconds||0)))
      setLoading(false)
    }
    load()
  }, [quizId])

  const studentGroups = useMemo(() => {
    if (!quiz) return []
    const groups = {}
    responses.forEach(r => {
      const key = `${r.class}-${padSeat(r.seat)}`
      if (!groups[key]) groups[key] = { class:r.class, seat:padSeat(r.seat), name:r.name, attempts:[] }
      groups[key].attempts.push(r)
    })
    return Object.values(groups)
      .sort((a,b)=>a.class.localeCompare(b.class)||a.seat.localeCompare(b.seat))
      .map(g => {
        const scores = g.attempts.map(a=>a.score)
        return { ...g, firstScore:scores[0]??'-', bestScore:Math.max(...scores), latestScore:scores[scores.length-1]??'-', count:g.attempts.length }
      })
  }, [responses, quiz])

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>
  if (!quiz) return <div className="empty-state"><div>找不到這份測驗</div></div>

  const allowMulti = quiz.settings?.allowMultipleAttempts !== false
  const classes = ['all',...new Set(studentGroups.map(g=>g.class).filter(Boolean))]
  const filtered = filter==='all' ? studentGroups : studentGroups.filter(g=>g.class===filter)
  const avgBest = filtered.length ? Math.round(filtered.reduce((s,g)=>s+(allowMulti?g.bestScore:g.firstScore),0)/filtered.length) : 0

  return (
    <div>
      <div className="page-header">
        <div>
          <div style={{fontSize:13,color:'var(--ink2)',cursor:'pointer',marginBottom:6}} onClick={()=>navigate('/')}>← 返回</div>
          <div className="page-title">{quiz.title}</div>
          <div className="page-sub">成績查詢 {allowMulti?'（可重複作答）':'（限答一次）'}</div>
        </div>
        <button className="btn btn-excel" onClick={()=>exportToExcel(quiz,responses)}>📥 匯出 Excel</button>
      </div>
      <div className="stats-row">
        {[
          {label:'學生人數',value:studentGroups.length,sub:'位'},
          {label:allowMulti?'平均最高分':'平均分數',value:avgBest,sub:'分',color:'var(--accent)'},
          {label:'最高分',value:studentGroups.length?Math.max(...studentGroups.map(g=>g.bestScore)):'-',sub:'分'},
          {label:'及格率',value:studentGroups.length?Math.round(studentGroups.filter(g=>g.bestScore>=60).length/studentGroups.length*100)+'%':'-',sub:'≥60分'},
        ].map(s=>(
          <div key={s.label} className="stat-box">
            <div className="stat-label">{s.label}</div>
            <div className="stat-value" style={{color:s.color||'var(--ink)'}}>{s.value}</div>
            <div className="stat-sub">{s.sub}</div>
          </div>
        ))}
      </div>
      <div className="card">
        <div style={{display:'flex',alignItems:'center',gap:8,marginBottom:16,flexWrap:'wrap'}}>
          <span style={{fontSize:13,fontWeight:600}}>篩選班級：</span>
          {classes.map(c=>(
            <button key={c} className={`btn btn-sm ${filter===c?'btn-primary':'btn-secondary'}`} onClick={()=>setFilter(c)}>
              {c==='all'?'全部':`${c}班`}
            </button>
          ))}
        </div>
        {filtered.length===0 ? <div className="empty-state" style={{padding:40}}><div className="empty-icon">📭</div><div>目前沒有作答紀錄</div></div>
        : <div style={{overflowX:'auto'}}>
          <table>
            <thead><tr>
              <th>班級</th><th>座號</th><th>姓名</th>
              {allowMulti ? <><th>第一次</th><th style={{color:'var(--accent)'}}>最高分</th><th>作答次數</th></> : <th>成績</th>}
              <th>答題狀況</th>
            </tr></thead>
            <tbody>
              {filtered.map((g,i)=>(
                <tr key={i}>
                  <td><span className="tag tag-class">{g.class}班</span></td>
                  <td style={{fontFamily:"'DM Mono',monospace"}}>{g.seat}</td>
                  <td style={{fontWeight:500}}>{g.name}</td>
                  {allowMulti ? <>
                    <td><span className={`score-badge ${scoreClass(g.firstScore)}`}>{g.firstScore}分</span></td>
                    <td><span className={`score-badge ${scoreClass(g.bestScore)}`} style={{fontWeight:700}}>{g.bestScore}分</span></td>
                    <td style={{color:'var(--ink2)',fontSize:13}}>{g.count}次</td>
                  </> : <td><span className={`score-badge ${scoreClass(g.firstScore)}`}>{g.firstScore}分</span></td>}
                  <td>
                    <div className="answer-dots">
                      {quiz.questions.map((q,qi)=>{
                        const lastAns = g.attempts[g.attempts.length-1]?.answers?.[qi]
                        return <div key={qi} className={`dot ${lastAns===q.correct?'dot-correct':'dot-wrong'}`} title={`第${qi+1}題`}>
                          {lastAns===q.correct?'✓':'✗'}
                        </div>
                      })}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>}
      </div>
    </div>
  )
}

// ─── Analytics ────────────────────────────────────────────────────────────────
function Analytics({ quizId }) {
  const [quiz, setQuiz] = useState(null)
  const [responses, setResponses] = useState([])
  const [loading, setLoading] = useState(true)
  const [tab, setTab] = useState('chart')
  const [shownStudents, setShownStudents] = useState({})

  useEffect(() => {
    async function load() {
      const qSnap = await getDoc(doc(db,'quizzes',quizId))
      if (!qSnap.exists()) { setLoading(false); return }
      setQuiz({id:qSnap.id,...qSnap.data()})
      const rSnap = await getDocs(query(collection(db,'responses'),where('quizId','==',quizId)))
      setResponses(rSnap.docs.map(d=>({id:d.id,...d.data()})))
      setLoading(false)
    }
    load()
  }, [quizId])

  const n = responses.length
  const stats = useMemo(() => {
    if (!quiz) return []
    return quiz.questions.map((q,qi) => {
      const correct = responses.filter(r=>r.answers?.[qi]===q.correct).length
      const rate = n ? Math.round(correct/n*100) : 0
      const errRate = 100 - rate
      const optCounts = [0,1,2,3].map(oi=>responses.filter(r=>r.answers?.[qi]===oi).length)
      const wrongStudents = responses.filter(r=>r.answers?.[qi]!==q.correct)
        .map(r=>({seat:padSeat(r.seat),class:r.class,name:r.name}))
        .sort((a,b)=>a.class.localeCompare(b.class)||a.seat.localeCompare(b.seat))
      return { q, rate, errRate, correct, optCounts, wrongStudents, colorClass:errorColor(errRate) }
    })
  }, [quiz, responses])

  // colorLabel unused, errEmoji handles display

  if (loading) return <div className="loading"><div className="spinner"/>載入中...</div>
  if (!quiz) return <div className="empty-state"><div>找不到這份測驗</div></div>

  const errRateClass = (cls) => {
    if (cls==='err-red') return 'err-rate-red'
    if (cls==='err-orange') return 'err-rate-orange'
    if (cls==='err-yellow') return 'err-rate-yellow'
    return 'err-rate-green'
  }
  const errEmoji = (cls) => ({
    'err-red':'🔴','err-orange':'🟠','err-yellow':'🟡','err-green':'🟢'
  }[cls]||'')

  return (
    <div>
      <div className="page-header">
        <div>
          <div style={{fontSize:13,color:'var(--ink2)',cursor:'pointer',marginBottom:6}} onClick={()=>navigate('/')}>← 返回</div>
          <div className="page-title">{quiz.title}</div>
          <div className="page-sub">答題分析 — {n} 人次作答</div>
        </div>
        <button className="btn btn-excel" onClick={()=>{
          getDocs(query(collection(db,'responses'),where('quizId','==',quizId))).then(s=>{
            exportToExcel(quiz, s.docs.map(d=>({id:d.id,...d.data()})))
          })
        }}>📥 匯出 Excel</button>
      </div>

      <div style={{display:'flex',gap:12,marginBottom:20,flexWrap:'wrap'}}>
        {[{label:'🔴 高錯誤（80%+）',count:stats.filter(s=>s.colorClass==='err-red').length,color:'var(--danger)'},
          {label:'🟠 中錯誤（50%+）',count:stats.filter(s=>s.colorClass==='err-orange').length,color:'var(--warn)'},
          {label:'🟡 低錯誤（30%+）',count:stats.filter(s=>s.colorClass==='err-yellow').length,color:'#e9c32a'},
          {label:'🟢 良好（<30%）',count:stats.filter(s=>s.colorClass==='err-green').length,color:'var(--accent2)'},
        ].map(s=>(
          <div key={s.label} style={{background:'white',borderRadius:8,padding:'8px 14px',border:'1px solid var(--border)',fontSize:13}}>
            {s.label}: <strong style={{color:s.color}}>{s.count}</strong> 題
          </div>
        ))}
      </div>

      <div className="tabs">
        <div className={`tab ${tab==='chart'?'active':''}`} onClick={()=>setTab('chart')}>📊 答題統計</div>
        <div className={`tab ${tab==='review'?'active':''}`} onClick={()=>setTab('review')}>🔍 錯題檢討</div>
      </div>

      {tab==='chart' && (
        <div style={{display:'grid',gridTemplateColumns:'1fr 1fr',gap:16,marginBottom:24}}>
          <div className="card">
            <div style={{fontSize:14,fontWeight:700,marginBottom:14}}>各題答對率</div>
            <div className="q-analytics">
              {stats.map((s,i)=>(
                <div key={i} className={`q-row ${s.colorClass}`}>
                  <div className="q-head">
                    <div className="q-text-sm">{errEmoji(s.colorClass)} 第{i+1}題：<MathText text={s.q.text.length>18?s.q.text.slice(0,18)+'...':s.q.text}/></div>
                    <div className={`err-rate ${errRateClass(s.colorClass)}`}>{s.rate}%</div>
                  </div>
                  <div className="q-bar-bg">
                    <div className={`q-bar ${s.rate>=70?'bar-green':'bar-red'}`} style={{width:`${s.rate}%`}}/>
                  </div>
                </div>
              ))}
            </div>
          </div>
          <div className="card">
            <div style={{fontSize:14,fontWeight:700,marginBottom:14}}>選項分佈</div>
            <div className="q-analytics">
              {stats.map((s,i)=>(
                <div key={i} className={`q-row ${s.colorClass}`}>
                  <div className="q-head" style={{marginBottom:8}}>
                    <div className="q-text-sm">第{i+1}題 {errEmoji(s.colorClass)}</div>
                    <div style={{fontSize:12,color:'var(--ink2)'}}>{s.correct}/{n} 答對</div>
                  </div>
                  <div className="q-opts-row">
                    {s.q.options.map((opt,oi)=>(
                      <div key={oi} className={`q-opt-chip ${oi===s.q.correct?'correct':''}`}>
                        {ABCD[oi]}.<MathText text={opt.slice(0,5)}/> ({s.optCounts[oi]}人)
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>
      )}

      {tab==='review' && (
        <div>
          <div style={{fontSize:13,color:'var(--ink2)',marginBottom:16,padding:'10px 14px',background:'white',borderRadius:8,border:'1px solid var(--border)'}}>
            🔴 80%以上答錯 ・ 🟠 50%以上答錯 ・ 🟡 30%以上答錯 ・ 🟢 答對率良好<br/>
            點擊「顯示答錯學生」可查看各班座號
          </div>
          <div className="q-analytics">
            {[...stats].sort((a,b)=>a.rate-b.rate).map((s,i)=>{
              const qi = quiz.questions.indexOf(s.q)
              const key = `q_${qi}`
              return (
                <div key={i} className={`q-row ${s.colorClass}`}>
                  <div className="q-head">
                    <div style={{flex:1}}>
                      <div style={{fontSize:12,color:'var(--ink2)',marginBottom:4}}>
                        {errEmoji(s.colorClass)} 第{qi+1}題 · {s.q.points}分
                        <span className={`err-rate ${errRateClass(s.colorClass)}`} style={{marginLeft:8}}>答錯率 {s.errRate}%</span>
                      </div>
                      <div style={{fontSize:14,fontWeight:600,lineHeight:1.6}}>
                        <MathText text={s.q.text}/>
                      </div>
                    </div>
                  </div>
                  <div className="q-opts-row" style={{marginBottom:8}}>
                    {s.q.options.map((opt,oi)=>(
                      <div key={oi} className={`q-opt-chip ${oi===s.q.correct?'correct':''}`} style={{padding:'4px 10px',fontSize:12}}>
                        {ABCD[oi]}. <MathText text={opt}/> ({s.optCounts[oi]}人)
                        {oi===s.q.correct && ' ✓'}
                      </div>
                    ))}
                  </div>
                  {s.q.explanation && (
                    <div style={{background:'white',borderRadius:6,padding:'8px 12px',fontSize:13,color:'var(--ink2)',marginBottom:8,border:'1px solid rgba(0,0,0,.08)'}}>
                      📖 <MathText text={s.q.explanation}/>
                    </div>
                  )}
                  <button className="btn btn-secondary btn-sm" onClick={()=>setShownStudents(prev=>({...prev,[key]:!prev[key]}))}>
                    {shownStudents[key]?'▼ 隱藏':'▶ 顯示答錯學生'} ({s.wrongStudents.length}人)
                  </button>
                  {shownStudents[key] && (
                    <div className="wrong-seats">
                      {s.wrongStudents.length===0 ? <span style={{fontSize:13,color:'var(--accent)'}}>全班答對！🎉</span>
                      : s.wrongStudents.map((w,wi)=>(
                        <span key={wi} className="wrong-seat-badge" title={w.name}>{w.class}-{w.seat}</span>
                      ))}
                    </div>
                  )}
                </div>
              )
            })}
          </div>
        </div>
      )}
    </div>
  )
}

// ─── Student Quiz ─────────────────────────────────────────────────────────────
function StudentQuiz({ quizId }) {
  const [step, setStep] = useState('info')
  const [info, setInfo] = useState({ name:'', class:'', seat:'' })
  const [quiz, setQuiz] = useState(null)
  const [roster, setRoster] = useState(null)
  const [loading, setLoading] = useState(true)
  const [answers, setAnswers] = useState({})
  const [shownHints, setShownHints] = useState({})
  const [score, setScore] = useState(0)
  const [saving, setSaving] = useState(false)
  const [sessionScores, setSessionScores] = useState([])
  const [alreadyAttempted, setAlreadyAttempted] = useState(false)

  useEffect(() => {
    async function load() {
      try {
        const snap = await getDoc(doc(db,'quizzes',quizId))
        if (snap.exists()) {
          const q = {id:snap.id,...snap.data()}
          setQuiz(q)
          if (q.settings?.useRoster && q.teacherId) {
            const rSnap = await getDoc(doc(db,'rosters',q.teacherId))
            if (rSnap.exists()) setRoster(rSnap.data().classes||{})
          }
        }
      } catch(e){console.error(e)}
      setLoading(false)
    }
    load()
  }, [quizId])

  // Auto-fill name from roster when seat changes
  useEffect(() => {
    if (roster && info.class && info.seat && quiz?.settings?.useRoster) {
      const student = roster[info.class]?.find(s=>s.seat===info.seat)
      if (student) setInfo(prev=>({...prev,name:student.name}))
    }
  }, [info.class, info.seat, roster])

  const s = quiz?.settings || {}
  const totalPoints = quiz?.questions?.reduce((sum,q)=>sum+q.points,0) || 0

  const availableClasses = useMemo(() => {
    if (!quiz) return []
    if (s.allowedClasses?.length) return s.allowedClasses
    if (s.useRoster && roster) return Object.keys(roster).sort()
    return []
  }, [quiz, roster])

  const availableSeats = useMemo(() => {
    if (s.useRoster && roster && info.class && roster[info.class]) {
      return roster[info.class].map(st=>st.seat)
    }
    return Array.from({length:50},(_,i)=>padSeat(i+1))
  }, [s.useRoster, roster, info.class])

  const useDropdowns = availableClasses.length > 0

  // localStorage key for this student+quiz
  const lsKey = (cls, seat) => `quizflow_submitted_${quizId}_${cls}_${seat}`

  const handleInfoSubmit = () => {
    const seatPadded = padSeat(info.seat)
    const finalInfo = {...info, seat: seatPadded}
    setInfo(finalInfo)

    // Single-attempt: check localStorage first (no DB read needed)
    if (s.allowMultipleAttempts === false) {
      if (localStorage.getItem(lsKey(finalInfo.class, seatPadded))) {
        setAlreadyAttempted(true)
        return
      }
    }
    setStep('quiz')
  }

  const handleSubmit = async () => {
    setSaving(true)
    let sc = 0
    quiz.questions.forEach((q,i) => { if (answers[i] === q.correct) sc += q.points })
    setScore(sc)
    setSessionScores(prev => [...prev, sc])

    try {
      const attemptNum = sessionScores.length + 1

      if (s.allowMultipleAttempts === false) {
        // Single-attempt: use fixed document ID to prevent duplicates at DB level
        // Firebase will reject a second write if the doc already exists via rules
        const fixedId = `${quizId}_${info.class}_${info.seat}`
        await setDoc(doc(db, 'responses', fixedId), {
          quizId, name: info.name, class: info.class, seat: info.seat,
          score: sc, answers: quiz.questions.map((_,i) => answers[i] ?? -1),
          submittedAt: serverTimestamp(), attemptNumber: 1,
        })
        // Mark in localStorage so same device is blocked instantly next time
        localStorage.setItem(lsKey(info.class, info.seat), '1')
      } else {
        // Multi-attempt: use addDoc with attempt number from session count
        await addDoc(collection(db, 'responses'), {
          quizId, name: info.name, class: info.class, seat: info.seat,
          score: sc, answers: quiz.questions.map((_,i) => answers[i] ?? -1),
          submittedAt: serverTimestamp(), attemptNumber: attemptNum,
        })
      }
    } catch(e) { console.error('儲存失敗', e) }

    setSaving(false)
    setStep('result')
    setTimeout(() => window.scrollTo({ top: 0, behavior: 'smooth' }), 50)
  }

  const handleRetry = () => {
    if (s.allowMultipleAttempts===false) { alert('此測驗只能作答一次'); return }
    setAnswers({}); setShownHints({}); setStep('quiz')
  }
  const handleGoHome = () => { setAnswers({}); setShownHints({}); setInfo({name:'',class:'',seat:''}); setSessionScores([]); setStep('info') }

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

  // Status guard
  const quizStatus = getStatus(quiz)
  if (quizStatus === 'draft') return (
    <div className="student-page">
      <div className="student-topbar"><div className="student-topbar-title">📋 {quiz.title}</div></div>
      <div className="student-body" style={{textAlign:'center',paddingTop:60}}>
        <div style={{fontSize:48,marginBottom:16}}>📝</div>
        <div style={{fontSize:20,fontWeight:700,marginBottom:8}}>測驗尚未開放</div>
        <div style={{fontSize:14,color:'var(--ink2)'}}>老師還在準備中，請稍後再試</div>
      </div>
    </div>
  )
  if (quizStatus === 'ended') return (
    <div className="student-page">
      <div className="student-topbar"><div className="student-topbar-title">📋 {quiz.title}</div></div>
      <div className="student-body" style={{textAlign:'center',paddingTop:60}}>
        <div style={{fontSize:48,marginBottom:16}}>🔴</div>
        <div style={{fontSize:20,fontWeight:700,marginBottom:8}}>測驗已結束</div>
        <div style={{fontSize:14,color:'var(--ink2)'}}>此測驗已關閉，無法作答</div>
      </div>
    </div>
  )

  // Time-based access guard
  const now = new Date()
  const { startTime, endTime } = quiz.settings || {}
  if (startTime && now < new Date(startTime)) {
    const fmt = startTime.replace('T',' ').slice(0,16)
    return (
      <div className="student-page">
        <div className="student-topbar"><div className="student-topbar-title">📋 {quiz.title}</div></div>
        <div className="student-body" style={{textAlign:'center',paddingTop:60}}>
          <div style={{fontSize:48,marginBottom:16}}>⏳</div>
          <div style={{fontSize:20,fontWeight:700,marginBottom:8}}>測驗尚未開始</div>
          <div style={{fontSize:14,color:'var(--ink2)'}}>開放時間：{fmt}</div>
        </div>
      </div>
    )
  }
  if (endTime && now > new Date(endTime)) {
    const fmt = endTime.replace('T',' ').slice(0,16)
    return (
      <div className="student-page">
        <div className="student-topbar"><div className="student-topbar-title">📋 {quiz.title}</div></div>
        <div className="student-body" style={{textAlign:'center',paddingTop:60}}>
          <div style={{fontSize:48,marginBottom:16}}>🔴</div>
          <div style={{fontSize:20,fontWeight:700,marginBottom:8}}>測驗已截止</div>
          <div style={{fontSize:14,color:'var(--ink2)'}}>截止時間：{fmt}</div>
        </div>
      </div>
    )
  }

  // Already attempted (single attempt mode)
  if (alreadyAttempted) return (
    <div className="student-page">
      <div className="student-topbar"><div className="student-topbar-title">{quiz.title}</div></div>
      <div className="student-body" style={{textAlign:'center',paddingTop:60}}>
        <div style={{fontSize:48,marginBottom:16}}>🔒</div>
        <div style={{fontSize:20,fontWeight:700,marginBottom:8}}>你已經作答過了</div>
        <div style={{fontSize:14,color:'var(--ink2)'}}>此測驗只能作答一次</div>
      </div>
    </div>
  )

  // Result page
  if (step==='result') {
    const correctCount = quiz.questions.filter((_,i)=>answers[i]===quiz.questions[i].correct).length
    const bestScore = sessionScores.length ? Math.max(...sessionScores) : score
    return (
      <>
      <div className="student-page">
        <div className="student-topbar">
          <div><div className="student-topbar-title">{quiz.title}</div></div>
        </div>
        <div className="student-body">
          <div className="result-card">
            <div style={{fontSize:16,fontWeight:700}}>✅ 作答完成！</div>
            <div style={{fontSize:14,color:'var(--ink2)',marginTop:6}}>{info.name} 同學</div>
            <div className="result-score">{score}</div>
            <div className="result-label">滿分 {totalPoints} 分</div>
            {sessionScores.length>1 && (
              <div style={{marginTop:12,fontSize:13,color:'var(--ink2)'}}>
                本次最高分：<strong style={{color:'var(--accent)'}}>{bestScore}分</strong>
                &nbsp;·&nbsp;已作答 {sessionScores.length} 次
              </div>
            )}
            <div className="result-breakdown">
              <div className="rb-item"><div className="rb-val" style={{color:'var(--accent)'}}>{correctCount}</div><div className="rb-lab">答對</div></div>
              <div className="rb-item"><div className="rb-val" style={{color:'var(--danger)'}}>{quiz.questions.length-correctCount}</div><div className="rb-lab">答錯</div></div>
              <div className="rb-item"><div className="rb-val">{totalPoints?Math.round(score/totalPoints*100):0}%</div><div className="rb-lab">正確率</div></div>
            </div>
            <div className="result-btns">
              {s.allowMultipleAttempts !== false && (
                <button onClick={handleRetry}
                  style={{flex:1,background:'var(--accent)',color:'white',padding:'12px',borderRadius:8,border:'none',fontSize:14,fontWeight:700,cursor:'pointer',fontFamily:"'Noto Sans TC',sans-serif"}}>
                  🔄 再作答一次
                </button>
              )}
              <button onClick={handleGoHome}
                style={{flex:1,background:'white',color:'var(--ink)',padding:'12px',borderRadius:8,border:'1.5px solid var(--border)',fontSize:14,fontWeight:600,cursor:'pointer',fontFamily:"'Noto Sans TC',sans-serif"}}>
                🏠 回到首頁
              </button>
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
                    <div className="sq-text"><MathText text={q.text}/></div>
                    <div className="sq-opts" style={{pointerEvents:'none'}}>
                      {q.options.map((opt,oi)=>{
                        let cls='revealed'
                        if(s.showCorrect!==false&&oi===q.correct) cls+=' correct-reveal'
                        else if(oi===answers[qi]) cls+=' wrong-reveal'
                        return (
                          <div key={oi} className={`sq-opt ${cls}`}>
                            <div className="opt-circle">{ABCD[oi]}</div>
                            <div className="opt-text"><MathText text={opt}/>
                              {s.showCorrect!==false&&oi===q.correct&&<span style={{fontSize:11,color:'var(--accent)',marginLeft:8,fontWeight:700}}>← 正確答案</span>}
                              {oi===answers[qi]&&oi!==q.correct&&<span style={{fontSize:11,color:'var(--danger)',marginLeft:8}}>← 你的選擇</span>}
                            </div>
                          </div>
                        )
                      })}
                    </div>
                    {s.showExplAfter!==false&&q.explanation&&(
                      isCorrect
                        ? <div className="expl-bubble"><div className="expl-bubble-title">📖 解析</div><div className="expl-bubble-text"><MathText text={q.explanation}/></div></div>
                        : <div className="wrong-expl-bubble"><div className="wrong-expl-title">📖 看看哪裡答錯了</div><div className="wrong-expl-text"><MathText text={q.explanation}/></div></div>
                    )}
                  </div>
                )
              })}
            </div>
          )}
        </div>
      </div>
      <button className="float-top-btn" onClick={()=>window.scrollTo({top:0,behavior:'smooth'})} title="回到頂部">↑</button>
      </>
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
              {/* Class */}
              <div>
                <label className="form-label" style={{fontSize:12}}>班級</label>
                {useDropdowns ? (
                  <select className="form-select" value={info.class} onChange={e=>setInfo({...info,class:e.target.value,seat:'',name:''})}>
                    <option value="">請選擇班級</option>
                    {availableClasses.map(c=><option key={c} value={c}>{c}班</option>)}
                  </select>
                ) : (
                  <input className="form-input" placeholder="例：103" value={info.class} onChange={e=>setInfo({...info,class:e.target.value})}/>
                )}
              </div>
              {/* Seat */}
              <div>
                <label className="form-label" style={{fontSize:12}}>座號</label>
                {useDropdowns ? (
                  <select className="form-select" value={info.seat} onChange={e=>setInfo({...info,seat:e.target.value})} disabled={!info.class}>
                    <option value="">請選擇座號</option>
                    {availableSeats.map(s=><option key={s} value={s}>{s}</option>)}
                  </select>
                ) : (
                  <input className="form-input" placeholder="例：01" value={info.seat} onChange={e=>setInfo({...info,seat:e.target.value})}/>
                )}
              </div>
              {/* Name */}
              <div>
                <label className="form-label" style={{fontSize:12}}>姓名</label>
                <input className="form-input" placeholder="你的名字" value={info.name}
                  onChange={e=>setInfo({...info,name:e.target.value})}
                  readOnly={s.useRoster && !s.allowNameEdit && !!info.name}
                  style={{background: s.useRoster&&!s.allowNameEdit&&info.name?'#f5f3ef':undefined}}/>
              </div>
            </div>
            <button onClick={handleInfoSubmit} disabled={!info.name||!info.class||!info.seat}
              style={{width:'100%',background:(!info.name||!info.class||!info.seat)?'#ccc':'var(--accent)',color:'white',padding:'13px 20px',borderRadius:8,border:'none',fontSize:15,fontWeight:700,cursor:(!info.name||!info.class||!info.seat)?'not-allowed':'pointer',fontFamily:"'Noto Sans TC',sans-serif"}}>
              開始作答 →
            </button>
          </div>
        )}

        {step==='quiz' && (
          <>
            <div className="progress-bar-wrap" id="quiz-top">
              <span style={{fontSize:13,color:'var(--ink2)',flexShrink:0}}>已作答 {Object.keys(answers).length}/{quiz.questions.length}</span>
              <div className="progress-track">
                <div className="progress-fill" style={{width:`${Object.keys(answers).length/quiz.questions.length*100}%`}}/>
              </div>
            </div>
            {quiz.questions.map((q,qi)=>(
              <div key={qi} id={`q-${qi}`} className="sq-card">
                <div className="sq-num">
                  <span>第 {qi+1} 題 · {q.points}分</span>
                  {s.allowHint!==false&&q.hint&&(
                    <button className="hint-btn" onClick={()=>setShownHints(h=>({...h,[qi]:!h[qi]}))}>
                      💡 {shownHints[qi]?'收起提示':'查看提示'}
                    </button>
                  )}
                </div>
                <div className="sq-text"><MathText text={q.text}/></div>
                {q.hint&&shownHints[qi]&&<div className="hint-bubble"><div className="hint-bubble-title">💡 提示</div><MathText text={q.hint}/></div>}
                <div className="sq-opts" style={{marginTop:q.hint&&shownHints[qi]?12:0}}>
                  {q.options.map((opt,oi)=>(
                    <div key={oi} className={`sq-opt ${answers[qi]===oi?'selected':''}`} onClick={()=>setAnswers({...answers,[qi]:oi})}>
                      <div className="opt-circle">{ABCD[oi]}</div>
                      <div className="opt-text"><MathText text={opt}/></div>
                    </div>
                  ))}
                </div>
              </div>
            ))}
            <button
              onClick={() => {
                const allAnswered = quiz.questions.every((_,i) => answers[i] !== undefined)
                if (!allAnswered) {
                  // Find first unanswered and scroll to it
                  const firstUnanswered = quiz.questions.findIndex((_,i) => answers[i] === undefined)
                  const el = document.getElementById(`q-${firstUnanswered}`)
                  if (el) {
                    el.scrollIntoView({ behavior:'smooth', block:'center' })
                    el.classList.add('unanswered-highlight')
                    setTimeout(() => el.classList.remove('unanswered-highlight'), 1500)
                  }
                  return
                }
                handleSubmit()
              }}
              disabled={saving}
              style={{width:'100%',background:saving?'#ccc':'var(--accent)',color:'white',padding:'15px 20px',borderRadius:10,border:'none',fontSize:16,fontWeight:700,cursor:saving?'not-allowed':'pointer',fontFamily:"'Noto Sans TC',sans-serif",marginBottom:32}}>
              {saving ? '儲存中...' : quiz.questions.every((_,i)=>answers[i]!==undefined) ? '提交答案 →' : `⬆ 跳到第 ${quiz.questions.findIndex((_,i)=>answers[i]===undefined)+1} 題（未作答）`}
            </button>
            {/* Floating scroll-to-top */}
            <button className="float-top-btn" onClick={()=>document.getElementById('quiz-top')?.scrollIntoView({behavior:'smooth'})} title="回到頂部">↑</button>
          </>
        )}
      </div>
    </div>
  )
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [user, setUser] = useState(undefined)
  const [displayName, setDisplayName] = useState('')
  const [showProfile, setShowProfile] = useState(false)
  const hash = useHash()

  useEffect(() => { return onAuthStateChanged(auth, u=>setUser(u||null)) }, [])
  useEffect(() => {
    if (!user) return
    getDoc(doc(db,'users',user.uid)).then(snap => {
      if (snap.exists()&&snap.data().displayName) setDisplayName(snap.data().displayName)
    }).catch(()=>{})
  }, [user])

  if (hash.startsWith('/s/')) return (
    <><style>{css}</style><StudentQuiz quizId={hash.replace('/s/','')} /></>
  )
  if (user===undefined) return (
    <><style>{css}</style><div className="loading" style={{minHeight:'100vh'}}><div className="spinner"/>載入中...</div></>
  )
  if (!user) return <><style>{css}</style><LoginPage/></>

  let content
  if (hash.startsWith('/results/'))        content = <Results quizId={hash.replace('/results/','')}/>
  else if (hash.startsWith('/analytics/')) content = <Analytics quizId={hash.replace('/analytics/','')}/>
  else if (hash==='/create')               content = <CreateQuiz user={user}/>
  else if (hash==='/roster')               content = <RosterManager user={user}/>
  else if (hash==='/admin'&&isAdmin(user)) content = <AdminPanel/>
  else                                     content = <Dashboard user={user}/>

  return (
    <><style>{css}</style>
      <div className="app">
        <div className="layout">
          <Sidebar hash={hash} user={user} displayName={displayName}
            onLogout={()=>signOut(auth)} onEditProfile={()=>setShowProfile(true)}/>
          <div className="main">{content}</div>
        </div>
      </div>
      {showProfile && <ProfileModal user={user} displayName={displayName}
        onSave={n=>{setDisplayName(n);setShowProfile(false)}} onCancel={()=>setShowProfile(false)}/>}
    </>
  )
}
