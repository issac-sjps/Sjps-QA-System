# QuizFlow 設定說明

## 上線步驟（不需要在自己電腦裝任何東西）

---

### Step 1：設定 Firebase

1. 前往 https://console.firebase.google.com
2. 選擇你的專案（或新建一個）
3. 點左側「Authentication」→「Sign-in method」→ 啟用 **Google**
4. 點左側「Firestore Database」→「建立資料庫」→ 選生產模式
5. 點「規則」分頁，把 `firestore.rules.txt` 的內容貼上去，儲存
6. 點左側齒輪「專案設定」→ 你的應用程式 → 找到 firebaseConfig 物件

把這些值填入 `src/firebase.js`：
```js
const firebaseConfig = {
  apiKey: "填你的值",
  authDomain: "填你的值",
  projectId: "填你的值",
  ...
}
```

---

### Step 2：設定 GitHub Repository

1. 前往 https://github.com → 右上角「+」→「New repository」
2. 名稱填 `quizflow`（或其他名稱）
3. 選 Public，點「Create repository」

---

### Step 3：上傳檔案到 GitHub

1. 在剛建好的 repo 頁面，點「uploading an existing file」或「Add file」
2. **把整個資料夾的內容拖曳進去**（包含 .github 資料夾！）
3. ⚠️ 注意：.github/workflows/deploy.yml 必須包含，這是自動部署的設定
4. 點「Commit changes」

---

### Step 4：開啟 GitHub Actions 權限

1. 在你的 repo，點上方「Settings」
2. 左側點「Actions」→「General」
3. 找到「Workflow permissions」→ 選 **Read and write permissions**
4. 儲存

---

### Step 5：等待自動部署

1. 點上方「Actions」分頁，看到「Deploy to GitHub Pages」正在執行
2. 等 1-2 分鐘，綠色 ✓ 表示成功

---

### Step 6：開啟 GitHub Pages

1. 在 repo，點「Settings」
2. 左側點「Pages」
3. Source 選 **Deploy from a branch**
4. Branch 選 **gh-pages**，資料夾選 `/ (root)`
5. 儲存

你的網址會是：
```
https://你的帳號名稱.github.io/quizflow/
```

---

### Step 7：設定 Firebase 允許你的網域

1. Firebase Console → Authentication → Settings → 已授權網域
2. 新增 `你的帳號名稱.github.io`

---

### Step 8：修改 vite.config.js（如果 repo 名稱不是 quizflow）

```js
base: '/你的repo名稱/',
```

修改後重新上傳這個檔案，GitHub Actions 會自動重新部署。

---

## 測試

1. 開啟你的網址，用 Google 帳號登入
2. 點「新增測驗」，建立一份測驗
3. 複製學生連結，在無痕視窗開啟測試學生作答
4. 回到老師頁面查看成績

---

## 常見問題

**Q: 登入時跳出錯誤**
A: 確認 Firebase Authentication 已啟用 Google，且已加入 GitHub Pages 網域

**Q: 找不到 .github 資料夾**
A: 這個資料夾是隱藏的，上傳時要確認它有被包含在內

**Q: Actions 執行失敗**
A: 確認 Settings → Actions → General → Workflow permissions 設為 Read and write
