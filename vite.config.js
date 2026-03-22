import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  plugins: [react()],
  // ⚠️ 改成你的 GitHub repo 名稱，例如 /quizflow/
  base: '/Sjps-QA-System/',
})
