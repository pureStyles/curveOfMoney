import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  base: '/curveOfMoney/', // GitHub Pages 子路径需要稳定的尾部斜杠
  plugins: [react()],
})
