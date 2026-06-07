import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

export default defineConfig({
  base: '/curveOfMoney', // 设置基础路径为当前目录
  plugins: [react()],
})