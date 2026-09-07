import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  server: {
    port: 3000,
    host: true
  },
  build: {
    rollupOptions: {
      output: {
        manualChunks(id) {
          if (id.includes('node_modules')) {
            // 拆分 FortuneSheet 极其相关的庞大核心依赖
            if (id.includes('@fortune-sheet')) {
              return 'fortune-sheet';
            }
            // 拆分庞大的 Excel 物理格式序列化与解析解析器 (ExcelJS 和 LuckyExcel)
            if (id.includes('luckyexcel') || id.includes('exceljs')) {
              return 'excel-parsers';
            }
            // 其余公共组件和基础库归入 vendor，实现极其优异的缓存机制
            return 'vendor';
          }
        }
      }
    }
  }
})