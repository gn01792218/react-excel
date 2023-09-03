import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import AutoImport from 'unplugin-auto-import/vite'
import { VitePWA } from 'vite-plugin-pwa'
const {resolve} = require('path')

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react(),
    VitePWA({
     manifest:{
      icons:[
        {
          src:"/apple-touch-icon.png",
          sizes:"180x180",
          type:"image/png",
          purpose:"any maskable"
        },
      ]
     } 
    }),
    AutoImport({
      imports:['react', 'react-router-dom'],
      dts:'src/auto-imports.d.ts'
    })
  ],
  define:{
    'process.env':{}
  },
  resolve:{
    alias:{
      '@':resolve(__dirname,'src'),
    }
  }
})
