import { defineConfig } from 'vite'
import pages from '@hono/vite-cloudflare-pages'
import { writeFileSync, cpSync, mkdirSync, existsSync } from 'fs'
import { resolve } from 'path'

// 빌드 후 _routes.json 자동 수정 + docs/ 복사 플러그인
function postBuildPlugin() {
  return {
    name: 'post-build',
    closeBundle() {
      // 1) _routes.json: /docs/* 와 /static/* 를 Worker에서 제외 → CF Pages 정적 서빙
      const routes = {
        version: 1,
        include: ['/*'],
        exclude: ['/docs/*', '/static/*']
      }
      writeFileSync('dist/_routes.json', JSON.stringify(routes, null, 2))
      console.log('✅ dist/_routes.json 업데이트 완료')

      // 2) docs/ → dist/docs/ 복사
      const src = resolve('./docs')
      const dest = resolve('./dist/docs')
      if (existsSync(src)) {
        mkdirSync(dest, { recursive: true })
        cpSync(src, dest, { recursive: true })
        console.log('✅ dist/docs/ 복사 완료')
      }

      // 3) public/static/ → dist/static/ 복사
      const staticSrc = resolve('./public/static')
      const staticDest = resolve('./dist/static')
      if (existsSync(staticSrc)) {
        mkdirSync(staticDest, { recursive: true })
        cpSync(staticSrc, staticDest, { recursive: true })
        console.log('✅ dist/static/ 복사 완료')
      }
    }
  }
}

export default defineConfig(({ mode }) => {
  if (mode === 'client') {
    return {
      build: {
        rollupOptions: {
          input: './src/client.ts',
          output: {
            entryFileNames: 'static/client.js',
          },
        },
      },
    }
  }

  return {
    plugins: [pages(), postBuildPlugin()],
    build: {
      outDir: 'dist',
    },
  }
})
