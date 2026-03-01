import { defineConfig, loadEnv } from 'vite'
import react from '@vitejs/plugin-react'
import { readFileSync } from 'node:fs'
import path from 'node:path'
import { fileURLToPath } from 'node:url'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

type PackageJson = {
  version?: string
  dependencies?: Record<string, string>
  devDependencies?: Record<string, string>
}

const pkgPath = path.resolve(__dirname, 'package.json')
const pkg = JSON.parse(readFileSync(pkgPath, 'utf-8')) as PackageJson

const APP_VERSION = pkg.version ?? '0.0.0'
const REACT_VERSION = pkg.dependencies?.react ?? ''
const VITE_VERSION = pkg.devDependencies?.vite ?? ''
const TS_VERSION = pkg.devDependencies?.typescript ?? ''
const BUILD_TIME = new Date().toISOString()

// https://vite.dev/config/
export default defineConfig(({ mode }) => {
  const env = loadEnv(mode, process.cwd(), '')
  const proxyTarget = env.VITE_API_PROXY_TARGET || 'http://127.0.0.1:8000'

  return {
    plugins: [
      react({
        babel: {
          plugins: [['babel-plugin-react-compiler']],
        },
      }),
    ],
    server: {
      host: 'localhost',
      port: 5173,
      strictPort: true,
      proxy: {
        '/api': proxyTarget,
      },
    },
    define: {
      __APP_VERSION__: JSON.stringify(APP_VERSION),
      __REACT_VERSION__: JSON.stringify(REACT_VERSION),
      __VITE_VERSION__: JSON.stringify(VITE_VERSION),
      __TS_VERSION__: JSON.stringify(TS_VERSION),
      __BUILD_TIME__: JSON.stringify(BUILD_TIME),
    },
  }
})
