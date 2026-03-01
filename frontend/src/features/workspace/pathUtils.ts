export const normalizePath = (path: string) =>
  path
    .replace(/\\/g, '/')
    .replace(/^\.\/+/, '')
    .replace(/^\/+/, '')
    .replace(/\/\.\//g, '/')

export const getParentDirPath = (path: string) => {
  const parts = path.split('/').filter(Boolean)
  parts.pop()
  return parts.join('/')
}
