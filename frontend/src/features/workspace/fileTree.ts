import type { FileSystemDirectoryHandle, FileSystemFileHandle } from './fileSystemAccess'

export type TreeNode = {
  name: string
  path: string
  type: 'folder' | 'file'
  children?: TreeNode[]
}

export type FileEntry = {
  path: string
  file: File
}

export const ensureHandlePermission = async (
  handle: FileSystemDirectoryHandle,
  mode: 'read' | 'readwrite'
): Promise<boolean> => {
  if (!handle.queryPermission || !handle.requestPermission) return true
  const granted = await handle.queryPermission({ mode } as any)
  if (granted === 'granted') return true
  return (await handle.requestPermission({ mode } as any)) === 'granted'
}

export const hasWritePermission = async (handle: FileSystemDirectoryHandle): Promise<boolean> => {
  if (!handle.queryPermission) return true
  try {
    const state = await handle.queryPermission({ mode: 'readwrite' } as any)
    return state === 'granted'
  } catch {
    return false
  }
}

export const readDirectoryEntries = async (handle: FileSystemDirectoryHandle): Promise<FileEntry[]> => {
  const entries: FileEntry[] = []

  const walk = async (dir: FileSystemDirectoryHandle, basePath: string) => {
    if (!dir.entries) return
    for await (const [name, entry] of dir.entries()) {
      const nextPath = basePath ? `${basePath}/${name}` : name
      if (entry.kind === 'file') {
        const file = await (entry as FileSystemFileHandle).getFile()
        entries.push({ path: nextPath, file })
      } else if (entry.kind === 'directory') {
        await walk(entry as FileSystemDirectoryHandle, nextPath)
      }
    }
  }

  await walk(handle, '')
  return entries
}

export const buildTree = (entries: FileEntry[]): TreeNode | null => {
  if (entries.length === 0) return null
  const root: TreeNode = { name: 'root', path: '', type: 'folder', children: [] }

  for (const entry of entries) {
    const relPath = entry.path
    const parts = relPath.split('/').filter(Boolean)
    let current = root

    parts.forEach((part, index) => {
      const isFile = index === parts.length - 1
      const currentPath = parts.slice(0, index + 1).join('/')
      if (!current.children) current.children = []
      let node = current.children.find((child) => child.name === part)

      if (!node) {
        node = {
          name: part,
          path: currentPath,
          type: isFile ? 'file' : 'folder',
          children: isFile ? undefined : [],
        }
        current.children.push(node)
      }

      if (!isFile) current = node
    })
  }

  return root
}
