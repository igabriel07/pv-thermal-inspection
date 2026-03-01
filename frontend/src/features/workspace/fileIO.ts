import type { FileSystemDirectoryHandle } from './fileSystemAccess'

export const readTextFile = async (dir: FileSystemDirectoryHandle, name: string) => {
  if (!dir.getFileHandle) return null
  try {
    const fileHandle = await dir.getFileHandle(name)
    const file = await fileHandle.getFile()
    return await file.text()
  } catch {
    return null
  }
}

export const readJsonFile = async <T,>(dir: FileSystemDirectoryHandle, name: string): Promise<T | null> => {
  const text = await readTextFile(dir, name)
  if (!text) return null
  try {
    return JSON.parse(text) as T
  } catch {
    return null
  }
}

export const writeTextFile = async (dir: FileSystemDirectoryHandle, name: string, content: string) => {
  if (!dir.getFileHandle) throw new Error('Folder write is not available in this browser.')
  const fileHandle = await dir.getFileHandle(name, { create: true })
  if (!fileHandle.createWritable) throw new Error('Unable to write files: File System Access API is unavailable or permission is missing.')
  const writable = await fileHandle.createWritable()
  await writable.write(content)
  await writable.close()
}

export const writeBinaryFile = async (dir: FileSystemDirectoryHandle, name: string, content: Blob | ArrayBuffer) => {
  if (!dir.getFileHandle) throw new Error('Folder write is not available in this browser.')
  const fileHandle = await dir.getFileHandle(name, { create: true })
  if (!fileHandle.createWritable) throw new Error('Unable to write files: File System Access API is unavailable or permission is missing.')
  const writable = await fileHandle.createWritable()
  await writable.write(content)
  await writable.close()
}

export const clearDirectory = async (dir: FileSystemDirectoryHandle) => {
  if (!dir.entries || !dir.removeEntry) return
  for await (const [name, entry] of dir.entries()) {
    await dir.removeEntry(name, { recursive: entry.kind === 'directory' })
  }
}
