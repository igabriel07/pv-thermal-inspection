import type { FileSystemDirectoryHandle } from './fileSystemAccess'

const STORAGE_DB = 'app-storage'
const STORAGE_STORE = 'folder-handles'
const STORAGE_KEY = 'last-folder'

export const openStorage = (): Promise<IDBDatabase> =>
  new Promise((resolve, reject) => {
    const request = indexedDB.open(STORAGE_DB, 1)
    request.onupgradeneeded = () => {
      const db = request.result
      if (!db.objectStoreNames.contains(STORAGE_STORE)) {
        db.createObjectStore(STORAGE_STORE)
      }
    }
    request.onsuccess = () => resolve(request.result)
    request.onerror = () => reject(request.error)
  })

export const storeFolderHandle = async (handle: FileSystemDirectoryHandle) => {
  const db = await openStorage()
  const tx = db.transaction(STORAGE_STORE, 'readwrite')
  tx.objectStore(STORAGE_STORE).put(handle, STORAGE_KEY)
  await new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
    tx.onabort = () => reject(tx.error)
  })
}

export const clearStoredFolderHandle = async () => {
  const db = await openStorage()
  const tx = db.transaction(STORAGE_STORE, 'readwrite')
  tx.objectStore(STORAGE_STORE).delete(STORAGE_KEY)
  await new Promise<void>((resolve, reject) => {
    tx.oncomplete = () => resolve()
    tx.onerror = () => reject(tx.error)
    tx.onabort = () => reject(tx.error)
  })
}

export const getStoredFolderHandle = async (): Promise<FileSystemDirectoryHandle | null> => {
  const db = await openStorage()
  const tx = db.transaction(STORAGE_STORE, 'readonly')
  const request = tx.objectStore(STORAGE_STORE).get(STORAGE_KEY)
  return new Promise((resolve) => {
    request.onsuccess = () => resolve((request.result as FileSystemDirectoryHandle) ?? null)
    request.onerror = () => resolve(null)
  })
}
