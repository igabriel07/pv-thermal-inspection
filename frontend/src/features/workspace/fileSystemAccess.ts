export type FileSystemHandlePermissionDescriptor = {
  mode?: 'read' | 'readwrite'
}

export type FileSystemPermissionState = 'granted' | 'denied' | 'prompt'

export interface FileSystemHandle {
  kind: 'file' | 'directory'
  name: string
  queryPermission?: (descriptor?: FileSystemHandlePermissionDescriptor) => Promise<FileSystemPermissionState>
  requestPermission?: (descriptor?: FileSystemHandlePermissionDescriptor) => Promise<FileSystemPermissionState>
}

export interface FileSystemFileHandle extends FileSystemHandle {
  kind: 'file'
  getFile: () => Promise<File>
  createWritable?: () => Promise<FileSystemWritableFileStream>
}

export interface FileSystemDirectoryHandle extends FileSystemHandle {
  kind: 'directory'
  entries?: () => AsyncIterableIterator<[string, FileSystemHandle]>
  getDirectoryHandle?: (name: string, options?: { create?: boolean }) => Promise<FileSystemDirectoryHandle>
  getFileHandle?: (name: string, options?: { create?: boolean }) => Promise<FileSystemFileHandle>
  removeEntry?: (name: string, options?: { recursive?: boolean }) => Promise<void>
}

export interface FileSystemWritableFileStream {
  write: (data: string | Blob | BufferSource) => Promise<void>
  close: () => Promise<void>
}

declare global {
  interface Window {
    showDirectoryPicker?: () => Promise<FileSystemDirectoryHandle>
  }
}

export {}
