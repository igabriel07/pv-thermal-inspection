import { useEffect, useMemo, useRef, useState } from 'react'
import type { IFloating, TableVerticalAlign } from 'docx'
import {
  AlignmentType,
  BorderStyle,
  Bookmark,
  Document,
  Footer,
  Header,
  HeightRule,
  InternalHyperlink,
  LeaderType,
  HorizontalPositionAlign,
  HorizontalPositionRelativeFrom,
  ImageRun,
  Packer,
  PageBreak,
  PageNumber,
  PageOrientation,
  Paragraph,
  TabStopType,
  Table,
  TableLayoutType,
  TableRow,
  TableCell,
  TextWrappingType,
  TextRun,
  VerticalAlignTable,
  VerticalPositionAlign,
  VerticalPositionRelativeFrom,
  WidthType,
  convertInchesToTwip,
} from 'docx'
import { renderAsync } from 'docx-preview'
import html2canvas from 'html2canvas'
import { jsPDF } from 'jspdf'
import './App.css'

type TreeNode = {
  name: string
  path: string
  type: 'folder' | 'file'
  children?: TreeNode[]
}

type FileEntry = {
  path: string
  file: File
}

type DocxDraftItem = {
  id: string
  path: string
  include: boolean
  caption: string
  tables: Array<{ id: string; title: string; rows: Array<{ id: string; name: string; description: string }> }>
  pageBreakBefore: boolean
}

type DocxDraftTableRow = { id: string; name: string; description: string }
type DocxDraftTable = { id: string; title: string; rows: DocxDraftTableRow[] }

type ReportDescriptionRow = {
  id: string
  title: string
  text: string
}

type ReportEquipmentItem = {
  id: string
  title: string
  text: string
  imageFile: File | null
  imagePreviewUrl: string | null
}

type FileSystemHandlePermissionDescriptor = {
  mode?: 'read' | 'readwrite'
}

type FileSystemPermissionState = 'granted' | 'denied' | 'prompt'

interface FileSystemHandle {
  kind: 'file' | 'directory'
  name: string
  queryPermission?: (descriptor?: FileSystemHandlePermissionDescriptor) => Promise<FileSystemPermissionState>
  requestPermission?: (descriptor?: FileSystemHandlePermissionDescriptor) => Promise<FileSystemPermissionState>
}

interface FileSystemFileHandle extends FileSystemHandle {
  kind: 'file'
  getFile: () => Promise<File>
  createWritable?: () => Promise<FileSystemWritableFileStream>
}

interface FileSystemDirectoryHandle extends FileSystemHandle {
  kind: 'directory'
  entries?: () => AsyncIterableIterator<[string, FileSystemHandle]>
  getDirectoryHandle?: (name: string, options?: { create?: boolean }) => Promise<FileSystemDirectoryHandle>
  getFileHandle?: (name: string, options?: { create?: boolean }) => Promise<FileSystemFileHandle>
  removeEntry?: (name: string, options?: { recursive?: boolean }) => Promise<void>
}

interface FileSystemWritableFileStream {
  write: (data: string | Blob | BufferSource) => Promise<void>
  close: () => Promise<void>
}

declare global {
  interface Window {
    showDirectoryPicker?: () => Promise<FileSystemDirectoryHandle>
  }
}

const STORAGE_DB = 'app-storage'
const STORAGE_STORE = 'folder-handles'
const STORAGE_KEY = 'last-folder'
// Multiplier applied on top of the user-controlled zoom so the default view
// renders larger without changing the zoom UI value.
const BASE_ZOOM = 1.44
const VIEWER_MAX_W = 1536
const VIEWER_MAX_H = 1229

const API_BASE = (import.meta.env.VITE_API_BASE_URL ?? '').replace(/\/+$/, '')
const apiUrl = (path: string) => {
  if (!API_BASE) return path
  return `${API_BASE}${path.startsWith('/') ? path : `/${path}`}`
}

function App() {
  const [message, setMessage] = useState('Loading...')
  const [theme, setTheme] = useState<'light' | 'dark'>('light')
  const [activeMenu, setActiveMenu] = useState('file')
  const [activeAction, setActiveAction] = useState<'view' | 'scan' | 'report' | ''>('')
  const [openMenu, setOpenMenu] = useState<string>('')
  const [fileTree, setFileTree] = useState<TreeNode | null>(null)
  const [fileMap, setFileMap] = useState<Record<string, File>>({})
  const [selectedPath, setSelectedPath] = useState<string>('')
  const [fileText, setFileText] = useState<string>('')
  const [fileUrl, setFileUrl] = useState<string>('')
  const [fileKind, setFileKind] = useState<'image' | 'text' | 'other' | ''>('')
  const [selectedLabels, setSelectedLabels] = useState<
    Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
  >([])
  const [folderHandle, setFolderHandle] = useState<FileSystemDirectoryHandle | null>(null)
  const [isScanning, setIsScanning] = useState(false)
  const [scanStatus, setScanStatus] = useState<string>('')
  const [scanError, setScanError] = useState<string>('')
  const [labelSaveError, setLabelSaveError] = useState<string>('')
  const [scanPromptOpen, setScanPromptOpen] = useState(false)
  const [scanProgress, setScanProgress] = useState<{ current: number; total: number }>({ current: 0, total: 0 })
  const [scanCompleted, setScanCompleted] = useState(false)
  const [faultsList, setFaultsList] = useState<string[]>([])
  const [reportText, setReportText] = useState('')
  const [reportStatus, setReportStatus] = useState<string>('')
  const [reportError, setReportError] = useState<string>('')
  const [docxDraft, setDocxDraft] = useState<DocxDraftItem[]>([])
  const [docxPreviewStatus, setDocxPreviewStatus] = useState<string>('')
  const [docxPreviewError, setDocxPreviewError] = useState<string>('')
  const [descriptionRows, setDescriptionRows] = useState<ReportDescriptionRow[]>([
    { id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, title: '', text: '' },
  ])
  const [equipmentItems, setEquipmentItems] = useState<ReportEquipmentItem[]>([
    { id: `${Date.now()}-${Math.random().toString(16).slice(2)}` , title: '', text: '', imageFile: null, imagePreviewUrl: null },
  ])
  const [reportDraftWidth, setReportDraftWidth] = useState<number>(420)
  const [isReportDraftResizing, setIsReportDraftResizing] = useState(false)
  const [imageMetrics, setImageMetrics] = useState<{ width: number; height: number; naturalWidth: number; naturalHeight: number }>({
    width: 0,
    height: 0,
    naturalWidth: 0,
    naturalHeight: 0,
  })
  const [selectedLabelIndex, setSelectedLabelIndex] = useState<number | null>(null)
  const [showLabels, setShowLabels] = useState(true)
  const [zoom, setZoom] = useState(1)
  const [pan, setPan] = useState<{ x: number; y: number }>({ x: 0, y: 0 })
  const [isPanning, setIsPanning] = useState(false)
  const [panStart, setPanStart] = useState<{ x: number; y: number } | null>(null)
  const [panOrigin, setPanOrigin] = useState<{ x: number; y: number } | null>(null)
  const [drawMode, setDrawMode] = useState<'select' | 'rect' | 'ellipse'>('select')
  const [isDrawing, setIsDrawing] = useState(false)
  const [drawStart, setDrawStart] = useState<{ x: number; y: number } | null>(null)
  const [draftLabel, setDraftLabel] = useState<{ shape: 'rect' | 'ellipse'; x: number; y: number; w: number; h: number } | null>(null)
  const [isDraggingLabel, setIsDraggingLabel] = useState(false)
  const [dragStart, setDragStart] = useState<{ x: number; y: number } | null>(null)
  const [dragOrigin, setDragOrigin] = useState<{ x: number; y: number; w: number; h: number } | null>(null)
  const [isResizingLabel, setIsResizingLabel] = useState(false)
  const [resizeHandle, setResizeHandle] = useState<'nw' | 'ne' | 'sw' | 'se' | null>(null)
  const [resizeOrigin, setResizeOrigin] = useState<{ x: number; y: number; w: number; h: number } | null>(null)
  const [labelHistory, setLabelHistory] = useState<
    Array<Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>>
  >([])
  const [labelHistoryIndex, setLabelHistoryIndex] = useState(0)
  const [expandedPaths, setExpandedPaths] = useState<Set<string>>(new Set(['']))
  const [explorerWidth, setExplorerWidth] = useState<number>(283)
  const [rightExplorerWidth, setRightExplorerWidth] = useState<number>(283)
  const [isResizing, setIsResizing] = useState(false)
  const [isRightResizing, setIsRightResizing] = useState(false)
  const [isExplorerMinimized, setIsExplorerMinimized] = useState(false)
  const [isRightExplorerMinimized, setIsRightExplorerMinimized] = useState(false)
  const [viewerHostSize, setViewerHostSize] = useState<{ width: number; height: number }>({ width: 0, height: 0 })
  const folderInputRef = useRef<HTMLInputElement | null>(null)
  const imageRef = useRef<HTMLImageElement | null>(null)
  const imageContainerRef = useRef<HTMLDivElement | null>(null)
  const viewerHostRef = useRef<HTMLDivElement | null>(null)
  const docxPreviewHostRef = useRef<HTMLDivElement | null>(null)
  const docxPreviewStylesRef = useRef<HTMLDivElement | null>(null)
  const reportSplitRef = useRef<HTMLDivElement | null>(null)
  const selectedLabelsRef = useRef(selectedLabels)
  const closeTimerRef = useRef<number | null>(null)
  const realtimePersistTimerRef = useRef<number | null>(null)
  const realtimePersistLabelsRef = useRef<
    Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }> | null
  >(null)
  const menuAreaRef = useRef<HTMLDivElement | null>(null)
  const [explorerContextMenu, setExplorerContextMenu] = useState<null | { x: number; y: number; node: TreeNode; source: 'tree' | 'faultsList' }>(null)

  useEffect(() => {
    fetch(apiUrl('/api/hello'))
      .then((res) => res.json())
      .then((data) => setMessage(data.message))
      .catch(() => setMessage('Backend unavailable. Start FastAPI.'))
  }, [])

  useEffect(() => {
    document.body.classList.remove('theme-light', 'theme-dark')
    document.body.classList.add(theme === 'dark' ? 'theme-dark' : 'theme-light')
  }, [theme])

  useEffect(() => {
    const restoreFolder = async () => {
      if (!('showDirectoryPicker' in window)) return
      const handle = await getStoredFolderHandle()
      if (!handle) return
      const canWrite = await ensureHandlePermission(handle, 'readwrite')
      if (!canWrite) {
        const canRead = await ensureHandlePermission(handle, 'read')
        if (!canRead) return
        setLabelSaveError('Folder opened read-only (write permission denied). Label saving is disabled until you reopen the folder and allow write access.')
      } else {
        setLabelSaveError('')
      }
      const entries = await readDirectoryEntries(handle)
      setFolderHandle(handle)
      applyEntries(entries)
    }

    restoreFolder().catch(() => undefined)
  }, [])

  useEffect(() => {
    return () => {
      if (fileUrl) URL.revokeObjectURL(fileUrl)
    }
  }, [fileUrl])

  useEffect(() => {
    if (!imageRef.current || !('ResizeObserver' in window)) return
    const image = imageRef.current
    const observer = new ResizeObserver((entries) => {
      const entry = entries[0]
      if (!entry) return
      const { width, height } = entry.contentRect
      setImageMetrics((prev) => ({
        ...prev,
        width,
        height,
      }))
    })
    observer.observe(image)
    return () => observer.disconnect()
  }, [fileUrl])

  useEffect(() => {
    selectedLabelsRef.current = selectedLabels
  }, [selectedLabels])

  useEffect(() => {
    return () => {
      if (realtimePersistTimerRef.current !== null) {
        window.clearTimeout(realtimePersistTimerRef.current)
        realtimePersistTimerRef.current = null
      }
    }
  }, [])

  useEffect(() => {
    const node = viewerHostRef.current
    if (!node || !('ResizeObserver' in window)) return

    const observer = new ResizeObserver(() => {
      setViewerHostSize({ width: node.clientWidth, height: node.clientHeight })
    })
    observer.observe(node)

    // Initialize immediately
    setViewerHostSize({ width: node.clientWidth, height: node.clientHeight })

    return () => observer.disconnect()
  }, [activeAction, selectedPath])


  useEffect(() => {
    if (activeAction !== 'view' || !folderHandle) return
    loadFaultsList().catch(() => undefined)
  }, [activeAction, folderHandle])

  useEffect(() => {
    if (activeAction !== 'report' || !folderHandle) return
    loadFaultsList().catch(() => undefined)
  }, [activeAction, folderHandle])

  useEffect(() => {
    if (!selectedPath) {
      setLabelHistory([])
      setLabelHistoryIndex(0)
      return
    }
    setLabelHistory([selectedLabels])
    setLabelHistoryIndex(0)
  }, [selectedPath])

  const openStorage = (): Promise<IDBDatabase> =>
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

  const storeFolderHandle = async (handle: FileSystemDirectoryHandle) => {
    const db = await openStorage()
    const tx = db.transaction(STORAGE_STORE, 'readwrite')
    tx.objectStore(STORAGE_STORE).put(handle, STORAGE_KEY)
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve()
      tx.onerror = () => reject(tx.error)
      tx.onabort = () => reject(tx.error)
    })
  }

  const clearStoredFolderHandle = async () => {
    const db = await openStorage()
    const tx = db.transaction(STORAGE_STORE, 'readwrite')
    tx.objectStore(STORAGE_STORE).delete(STORAGE_KEY)
    await new Promise<void>((resolve, reject) => {
      tx.oncomplete = () => resolve()
      tx.onerror = () => reject(tx.error)
      tx.onabort = () => reject(tx.error)
    })
  }

  const getStoredFolderHandle = async (): Promise<FileSystemDirectoryHandle | null> => {
    const db = await openStorage()
    const tx = db.transaction(STORAGE_STORE, 'readonly')
    const request = tx.objectStore(STORAGE_STORE).get(STORAGE_KEY)
    return new Promise((resolve) => {
      request.onsuccess = () => resolve(request.result ?? null)
      request.onerror = () => resolve(null)
    })
  }

  const ensureHandlePermission = async (handle: FileSystemDirectoryHandle, mode: 'read' | 'readwrite'): Promise<boolean> => {
    if (!handle.queryPermission || !handle.requestPermission) return true
    const granted = await handle.queryPermission({ mode } as any)
    if (granted === 'granted') return true
    return (await handle.requestPermission({ mode } as any)) === 'granted'
  }

  const readDirectoryEntries = async (handle: FileSystemDirectoryHandle): Promise<FileEntry[]> => {
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

  const buildTree = (entries: FileEntry[]): TreeNode | null => {
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

  const refreshFolderEntries = async (handle: FileSystemDirectoryHandle) => {
    const entries = await readDirectoryEntries(handle)
    const nextMap: Record<string, File> = {}
    entries.forEach((entry) => {
      nextMap[entry.path] = entry.file
    })
    setFileMap(nextMap)
    setFileTree(buildTree(entries))

    if (selectedPath && !nextMap[selectedPath]) {
      setSelectedPath('')
      setFileText('')
      if (fileUrl) URL.revokeObjectURL(fileUrl)
      setFileUrl('')
      setFileKind('')
      setSelectedLabels([])
      setSelectedLabelIndex(null)
      setShowLabels(true)
      setZoom(1)
      setPan({ x: 0, y: 0 })
      setDrawMode('select')
    }
  }

  const applyEntries = (entries: FileEntry[]) => {
    const map: Record<string, File> = {}
    entries.forEach((entry) => {
      map[entry.path] = entry.file
    })
    setFileMap(map)
    setFileTree(buildTree(entries))
    setExpandedPaths(new Set(['']))
    setSelectedPath('')
    setFileText('')
    if (fileUrl) URL.revokeObjectURL(fileUrl)
    setFileUrl('')
    setFileKind('')
    setSelectedLabels([])
    setSelectedLabelIndex(null)
    setShowLabels(true)
    setZoom(1)
    setPan({ x: 0, y: 0 })
    setDrawMode('select')
    setFaultsList([])
    setScanCompleted(false)
  }

  const handleOpenFolder = (event: React.ChangeEvent<HTMLInputElement>) => {
    const files = Array.from(event.target.files || [])
    const entries = files.map((file) => ({
      file,
      path: file.webkitRelativePath || file.name,
    }))
    applyEntries(entries)
  }

  const handleChooseFolder = async () => {
    if (!('showDirectoryPicker' in window)) {
      folderInputRef.current?.click()
      return
    }

    const handle = await (window as Window & { showDirectoryPicker: () => Promise<FileSystemDirectoryHandle> }).showDirectoryPicker()
    const canWrite = await ensureHandlePermission(handle, 'readwrite')
    if (!canWrite) {
      const canRead = await ensureHandlePermission(handle, 'read')
      if (!canRead) return
      setLabelSaveError('Folder opened read-only (write permission denied). Label saving is disabled until you reopen the folder and allow write access.')
    } else {
      setLabelSaveError('')
    }
    await storeFolderHandle(handle)
    const entries = await readDirectoryEntries(handle)
    setFolderHandle(handle)
    applyEntries(entries)
  }

  const handleRemoveFolder = async () => {
    setFileMap({})
    setFileTree(null)
    setExpandedPaths(new Set(['']))
    setSelectedPath('')
    setFileText('')
    if (fileUrl) URL.revokeObjectURL(fileUrl)
    setFileUrl('')
    setFileKind('')
    setActiveMenu('file')
    setActiveAction('')
    setSelectedLabels([])
    setFaultsList([])
    setScanCompleted(false)
    setFolderHandle(null)
    setLabelSaveError('')
    await clearStoredFolderHandle().catch(() => undefined)
  }

  const isImageName = (name: string) =>
    /\.(png|jpe?g|gif|webp|bmp|svg|ico)$/i.test(name)

  const isImagePath = (path: string) => isImageName(path.split('/').pop() || path)

  const normalizePath = (path: string) => path.replace(/\\/g, '/').replace(/^\/+/, '')

  const resolvePathInFileMap = (rawPath: string) => {
    const normalized = normalizePath(rawPath.trim())
    if (!normalized) return ''
    if (fileMap[normalized]) return normalized
    if (thermalFolderPath) {
      const candidate = normalized.startsWith(`${thermalFolderPath}/`)
        ? normalized
        : `${thermalFolderPath}/${normalized}`
      if (fileMap[candidate]) return candidate
    }
    return normalized
  }

  const thermalFolderPath = useMemo(() => {
    if (!fileTree?.children?.length) return null
    const target = 'thermal'
    const stack: TreeNode[] = [...fileTree.children]
    while (stack.length > 0) {
      const node = stack.shift()!
      if (node.type === 'folder' && node.name.toLowerCase() === target) {
        return node.path
      }
      if (node.type === 'folder' && node.children && node.children.length > 0) {
        stack.unshift(...node.children)
      }
    }
    return null
  }, [fileTree])

  const isInThermalFolder = (path: string) => {
    if (!thermalFolderPath) return true
    return path === thermalFolderPath || path.startsWith(`${thermalFolderPath}/`)
  }

  const getParentDirPath = (path: string) => {
    const parts = path.split('/').filter(Boolean)
    parts.pop()
    return parts.join('/')
  }

  const [navScope, setNavScope] = useState<'tree' | 'faultsList' | null>(null)

  const handleSelectFileFromScope = (scope: 'tree' | 'faultsList', path: string) => {
    setNavScope(scope)
    void handleSelectFile(path)
  }

  const currentDirImagePaths = useMemo(() => {
    if (!selectedPath) return []
    const dir = getParentDirPath(selectedPath)
    return Object.keys(fileMap)
      .filter((path) => isImagePath(path) && getParentDirPath(path) === dir)
      .sort((a, b) => a.localeCompare(b))
  }, [fileMap, selectedPath])

  const treeVisibleImagePaths = useMemo(() => {
    if (!expandedPaths.has('')) return []
    if (!fileTree?.children?.length) return []

    const findThermal = (nodes: TreeNode[]) => {
      const stack: TreeNode[] = [...nodes]
      while (stack.length > 0) {
        const node = stack.shift()!
        if (node.type === 'folder' && node.name.toLowerCase() === 'thermal') return node
        if (node.type === 'folder' && node.children && node.children.length > 0) {
          stack.unshift(...node.children)
        }
      }
      return null
    }

    const rootNode = findThermal(fileTree.children) ?? (fileTree.children.find((node) => node.type === 'folder') ?? null)
    const rootChildren = rootNode ? (rootNode.children ?? []) : fileTree.children

    const sortLikeExplorer = (nodes: TreeNode[]) =>
      [...nodes].sort((a, b) => {
        const priority = (node: TreeNode) => {
          if (node.type === 'folder') return 0
          return isImageName(node.name) ? 2 : 1
        }
        const priorityDiff = priority(a) - priority(b)
        if (priorityDiff !== 0) return priorityDiff
        return a.name.localeCompare(b.name)
      })

    const collect = (nodes: TreeNode[]): string[] => {
      const paths: string[] = []
      for (const node of sortLikeExplorer(nodes)) {
        if (node.type === 'file') {
          if (isImagePath(node.path) && fileMap[node.path]) paths.push(node.path)
          continue
        }
        if (expandedPaths.has(node.path) && node.children && node.children.length > 0) {
          paths.push(...collect(node.children))
        }
      }
      return paths
    }

    return collect(rootChildren)
  }, [expandedPaths, fileMap, fileTree])

  const faultsListImagePaths = useMemo(() => {
    return faultsList
      .filter((path) => Boolean(fileMap[path]))
      .filter((path) => isImagePath(path))
  }, [faultsList, fileMap])

  const scopedImagePaths = useMemo(() => {
    if (navScope === 'tree') return treeVisibleImagePaths
    if (navScope === 'faultsList') return faultsListImagePaths
    return []
  }, [faultsListImagePaths, navScope, treeVisibleImagePaths])

  const canNavigateScopedImages = scopedImagePaths.length > 1 && selectedPath ? scopedImagePaths.includes(selectedPath) : false

  const currentDirImageIndex = useMemo(() => {
    if (!selectedPath) return -1
    return currentDirImagePaths.indexOf(selectedPath)
  }, [currentDirImagePaths, selectedPath])

  const canNavigateDirImages = currentDirImagePaths.length > 1 && currentDirImageIndex >= 0

  const prevImagePath = useMemo(() => {
    if (canNavigateScopedImages) {
      const idx = scopedImagePaths.indexOf(selectedPath)
      const len = scopedImagePaths.length
      if (idx < 0 || len < 2) return null
      const prevIndex = (idx - 1 + len) % len
      return scopedImagePaths[prevIndex] ?? null
    }
    if (!canNavigateDirImages) return null
    const len = currentDirImagePaths.length
    const prevIndex = (currentDirImageIndex - 1 + len) % len
    return currentDirImagePaths[prevIndex] ?? null
  }, [canNavigateDirImages, canNavigateScopedImages, currentDirImageIndex, currentDirImagePaths, scopedImagePaths, selectedPath])

  const nextImagePath = useMemo(() => {
    if (canNavigateScopedImages) {
      const idx = scopedImagePaths.indexOf(selectedPath)
      const len = scopedImagePaths.length
      if (idx < 0 || len < 2) return null
      const nextIndex = (idx + 1) % len
      return scopedImagePaths[nextIndex] ?? null
    }
    if (!canNavigateDirImages) return null
    const len = currentDirImagePaths.length
    const nextIndex = (currentDirImageIndex + 1) % len
    return currentDirImagePaths[nextIndex] ?? null
  }, [canNavigateDirImages, canNavigateScopedImages, currentDirImageIndex, currentDirImagePaths, scopedImagePaths, selectedPath])

  const getFaultsDir = async (handle: FileSystemDirectoryHandle, create: boolean) => {
    if (!handle.getDirectoryHandle) return null
    return handle.getDirectoryHandle('faults', { create })
  }

  const getWorkflowRootDir = async (create: boolean) => {
    if (!folderHandle) return null
    if (!thermalFolderPath) return folderHandle
    if (!folderHandle.getDirectoryHandle) return null
    const parts = thermalFolderPath.split('/').filter(Boolean)
    let dir: FileSystemDirectoryHandle = folderHandle
    for (const part of parts) {
      if (!dir.getDirectoryHandle) return null
      dir = await dir.getDirectoryHandle(part, { create })
    }
    return dir
  }

  const getWorkflowFaultsDir = async (create: boolean) => {
    const root = await getWorkflowRootDir(create)
    if (!root) return null
    return getFaultsDir(root, create)
  }

  const readTextFile = async (dir: FileSystemDirectoryHandle, name: string) => {
    if (!dir.getFileHandle) return null
    try {
      const fileHandle = await dir.getFileHandle(name)
      const file = await fileHandle.getFile()
      return await file.text()
    } catch {
      return null
    }
  }

  const loadFaultsList = async () => {
    if (!folderHandle) return
    const workflowRoot = await getWorkflowRootDir(false)
    const text = workflowRoot ? await readTextFile(workflowRoot, 'faults.txt') : null
    if (!text) {
      const fallback = await readTextFile(folderHandle, 'faults.txt')
      if (!fallback) {
        setFaultsList([])
        return
      }
      const raw = fallback.split(/\r?\n/).map((line) => line.trim()).filter(Boolean)
      const seen = new Set<string>()
      const resolved: string[] = []
      for (const item of raw) {
        const candidate = resolvePathInFileMap(item)
        if (!candidate || seen.has(candidate)) continue
        seen.add(candidate)
        resolved.push(candidate)
      }
      setFaultsList(resolved)
      return
    }
    const raw = text.split(/\r?\n/).map((line) => line.trim()).filter(Boolean)
    const seen = new Set<string>()
    const resolved: string[] = []
    for (const item of raw) {
      const candidate = resolvePathInFileMap(item)
      if (!candidate || seen.has(candidate)) continue
      seen.add(candidate)
      resolved.push(candidate)
    }
    setFaultsList(resolved)
  }

  const normalizeFaultsText = (text: string) => {
    const lines = text
      .replace(/\r\n/g, '\n')
      .split('\n')
      .map((line) => line.trim())
      .filter(Boolean)
    // Deduplicate while preserving order.
    const seen = new Set<string>()
    const unique: string[] = []
    for (const line of lines) {
      if (seen.has(line)) continue
      seen.add(line)
      unique.push(line)
    }
    return unique
  }

  const buildDocxDraftFromPaths = (paths: string[], prev: DocxDraftItem[]) => {
    const prevByPath = new Map(prev.map((item) => [item.path, item]))
    return paths.map((path, idx) => {
      const existing = prevByPath.get(path) as (DocxDraftItem & { notes?: string; fields?: DocxDraftTableRow[]; tables?: DocxDraftTable[] }) | undefined

      const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`
      const defaultRow = (): DocxDraftTableRow => ({ id: newId(), name: '', description: '' })
      const defaultTable = (): DocxDraftTable => ({ id: newId(), title: '', rows: [defaultRow()] })

      const existingTables = Array.isArray(existing?.tables) ? existing?.tables : null
      const legacyRows = Array.isArray(existing?.fields) ? existing?.fields : null
      const legacyNotes = typeof existing?.notes === 'string' && existing.notes.trim()
        ? [{ id: newId(), name: 'Notes', description: existing.notes.trim() }]
        : null

      const migratedTables: DocxDraftTable[] | null = legacyRows && legacyRows.length
        ? [{ id: newId(), title: '', rows: legacyRows }]
        : legacyNotes
          ? [{ id: newId(), title: '', rows: legacyNotes }]
          : null

      const tables = existingTables && existingTables.length
        ? existingTables
        : migratedTables && migratedTables.length
          ? migratedTables
          : [defaultTable()]
      return {
        id: existing?.id ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        path,
        include: existing?.include ?? true,
        caption: existing?.caption ?? path,
        tables,
        pageBreakBefore: existing?.pageBreakBefore ?? idx > 0,
      }
    })
  }

  const addDocxDraftTable = (draftId: string) => {
    const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`
    setDocxDraft((prev) =>
      prev.map((item) =>
        item.id !== draftId
          ? item
          : {
              ...item,
              tables: [
                ...(Array.isArray(item.tables) ? item.tables : []),
                { id: newId(), title: '', rows: [{ id: newId(), name: '', description: '' }] },
              ],
            }
      )
    )
  }

  const removeDocxDraftTable = (draftId: string, tableId: string) => {
    const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`
    setDocxDraft((prev) =>
      prev.map((item) => {
        if (item.id !== draftId) return item
        const nextTables = (item.tables || []).filter((t) => t.id !== tableId)
        return {
          ...item,
          tables: nextTables.length ? nextTables : [{ id: newId(), title: '', rows: [{ id: newId(), name: '', description: '' }] }],
        }
      })
    )
  }

  const updateDocxDraftTable = (draftId: string, tableId: string, patch: Partial<Pick<DocxDraftTable, 'title'>>) => {
    setDocxDraft((prev) =>
      prev.map((item) => {
        if (item.id !== draftId) return item
        const nextTables = (item.tables || []).map((t) => (t.id === tableId ? { ...t, ...patch } : t))
        return { ...item, tables: nextTables }
      })
    )
  }

  const addDocxDraftTableRow = (draftId: string, tableId: string) => {
    const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`
    setDocxDraft((prev) =>
      prev.map((item) => {
        if (item.id !== draftId) return item
        const nextTables = (item.tables || []).map((t) =>
          t.id !== tableId
            ? t
            : { ...t, rows: [...(t.rows || []), { id: newId(), name: '', description: '' }] }
        )
        return { ...item, tables: nextTables }
      })
    )
  }

  const updateDocxDraftTableRow = (draftId: string, tableId: string, rowId: string, patch: Partial<DocxDraftTableRow>) => {
    setDocxDraft((prev) =>
      prev.map((item) => {
        if (item.id !== draftId) return item
        const nextTables = (item.tables || []).map((t) => {
          if (t.id !== tableId) return t
          const nextRows = (t.rows || []).map((r) => (r.id === rowId ? { ...r, ...patch } : r))
          return { ...t, rows: nextRows }
        })
        return { ...item, tables: nextTables }
      })
    )
  }

  const removeDocxDraftTableRow = (draftId: string, tableId: string, rowId: string) => {
    const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`
    setDocxDraft((prev) =>
      prev.map((item) => {
        if (item.id !== draftId) return item
        const nextTables = (item.tables || []).map((t) => {
          if (t.id !== tableId) return t
          const nextRows = (t.rows || []).filter((r) => r.id !== rowId)
          return { ...t, rows: nextRows.length ? nextRows : [{ id: newId(), name: '', description: '' }] }
        })
        return { ...item, tables: nextTables }
      })
    )
  }

  const syncDocxDraftFromEditor = () => {
    const paths = normalizeFaultsText(reportText || faultsList.join('\n'))
    setDocxDraft((prev) => buildDocxDraftFromPaths(paths, prev))
    setDocxPreviewStatus(`Draft synced (${paths.length} item(s)).`)
    setDocxPreviewError('')
  }

  const moveDocxDraftItem = (index: number, delta: -1 | 1) => {
    setDocxDraft((prev) => {
      const next = [...prev]
      const target = index + delta
      if (target < 0 || target >= next.length) return prev
      const tmp = next[index]
      next[index] = next[target]
      next[target] = tmp
      return next
    })
  }

  const updateDocxDraftItem = (id: string, patch: Partial<DocxDraftItem>) => {
    setDocxDraft((prev) => prev.map((item) => (item.id === id ? { ...item, ...patch } : item)))
  }

  const removeDocxDraftItem = (id: string) => {
    setDocxDraft((prev) => prev.filter((item) => item.id !== id))
  }

  const loadReportFromDisk = async () => {
    setReportError('')
    setReportStatus('Loading…')

    if (!folderHandle) {
      setReportError('Choose a folder with File → Open Folder…')
      setReportStatus('')
      return
    }
    const workflowRoot = await getWorkflowRootDir(false)
    const text = (workflowRoot ? await readTextFile(workflowRoot, 'faults.txt') : null) || (await readTextFile(folderHandle, 'faults.txt')) || ''
    const normalized = normalizeFaultsText(text)
    setReportText(normalized.join('\n'))
    setDocxDraft((prev) => buildDocxDraftFromPaths(normalized, prev))
    setReportStatus(`Loaded ${normalized.length} item(s).`)
  }

  const saveReportToDisk = async () => {
    setReportError('')
    setReportStatus('Saving…')

    if (!folderHandle) {
      setReportError('Choose a folder with File → Open Folder…')
      setReportStatus('')
      return
    }
    if (!folderHandle.getDirectoryHandle) {
      setReportError('Folder access is not available in this browser.')
      setReportStatus('')
      return
    }

    const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
    if (!canWrite) {
      setReportError('Write permission is required. Reopen the folder and allow write access.')
      setReportStatus('')
      return
    }

    try {
      const draftPaths = docxDraft.length ? docxDraft.map((d) => d.path) : normalizeFaultsText(reportText).map((p) => p)
      const normalized = normalizeFaultsText(draftPaths.join('\n'))
      const content = normalized.join('\n')

      const workflowRoot = await getWorkflowRootDir(true)
      if (!workflowRoot) throw new Error('Unable to access thermal folder')

      const faultsDir = await getFaultsDir(workflowRoot, true)
      if (faultsDir) {
        await writeTextFile(faultsDir, 'faults.txt', content)
      }
      await writeTextFile(workflowRoot, 'faults.txt', content)

      setFaultsList(normalized)
      setReportText(content)
      setDocxDraft((prev) => buildDocxDraftFromPaths(normalized, prev))
      setReportStatus(`Saved ${normalized.length} item(s).`)
    } catch (error) {
      setReportError(error instanceof Error ? error.message : 'Failed to save faults.txt')
      setReportStatus('')
    }
  }

  const buildWordReportBlobFromDraft = async (draft: DocxDraftItem[]) => {
    const paragraphRunsFromMultiline = (value: string) => {
      const parts = value.replace(/\r\n/g, '\n').split('\n')
      return parts.map((part, idx) => new TextRun({ text: part, break: idx === 0 ? 0 : 1 }))
    }

    const getImageSize = async (blob: Blob): Promise<{ width: number; height: number }> => {
      try {
        const bitmap = await createImageBitmap(blob)
        const size = { width: bitmap.width, height: bitmap.height }
        bitmap.close()
        return size
      } catch {
        return { width: 1200, height: 800 }
      }
    }

    const fetchPublicImageRun = async (
      name: string,
      type: 'png' | 'jpg',
      maxWidth: number,
      maxHeight: number,
      allowUpscale: boolean = false,
      floating?: IFloating
    ): Promise<ImageRun | null> => {
      try {
        const base = import.meta.env.BASE_URL || '/'
        const url = `${base}${name}`
        const res = await fetch(url)
        if (!res.ok) return null
        const blob = await res.blob()
        const buf = await blob.arrayBuffer()
        const { width, height } = await getImageSize(blob)
        let scale = Math.min(maxWidth / width, maxHeight / height)
        if (!allowUpscale) scale = Math.min(scale, 1)
        const w = Math.max(1, Math.round(width * scale))
        const h = Math.max(1, Math.round(height * scale))
        return new ImageRun({
          type,
          data: new Uint8Array(buf),
          transformation: { width: w, height: h },
          floating,
        })
      } catch {
        return null
      }
    }

    const maxW = 650
    const maxH = 900
    const children: Array<Paragraph | Table> = []

    const fileToImageRun = async (file: File, maxWidth: number, maxHeight: number): Promise<ImageRun | null> => {
      if (!file.type.startsWith('image/')) return null

      const imageType = (() => {
        const t = (file.type || '').toLowerCase()
        if (t === 'image/png') return 'png'
        if (t === 'image/jpeg' || t === 'image/jpg') return 'jpg'
        if (t === 'image/gif') return 'gif'
        if (t === 'image/bmp') return 'bmp'
        return null
      })()

      if (!imageType) return null
      const buf = await file.arrayBuffer()
      const { width, height } = await getImageSize(file)
      const scale = Math.min(maxWidth / width, maxHeight / height, 1)
      const w = Math.max(1, Math.round(width * scale))
      const h = Math.max(1, Math.round(height * scale))
      return new ImageRun({
        type: imageType,
        data: new Uint8Array(buf),
        transformation: { width: w, height: h },
      })
    }

    const blobToUint8Array = async (blob: Blob) => new Uint8Array(await blob.arrayBuffer())

    const renderLabeledPngBytes = async (
      file: File,
      labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse' }>,
      targetWidth: number,
      targetHeight: number
    ): Promise<Uint8Array> => {
      const canvas = document.createElement('canvas')
      canvas.width = Math.max(1, Math.round(targetWidth))
      canvas.height = Math.max(1, Math.round(targetHeight))

      const ctx = canvas.getContext('2d')
      if (!ctx) throw new Error('Canvas 2D context is not available')

      // Draw source image scaled to the final DOCX transformation size.
      try {
        const bitmap = await createImageBitmap(file)
        ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height)
      } catch {
        const url = URL.createObjectURL(file)
        try {
          const img = new Image()
          img.decoding = 'async'
          img.src = url
          await new Promise<void>((resolve, reject) => {
            img.onload = () => resolve()
            img.onerror = () => reject(new Error('Failed to load image'))
          })
          ctx.drawImage(img, 0, 0, canvas.width, canvas.height)
        } finally {
          URL.revokeObjectURL(url)
        }
      }

      if (labels.length > 0) {
        // Draw high-contrast outlines (black under-stroke + intense green stroke)
        // so labels remain visible on any background. No text is rendered.
        const lineWidth = Math.max(3, Math.round(Math.min(canvas.width, canvas.height) / 180))
        const green = 'rgba(0, 255, 0, 0.98)'
        const black = 'rgba(0, 0, 0, 0.75)'
        ctx.lineJoin = 'round'
        ctx.lineCap = 'round'

        for (const label of labels) {
          const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
          const cx = (isNormalized ? label.x : label.x / canvas.width) * canvas.width
          const cy = (isNormalized ? label.y : label.y / canvas.height) * canvas.height
          const bw = (isNormalized ? label.w : label.w / canvas.width) * canvas.width
          const bh = (isNormalized ? label.h : label.h / canvas.height) * canvas.height

          const left = cx - bw / 2
          const top = cy - bh / 2
          const width = bw
          const height = bh

          if (label.shape === 'ellipse') {
            const rx = Math.max(1, width / 2)
            const ry = Math.max(1, height / 2)

            ctx.beginPath()
            ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2)
            ctx.lineWidth = lineWidth + 2
            ctx.strokeStyle = black
            ctx.stroke()

            ctx.beginPath()
            ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2)
            ctx.lineWidth = lineWidth
            ctx.strokeStyle = green
            ctx.stroke()
          } else {
            ctx.lineWidth = lineWidth + 2
            ctx.strokeStyle = black
            ctx.strokeRect(left, top, width, height)

            ctx.lineWidth = lineWidth
            ctx.strokeStyle = green
            ctx.strokeRect(left, top, width, height)
          }
        }
      }

      const blob = await new Promise<Blob>((resolve, reject) => {
        canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('Failed to encode PNG'))), 'image/png')
      })

      return blobToUint8Array(blob)
    }

    // Cover page (always first page)
    // Use a borderless table with fixed row heights so placement is stable in Word and docx-preview.
    const toposolLogo = await fetchPublicImageRun('toposol-logo.png', 'png', 700, 300, true)
    const thermalLogo = await fetchPublicImageRun('thermal-logo.jpg', 'jpg', 580, 320, true)

    const coverContentHeightTwip = Math.round(convertInchesToTwip(11.69 - 1.0)) // A4 height minus 0.5" top+bottom margins
    const coverHeights = {
      topOffset: Math.round(convertInchesToTwip(0.35)),
      topLogo: Math.round(convertInchesToTwip(1.35)),
      spacer1: Math.round(convertInchesToTwip(1.85)),
      title: Math.round(convertInchesToTwip(1.0)),
      bottomLogo: Math.round(convertInchesToTwip(5)),
    }
    const used =
      coverHeights.topOffset +
      coverHeights.topLogo +
      coverHeights.spacer1 +
      coverHeights.title +
      coverHeights.bottomLogo
    const spacer2 = Math.max(Math.round(convertInchesToTwip(1.0)), coverContentHeightTwip - used)

    const borderlessCell = (cellChildren: Paragraph[], verticalAlign: TableVerticalAlign = VerticalAlignTable.CENTER) =>
      new TableCell({
        verticalAlign,
        borders: {
          top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
          right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        },
        children: cellChildren,
      })

    const spacerParagraph = () => new Paragraph({ children: [new TextRun({ text: '' })] })

    const coverTable = new Table({
      layout: TableLayoutType.FIXED,
      width: { size: 100, type: WidthType.PERCENTAGE },
      borders: {
        top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        insideHorizontal: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
        insideVertical: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      },
      rows: [
        new TableRow({
          height: { value: coverHeights.topOffset, rule: HeightRule.EXACT },
          children: [borderlessCell([spacerParagraph()])],
        }),
        new TableRow({
          height: { value: coverHeights.topLogo, rule: HeightRule.EXACT },
          children: [
            borderlessCell(
              [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: toposolLogo
                  ? [toposolLogo]
                  : [new TextRun({ text: '[Missing toposol-logo.png]', color: 'B91C1C' })],
              }),
              ],
              VerticalAlignTable.BOTTOM
            ),
          ],
        }),
        new TableRow({
          height: { value: coverHeights.spacer1, rule: HeightRule.EXACT },
          children: [borderlessCell([spacerParagraph()])],
        }),
        new TableRow({
          height: { value: coverHeights.title, rule: HeightRule.EXACT },
          children: [
            borderlessCell([
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: 'THERMOGRAPHY REPORT', bold: true, size: 56 })],
              }),
            ]),
          ],
        }),
        new TableRow({
          height: { value: spacer2, rule: HeightRule.EXACT },
          children: [borderlessCell([spacerParagraph()])],
        }),
        new TableRow({
          height: { value: coverHeights.bottomLogo, rule: HeightRule.EXACT },
          children: [
            borderlessCell([
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: thermalLogo
                  ? [thermalLogo]
                  : [new TextRun({ text: '[Missing thermal-logo.jpg]', color: 'B91C1C' })],
              }),
            ]),
          ],
        }),
      ],
    })

    children.push(coverTable)

    const items = draft.filter((d) => d.include)

    const BOOKMARK_DESCRIPTION = 'section_description'
    const BOOKMARK_EQUIPMENT = 'section_equipment'
    const BOOKMARK_IMAGES = 'section_images'
    const BOOKMARK_REPORTS = 'section_reports'

    const sectionTitle = (title: string, bookmarkId: string) =>
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 520 },
        children: [
          new Bookmark({
            id: bookmarkId,
            children: [new TextRun({ text: title, bold: true, size: 32 })],
          }),
        ],
      })

    // Page 2: Table of Contents
    children.push(new Paragraph({ children: [new PageBreak()] }))

    const descriptionPage = 2
    const equipmentPage = 3
    const imagesPage = 4
    // With our layout: Images title + first image are on page 4, then one page per remaining image,
    // then Reports on the next page.
    const reportsPage = 4 + items.length

    const tocRightStop = Math.round(convertInchesToTwip(7.1))
    const tocTitleStop = Math.round(convertInchesToTwip(0.7))
    const tocNumberStop = Math.round(convertInchesToTwip(0.55))

    const tocLine = (index: number, title: string, bookmarkId: string, page: number) =>
      new Paragraph({
        alignment: AlignmentType.LEFT,
        tabStops: [
          // Right-aligned column for the index (1., 2., 10., ...)
          { type: TabStopType.RIGHT, position: tocNumberStop },
          { type: TabStopType.LEFT, position: tocTitleStop },
          { type: TabStopType.RIGHT, position: tocRightStop, leader: LeaderType.DOT },
        ],
        spacing: { before: 80, after: 80 },
        children: [
          new TextRun({ text: `${index}.` }),
          // First tab right-aligns the number, second tab jumps to the title column.
          new TextRun({ text: '\t\t' }),
          new InternalHyperlink({
            anchor: bookmarkId,
            children: [new TextRun({ text: title, color: '000000' })],
          }),
          new TextRun({ text: '\t' }),
          new TextRun({ text: String(page) }),
        ],
      })

    children.push(
      new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 260 } }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 520 },
        children: [new TextRun({ text: 'Table of Contents', bold: true, size: 32 })],
      }),
      tocLine(1, 'Description', BOOKMARK_DESCRIPTION, descriptionPage),
      tocLine(2, 'Equipment', BOOKMARK_EQUIPMENT, equipmentPage),
      tocLine(3, 'Images', BOOKMARK_IMAGES, imagesPage),
      tocLine(4, 'Reports (anomaly detected)', BOOKMARK_REPORTS, reportsPage)
    )

    // Page 3: Description
    children.push(new Paragraph({ children: [new PageBreak()] }))
    children.push(sectionTitle('Description', BOOKMARK_DESCRIPTION))

    const cleanDescriptionRows = descriptionRows
      .map((row) => ({ ...row, title: row.title.trim(), text: row.text.trim() }))
      .filter((row) => row.title || row.text)

    if (cleanDescriptionRows.length === 0) {
      children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '' })] }))
    } else {
      for (const row of cleanDescriptionRows) {
        if (row.title) {
          children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 120, after: 80 },
              children: [new TextRun({ text: row.title, bold: true, size: 24 })],
            })
          )
        }
        if (row.text) {
          children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 140 },
              children: paragraphRunsFromMultiline(row.text),
            })
          )
        }
      }
    }

    // Page 4: Equipment
    children.push(new Paragraph({ children: [new PageBreak()] }))
    children.push(sectionTitle('Equipment', BOOKMARK_EQUIPMENT))

    const cleanEquipmentItems = equipmentItems
      .map((item) => ({ ...item, title: item.title.trim(), text: item.text.trim() }))
      .filter((item) => item.title || item.text || item.imageFile)

    if (cleanEquipmentItems.length === 0) {
      children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '' })] }))
    } else {
      for (const item of cleanEquipmentItems) {
        if (item.title) {
          children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 120, after: 80 },
              children: [new TextRun({ text: item.title, bold: true, size: 24 })],
            })
          )
        }
        if (item.text) {
          children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 0, after: 120 },
              children: paragraphRunsFromMultiline(item.text),
            })
          )
        }

        if (item.imageFile) {
          const run = await fileToImageRun(item.imageFile, 520, 320)
          if (run) {
            children.push(
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 240 },
                children: [run],
              })
            )
          } else {
            children.push(
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 240 },
                children: [new TextRun({ text: `Unsupported equipment image: ${item.imageFile.type || 'unknown type'}`, color: 'B91C1C' })],
              })
            )
          }
        }
      }
    }

    // Images section: start a new page, show the title, then start the first image on the same page.
    children.push(new Paragraph({ children: [new PageBreak()] }))
    children.push(sectionTitle('Images', BOOKMARK_IMAGES))

    for (let idx = 0; idx < items.length; idx += 1) {
      const item = items[idx]
      const file = fileMap[item.path]

      if (idx > 0 && item.pageBreakBefore) {
        children.push(
          new Paragraph({
            children: [new PageBreak()],
          })
        )
      }

      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: item.caption || item.path, bold: true })],
        })
      )

      if (!file) {
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Missing file: ${item.path}`, color: 'B91C1C' })] }))
        continue
      }

      if (!file.type.startsWith('image/')) {
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Not an image: ${file.type || 'unknown type'}`, color: 'B91C1C' })] }))
        continue
      }

      const imageType = (() => {
        const t = (file.type || '').toLowerCase()
        if (t === 'image/png') return 'png'
        if (t === 'image/jpeg' || t === 'image/jpg') return 'jpg'
        if (t === 'image/gif') return 'gif'
        if (t === 'image/bmp') return 'bmp'
        return null
      })()

      if (!imageType) {
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `Unsupported image type for Word export: ${file.type}`, color: 'B91C1C' })] }))
        continue
      }

      const buf = await file.arrayBuffer()
      const { width, height } = await getImageSize(file)
      const scale = Math.min(maxW / width, maxH / height, 1)
      const w = Math.max(1, Math.round(width * scale))
      const h = Math.max(1, Math.round(height * scale))

      const labels = await readLabelsForPath(item.path)

      const pushImageParagraph = (run: ImageRun) => {
        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [run],
          })
        )
      }

      let usedLabeled = false
      if (labels.length > 0) {
        try {
          const labeledBytes = await renderLabeledPngBytes(file, labels, w, h)
          pushImageParagraph(
            new ImageRun({
              type: 'png',
              data: labeledBytes,
              transformation: { width: w, height: h },
            })
          )
          usedLabeled = true
        } catch {
          usedLabeled = false
        }
      }

      if (!usedLabeled) {
        pushImageParagraph(
          new ImageRun({
            type: imageType,
            data: new Uint8Array(buf),
            transformation: { width: w, height: h },
          })
        )
      }

      // One or more user-defined tables under the image (centered). No column header row.
      const cellBorders = {
        top: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
        left: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
        right: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
      }

      const tablesToRender = Array.isArray(item.tables) && item.tables.length ? item.tables : []
      for (const t of tablesToRender) {
        const title = (t.title || '').trim()
        if (title) {
          children.push(
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { before: 160, after: 80 },
              children: [new TextRun({ text: title, bold: true })],
            })
          )
        } else {
          children.push(new Paragraph({ children: [new TextRun({ text: '' })] }))
        }

        const rows = (t.rows || []).filter((r) => (r.name || '').trim() || (r.description || '').trim())
        const rowsToRender: Array<Pick<DocxDraftTableRow, 'name' | 'description'>> = rows.length
          ? rows
          : [{ name: '', description: '' }]

        children.push(
          new Table({
            // TS-safe centering fallback: percentage width + fixed layout; Word generally centers tables with percentage width.
            // If you want hard centering in Word, we can switch to explicit Table alignment once confirmed in your docx version.
            width: { size: 86, type: WidthType.PERCENTAGE },
            layout: TableLayoutType.FIXED,
            borders: {
              top: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
              bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
              left: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
              right: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
              insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
              insideVertical: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
            },
            rows: rowsToRender.map(
              (r) =>
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 33, type: WidthType.PERCENTAGE },
                      borders: cellBorders,
                      children: [new Paragraph({ children: [new TextRun({ text: (r.name || '').trim(), bold: true })] })],
                    }),
                    new TableCell({
                      width: { size: 67, type: WidthType.PERCENTAGE },
                      borders: cellBorders,
                      children: [new Paragraph({ children: paragraphRunsFromMultiline((r.description || '').trim()) })],
                    }),
                  ],
                })
            ),
          })
        )
      }
    }

    // Reports section (starts after all image pages)
    children.push(new Paragraph({ children: [new PageBreak()] }))
    children.push(sectionTitle('Reports (anomaly detected)', BOOKMARK_REPORTS))
    children.push(new Paragraph({ children: [new TextRun({ text: '' })] }))

    const coverFooter = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [
            new TextRun({ text: 'Estias 3, Kozani', break: 1 }),
            new TextRun({ text: '+30 24610 32495', break: 1 }),
            new TextRun({ text: 'toposolike@ gmail.com', break: 1 }),
          ],
        }),
      ],
    })

    // Floating offsets are expressed in EMUs (English Metric Units).
    // docx image pixel sizes are converted using ~9525 EMU per px (96 DPI).
    const EMU_PER_PX = 9525
    const A4_HEIGHT_EMU = 10689336 // 11.69in * 914400
    const watermarkPx = 760
    const watermarkHeightEmu = watermarkPx * EMU_PER_PX
    const watermarkDownShiftEmu = 200000 // ~0.22in
    const watermarkTopOffsetEmu = Math.round((A4_HEIGHT_EMU - watermarkHeightEmu) / 2 + watermarkDownShiftEmu)

    const watermarkFloating: IFloating = {
      horizontalPosition: {
        relative: HorizontalPositionRelativeFrom.PAGE,
        align: HorizontalPositionAlign.CENTER,
      },
      verticalPosition: {
        relative: VerticalPositionRelativeFrom.PAGE,
        offset: watermarkTopOffsetEmu,
      },
      behindDocument: true,
      allowOverlap: true,
      wrap: {
        type: TextWrappingType.NONE,
      },
      zIndex: 0,
    }

    const tinyLogoFloating: IFloating = {
      horizontalPosition: {
        relative: HorizontalPositionRelativeFrom.MARGIN,
        align: HorizontalPositionAlign.LEFT,
      },
      verticalPosition: {
        relative: VerticalPositionRelativeFrom.MARGIN,
        align: VerticalPositionAlign.TOP,
      },
      behindDocument: false,
      allowOverlap: true,
      wrap: {
        type: TextWrappingType.NONE,
      },
      zIndex: 10,
    }

    const buildWatermarkLogoRun = async (): Promise<ImageRun | null> => {
      try {
        const base = import.meta.env.BASE_URL || '/'
        const url = `${base}toposol-logo.png`
        const res = await fetch(url)
        if (!res.ok) return null

        const srcBlob = await res.blob()
        const bitmap = await createImageBitmap(srcBlob)

        const canvasSize = 2200
        const canvas = document.createElement('canvas')
        canvas.width = canvasSize
        canvas.height = canvasSize
        const ctx = canvas.getContext('2d')
        if (!ctx) {
          bitmap.close()
          return null
        }

        // Transparent background; draw a very light rotated logo.
        ctx.clearRect(0, 0, canvasSize, canvasSize)
        ctx.save()
        ctx.translate(canvasSize / 2, canvasSize / 2)
        ctx.rotate(Math.PI / 4)

        const maxSide = 1700
        const scale = Math.min(maxSide / bitmap.width, maxSide / bitmap.height)
        const w = Math.max(1, Math.round(bitmap.width * scale))
        const h = Math.max(1, Math.round(bitmap.height * scale))

        ctx.globalAlpha = 0.12
        ctx.drawImage(bitmap, -w / 2, -h / 2, w, h)
        ctx.restore()
        bitmap.close()

        const outBlob = await new Promise<Blob>((resolve, reject) => {
          canvas.toBlob((b) => {
            if (b) resolve(b)
            else reject(new Error('Failed to generate watermark image.'))
          }, 'image/png')
        })

        const buf = await outBlob.arrayBuffer()
        return new ImageRun({
          type: 'png',
          data: new Uint8Array(buf),
          // NOTE: docx uses pixel sizes which are converted to EMUs; keep within A4 width.
          transformation: { width: 760, height: 760 },
          floating: watermarkFloating,
        })
      } catch {
        return null
      }
    }

    const watermarkLogo = await buildWatermarkLogoRun()
    const tinyToposolLogo = await fetchPublicImageRun('toposol-logo.png', 'png', 70, 70, true, tinyLogoFloating)

    const headerChildren: Paragraph[] = []

    if (tinyToposolLogo) {
      headerChildren.push(
        new Paragraph({
          alignment: AlignmentType.LEFT,
          children: [tinyToposolLogo],
        })
      )
    }

    headerChildren.push(
      new Paragraph({
        alignment: AlignmentType.RIGHT,
        children: [new TextRun({ text: 'Thermographic Report' })],
      })
    )

    if (watermarkLogo) {
      headerChildren.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [watermarkLogo],
        })
      )
    }

    const watermarkHeader = new Header({
      children: headerChildren,
    })

    const emptyFirstHeader = new Header({
      children: [new Paragraph({ children: [new TextRun({ text: '' })] })],
    })

    const pageNumberFooter = new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ children: ['Page | ', PageNumber.CURRENT] })],
        }),
      ],
    })

    const doc = new Document({
      sections: [
        {
          properties: {
            titlePage: true,
            page: {
              size: {
                orientation: PageOrientation.PORTRAIT,
                width: Math.round(convertInchesToTwip(8.27)),
                height: Math.round(convertInchesToTwip(11.69)),
              },
              pageNumbers: {
                // Start numbering at 0 so the first non-cover page displays as Page | 1.
                start: 0,
              },
              margin: {
                top: Math.round(convertInchesToTwip(0.5)),
                bottom: Math.round(convertInchesToTwip(0.5)),
                left: Math.round(convertInchesToTwip(0.5)),
                right: Math.round(convertInchesToTwip(0.5)),
              },
            },
          },
          headers: {
            first: emptyFirstHeader,
            default: watermarkHeader,
          },
          footers: {
            first: coverFooter,
            default: pageNumberFooter,
          },
          children,
        },
      ],
    })

    return await Packer.toBlob(doc)
  }

  const previewWordReport = async () => {
    setDocxPreviewError('')
    setDocxPreviewStatus('Building preview…')

    const host = docxPreviewHostRef.current
    if (!host) {
      setDocxPreviewError('Preview host is not available.')
      setDocxPreviewStatus('')
      return
    }

    try {
      const draftToUse = docxDraft.length ? docxDraft : buildDocxDraftFromPaths(normalizeFaultsText(reportText || faultsList.join('\n')), [])
      const blob = await buildWordReportBlobFromDraft(draftToUse)
      const arrayBuffer = await blob.arrayBuffer()
      host.innerHTML = ''
      const styleHost = docxPreviewStylesRef.current ?? undefined
      await renderAsync(arrayBuffer, host, styleHost, {
        inWrapper: true,
        ignoreWidth: false,
        ignoreHeight: false,
        breakPages: true,
      })

      // docx-preview does not reliably render DOCX headers/footers across all documents.
      // To ensure pagination is visible in preview (and in PDF export), overlay page numbers on each page.
      try {
        const wrapper = host.querySelector('.docx-wrapper') as HTMLElement | null
        const pages = wrapper ? (Array.from(wrapper.querySelectorAll('section.docx')) as HTMLElement[]) : []
        for (const page of pages) {
          page.querySelectorAll('.preview-page-number').forEach((n) => n.remove())
        }
        for (let i = 0; i < pages.length; i += 1) {
          if (i === 0) continue // cover page: no pagination
          const pageNumber = i // cover is page 0; TOC becomes Page | 1
          const el = document.createElement('div')
          el.className = 'preview-page-number'
          el.textContent = `Page | ${pageNumber}`
          pages[i].appendChild(el)
        }
      } catch {
        // Best-effort only; don't fail preview.
      }

      setDocxPreviewStatus('Preview updated.')
    } catch (error) {
      setDocxPreviewError(error instanceof Error ? error.message : 'Failed to render DOCX preview')
      setDocxPreviewStatus('')
    }
  }

  const downloadWordReport = async () => {
    setReportError('')
    setReportStatus('Generating Word report…')

    try {
      const draftToUse = docxDraft.length ? docxDraft : buildDocxDraftFromPaths(normalizeFaultsText(reportText || faultsList.join('\n')), [])
      const includedCount = draftToUse.filter((d) => d.include).length

      const blob = await buildWordReportBlobFromDraft(draftToUse)
      const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
      const filename = `faults-report-${stamp}.docx`

      const url = URL.createObjectURL(blob)
      const a = document.createElement('a')
      a.href = url
      a.download = filename
      document.body.appendChild(a)
      a.click()
      a.remove()
      window.setTimeout(() => URL.revokeObjectURL(url), 1500)

      setReportStatus(includedCount === 0 ? `Generated ${filename} (cover only)` : `Generated ${filename}`)
    } catch (error) {
      setReportError(error instanceof Error ? error.message : 'Failed to generate Word report')
      setReportStatus('')
    }
  }

  const downloadPdfFromPreview = async () => {
    setReportError('')
    setReportStatus('Generating PDF from preview…')

    try {
      // Ensure the preview is up-to-date so the PDF matches what the user sees.
      await previewWordReport()

      const host = docxPreviewHostRef.current
      if (!host) throw new Error('Preview host is not available.')

      const wrapper = host.querySelector('.docx-wrapper') as HTMLElement | null
      if (!wrapper) throw new Error('DOCX preview is not ready yet. Click “Preview DOCX” first.')

      const pages = Array.from(wrapper.querySelectorAll('section.docx')) as HTMLElement[]
      const targets = pages.length ? pages : [wrapper]

      const pdf = new jsPDF({ orientation: 'p', unit: 'pt', format: 'a4' })
      const pdfWidth = pdf.internal.pageSize.getWidth()
      const pdfHeight = pdf.internal.pageSize.getHeight()

      for (let i = 0; i < targets.length; i++) {
        const el = targets[i]

        const canvas = await html2canvas(el, {
          backgroundColor: '#ffffff',
          scale: 2,
          useCORS: true,
          logging: false,
        })

        const imgData = canvas.toDataURL('image/jpeg', 0.92)
        const imgWidth = pdfWidth
        const imgHeight = (canvas.height * imgWidth) / canvas.width

        if (i > 0) pdf.addPage()

        // Fit to page width; if height overflows, fit to height instead.
        if (imgHeight <= pdfHeight) {
          pdf.addImage(imgData, 'JPEG', 0, 0, imgWidth, imgHeight)
        } else {
          const fitWidth = (canvas.width * pdfHeight) / canvas.height
          const x = Math.max(0, (pdfWidth - fitWidth) / 2)
          pdf.addImage(imgData, 'JPEG', x, 0, fitWidth, pdfHeight)
        }
      }

      const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
      pdf.save(`faults-report-${stamp}.pdf`)
      setReportStatus('PDF downloaded.')
    } catch (error) {
      setReportError(error instanceof Error ? error.message : 'Failed to generate PDF from preview')
      setReportStatus('')
    }
  }

  const addDescriptionRow = () => {
    setDescriptionRows((prev) => [...prev, { id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, title: '', text: '' }])
  }

  const updateDescriptionRow = (id: string, patch: Partial<ReportDescriptionRow>) => {
    setDescriptionRows((prev) => prev.map((row) => (row.id === id ? { ...row, ...patch } : row)))
  }

  const removeDescriptionRow = (id: string) => {
    setDescriptionRows((prev) => {
      const next = prev.filter((row) => row.id !== id)
      return next.length ? next : [{ id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, title: '', text: '' }]
    })
  }

  const addEquipmentItem = () => {
    setEquipmentItems((prev) => [...prev, { id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, title: '', text: '', imageFile: null, imagePreviewUrl: null }])
  }

  const updateEquipmentItem = (id: string, patch: Partial<ReportEquipmentItem>) => {
    setEquipmentItems((prev) => prev.map((item) => (item.id === id ? { ...item, ...patch } : item)))
  }

  const removeEquipmentItem = (id: string) => {
    setEquipmentItems((prev) => {
      const next = prev.filter((item) => item.id !== id)
      return next.length ? next : [{ id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, title: '', text: '', imageFile: null, imagePreviewUrl: null }]
    })
  }

  const readFileAsDataUrl = (file: File): Promise<string> => {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onerror = () => reject(new Error('Failed to read image'))
      reader.onload = () => resolve(String(reader.result || ''))
      reader.readAsDataURL(file)
    })
  }

  const setEquipmentItemImage = async (id: string, file: File | null) => {
    if (!file) {
      updateEquipmentItem(id, { imageFile: null, imagePreviewUrl: null })
      return
    }
    const preview = await readFileAsDataUrl(file)
    updateEquipmentItem(id, { imageFile: file, imagePreviewUrl: preview })
  }

  const readLabelsForPath = async (path: string) => {
    try {
      const faultsDir = await getWorkflowFaultsDir(false)
      if (!faultsDir) return []
      const labelsDir = await faultsDir.getDirectoryHandle?.('labels')
      if (!labelsDir) return []
      const stem = path.split('/').pop()?.replace(/\.[^/.]+$/, '') || path
      const labelText = await readTextFile(labelsDir, `${stem}.txt`)
      if (!labelText) return []
      return labelText
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
          const [classId, x, y, w, h, conf] = line.split(/\s+/).map(Number)
          return {
            classId: Number.isFinite(classId) ? classId : 0,
            x: Number.isFinite(x) ? x : 0,
            y: Number.isFinite(y) ? y : 0,
            w: Number.isFinite(w) ? w : 0,
            h: Number.isFinite(h) ? h : 0,
            conf: Number.isFinite(conf) ? conf : 0,
            shape: 'rect' as const,
            source: 'auto' as const,
          }
        })
    } catch {
      return []
    }
  }

  const writeTextFile = async (dir: FileSystemDirectoryHandle, name: string, content: string) => {
    if (!dir.getFileHandle) throw new Error('Folder write is not available in this browser.')
    const fileHandle = await dir.getFileHandle(name, { create: true })
    if (!fileHandle.createWritable) throw new Error('Unable to write files: File System Access API is unavailable or permission is missing.')
    const writable = await fileHandle.createWritable()
    await writable.write(content)
    await writable.close()
  }

  const clearDirectory = async (dir: FileSystemDirectoryHandle) => {
    if (!dir.entries || !dir.removeEntry) return
    for await (const [name, entry] of dir.entries()) {
      await dir.removeEntry(name, { recursive: entry.kind === 'directory' })
    }
  }

  const runScan = async (overwrite: boolean) => {
    if (!folderHandle) {
      setScanError('Choose a folder with File → Open Folder…')
      return
    }
    if (!folderHandle.getDirectoryHandle) {
      setScanError('Folder access is not available in this browser.')
      return
    }

    const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
    if (!canWrite) {
      setScanError('Write permission is required to run a scan and save labels. Reopen the folder and allow write access.')
      return
    }

    setIsScanning(true)
    setScanCompleted(false)
    setScanError('')
    setScanStatus('Preparing scan…')

    try {
      const workflowRoot = await getWorkflowRootDir(true)
      if (!workflowRoot) throw new Error('Unable to access thermal folder')

      const faultsDir = await getFaultsDir(workflowRoot, true)
      if (!faultsDir) throw new Error('Unable to access faults folder')

      if (overwrite) {
        await clearDirectory(faultsDir)
      }

      const labelsDir = await faultsDir.getDirectoryHandle?.('labels', { create: true })
      if (!labelsDir) throw new Error('Unable to create labels folder')

      const existingText = overwrite ? '' : (await readTextFile(faultsDir, 'faults.txt')) || ''
      const existingSet = new Set(
        existingText
          .split(/\r?\n/)
          .map((line) => resolvePathInFileMap(line))
          .filter(Boolean)
      )
      const images = Object.entries(fileMap)
        .filter(([path, file]) => isImagePath(path) && file.type !== '' && isInThermalFolder(path))
        .map(([path, file]) => ({ path, file }))

      setScanProgress({ current: 0, total: images.length })

      const detected: string[] = []

      for (let idx = 0; idx < images.length; idx += 1) {
        const { path, file } = images[idx]
        setScanStatus(`Scanning ${idx + 1}/${images.length}: ${path}`)
        setScanProgress({ current: idx + 1, total: images.length })

        const formData = new FormData()
        formData.append('file', file, path)

        let response: Response
        try {
          response = await fetch(apiUrl('/api/scan/detect'), {
            method: 'POST',
            body: formData,
          })
        } catch {
          throw new Error('Backend is not reachable. Start the FastAPI server on http://127.0.0.1:8000 and try again.')
        }

        if (!response.ok) {
          const err = await response.json().catch(() => ({}))
          throw new Error(err.detail || 'Scan failed')
        }

        const data = await response.json()
        if (data?.hasFaults && Array.isArray(data.labels)) {
          if (!existingSet.has(path)) {
            const stem = path.split('/').pop()?.replace(/\.[^/.]+$/, '') || path
            const labelText = data.labels
              .map((label: any) => `${label.classId} ${label.x} ${label.y} ${label.w} ${label.h} ${label.conf}`)
              .join('\n')
            await writeTextFile(labelsDir, `${stem}.txt`, labelText)
            detected.push(path)
            existingSet.add(path)
          }
        }
      }

      const merged = [...existingSet]
      await writeTextFile(faultsDir, 'faults.txt', merged.join('\n'))
      await writeTextFile(workflowRoot, 'faults.txt', merged.join('\n'))
      setFaultsList(merged)
      setScanCompleted(true)
      setScanStatus(`Scan complete. Faults found: ${detected.length}`)
    } catch (error) {
      setScanError(error instanceof Error ? error.message : 'Scan failed')
      setScanStatus('')
    } finally {
      setIsScanning(false)
      setScanPromptOpen(false)
    }
  }

  const handleScanClick = async () => {
    setScanError('')
    setScanStatus('')
    if (!folderHandle) {
      setScanError('Choose a folder with File → Open Folder…')
      return
    }
    if (!folderHandle.getDirectoryHandle) {
      setScanError('Folder access is not available in this browser.')
      return
    }

    try {
      const existing = await getWorkflowFaultsDir(false)
      if (existing) {
        setScanPromptOpen(true)
        return
      }
      await runScan(false)
    } catch {
      await runScan(false)
    }
  }

  const handleSelectFile = async (path: string) => {
    const file = fileMap[path]
    if (!file) return
    setSelectedPath(path)
    if (fileUrl) URL.revokeObjectURL(fileUrl)
    setFileUrl('')
    setFileText('')
    setSelectedLabels([])
    setSelectedLabelIndex(null)
    setShowLabels(true)
    setZoom(1)
    setPan({ x: 0, y: 0 })
    setDrawMode('select')

    if (file.type.startsWith('image/')) {
      setFileKind('image')
      setFileUrl(URL.createObjectURL(file))
      const labels = await readLabelsForPath(path)
      setSelectedLabels(labels)
      setSelectedLabelIndex(null)
      setShowLabels(true)
      setZoom(1)
      setPan({ x: 0, y: 0 })
      setDrawMode('select')
      setLabelHistory([labels])
      setLabelHistoryIndex(0)
      return
    }

    if (file.type.startsWith('text/') || file.name.match(/\.(md|txt|json|csv|ts|tsx|js|jsx|py|html|css)$/i)) {
      setFileKind('text')
      const text = await file.text()
      setFileText(text)
      return
    }

    setFileKind('other')
  }

  const rootNode = useMemo(() => {
    if (!fileTree?.children?.length) return null
    const stack: TreeNode[] = [...fileTree.children]
    while (stack.length > 0) {
      const node = stack.shift()!
      if (node.type === 'folder' && node.name.toLowerCase() === 'thermal') return node
      if (node.type === 'folder' && node.children && node.children.length > 0) {
        stack.unshift(...node.children)
      }
    }
    return fileTree.children.find((node) => node.type === 'folder') ?? null
  }, [fileTree])

  const rootLabel = useMemo(() => {
    if (!fileTree?.children?.length) return 'No folder selected'
    return rootNode ? rootNode.name : 'Folder'
  }, [fileTree, rootNode])

  const rootChildren = useMemo(() => {
    if (rootNode) return rootNode.children ?? []
    return fileTree?.children ?? []
  }, [fileTree, rootNode])

  const hasFolderSelected = useMemo(
    () => Boolean(fileTree && (rootNode || (fileTree.children && fileTree.children.length > 0))),
    [fileTree, rootNode]
  )

  const showLeftExplorer = hasFolderSelected && (activeMenu === 'file' || (activeMenu === 'actions' && activeAction === 'view'))
  const showRightExplorer = hasFolderSelected && activeMenu === 'actions' && activeAction === 'view'

  const gridTemplateColumns = showLeftExplorer && showRightExplorer
    ? `${isExplorerMinimized ? 0 : explorerWidth}px 1fr ${isRightExplorerMinimized ? 0 : rightExplorerWidth}px`
    : showLeftExplorer
      ? `${isExplorerMinimized ? 0 : explorerWidth}px 1fr`
      : showRightExplorer
        ? `1fr ${isRightExplorerMinimized ? 0 : rightExplorerWidth}px`
        : '1fr'

  const getFileIcon = (name: string) => {
    if (isImageName(name)) return '📷'
    if (/\.(md|txt)$/i.test(name)) return '🖹'
    if (/\.(json|csv)$/i.test(name)) return '📋'
    if (/\.(ts|tsx|js|jsx)$/i.test(name)) return '🧩'
    if (/\.(py)$/i.test(name)) return '🐍'
    if (/\.(html|css)$/i.test(name)) return '🌐'
    return '📄'
  }

  const getNodePriority = (node: TreeNode) => {
    if (node.type === 'folder') return 0
    return isImageName(node.name) ? 2 : 1
  }

  const getLabelFileName = (path: string) => {
    const stem = path.split('/').pop()?.replace(/\.[^/.]+$/, '') || path
    return `${stem}.txt`
  }

  const openExplorerContextMenu = (event: React.MouseEvent, node: TreeNode, source: 'tree' | 'faultsList') => {
    event.preventDefault()
    event.stopPropagation()
    const estimatedWidth = 220
    const estimatedHeight = 92
    const x = Math.min(event.clientX, Math.max(0, window.innerWidth - estimatedWidth))
    const y = Math.min(event.clientY, Math.max(0, window.innerHeight - estimatedHeight))
    setExplorerContextMenu({ x, y, node, source })
  }

  const removeFaultFromList = async (path: string) => {
    if (!folderHandle) {
      window.alert('This action requires opening a folder via “Choose Folder…”.')
      return
    }
    if (!folderHandle.getDirectoryHandle) {
      window.alert('Folder write is not available in this browser.')
      return
    }

    const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
    if (!canWrite) {
      window.alert('Write permission is required. Reopen the folder and allow write access.')
      return
    }

    const ok = window.confirm(`Remove “${path.split('/').pop() || path}” from faults.txt? (This will NOT delete the image file.)`)
    if (!ok) return

    const nextFaults = faultsList.filter((p) => p !== path)
    setFaultsList(nextFaults)

    try {
      const workflowRoot = await getWorkflowRootDir(true)
      if (!workflowRoot) throw new Error('Unable to access thermal folder')
      const faultsDir = await getFaultsDir(workflowRoot, true)
      if (faultsDir) {
        await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))
      }
      await writeTextFile(workflowRoot, 'faults.txt', nextFaults.join('\n'))
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Failed to update faults list')
    }

    if (selectedPath === path) {
      setSelectedPath('')
      setFileText('')
      if (fileUrl) URL.revokeObjectURL(fileUrl)
      setFileUrl('')
      setFileKind('')
      setSelectedLabels([])
      setSelectedLabelIndex(null)
    }
  }

  const deleteExplorerNode = async (node: TreeNode) => {
    if (!folderHandle) {
      window.alert('Delete is only available when a folder is opened via “Choose Folder…”.')
      return
    }
    if (!folderHandle.getDirectoryHandle) {
      window.alert('Folder delete is not available in this browser.')
      return
    }

    const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
    if (!canWrite) {
      window.alert('Write permission is required to delete files. Reopen the folder and allow write access.')
      return
    }

    const isFolder = node.type === 'folder'
    const ok = window.confirm(
      isFolder ? `Delete folder “${node.name}” and all its contents?` : `Delete file “${node.name}”?`
    )
    if (!ok) return

    try {
      const parentPath = getParentDirPath(node.path)
      const parts = parentPath ? parentPath.split('/').filter(Boolean) : []
      let parentDir: FileSystemDirectoryHandle = folderHandle
      for (const part of parts) {
        if (!parentDir.getDirectoryHandle) throw new Error('Folder traversal is not available in this browser.')
        parentDir = await parentDir.getDirectoryHandle(part)
      }
      if (!parentDir.removeEntry) throw new Error('Delete is not supported by this browser (removeEntry unavailable).')

      await parentDir.removeEntry(node.name, { recursive: isFolder })

      if (!isFolder && faultsList.includes(node.path)) {
        const nextFaults = faultsList.filter((p) => p !== node.path)
        setFaultsList(nextFaults)
        const workflowRoot = await getWorkflowRootDir(true)
        const faultsDir = workflowRoot ? await getFaultsDir(workflowRoot, true) : null
        if (faultsDir) {
          await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))
        }
        if (workflowRoot) await writeTextFile(workflowRoot, 'faults.txt', nextFaults.join('\n'))
      }

      if (!isFolder) {
        const faultsDir = await getWorkflowFaultsDir(false)
        const labelsDir = faultsDir ? await faultsDir.getDirectoryHandle?.('labels') : null
        if (labelsDir?.removeEntry) {
          await labelsDir.removeEntry(getLabelFileName(node.path)).catch(() => undefined)
        }
      }

      if (isFolder) {
        setExpandedPaths((prev) => {
          const next = new Set(Array.from(prev).filter((p) => !(p && (p === node.path || p.startsWith(`${node.path}/`)))))
          next.add('')
          return next
        })
      }

      await refreshFolderEntries(folderHandle)
    } catch (error) {
      window.alert(error instanceof Error ? error.message : 'Delete failed')
    }
  }

  const pushLabelHistory = (labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>) => {
    setLabelHistory((prev) => {
      const trimmed = prev.slice(0, labelHistoryIndex + 1)
      const next = [...trimmed, labels]
      setLabelHistoryIndex(next.length - 1)
      return next
    })
  }

  const applyHistoryIndex = async (nextIndex: number) => {
    const snapshot = labelHistory[nextIndex]
    if (!snapshot) return
    setSelectedLabels(snapshot)
    setSelectedLabelIndex(null)
    setLabelHistoryIndex(nextIndex)
    await persistLabels(snapshot)
  }

  const persistLabels = async (labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; source?: 'auto' | 'manual' }>) => {
    try {
      if (!folderHandle || !folderHandle.getDirectoryHandle || !selectedPath) return

      const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
      if (!canWrite) {
        setLabelSaveError('Write permission denied. Reopen the folder and allow write access to save labels.')
        return
      }

      const workflowRoot = await getWorkflowRootDir(true)
      if (!workflowRoot) throw new Error('Unable to access thermal folder')
      const faultsDir = await getFaultsDir(workflowRoot, true)
      if (!faultsDir) throw new Error('Unable to access faults folder')
      const labelsDir = await faultsDir.getDirectoryHandle?.('labels', { create: true })
      if (!labelsDir) throw new Error('Unable to access labels folder')

      const lines = labels.map((label) =>
        `${label.classId} ${label.x} ${label.y} ${label.w} ${label.h} ${Number.isFinite(label.conf) ? label.conf : 1}`
      )
      await writeTextFile(labelsDir, getLabelFileName(selectedPath), lines.join('\n'))

      if (labels.length > 0) {
        if (!faultsList.includes(selectedPath)) {
          const nextFaults = [...faultsList, selectedPath]
          setFaultsList(nextFaults)
          await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))
          await writeTextFile(workflowRoot, 'faults.txt', nextFaults.join('\n'))
        }
      } else {
        if (faultsList.includes(selectedPath)) {
          const nextFaults = faultsList.filter((path) => path !== selectedPath)
          setFaultsList(nextFaults)
          await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))
          await writeTextFile(workflowRoot, 'faults.txt', nextFaults.join('\n'))
        }
      }

      setLabelSaveError('')
    } catch (error) {
      setLabelSaveError(error instanceof Error ? error.message : 'Failed to save labels')
    }
  }

  const scheduleRealtimePersist = (
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
  ) => {
    if (!folderHandle || !selectedPath) return
    realtimePersistLabelsRef.current = labels
    if (realtimePersistTimerRef.current !== null) return
    realtimePersistTimerRef.current = window.setTimeout(() => {
      realtimePersistTimerRef.current = null
      const latest = realtimePersistLabelsRef.current
      if (!latest) return
      void persistLabels(latest)
    }, 150)
  }

  const clamp01 = (value: number) => Math.min(Math.max(value, 0), 1)

  // Maps a pointer event to normalized image coordinates (0..1) based on the
  // actual rendered <img> rect. This stays correct under zoom/pan transforms.
  const toImageCoords = (event: React.MouseEvent) => {
    const img = imageRef.current
    if (!img) return null
    const rect = img.getBoundingClientRect()
    if (!rect.width || !rect.height) return null

    const nx = (event.clientX - rect.left) / rect.width
    const ny = (event.clientY - rect.top) / rect.height

    return {
      x: clamp01(nx),
      y: clamp01(ny),
    }
  }

  const clampPanToBounds = (nextPan: { x: number; y: number }) => {
    // No panning at 100% zoom.
    if (zoom <= 1) return { x: 0, y: 0 }

    const viewport = imageContainerRef.current
    const baseW = imageMetrics.width
    const baseH = imageMetrics.height
    if (!viewport || !baseW || !baseH) return nextPan

    const viewportW = viewport.clientWidth
    const viewportH = viewport.clientHeight
    if (!viewportW || !viewportH) return nextPan

    const scale = zoom * BASE_ZOOM
    const scaledW = baseW * scale
    const scaledH = baseH * scale

    const maxX = Math.max(0, (scaledW - viewportW) / 2)
    const maxY = Math.max(0, (scaledH - viewportH) / 2)

    return {
      x: Math.min(maxX, Math.max(-maxX, nextPan.x)),
      y: Math.min(maxY, Math.max(-maxY, nextPan.y)),
    }
  }

  const getViewerSize = () => {
    const naturalW = imageMetrics.naturalWidth
    const naturalH = imageMetrics.naturalHeight

    // If we don't know the image yet, keep it flexible.
    if (!naturalW || !naturalH) return null

    const availableW = Math.min(VIEWER_MAX_W, viewerHostSize.width || VIEWER_MAX_W)
    const availableH = Math.min(VIEWER_MAX_H, viewerHostSize.height || VIEWER_MAX_H)
    const scale = Math.min(availableW / naturalW, availableH / naturalH)
    return {
      width: Math.round(naturalW * scale),
      height: Math.round(naturalH * scale),
    }
  }

  useEffect(() => {
    // When zoom changes or the image measures change, keep the current pan valid.
    setPan((prev) => clampPanToBounds(prev))
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [zoom, imageMetrics.width, imageMetrics.height])

  const hitTestLabel = (point: { x: number; y: number }) => {
    for (let i = selectedLabels.length - 1; i >= 0; i -= 1) {
      const label = selectedLabels[i]
      const left = label.x - label.w / 2
      const right = label.x + label.w / 2
      const top = label.y - label.h / 2
      const bottom = label.y + label.h / 2
      if (point.x >= left && point.x <= right && point.y >= top && point.y <= bottom) {
        return i
      }
    }
    return null
  }

  const resizeLabel = (
    origin: { x: number; y: number; w: number; h: number },
    handle: 'nw' | 'ne' | 'sw' | 'se',
    point: { x: number; y: number }
  ) => {
    const clamp = (value: number, min: number, max: number) => Math.min(Math.max(value, min), max)
    const minSize = 0.01

    // Start from the current label bounds (normalized 0..1).
    let left = origin.x - origin.w / 2
    let right = origin.x + origin.w / 2
    let top = origin.y - origin.h / 2
    let bottom = origin.y + origin.h / 2

    // Clamp the origin bounds first (defensive: avoids negative sizes if input is slightly out-of-range).
    left = clamp(left, 0, 1)
    right = clamp(right, 0, 1)
    top = clamp(top, 0, 1)
    bottom = clamp(bottom, 0, 1)

    // Ensure origin has at least min size.
    if (right - left < minSize) {
      const mid = (left + right) / 2
      left = clamp(mid - minSize / 2, 0, 1 - minSize)
      right = left + minSize
    }
    if (bottom - top < minSize) {
      const mid = (top + bottom) / 2
      top = clamp(mid - minSize / 2, 0, 1 - minSize)
      bottom = top + minSize
    }

    // Move only the dragged corner, keeping the opposite corner fixed.
    switch (handle) {
      case 'nw':
        left = clamp(point.x, 0, right - minSize)
        top = clamp(point.y, 0, bottom - minSize)
        break
      case 'ne':
        right = clamp(point.x, left + minSize, 1)
        top = clamp(point.y, 0, bottom - minSize)
        break
      case 'sw':
        left = clamp(point.x, 0, right - minSize)
        bottom = clamp(point.y, top + minSize, 1)
        break
      case 'se':
        right = clamp(point.x, left + minSize, 1)
        bottom = clamp(point.y, top + minSize, 1)
        break
    }

    const w = right - left
    const h = bottom - top
    return {
      x: left + w / 2,
      y: top + h / 2,
      w: clamp(w, minSize, 1),
      h: clamp(h, minSize, 1),
    }
  }

  const sortNodes = (nodes: TreeNode[]) =>
    [...nodes].sort((a, b) => {
      const priorityDiff = getNodePriority(a) - getNodePriority(b)
      if (priorityDiff !== 0) return priorityDiff
      return a.name.localeCompare(b.name)
    })

  const toggleFolder = (path: string) => {
    setExpandedPaths((prev) => {
      const next = new Set(prev)
      if (next.has(path)) {
        next.delete(path)
      } else {
        next.add(path)
      }
      return next
    })
  }

  const renderTree = (node: TreeNode, depth = 0) => {
    const paddingLeft = depth === 0 ? 2 : 2 + depth * 12
    if (node.type === 'file') {
      return (
        <button
          key={node.path}
          className={`tree-item file ${selectedPath === node.path ? 'active' : ''}`}
          style={{ paddingLeft }}
          onClick={() => handleSelectFileFromScope('tree', node.path)}
          onContextMenu={(event) => openExplorerContextMenu(event, node, 'tree')}
        >
          <span className="file-label">
            <span className="file-icon">{getFileIcon(node.name)}</span>
            <span className="file-name">{node.name}</span>
          </span>
        </button>
      )
    }

    const isExpanded = expandedPaths.has(node.path)
    return (
      <div key={node.path || node.name} className="tree-group">
        <button
          className="tree-item folder"
          style={{ paddingLeft }}
          onClick={() => toggleFolder(node.path)}
          onContextMenu={(event) => openExplorerContextMenu(event, node, 'tree')}
        >
          <span className="file-label">
            <span className="file-icon">{isExpanded ? '▾' : '▸'} 📁</span>
            <span className="file-name">{node.name}</span>
          </span>
        </button>
        {isExpanded && node.children && node.children.length > 0 && (
          <div className="tree-children">
            {sortNodes(node.children).map((child) => renderTree(child, depth + 1))}
          </div>
        )}
      </div>
    )
  }

  useEffect(() => {
    if (!openMenu) return

    const handleMove = (event: MouseEvent) => {
      if (!menuAreaRef.current) return
      const rect = menuAreaRef.current.getBoundingClientRect()
      const dynamicBuffer = Math.max(rect.width, rect.height)
      const buffer = Math.max(120, dynamicBuffer)
      const left = rect.left - buffer
      const right = rect.right + buffer
      const top = rect.top - buffer
      const bottom = rect.bottom + buffer

      if (
        event.clientX < left ||
        event.clientX > right ||
        event.clientY < top ||
        event.clientY > bottom
      ) {
        setOpenMenu('')
      }
    }

    window.addEventListener('mousemove', handleMove)
    return () => window.removeEventListener('mousemove', handleMove)
  }, [openMenu])

  useEffect(() => {
    if (!explorerContextMenu) return

    const close = () => setExplorerContextMenu(null)
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') close()
    }

    window.addEventListener('keydown', handleKeyDown)
    window.addEventListener('click', close)
    window.addEventListener('blur', close)
    window.addEventListener('resize', close)
    window.addEventListener('scroll', close, true)
    return () => {
      window.removeEventListener('keydown', handleKeyDown)
      window.removeEventListener('click', close)
      window.removeEventListener('blur', close)
      window.removeEventListener('resize', close)
      window.removeEventListener('scroll', close, true)
    }
  }, [explorerContextMenu])

  useEffect(() => {
    if (activeAction !== 'view') return
    if (!selectedPath || fileKind !== 'image') return

    const shouldIgnoreKeyTarget = (event: KeyboardEvent) => {
      const target = event.target as (HTMLElement | null)
      if (!target) return false
      if ((target as HTMLElement).isContentEditable) return true
      const tag = target.tagName?.toLowerCase()
      return tag === 'input' || tag === 'textarea' || tag === 'select'
    }

    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.altKey || event.ctrlKey || event.metaKey) return
      if (shouldIgnoreKeyTarget(event)) return
      if (event.key !== 'ArrowLeft' && event.key !== 'ArrowRight') return

      if (event.key === 'ArrowLeft') {
        if (!prevImagePath) return
        event.preventDefault()
        void handleSelectFile(prevImagePath)
        return
      }

      if (!nextImagePath) return
      event.preventDefault()
      void handleSelectFile(nextImagePath)
    }

    window.addEventListener('keydown', handleKeyDown)
    return () => window.removeEventListener('keydown', handleKeyDown)
  }, [activeAction, fileKind, nextImagePath, prevImagePath, selectedPath])

  useEffect(() => {
    if (!isResizing) return

    const handleMove = (event: MouseEvent) => {
      const min = 200
      const max = 520
      const nextWidth = Math.min(max, Math.max(min, event.clientX))
      setExplorerWidth(nextWidth)
      if (isExplorerMinimized) {
        setIsExplorerMinimized(false)
      }
    }

    const handleUp = () => setIsResizing(false)

    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    return () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
    }
  }, [isResizing])

  useEffect(() => {
    if (!isRightResizing) return

    const handleMove = (event: MouseEvent) => {
      const min = 200
      const max = 520
      const viewportWidth = window.innerWidth
      const nextWidth = Math.min(max, Math.max(min, viewportWidth - event.clientX))
      setRightExplorerWidth(nextWidth)
      if (isRightExplorerMinimized) {
        setIsRightExplorerMinimized(false)
      }
    }

    const handleUp = () => setIsRightResizing(false)

    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    return () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
    }
  }, [isRightResizing, isRightExplorerMinimized])

  useEffect(() => {
    if (!isReportDraftResizing) return

    const handleMove = (event: MouseEvent) => {
      const node = reportSplitRef.current
      if (!node) return
      const rect = node.getBoundingClientRect()

      const minSidebar = 320
      const maxSidebar = 720
      const nextWidth = Math.min(maxSidebar, Math.max(minSidebar, rect.right - event.clientX))
      setReportDraftWidth(nextWidth)
    }

    const handleUp = () => setIsReportDraftResizing(false)

    window.addEventListener('mousemove', handleMove)
    window.addEventListener('mouseup', handleUp)
    return () => {
      window.removeEventListener('mousemove', handleMove)
      window.removeEventListener('mouseup', handleUp)
    }
  }, [isReportDraftResizing])

  return (
    <div className="app">
      <header className="topbar">
        <div
          ref={menuAreaRef}
          className="topbar-left"
          onMouseEnter={() => {
            if (closeTimerRef.current) {
              window.clearTimeout(closeTimerRef.current)
              closeTimerRef.current = null
            }
          }}
          onMouseLeave={() => {
            closeTimerRef.current = window.setTimeout(() => {
              setOpenMenu('')
            }, 1000)
          }}
        >
          <nav className="topmenu">
            <div className="menu-item">
              <button
                onMouseEnter={() => setOpenMenu('')}
                onClick={() => {
                  setOpenMenu('')
                  setActiveMenu('file')
                  setActiveAction('')
                }}
                title="Home"
              >
                Home
              </button>
            </div>
            <div className="menu-item">
              <button
                className={openMenu === 'actions' ? 'active' : ''}
                onMouseEnter={() => setOpenMenu('actions')}
                onClick={() => setOpenMenu(openMenu === 'actions' ? '' : 'actions')}
              >
                Actions
              </button>
              {openMenu === 'actions' && (
                <div className="dropdown">
                  <button
                    className={`dropdown-item ${!hasFolderSelected ? 'disabled' : ''}`}
                    disabled={!hasFolderSelected}
                    onClick={() => {
                      setActiveMenu('actions')
                      setActiveAction('view')
                    }}
                  >
                    View
                  </button>
                  <button
                    className={`dropdown-item ${!hasFolderSelected ? 'disabled' : ''}`}
                    disabled={!hasFolderSelected}
                    onClick={() => {
                      setActiveMenu('actions')
                      setActiveAction('scan')
                    }}
                  >
                    Scan
                  </button>
                  <button
                    className={`dropdown-item ${!hasFolderSelected ? 'disabled' : ''}`}
                    disabled={!hasFolderSelected}
                    onClick={() => {
                      setActiveMenu('actions')
                      setActiveAction('report')
                      void loadReportFromDisk()
                    }}
                  >
                    Report
                  </button>
                </div>
              )}
            </div>
            <div className="menu-item">
              <button
                className={openMenu === 'help' ? 'active' : ''}
                onMouseEnter={() => setOpenMenu('help')}
                onClick={() => setOpenMenu(openMenu === 'help' ? '' : 'help')}
              >
                Help
              </button>
              {openMenu === 'help' && (
                <div className="dropdown">
                  <button className="dropdown-item" onClick={() => setActiveMenu('help')}>
                    Documentation
                  </button>
                  <button className="dropdown-item" onClick={() => setActiveMenu('help')}>
                    About
                  </button>
                </div>
              )}
            </div>
          </nav>
        </div>
        <div className="topbar-right">
          <button
            className="theme-toggle"
            onClick={() => setTheme((prev) => (prev === 'dark' ? 'light' : 'dark'))}
            title={theme === 'dark' ? 'Switch to light mode' : 'Switch to dark mode'}
          >
            {theme === 'dark' ? '☀️' : '🌙'}
          </button>
        </div>
      </header>

      <div
        className="body"
        style={{
          gridTemplateColumns,
        }}
      >
        {showLeftExplorer && (
          <>
            <aside className={`explorer ${isExplorerMinimized ? 'minimized' : ''}`}>
              <div className="explorer-header">
                <div className="explorer-title">Explorer</div>
                <button
                  className="explorer-toggle"
                  onClick={() => setIsExplorerMinimized((prev) => !prev)}
                  title={isExplorerMinimized ? 'Restore explorer' : 'Minimize explorer'}
                >
                  {isExplorerMinimized ? '▸' : '▾'}
                </button>
              </div>
              {!isExplorerMinimized && (
                <>
                  <div className="explorer-pane">
                    <div className="explorer-pane-title">Temp title A</div>
                    <div className="tree">
                      <button
                        className="tree-item folder root"
                        style={{ paddingLeft: 2 }}
                        onClick={() => toggleFolder('')}
                      >
                        {expandedPaths.has('') ? '▾' : '▸'} {rootLabel}
                      </button>
                      {expandedPaths.has('') && sortNodes(rootChildren).map((child) => renderTree(child, 1))}
                    </div>
                  </div>
                  <div className="explorer-pane">
                    <div className="explorer-pane-title">Temp title B</div>
                    <div className="tree">
                      {activeAction === 'view' ? (
                        faultsList.length > 0 ? (
                          faultsList.map((path) => {
                            const name = path.split('/').pop() || path
                            const isMissing = !fileMap[path]
                            return (
                              <button
                                key={path}
                                className={`tree-item file ${selectedPath === path ? 'active' : ''}`}
                                style={{ paddingLeft: 2 }}
                                onClick={() => handleSelectFileFromScope('faultsList', path)}
                                onContextMenu={(event) => openExplorerContextMenu(event, { name, path, type: 'file' }, 'faultsList')}
                                disabled={isMissing}
                                title={isMissing ? 'File not found in folder' : path}
                              >
                                <span className="file-label">
                                  <span className="file-icon">📷</span>
                                  <span className="file-name">{name}</span>
                                </span>
                              </button>
                            )
                          })
                        ) : (
                          <div className="explorer-empty">No faults listed yet.</div>
                        )
                      ) : (
                        <div className="tree-item file" style={{ paddingLeft: 2 }}>
                          <span className="file-label">
                            <span className="file-icon">📄</span>
                            <span className="file-name">Placeholder</span>
                          </span>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="explorer-resizer" onMouseDown={() => setIsResizing(true)} />
                </>
              )}
            </aside>
            {isExplorerMinimized && (
              <button
                className="explorer-toggle-floating"
                onClick={() => setIsExplorerMinimized(false)}
                title="Restore explorer"
              >
                ▸
              </button>
            )}
          </>
        )}
        <main className="content">
          {/* <header className="page-header">
            <h1>Simple One‑Page App</h1>
            <p>React frontend calling a FastAPI backend.</p>
          </header> */}

          {activeMenu === 'file' && (
            <section id="welcome-section" className="card home-card">
              <div className="home-hero">
                <div className="home-hero-text">
                  <span className="home-badge">Workspace</span>
                  <h1 id="welcome-header">Welcome</h1>
                  <p style={{ color: '#e2e8f0' }}>Open a folder to browse files and preview their contents.</p>
                  <div className="home-actions">
                    <button
                      className="link-button"
                      onClick={handleChooseFolder}
                    >
                      {hasFolderSelected ? 'Change Folder...' : 'Choose Folder...'}
                    </button>
                    {hasFolderSelected && (
                      <button className="link-button" onClick={() => void handleRemoveFolder()}>
                        Close Folder
                      </button>
                    )}
                    {hasFolderSelected && (
                      <button
                        className="link-button"
                        onClick={() => {
                          setActiveMenu('actions')
                          setActiveAction('scan')
                        }}
                      >
                        Scan for faults
                      </button>
                    )}
                  </div>
                  <input
                    ref={(node) => {
                      folderInputRef.current = node
                      node?.setAttribute('webkitdirectory', '')
                    }}
                    className="hidden-input"
                    type="file"
                    onChange={handleOpenFolder}
                  />
                </div>
                <div className="home-hero-card">
                  <div className="home-hero-title">Quick tips</div>
                  <ul>
                    <li>Use File → Open Folder…</li>
                    <li>Explore images and text files</li>
                    <li>Preview updates instantly</li>
                  </ul>
                </div>
              </div>
            </section>
          )}

          {activeMenu === 'actions' && (
            <section className={`card ${activeAction === 'view' ? 'card-view' : ''} ${activeAction === 'scan' ? 'card-scan' : ''} ${activeAction === 'report' ? 'card-report' : ''}`}>
              <div className="page-header-row">
                <h2>
                  {activeAction === 'view' && 'View'}
                  {activeAction === 'scan' && 'Scan'}
                  {activeAction === 'report' && 'Report'}
                  {!activeAction && 'Actions'}
                </h2>
                {activeAction === 'view' && selectedPath && fileKind === 'image' && fileUrl && (
                  <div className="view-toolbar">
                    <div className="view-zoom">
                      <button
                        className="link-button"
                        onClick={() => setZoom((prev) => Math.max(0.5, Number((prev - 0.1).toFixed(2))))}
                      >
                        −
                      </button>
                      <button
                        className="link-button"
                        onClick={() => {
                          setZoom(1)
                          setPan({ x: 0, y: 0 })
                        }}
                      >
                        {Math.round(zoom * 100)}%
                      </button>
                      <button
                        className="link-button"
                        onClick={() => setZoom((prev) => Math.min(3, Number((prev + 0.1).toFixed(2))))}
                      >
                        +
                      </button>
                    </div>
                    <button
                      className="link-button"
                      onClick={() => setShowLabels((prev) => !prev)}
                    >
                      {showLabels ? 'Hide labels' : 'Show labels'}
                    </button>
                  </div>
                )}
              </div>
              {activeAction === 'view' && labelSaveError && (
                <div className="view-error">{labelSaveError}</div>
              )}
              {activeAction === 'view' && (
                <div className="view-preview">
                  {!selectedPath && <p>Select a file from the explorer.</p>}
                  {selectedPath && fileKind === 'image' && fileUrl && (
                    <div className="image-preview" ref={viewerHostRef}>
                      <div className="viewer-frame">
                      {(() => {
                        const size = getViewerSize()
                        return (
                      <div className="viewer-shell">
                        <button
                          type="button"
                          className="viewer-shell-nav viewer-shell-nav-left"
                          onClick={() => {
                            if (prevImagePath) void handleSelectFile(prevImagePath)
                          }}
                          disabled={!prevImagePath}
                          aria-label="Previous image"
                          title="Previous image"
                        >
                          ‹
                        </button>

                        <div
                          className={`image-viewport ${isPanning ? 'panning' : zoom > 1 ? 'can-pan' : ''}`}
                          ref={imageContainerRef}
                          style={size ? { width: `${size.width}px`, height: `${size.height}px` } : undefined}
                        onMouseMove={(event) => {
                          const point = toImageCoords(event)
                          if (!point) return
                          if (isPanning && panStart && panOrigin) {
                            const dx = event.clientX - panStart.x
                            const dy = event.clientY - panStart.y
                            setPan(clampPanToBounds({ x: panOrigin.x + dx, y: panOrigin.y + dy }))
                            return
                          }
                          if (isResizingLabel && resizeHandle && resizeOrigin) {
                            const resized = resizeLabel(resizeOrigin, resizeHandle, point)
                            const next = selectedLabels.map((label, index) => {
                              if (index !== selectedLabelIndex) return label
                              return {
                                ...label,
                                x: resized.x,
                                y: resized.y,
                                w: resized.w,
                                h: resized.h,
                              }
                            })
                            selectedLabelsRef.current = next
                            setSelectedLabels(next)
                            scheduleRealtimePersist(next)
                            return
                          }
                          if (isDrawing && drawStart) {
                            const left = Math.min(drawStart.x, point.x)
                            const right = Math.max(drawStart.x, point.x)
                            const top = Math.min(drawStart.y, point.y)
                            const bottom = Math.max(drawStart.y, point.y)
                            const w = right - left
                            const h = bottom - top
                            setDraftLabel((prev) =>
                              prev
                                ? {
                                    ...prev,
                                    x: left + w / 2,
                                    y: top + h / 2,
                                    w,
                                    h,
                                  }
                                : null
                            )
                          }
                          if (isDraggingLabel && dragStart && dragOrigin !== null) {
                            const dx = point.x - dragStart.x
                            const dy = point.y - dragStart.y
                            const next = selectedLabels.map((label, index) => {
                              if (index !== selectedLabelIndex) return label
                              return {
                                ...label,
                                x: Math.min(Math.max(dragOrigin.x + dx, 0), 1),
                                y: Math.min(Math.max(dragOrigin.y + dy, 0), 1),
                              }
                            })
                            selectedLabelsRef.current = next
                            setSelectedLabels(next)
                            scheduleRealtimePersist(next)
                          }
                        }}
                        onMouseUp={() => {
                          if (isDrawing && draftLabel) {
                            const minSize = 0.01
                            if (draftLabel.w > minSize && draftLabel.h > minSize) {
                              const next = [
                                ...selectedLabels,
                                {
                                  classId: 0,
                                  x: draftLabel.x,
                                  y: draftLabel.y,
                                  w: draftLabel.w,
                                  h: draftLabel.h,
                                  conf: 1,
                                  shape: draftLabel.shape,
                                  source: 'manual' as const,
                                },
                              ]
                              setSelectedLabels(next)
                              setSelectedLabelIndex(next.length - 1)
                              pushLabelHistory(next)
                              persistLabels(next)
                            }
                          }
                          if (isResizingLabel) {
                            const current = selectedLabelsRef.current
                            pushLabelHistory(current)
                            persistLabels(current)
                          }
                          if (isDraggingLabel) {
                            const current = selectedLabelsRef.current
                            pushLabelHistory(current)
                            persistLabels(current)
                          }
                          if (isPanning) {
                            setIsPanning(false)
                            setPanStart(null)
                            setPanOrigin(null)
                          }
                          setIsDrawing(false)
                          setDrawStart(null)
                          setDraftLabel(null)
                          setIsDraggingLabel(false)
                          setDragStart(null)
                          setDragOrigin(null)
                          setIsResizingLabel(false)
                          setResizeHandle(null)
                          setResizeOrigin(null)
                        }}
                        onMouseLeave={() => {
                          if (isDrawing || isDraggingLabel || isResizingLabel || isPanning) {
                            setIsDrawing(false)
                            setDrawStart(null)
                            setDraftLabel(null)
                            setIsDraggingLabel(false)
                            setDragStart(null)
                            setDragOrigin(null)
                            setIsResizingLabel(false)
                            setResizeHandle(null)
                            setResizeOrigin(null)
                            setIsPanning(false)
                            setPanStart(null)
                            setPanOrigin(null)
                          }
                        }}
                        >
                        <div
                          className="image-zoom"
                          style={{ transform: `translate(${pan.x}px, ${pan.y}px) scale(${zoom * BASE_ZOOM})` }}
                        >
                          <img
                            ref={imageRef}
                            src={fileUrl}
                            alt={selectedPath}
                            className="preview-image"
                            draggable={false}
                            onLoad={(event) => {
                              const img = event.currentTarget
                              setImageMetrics({
                                width: img.clientWidth,
                                height: img.clientHeight,
                                naturalWidth: img.naturalWidth,
                                naturalHeight: img.naturalHeight,
                              })
                            }}
                          />
                          <div
                            className={`draw-surface ${drawMode}`}
                            onMouseDown={(event) => {
                          if (!fileUrl) return
                          const point = toImageCoords(event)
                          if (!point) return
                          if (drawMode === 'rect' || drawMode === 'ellipse') {
                            setIsDrawing(true)
                            setDrawStart(point)
                            setDraftLabel({
                              shape: drawMode,
                              x: point.x,
                              y: point.y,
                              w: 0,
                              h: 0,
                            })
                          }
                          if (drawMode === 'select') {
                            const hitIndex = hitTestLabel(point)
                            if (hitIndex !== null) {
                              setSelectedLabelIndex(hitIndex)
                              setIsDraggingLabel(true)
                              setDragStart(point)
                              setDragOrigin({ ...selectedLabels[hitIndex] })
                            } else {
                              setSelectedLabelIndex(null)
                                  if (zoom > 1) {
                                    setIsPanning(true)
                                    setPanStart({ x: event.clientX, y: event.clientY })
                                    setPanOrigin({ ...pan })
                                  }
                            }
                          }
                          }}
                          />
                          {showLabels && selectedLabels.length > 0 && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div
                              className="label-overlay"
                              style={{ width: imageMetrics.width, height: imageMetrics.height }}
                            >
                          {selectedLabels.map((label, index) => {
                            const { width: renderW, height: renderH, naturalWidth, naturalHeight } = imageMetrics
                            const safeNaturalW = naturalWidth || renderW
                            const safeNaturalH = naturalHeight || renderH
                            const scale = Math.min(renderW / safeNaturalW, renderH / safeNaturalH)
                            const contentW = safeNaturalW * scale
                            const contentH = safeNaturalH * scale
                            const offsetX = (renderW - contentW) / 2
                            const offsetY = (renderH - contentH) / 2

                            const isNormalized =
                              label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
                            const centerX = isNormalized ? label.x * safeNaturalW : label.x
                            const centerY = isNormalized ? label.y * safeNaturalH : label.y
                            const boxW = isNormalized ? label.w * safeNaturalW : label.w
                            const boxH = isNormalized ? label.h * safeNaturalH : label.h

                            const left = offsetX + (centerX - boxW / 2) * scale
                            const top = offsetY + (centerY - boxH / 2) * scale
                            const width = boxW * scale
                            const height = boxH * scale

                            return (
                              <div
                                key={`${label.classId}-${index}`}
                                className={`label-box ${label.shape === 'ellipse' ? 'label-box-ellipse' : ''} ${selectedLabelIndex === index ? 'label-box-selected' : ''}`}
                                style={{
                                  left,
                                  top,
                                  width,
                                  height,
                                }}
                              />
                            )
                          })}
                            </div>
                          )}
                          {showLabels && drawMode === 'select' && selectedLabelIndex !== null && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div
                              className="label-handles"
                              style={{ width: imageMetrics.width, height: imageMetrics.height }}
                            >
                          {(() => {
                            const label = selectedLabels[selectedLabelIndex]
                            if (!label) return null

                            const { width: renderW, height: renderH, naturalWidth, naturalHeight } = imageMetrics
                            const safeNaturalW = naturalWidth || renderW
                            const safeNaturalH = naturalHeight || renderH

                            const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
                            const x = isNormalized ? label.x : label.x / safeNaturalW
                            const y = isNormalized ? label.y : label.y / safeNaturalH
                            const w = isNormalized ? label.w : label.w / safeNaturalW
                            const h = isNormalized ? label.h : label.h / safeNaturalH

                            const left = (x - w / 2) * renderW
                            const top = (y - h / 2) * renderH
                            const width = w * renderW
                            const height = h * renderH
                            const right = left + width
                            const bottom = top + height
                            const handleProps = (handle: 'nw' | 'ne' | 'sw' | 'se') => ({
                              onMouseDown: (event: React.MouseEvent) => {
                                event.preventDefault()
                                event.stopPropagation()
                                const point = toImageCoords(event)
                                if (!point) return
                                setIsResizingLabel(true)
                                setResizeHandle(handle)
                                setResizeOrigin({ ...label })
                              },
                            })
                            return (
                              <>
                                <div className="label-handle nw" style={{ left, top }} {...handleProps('nw')} />
                                <div className="label-handle ne" style={{ left: right, top }} {...handleProps('ne')} />
                                <div className="label-handle sw" style={{ left, top: bottom }} {...handleProps('sw')} />
                                <div className="label-handle se" style={{ left: right, top: bottom }} {...handleProps('se')} />
                              </>
                            )
                          })()}
                            </div>
                          )}
                          {showLabels && draftLabel && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div
                              className="manual-overlay"
                              style={{ width: imageMetrics.width, height: imageMetrics.height }}
                            >
                          {(() => {
                            const { width: renderW, height: renderH } = imageMetrics
                            const left = (draftLabel.x - draftLabel.w / 2) * renderW
                            const top = (draftLabel.y - draftLabel.h / 2) * renderH
                            const width = draftLabel.w * renderW
                            const height = draftLabel.h * renderH
                            return (
                              <div
                                className={`manual-label ${draftLabel.shape === 'ellipse' ? 'ellipse' : 'rect'} draft`}
                                style={{ left, top, width, height }}
                              />
                            )
                          })()}
                            </div>
                          )}
                        </div>
                        </div>

                        <button
                          type="button"
                          className="viewer-shell-nav viewer-shell-nav-right"
                          onClick={() => {
                            if (nextImagePath) void handleSelectFile(nextImagePath)
                          }}
                          disabled={!nextImagePath}
                          aria-label="Next image"
                          title="Next image"
                        >
                          ›
                        </button>
                      </div>
                        )
                      })()}
                      </div>
                    </div>
                  )}
                  {selectedPath && fileKind === 'text' && (
                    <pre className="preview-text">{fileText}</pre>
                  )}
                  {selectedPath && fileKind === 'other' && (
                    <p>Preview not supported for this file type.</p>
                  )}
                </div>
              )}
              {activeAction === 'scan' && (
                <div className="scan-panel">
                  <div className="scan-toolbar">
                    <button
                      className="link-button"
                      onClick={handleScanClick}
                      disabled={isScanning}
                    >
                      {isScanning ? 'Scanning…' : scanCompleted ? 'Restart scan' : 'Start scan'}
                    </button>
                    {scanCompleted && (
                      <button
                        className="link-button"
                        onClick={() => {
                          setActiveMenu('actions')
                          setActiveAction('view')
                        }}
                        disabled={isScanning}
                      >
                        View
                      </button>
                    )}
                  </div>
                  {scanPromptOpen && (
                    <div className="scan-prompt">
                      <p>There is folder with faults. Do you want to overwrite?</p>
                      <div className="scan-prompt-actions">
                        <button
                          className="link-button"
                          onClick={() => runScan(true)}
                          disabled={isScanning}
                        >
                          Yes, overwrite
                        </button>
                        <button
                          className="link-button"
                          onClick={() => setScanPromptOpen(false)}
                          disabled={isScanning}
                        >
                          No, cancel
                        </button>
                      </div>
                    </div>
                  )}
                  {scanStatus && <div className="scan-status">{scanStatus}</div>}
                  {scanError && <div className="scan-error">{scanError}</div>}
                  {isScanning && (
                    <div className="scan-progress">
                      Processing {scanProgress.current} / {scanProgress.total}
                    </div>
                  )}
                </div>
              )}
              {activeAction === 'report' && (
                <div className="report-panel">
                  <p style={{ marginTop: 0 }}>
                    Edit <strong>faults.txt</strong> (one image path per line). This does not delete images; it only updates the list.
                  </p>
                  <div className="report-toolbar">
                    <button className="link-button" onClick={() => loadReportFromDisk()}>
                      Reload
                    </button>
                    <button
                      className="link-button"
                      onClick={() => {
                        const normalized = normalizeFaultsText((docxDraft.length ? docxDraft.map((d) => d.path) : normalizeFaultsText(reportText)).join('\n'))
                        setReportText(normalized.join('\n'))
                        setDocxDraft((prev) => buildDocxDraftFromPaths(normalized, prev))
                      }}
                    >
                      Normalize
                    </button>
                    <button className="link-button" onClick={() => saveReportToDisk()}>
                      Save
                    </button>
                    <button className="link-button" onClick={() => syncDocxDraftFromEditor()}>
                      Sync DOCX draft
                    </button>
                    <button className="link-button" onClick={() => previewWordReport()}>
                      Preview DOCX
                    </button>
                    <button className="link-button" onClick={() => downloadWordReport()}>
                      Download DOCX
                    </button>
                    <button className="link-button" onClick={() => downloadPdfFromPreview()}>
                      Download PDF
                    </button>
                  </div>
                  <div ref={reportSplitRef} className="report-split">
                    <div className="report-left">
                      <div className="report-left-stack">
                        <div className="report-left-window">
                          <div className="report-sidebar-title">Description</div>
                          <div className="report-sidebar-subtitle">Add one or more description blocks (included in the DOCX)</div>

                          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
                            <button className="link-button" onClick={() => addDescriptionRow()}>
                              + Add description
                            </button>
                            <div className="docx-draft-meta">{descriptionRows.length} row(s)</div>
                          </div>

                          <div className="description-rows">
                            {descriptionRows.map((row, idx) => (
                              <div key={row.id} className="description-row">
                                <div className="description-row-header">
                                  <div className="description-row-title">{idx + 1}.</div>
                                  <button className="tiny-button danger" onClick={() => removeDescriptionRow(row.id)} title="Remove">
                                    ✕
                                  </button>
                                </div>

                                <label className="description-label">
                                  Title
                                  <input
                                    className="description-input"
                                    value={row.title}
                                    onChange={(e) => updateDescriptionRow(row.id, { title: e.target.value })}
                                    placeholder="e.g. Site / Project / Summary"
                                  />
                                </label>

                                <label className="description-label">
                                  Text
                                  <textarea
                                    className="description-textarea"
                                    value={row.text}
                                    onChange={(e) => updateDescriptionRow(row.id, { text: e.target.value })}
                                    placeholder="Write the description…"
                                    rows={5}
                                  />
                                </label>
                              </div>
                            ))}
                          </div>
                        </div>

                        <div className="report-left-window">
                          <div className="report-sidebar-title">Equipment</div>
                          <div className="report-sidebar-subtitle">Add equipment entries (title, description, optional image) (included in the DOCX)</div>

                          <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
                            <button className="link-button" onClick={() => addEquipmentItem()}>
                              + Add equipment
                            </button>
                            <div className="docx-draft-meta">{equipmentItems.length} item(s)</div>
                          </div>

                          <div className="equipment-rows">
                            {equipmentItems.map((item, idx) => (
                              <div key={item.id} className="equipment-row">
                                <div className="equipment-row-top">
                                  <input
                                    className="equipment-row-title"
                                    value={item.title}
                                    onChange={(e) => updateEquipmentItem(item.id, { title: e.target.value })}
                                    placeholder={`Equipment ${idx + 1} title`}
                                  />
                                  <button
                                    className="tiny-button danger"
                                    onClick={() => removeEquipmentItem(item.id)}
                                    title="Remove"
                                  >
                                    ✕
                                  </button>
                                </div>

                                <textarea
                                  className="equipment-row-text"
                                  value={item.text}
                                  onChange={(e) => updateEquipmentItem(item.id, { text: e.target.value })}
                                  placeholder="Equipment description…"
                                  rows={4}
                                />

                                <div className="equipment-row-image">
                                  <input
                                    className="equipment-row-file"
                                    type="file"
                                    accept="image/*"
                                    onChange={(e) => {
                                      const f = e.target.files?.[0] || null
                                      void setEquipmentItemImage(item.id, f)
                                    }}
                                  />
                                  {item.imagePreviewUrl ? (
                                    <div className="equipment-row-image-preview">
                                      <img src={item.imagePreviewUrl} alt="equipment" />
                                      <button
                                        className="equipment-row-clear-image"
                                        onClick={() => void setEquipmentItemImage(item.id, null)}
                                        title="Remove image"
                                      >
                                        Remove image
                                      </button>
                                    </div>
                                  ) : null}
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>

                    <div className="report-main">
                      <div className="docx-preview">
                        <div className="docx-preview-header">
                          <strong>DOCX preview</strong>
                          <span className="docx-preview-hint">(rendered in browser; final layout may differ slightly in Word)</span>
                        </div>
                        {docxPreviewStatus && <div className="report-status">{docxPreviewStatus}</div>}
                        {docxPreviewError && <div className="report-error">{docxPreviewError}</div>}
                        <div ref={docxPreviewStylesRef} style={{ display: 'none' }} />
                        <div ref={docxPreviewHostRef} className="docx-preview-host" />
                      </div>
                    </div>

                    <div className="report-splitter" onMouseDown={() => setIsReportDraftResizing(true)} />

                    <div className="report-sidebar" style={{ width: reportDraftWidth }}>
                      <div className="report-sidebar-window">
                        <div className="report-sidebar-title">DOCX draft</div>
                        <div className="report-sidebar-subtitle">Edit captions/fields/order before export</div>
                        <div className="docx-draft-meta" style={{ marginBottom: 10 }}>
                          {docxDraft.length ? `${docxDraft.filter((d) => d.include).length}/${docxDraft.length} included` : 'No draft yet'}
                        </div>

                        {docxDraft.length === 0 ? (
                          <div className="docx-draft-empty">
                            Click <strong>Sync DOCX draft</strong> to create an editable draft.
                          </div>
                        ) : (
                          <div className="docx-draft-list">
                            {docxDraft.map((item, idx) => (
                              <div key={item.id} className="docx-draft-item">
                                <div className="docx-draft-row">
                                  <label className="docx-draft-include" title="Include in DOCX export">
                                    <input
                                      type="checkbox"
                                      checked={item.include}
                                      onChange={(e) => updateDocxDraftItem(item.id, { include: e.target.checked })}
                                    />
                                    Include
                                  </label>
                                  <div className="docx-draft-path" title={item.path}>
                                    {idx + 1}. {item.path}
                                  </div>
                                  <div className="docx-draft-actions">
                                    <button className="tiny-button" onClick={() => moveDocxDraftItem(idx, -1)} disabled={idx === 0}>
                                      ↑
                                    </button>
                                    <button
                                      className="tiny-button"
                                      onClick={() => moveDocxDraftItem(idx, 1)}
                                      disabled={idx === docxDraft.length - 1}
                                    >
                                      ↓
                                    </button>
                                    <button className="tiny-button danger" onClick={() => removeDocxDraftItem(item.id)}>
                                      Remove
                                    </button>
                                  </div>
                                </div>
                                <div className="docx-draft-fields">
                                  <label className="docx-field">
                                    <div className="docx-field-label">Caption</div>
                                    <input
                                      className="docx-field-input"
                                      value={item.caption}
                                      onChange={(e) => updateDocxDraftItem(item.id, { caption: e.target.value })}
                                      placeholder="Caption shown above the image"
                                    />
                                  </label>
                                  <div className="docx-field">
                                    <div className="docx-field-label">Tables under image</div>
                                    <div className="docx-kv-editor">
                                      {(item.tables || []).map((table, tableIdx) => (
                                        <div key={table.id} className="docx-kv-table">
                                          <div className="docx-kv-table-header">
                                            <div className="docx-kv-table-title">
                                              <div className="docx-kv-table-title-label">Table title</div>
                                              <input
                                                className="docx-field-input"
                                                value={table.title}
                                                onChange={(e) => updateDocxDraftTable(item.id, table.id, { title: e.target.value })}
                                                placeholder={`Table ${tableIdx + 1} title (optional)`}
                                              />
                                            </div>
                                            <button
                                              className="tiny-button danger"
                                              onClick={() => removeDocxDraftTable(item.id, table.id)}
                                              type="button"
                                              title="Remove table"
                                            >
                                              Remove table
                                            </button>
                                          </div>

                                          {(table.rows || []).map((row) => (
                                            <div key={row.id} className="docx-kv-row">
                                              <input
                                                className="docx-field-input docx-kv-input"
                                                value={row.name}
                                                onChange={(e) => updateDocxDraftTableRow(item.id, table.id, row.id, { name: e.target.value })}
                                                placeholder="Field name"
                                              />
                                              <input
                                                className="docx-field-input docx-kv-input"
                                                value={row.description}
                                                onChange={(e) => updateDocxDraftTableRow(item.id, table.id, row.id, { description: e.target.value })}
                                                placeholder="Description"
                                              />
                                              <button
                                                className="tiny-button danger"
                                                onClick={() => removeDocxDraftTableRow(item.id, table.id, row.id)}
                                                title="Remove row"
                                                type="button"
                                              >
                                                ✕
                                              </button>
                                            </div>
                                          ))}

                                          <div className="docx-kv-actions">
                                            <button className="tiny-button" onClick={() => addDocxDraftTableRow(item.id, table.id)} type="button">
                                              + Add row
                                            </button>
                                          </div>
                                        </div>
                                      ))}

                                      <div className="docx-kv-actions docx-kv-actions-global">
                                        <button className="tiny-button" onClick={() => addDocxDraftTable(item.id)} type="button">
                                          + Add table
                                        </button>
                                      </div>
                                    </div>
                                  </div>
                                  <label className="docx-draft-break">
                                    <input
                                      type="checkbox"
                                      checked={item.pageBreakBefore}
                                      onChange={(e) => updateDocxDraftItem(item.id, { pageBreakBefore: e.target.checked })}
                                    />
                                    Page break before
                                  </label>
                                </div>
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                  {reportStatus && <div className="report-status">{reportStatus}</div>}
                  {reportError && <div className="report-error">{reportError}</div>}
                </div>
              )}
              {activeAction !== 'view' && activeAction !== 'scan' && activeAction !== 'report' && (
                <div className="empty-page">(empty)</div>
              )}
            </section>
          )}

          {activeMenu !== 'file' && activeMenu !== 'actions' && (
            <section className="card">
              <h2>Status</h2>
              <p>{message}</p>
            </section>
          )}

          

          {activeMenu !== 'actions' && activeMenu !== 'file' && (
            <section className="card viewer">
              <h3>Preview</h3>
              {!selectedPath && <p>Select a file from the explorer.</p>}
              {selectedPath && fileKind === 'image' && fileUrl && (
                <img src={fileUrl} alt={selectedPath} className="preview-image" />
              )}
              {selectedPath && fileKind === 'text' && (
                <pre className="preview-text">{fileText}</pre>
              )}
              {selectedPath && fileKind === 'other' && (
                <p>Preview not supported for this file type.</p>
              )}
            </section>
          )}
        </main>

        {showRightExplorer && (
          <>
            <aside className={`explorer explorer-right ${isRightExplorerMinimized ? 'minimized' : ''}`}>
              <div className="explorer-header">
                {/* <div className="explorer-title">Explorer</div> */}
                <button
                  className="explorer-toggle"
                  onClick={() => setIsRightExplorerMinimized((prev) => !prev)}
                  title={isRightExplorerMinimized ? 'Restore explorer' : 'Minimize explorer'}
                >
                  {isRightExplorerMinimized ? '▸' : '▾'}
                </button>
              </div>
              {!isRightExplorerMinimized && (
                <>
                  <div className="explorer-pane">
                    <div className="explorer-pane-title">Faults</div>
                    <div className="tree">
                      {selectedLabels.length === 0 && (
                        <div className="explorer-empty">No faults for this image.</div>
                      )}
                      {selectedLabels.map((label, index) => (
                        <button
                          key={`${label.classId}-${index}`}
                          className={`tree-item file ${selectedLabelIndex === index ? 'active' : ''}`}
                          style={{ paddingLeft: 2 }}
                          onClick={() =>
                            setSelectedLabelIndex((prev) => (prev === index ? null : index))
                          }
                        >
                          <span className="file-label">
                            <span className="file-icon">⚡</span>
                            <span className="file-name">
                              {label.source === 'manual' ? `Manual ${index + 1}` : `Fault ${index + 1}`}
                              {label.source !== 'manual' && Number.isFinite(label.conf)
                                ? ` — ${Math.round(label.conf * 100)}%`
                                : ''}
                            </span>
                          </span>
                        </button>
                      ))}
                    </div>
                  </div>
                  <div className="explorer-pane">
                    {/* <div className="explorer-pane-title">Temp title D</div> */}
                    <div className="tree">
                      <div className="tool-group">
                        <button
                          className={`tool-button ${drawMode === 'select' ? 'active' : ''}`}
                          onClick={() => setDrawMode('select')}
                          disabled={!selectedPath || fileKind !== 'image'}
                        >
                          Select / Edit
                        </button>
                        <button
                          className={`tool-button ${drawMode === 'rect' ? 'active' : ''}`}
                          onClick={() => setDrawMode('rect')}
                          disabled={!selectedPath || fileKind !== 'image'}
                        >
                          Draw Rectangle
                        </button>
                        <button
                          className={`tool-button ${drawMode === 'ellipse' ? 'active' : ''}`}
                          onClick={() => setDrawMode('ellipse')}
                          disabled={!selectedPath || fileKind !== 'image'}
                        >
                          Draw Circle
                        </button>
                        <button
                          className="tool-button danger"
                          onClick={() => {
                            if (selectedLabelIndex === null) return
                            const next = selectedLabels.filter((_, idx) => idx !== selectedLabelIndex)
                            setSelectedLabelIndex(null)
                            setSelectedLabels(next)
                            pushLabelHistory(next)
                            persistLabels(next)
                          }}
                          disabled={selectedLabelIndex === null}
                        >
                          Delete Selected
                        </button>
                        <div className="tool-row">
                          <button
                            className="tool-button"
                            onClick={() => applyHistoryIndex(labelHistoryIndex - 1)}
                            disabled={labelHistoryIndex <= 0}
                          >
                            Undo
                          </button>
                          <button
                            className="tool-button"
                            onClick={() => applyHistoryIndex(labelHistoryIndex + 1)}
                            disabled={labelHistoryIndex >= labelHistory.length - 1}
                          >
                            Redo
                          </button>
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="explorer-resizer explorer-resizer-right" onMouseDown={() => setIsRightResizing(true)} />
                </>
              )}
            </aside>
            {isRightExplorerMinimized && (
              <button
                className="explorer-toggle-floating explorer-toggle-floating-right"
                onClick={() => setIsRightExplorerMinimized(false)}
                title="Restore explorer"
              >
                ▸
              </button>
            )}
          </>
        )}
      </div>

      {explorerContextMenu && (
        <div
          className="context-menu"
          style={{ top: explorerContextMenu.y, left: explorerContextMenu.x }}
          onClick={(event) => event.stopPropagation()}
          onContextMenu={(event) => event.preventDefault()}
        >
          <div className="context-menu-title">{explorerContextMenu.node.name}</div>
          <button
            className="context-menu-item danger"
            onClick={() => {
              const { node, source } = explorerContextMenu
              setExplorerContextMenu(null)
              if (source === 'faultsList') {
                void removeFaultFromList(node.path)
                return
              }
              void deleteExplorerNode(node)
            }}
          >
            {explorerContextMenu.source === 'faultsList' ? 'Remove from list' : 'Delete'}
          </button>
        </div>
      )}
    </div>
  )
}

export default App
