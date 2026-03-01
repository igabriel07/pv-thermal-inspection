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
  ShadingType,
  TextWrappingType,
  TextRun,
  VerticalAlignTable,
  VerticalPositionRelativeFrom,
  WidthType,
  convertInchesToTwip,
} from 'docx'
import './App.css'

type DraftTextInputProps = {
  value: string
  className?: string
  placeholder?: string
  onCommit: (value: string) => void
}

const DraftTextInput = ({ value, className, placeholder, onCommit }: DraftTextInputProps) => {
  const [draft, setDraft] = useState<string>(value)
  const isFocusedRef = useRef(false)
  const lastPropRef = useRef(value)

  useEffect(() => {
    // Keep the input in sync with external state changes when not actively editing.
    if (isFocusedRef.current) return
    if (value === lastPropRef.current) return
    lastPropRef.current = value
    setDraft(value)
  }, [value])

  const commit = () => {
    const next = draft
    if (next === value) return
    onCommit(next)
  }

  return (
    <input
      className={className}
      value={draft}
      placeholder={placeholder}
      onFocus={() => {
        isFocusedRef.current = true
      }}
      onBlur={() => {
        isFocusedRef.current = false
        commit()
      }}
      onKeyDown={(e) => {
        if (e.key === 'Enter') {
          ;(e.currentTarget as HTMLInputElement).blur()
        }
      }}
      onChange={(e) => setDraft(e.target.value)}
    />
  )
}

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

type ReportCustomChapterSection = {
  id: string
  title: string // optional
  text: string // optional
  imageFile: File | null // optional
  imagePreviewUrl: string | null
}

type ReportCustomChapter = {
  id: string
  chapterTitle: string // required
  sections: ReportCustomChapterSection[]
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
// When viewing an RGB image, render it slightly smaller for better framing.
const RGB_VIEW_IMAGE_SCALE = 0.98
const VIEWER_MAX_W = 1536
const VIEWER_MAX_H = 1229

const FAULT_TYPE_OPTIONS: Array<{ id: number; label: string }> = [
  { id: 0, label: 'Multi ByPassed' },
  { id: 1, label: 'Multi Diode' },
  { id: 2, label: 'Multi HotSpot' },
  { id: 3, label: 'Single ByPassed' },
  { id: 4, label: 'Single Diode' },
  { id: 5, label: 'Single HotSpot' },
  { id: 6, label: 'String Open Circuit' },
  { id: 7, label: 'String Reversed Polarity' },
  { id: 9, label: 'Unknown' },
]

const normalizeFaultTypeId = (value: unknown): number => {
  const n = typeof value === 'number' ? value : Number(value)
  if (!Number.isFinite(n)) return 9
  const id = Math.trunc(n)
  if (FAULT_TYPE_OPTIONS.some((o) => o.id === id)) return id
  return 9
}

const getFaultTypeLabel = (value: unknown): string => {
  const id = normalizeFaultTypeId(value)
  return FAULT_TYPE_OPTIONS.find((o) => o.id === id)?.label ?? 'Unknown'
}

const formatPossibleFaultType = (value: unknown): string => `Possible ${getFaultTypeLabel(value)}`

// Bucket sizes for pitch/altitude-conditioned RGB alignment.
// These are used to select a stored scale override (learned via Alt+wheel)
// without introducing any new UI.
const RGB_ALIGN_PITCH_BUCKET_DEG = 5
const RGB_ALIGN_ALT_BUCKET_M = 5

const bucketRound = (value: number, step: number) => {
  if (!Number.isFinite(value) || !Number.isFinite(step) || step <= 0) return value
  return Math.round(value / step) * step
}

const makeRgbAlignScaleContextKey = (gimbalPitchDeg: number | null, relativeAltM: number | null) => {
  if (gimbalPitchDeg === null || relativeAltM === null) return null
  const p = bucketRound(gimbalPitchDeg, RGB_ALIGN_PITCH_BUCKET_DEG)
  const a = bucketRound(relativeAltM, RGB_ALIGN_ALT_BUCKET_M)
  if (!Number.isFinite(p) || !Number.isFinite(a)) return null
  return `p${p.toFixed(0)}_a${a.toFixed(0)}`
}

const getRgbAlignScaleStorageKey = (contextKey: string | null) => {
  if (!contextKey) return 'rgb-align-scale-v1'
  return `rgb-align-scale-v1:${contextKey}`
}

const parseStoredScale = (raw: string | null) => {
  const n = raw ? Number(raw) : NaN
  return Number.isFinite(n) ? Math.min(2, Math.max(0.5, n)) : null
}

// Yaw-based pixel tweaks for RGB label placement within the locked viewer frame.
// Negative X moves left; positive Y moves down.
//
// IMPORTANT:
// - `FlightYawDegree` is normalized to [0, 360) before matching.
// - The -90° “anchor” is intentionally at -92.30 (i.e. 267.70°).
// - Defaults are set to the current calibrated offset so behavior does not
//   regress until other anchors are calibrated.
const DEFAULT_RGB_YAW_OFFSET_PX = { x: -15, y: 5 }
// If the flight yaw is within this many degrees of an anchor, use the anchor
// offset exactly (no interpolation).
const RGB_YAW_SNAP_DEG = 10
const RGB_YAW_ANCHORS = [
  { angleDeg: 0, offsetPx: { ...DEFAULT_RGB_YAW_OFFSET_PX } },
  // +90° bucket: move labels 12px up and 10px left.
  { angleDeg: 90, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 10, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 12 } },
  // +90.80° bucket: move labels 5px further up (relative to the +90° bucket).
  { angleDeg: 90.8, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 10, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 20 } },
  // +93.70° bucket: move labels 2px down (relative to the +90° bucket).
  { angleDeg: 93.7, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 10, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 13 } },
  // +133.40° bucket: move labels 22px up and 5px right.
  { angleDeg: 133.4, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x + 5, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 22 } },
  // +167.20° bucket: move labels 15px down, 10px left (relative to the nearby +174.40° behavior).
  { angleDeg: 167.2, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 1, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 10 } },
  // +174.40° bucket: move labels 9px right and 20px up.
  { angleDeg: 174.4, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x + 9, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 25 } },
  // -180° bucket: for FlightYawDegree ≈ -177, move labels 35px up.
  { angleDeg: 180, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 35 } },
  // -90.10° bucket (normalized to 269.9°): move labels 14px right and 4px down (relative to the -90° bucket).
  { angleDeg: 269.9, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x + 2, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 3 } },
  // Some datasets may report yaw as 270° (equivalent to -90°).
  { angleDeg: 270, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 12, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 7 } },
  // -89.90° bucket (normalized to 270.1°): move labels 5px right and 2px down (relative to the -90° bucket).
  { angleDeg: 270.1, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 7, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 5 } },
  // -98.5° bucket (normalized to 261.5°): move labels 20px up and 5px left.
  { angleDeg: 261.5, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 5, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 20 } },
  // -93.9° bucket (normalized to 266.1°): move labels 35px up.
  { angleDeg: 266.1, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 35 } },
  // -92.30° normalized into [0, 360) => 267.70°
  { angleDeg: 267.7, offsetPx: { ...DEFAULT_RGB_YAW_OFFSET_PX } },
] as const

const normAngle360 = (deg: number) => {
  const n = deg % 360
  return n < 0 ? n + 360 : n
}

const angleDistanceDeg = (aDeg: number, bDeg: number) => {
  const a = normAngle360(aDeg)
  const b = normAngle360(bDeg)
  const diff = Math.abs(a - b)
  return Math.min(diff, 360 - diff)
}

const toFiniteNumberOrNull = (value: unknown): number | null => {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return null
    const parsed = Number(trimmed)
    return Number.isFinite(parsed) ? parsed : null
  }
  return null
}

const extractFlightYawDegree = (report: any | null): number | null => {
  if (!report) return null

  // Preferred structured location from backend/check_for_metada_tiff.py
  const v1 = toFiniteNumberOrNull(report?.categories?.flight?.FlightYawDegree)
  if (v1 !== null) return v1

  // Backup locations (summary also contains DJI_FlightYawDegree)
  const v2 = toFiniteNumberOrNull(report?.summary?.DJI_FlightYawDegree)
  if (v2 !== null) return v2
  const v3 = toFiniteNumberOrNull(report?.summary?.FlightYawDegree)
  if (v3 !== null) return v3

  // Some exiftool dumps may expose flattened keys.
  const v4 = toFiniteNumberOrNull(report?.['XMP-drone-dji:FlightYawDegree'])
  if (v4 !== null) return v4
  const v5 = toFiniteNumberOrNull(report?.DJI_FlightYawDegree)
  if (v5 !== null) return v5

  return null
}

const extractGimbalYawDegree = (report: any | null): number | null => {
  if (!report) return null

  // Preferred structured location from backend/check_for_metada_tiff.py
  const v1 = toFiniteNumberOrNull(report?.categories?.flight?.GimbalYawDegree)
  if (v1 !== null) return v1

  // Backup locations (summary also contains DJI_GimbalYawDegree)
  const v2 = toFiniteNumberOrNull(report?.summary?.DJI_GimbalYawDegree)
  if (v2 !== null) return v2
  const v3 = toFiniteNumberOrNull(report?.summary?.GimbalYawDegree)
  if (v3 !== null) return v3

  // Some exiftool dumps may expose flattened keys.
  const v4 = toFiniteNumberOrNull(report?.['XMP-drone-dji:GimbalYawDegree'])
  if (v4 !== null) return v4
  const v5 = toFiniteNumberOrNull(report?.DJI_GimbalYawDegree)
  if (v5 !== null) return v5

  return null
}

const extractGimbalPitchDegree = (report: any | null): number | null => {
  if (!report) return null

  const v1 = toFiniteNumberOrNull(report?.categories?.flight?.GimbalPitchDegree)
  if (v1 !== null) return v1

  const v2 = toFiniteNumberOrNull(report?.summary?.DJI_GimbalPitchDegree)
  if (v2 !== null) return v2
  const v3 = toFiniteNumberOrNull(report?.summary?.GimbalPitchDegree)
  if (v3 !== null) return v3

  const v4 = toFiniteNumberOrNull(report?.['XMP-drone-dji:GimbalPitchDegree'])
  if (v4 !== null) return v4
  const v5 = toFiniteNumberOrNull(report?.DJI_GimbalPitchDegree)
  if (v5 !== null) return v5

  return null
}

const extractRelativeAltitude = (report: any | null): number | null => {
  if (!report) return null

  const v1 = toFiniteNumberOrNull(report?.categories?.flight?.RelativeAltitude)
  if (v1 !== null) return v1

  const v2 = toFiniteNumberOrNull(report?.summary?.DJI_RelativeAltitude)
  if (v2 !== null) return v2
  const v3 = toFiniteNumberOrNull(report?.summary?.RelativeAltitude)
  if (v3 !== null) return v3

  const v4 = toFiniteNumberOrNull(report?.['XMP-drone-dji:RelativeAltitude'])
  if (v4 !== null) return v4
  const v5 = toFiniteNumberOrNull(report?.DJI_RelativeAltitude)
  if (v5 !== null) return v5

  return null
}

// Select a yaw angle for alignment based on similarities/differences between
// flight yaw and gimbal yaw. In practice, gimbal yaw usually correlates better
// with the camera view direction when they diverge.
const extractYawForAlignment = (report: any | null): number | null => {
  const flightYaw = extractFlightYawDegree(report)
  const gimbalYaw = extractGimbalYawDegree(report)

  if (flightYaw === null && gimbalYaw === null) return null
  if (flightYaw === null) return gimbalYaw
  if (gimbalYaw === null) return flightYaw

  const a = normAngle360(flightYaw)
  const b = normAngle360(gimbalYaw)
  const diff = Math.min(Math.abs(a - b), 360 - Math.abs(a - b))

  // If they broadly agree, use the circular mean for stability.
  if (diff <= 15) {
    const ra = (a * Math.PI) / 180
    const rb = (b * Math.PI) / 180
    const x = Math.cos(ra) + Math.cos(rb)
    const y = Math.sin(ra) + Math.sin(rb)
    const mean = Math.atan2(y, x) * (180 / Math.PI)
    return normAngle360(mean)
  }

  // If they disagree a lot, prefer gimbal yaw (camera pointing direction).
  return gimbalYaw
}

// Optional per-(flightYaw,gimbalYaw) correction rules.
// Use this when small flight/gimbal differences correlate with a consistent
// pixel shift that yaw-only interpolation can’t capture.
const RGB_FLIGHT_GIMBAL_YAW_ADJUSTMENTS: Array<{
  flightYawDeg: number
  gimbalYawDeg: number
  tolDeg: number
  deltaPx: { x: number; y: number }
}> = [
  // FlightYaw=-90.00 and GimbalYaw=-91.30 => move 5px left and 5px up.
  { flightYawDeg: 270, gimbalYawDeg: 268.7, tolDeg: 0.6, deltaPx: { x: -5, y: -5 } },
  // FlightYaw=-177.00 (normalized to 183.00) and GimbalYaw=+9.90 => move 40px up.
  { flightYawDeg: 183, gimbalYawDeg: 9.9, tolDeg: 0.6, deltaPx: { x: 0, y: -40 } },
  // FlightYaw=+90.80 and GimbalYaw=-90.80 (normalized to 269.20) => move 20px left and 30px up.
  { flightYawDeg: 90.8, gimbalYawDeg: 269.2, tolDeg: 0.6, deltaPx: { x: -20, y: -30 } },
  // FlightYaw=-98.50 (normalized to 261.50) and GimbalYaw=+85.20 => move 20px up.
  { flightYawDeg: 261.5, gimbalYawDeg: 85.2, tolDeg: 0.6, deltaPx: { x: 0, y: -20 } },
  // FlightYaw=-90.10 (normalized to 269.90) and GimbalYaw=-89.80 (normalized to 270.20) => move 20px right and 3px down.
  { flightYawDeg: 269.9, gimbalYawDeg: 270.2, tolDeg: 0.6, deltaPx: { x: 20, y: 3 } },
  // FlightYaw=-90.00 (normalized to 270.00) and GimbalYaw=-89.40 (normalized to 270.60) => move 20px left.
  { flightYawDeg: 270, gimbalYawDeg: 270.6, tolDeg: 0.6, deltaPx: { x: -20, y: 0 } },
]

const getFlightGimbalYawAdjustmentPx = (flightYaw: number | null, gimbalYaw: number | null) => {
  if (flightYaw === null || gimbalYaw === null) return { x: 0, y: 0 }
  const f = normAngle360(flightYaw)
  const g = normAngle360(gimbalYaw)

  for (const rule of RGB_FLIGHT_GIMBAL_YAW_ADJUSTMENTS) {
    if (angleDistanceDeg(f, rule.flightYawDeg) > rule.tolDeg) continue
    if (angleDistanceDeg(g, rule.gimbalYawDeg) > rule.tolDeg) continue
    return { ...rule.deltaPx }
  }

  return { x: 0, y: 0 }
}

const getYawInterpolatedOffsetPx = (flightYawDegree: number | null) => {
  if (flightYawDegree === null) return { ...DEFAULT_RGB_YAW_OFFSET_PX }

  const yaw = normAngle360(flightYawDegree)
  const anchors = [...RGB_YAW_ANCHORS].sort((a, b) => a.angleDeg - b.angleDeg)
  if (anchors.length === 0) return { ...DEFAULT_RGB_YAW_OFFSET_PX }
  if (anchors.length === 1) return { ...anchors[0].offsetPx }

  // Snap to the nearest anchor when close enough.
  const angleDist = (aDeg: number, bDeg: number) => {
    const diff = Math.abs(normAngle360(aDeg) - normAngle360(bDeg))
    return Math.min(diff, 360 - diff)
  }
  let nearest: (typeof anchors)[number] | null = null
  let nearestDist = Number.POSITIVE_INFINITY
  for (const a of anchors) {
    const d = angleDist(yaw, a.angleDeg)
    if (d < nearestDist) {
      nearestDist = d
      nearest = a
    }
  }
  if (nearest && nearestDist <= RGB_YAW_SNAP_DEG) return { ...nearest.offsetPx }

  // Find the clockwise segment [a, b) that contains yaw (with wrap support).
  for (let i = 0; i < anchors.length; i++) {
    const a = anchors[i]
    const b = anchors[(i + 1) % anchors.length]
    const aDeg = a.angleDeg
    const bDeg = b.angleDeg

    if (i < anchors.length - 1) {
      if (yaw >= aDeg && yaw < bDeg) {
        const t = (yaw - aDeg) / (bDeg - aDeg)
        return {
          x: a.offsetPx.x + (b.offsetPx.x - a.offsetPx.x) * t,
          y: a.offsetPx.y + (b.offsetPx.y - a.offsetPx.y) * t,
        }
      }
      continue
    }

    // Wrap segment: last -> first
    const bWrap = bDeg + 360
    const yawWrap = yaw < aDeg ? yaw + 360 : yaw
    if (yawWrap >= aDeg && yawWrap < bWrap) {
      const t = (yawWrap - aDeg) / (bWrap - aDeg)
      return {
        x: a.offsetPx.x + (b.offsetPx.x - a.offsetPx.x) * t,
        y: a.offsetPx.y + (b.offsetPx.y - a.offsetPx.y) * t,
      }
    }
  }

  // Fallback (shouldn't happen)
  return { ...DEFAULT_RGB_YAW_OFFSET_PX }
}

const CENTER_BOX_W = 250
const CENTER_BOX_H = 250
const CENTER_FLAGS_FILE = 'center_flags.json'
const STARRED_FILE = 'starred.txt'

type CenterFlagsDiskV2 = {
  meta: { version: 2; boxW: number; boxH: number }
  flags: Record<string, 0 | 1>
}

const CURRENT_CENTER_FLAGS_META: CenterFlagsDiskV2['meta'] = {
  version: 2,
  boxW: CENTER_BOX_W,
  boxH: CENTER_BOX_H,
}

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
  const [activeHelpPage, setActiveHelpPage] = useState<'documentation' | 'about'>('documentation')
  const [helpLang, setHelpLang] = useState<'en' | 'el'>('en')
  const [backendHealth, setBackendHealth] = useState<{ status?: string; message?: string; versions?: Record<string, string> } | null>(null)
  const [backendHealthError, setBackendHealthError] = useState<string>('')
  const [fileTree, setFileTree] = useState<TreeNode | null>(null)
  const [fileMap, setFileMap] = useState<Record<string, File>>({})
  const [workspaceIndexTick, setWorkspaceIndexTick] = useState(0)
  // `selectedPath` is the *context* path (thermal) used for labels/metadata.
  const [selectedPath, setSelectedPath] = useState<string>('')
  // `selectedViewPath` is the *display* path (thermal or rgb) shown in the View.
  const [selectedViewPath, setSelectedViewPath] = useState<string>('')
  const [viewVariant, setViewVariant] = useState<'thermal' | 'rgb'>('thermal')
  const [contextImageNatural, setContextImageNatural] = useState<{ width: number; height: number } | null>(null)
  const [rgbAlignOffset, setRgbAlignOffset] = useState<{ x: number; y: number }>(() => {
    try {
      const raw = window.localStorage.getItem('rgb-align-offset-v1')
      if (!raw) return { x: 0, y: 0 }
      const parsed = JSON.parse(raw) as { x?: unknown; y?: unknown }
      const x = typeof parsed.x === 'number' && Number.isFinite(parsed.x) ? parsed.x : 0
      const y = typeof parsed.y === 'number' && Number.isFinite(parsed.y) ? parsed.y : 0
      return { x, y }
    } catch {
      return { x: 0, y: 0 }
    }
  })
  const [rgbAlignScale, setRgbAlignScale] = useState<number>(() => {
    try {
      const raw = window.localStorage.getItem('rgb-align-scale-v1')
      const n = raw ? Number(raw) : 1
      return Number.isFinite(n) ? n : 1
    } catch {
      return 1
    }
  })
  const [rgbAlignScaleContextTick, setRgbAlignScaleContextTick] = useState(0)
  const [isRgbAlignDragging, setIsRgbAlignDragging] = useState(false)
  const [rgbAlignDragStart, setRgbAlignDragStart] = useState<{ x: number; y: number } | null>(null)
  const [rgbAlignDragOrigin, setRgbAlignDragOrigin] = useState<{ x: number; y: number } | null>(null)
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
  const [rgbLabelsExportDirty, setRgbLabelsExportDirty] = useState(false)
  const [rgbLabelsExportFileExists, setRgbLabelsExportFileExists] = useState(false)
  const [rgbLabelsExportFileEmpty, setRgbLabelsExportFileEmpty] = useState(false)
  const [rgbLabelsExportStatus, setRgbLabelsExportStatus] = useState<'idle' | 'checking' | 'exists' | 'missing'>('idle')
  const [rgbLabelsExportSaving, setRgbLabelsExportSaving] = useState(false)
  const [rgbLabelsAreImageSpace, setRgbLabelsAreImageSpace] = useState(false)
  const [rgbPendingCropFileLabels, setRgbPendingCropFileLabels] = useState<
    Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }> | null
  >(null)
  const [rgbLabelsDirEnsuredOnOpen, setRgbLabelsDirEnsuredOnOpen] = useState(false)
  const [scanPromptOpen, setScanPromptOpen] = useState(false)
  const [scanProgress, setScanProgress] = useState<{ current: number; total: number }>({ current: 0, total: 0 })
  const [scanCompleted, setScanCompleted] = useState(false)
  const [faultsList, setFaultsList] = useState<string[]>([])
  const [centerFaultFlags, setCenterFaultFlags] = useState<Record<string, 0 | 1>>({})
  const [onlyCenterBoxFaults, setOnlyCenterBoxFaults] = useState(false)
  const [starredFaults, setStarredFaults] = useState<Record<string, true>>({})
  const [onlyStarredInReport, setOnlyStarredInReport] = useState(false)
  const [reportAdvancedOpen, setReportAdvancedOpen] = useState(false)
  const [reportText, setReportText] = useState('')
  const [reportStatus, setReportStatus] = useState<string>('')
  const [reportError, setReportError] = useState<string>('')
  const [docxDraft, setDocxDraft] = useState<DocxDraftItem[]>([])
  const [includeRgbInDocx, setIncludeRgbInDocx] = useState(true)
  const [docxPreviewStatus, setDocxPreviewStatus] = useState<string>('')
  const [docxPreviewError, setDocxPreviewError] = useState<string>('')
  const [descriptionRows, setDescriptionRows] = useState<ReportDescriptionRow[]>([
    { id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, title: '', text: '' },
  ])
  const [equipmentItems, setEquipmentItems] = useState<ReportEquipmentItem[]>([
    { id: `${Date.now()}-${Math.random().toString(16).slice(2)}` , title: '', text: '', imageFile: null, imagePreviewUrl: null },
  ])

  type ReportLeftTabId = 'description' | 'equipment' | `chapter:${string}`
  const [reportLeftTab, setReportLeftTab] = useState<ReportLeftTabId>('description')
  const [reportChapters, setReportChapters] = useState<ReportCustomChapter[]>([])
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
  const [rightExplorerWidth, setRightExplorerWidth] = useState<number>(310)
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
  const docxPreviewPdfUrlRef = useRef<string | null>(null)
  const docxPreviewPdfBlobRef = useRef<Blob | null>(null)
  const docxPreviewLastErrorRef = useRef<string>('')
  const reportSplitRef = useRef<HTMLDivElement | null>(null)
  const selectedLabelsRef = useRef(selectedLabels)
  const workspaceIndexSignatureRef = useRef<string>('')
  const workspacePollInFlightRef = useRef<Promise<void> | null>(null)

  useEffect(() => {
    return () => {
      const url = docxPreviewPdfUrlRef.current
      if (url) URL.revokeObjectURL(url)
      docxPreviewPdfUrlRef.current = null
      docxPreviewPdfBlobRef.current = null
    }
  }, [])
  type RgbWorkingLabelsCache = {
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
    areImageSpace: boolean
  }
  const rgbWorkingLabelsByThermalPathRef = useRef<
    Record<string, RgbWorkingLabelsCache>
  >({})
  const closeTimerRef = useRef<number | null>(null)
  const realtimePersistTimerRef = useRef<number | null>(null)
  const realtimePersistLabelsRef = useRef<
    Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }> | null
  >(null)
  const menuAreaRef = useRef<HTMLDivElement | null>(null)
  const [explorerContextMenu, setExplorerContextMenu] = useState<null | { x: number; y: number; node: TreeNode; source: 'tree' | 'faultsList' }>(null)
  const centerFlagsEnsureInFlightRef = useRef<Promise<Record<string, 0 | 1>> | null>(null)
  const reportResyncTimerRef = useRef<number | null>(null)
  const reportDefaultStarredAppliedRef = useRef(false)
  const reportResyncPendingRef = useRef<
    | null
    | {
        flagsOverride?: Record<string, 0 | 1>
        onlyCenterOverride?: boolean
        onlyStarredOverride?: boolean
        starredOverride?: Record<string, true>
      }
  >(null)
  const lastReportPrintablePathsRef = useRef<string[] | null>(null)

  const [viewMetadata, setViewMetadata] = useState<any | null>(null)
  const [viewMetadataStatus, setViewMetadataStatus] = useState<'idle' | 'loading' | 'ready' | 'error'>('idle')
  const [viewMetadataError, setViewMetadataError] = useState<string>('')
  const viewMetadataCacheRef = useRef<Map<string, any>>(new Map())
  const [viewMetadataSelections, setViewMetadataSelections] = useState<Record<string, Record<string, boolean>>>({})
  const [viewMetadataOverrides, setViewMetadataOverrides] = useState<Record<string, Record<string, any>>>({})
  const [viewMetadataEditing, setViewMetadataEditing] = useState<null | { tiffKey: string; id: string; draft: string }>(null)

  const viewFlightYawDegree = useMemo(() => extractFlightYawDegree(viewMetadata), [viewMetadata])
  const viewGimbalYawDegree = useMemo(() => extractGimbalYawDegree(viewMetadata), [viewMetadata])
  const viewGimbalPitchDegree = useMemo(() => extractGimbalPitchDegree(viewMetadata), [viewMetadata])
  const viewRelativeAltitudeM = useMemo(() => extractRelativeAltitude(viewMetadata), [viewMetadata])
  const viewAlignYawDegree = useMemo(() => extractYawForAlignment(viewMetadata), [viewMetadata])

  const rgbAlignScaleContextKey = useMemo(
    () => makeRgbAlignScaleContextKey(viewGimbalPitchDegree, viewRelativeAltitudeM),
    [viewGimbalPitchDegree, viewRelativeAltitudeM]
  )

  const rgbAlignScaleEffective = useMemo(() => {
    const globalScale = Number.isFinite(rgbAlignScale) ? Math.min(2, Math.max(0.5, rgbAlignScale)) : 1
    if (!rgbAlignScaleContextKey) return globalScale

    try {
      const stored = parseStoredScale(window.localStorage.getItem(getRgbAlignScaleStorageKey(rgbAlignScaleContextKey)))
      return stored ?? globalScale
    } catch {
      return globalScale
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [rgbAlignScale, rgbAlignScaleContextKey, rgbAlignScaleContextTick])

  const [viewLabelTemps, setViewLabelTemps] = useState<any | null>(null)
  const [viewLabelTempsStatus, setViewLabelTempsStatus] = useState<'idle' | 'loading' | 'ready' | 'error'>('idle')
  const [viewLabelTempsError, setViewLabelTempsError] = useState<string>('')
  const viewLabelTempsAbortRef = useRef<AbortController | null>(null)
  const viewLabelTempsDebounceRef = useRef<number | null>(null)

  type WideTempsGrid = { w: number; h: number; data: Float32Array }
  const wideTempsCacheRef = useRef<Map<string, WideTempsGrid>>(new Map())
  const wideTempsInFlightRef = useRef<Map<string, Promise<WideTempsGrid | null>>>(new Map())
  const viewHoverTempsCsvNameRef = useRef<string>('')
  const viewHoverPointRef = useRef<{ x: number; y: number } | null>(null)
  const viewHoverRafRef = useRef<number | null>(null)
  const viewHoverLastRef = useRef<{ px: number; py: number; v: number } | null>(null)
  const [viewHoverPixelTempC, setViewHoverPixelTempC] = useState<number | null>(null)
  const [viewHoverPixel, setViewHoverPixel] = useState<{ x: number; y: number } | null>(null)
  const [viewHoverTempsStatus, setViewHoverTempsStatus] = useState<'idle' | 'loading' | 'ready' | 'missing' | 'error'>('idle')

  const computeDefaultMetadataSelections = (report: any, selectedImagePath: string) => {
    const defaults: Record<string, boolean> = {}

    const selectedImageName = selectedImagePath.split('/').pop() || selectedImagePath
    void selectedImageName

    const keysToDefaultChecked = [
      // Temperature measurements
      'measurement_temperatures.pixel_stats.min',
      'measurement_temperatures.pixel_stats.mean',
      'measurement_temperatures.pixel_stats.max',

      // Thermal parameters
      'measurement_params.Distance',
      'measurement_params.RelativeHumidity',
      'measurement_params.Emissivity',
      'measurement_params.AmbientTemperature',
      'measurement_params.WindSpeed',
      'measurement_params.Irradiance',

      // Image info
      'image.file_name',
      'image.tiff_file',
      'image.camera_model',
      'image.serial_number',
      'image.focal_length',
      'image.f_number',
      'image.width',
      'image.height',
      'image.timestamp_created',
      'image.latitude',
      'image.longitude',

      // Geolocation (incl. QR)
      'geolocation.latitude',
      'geolocation.longitude',
      'geolocation.qr_code',
    ]

    for (const k of keysToDefaultChecked) defaults[k] = true

    // If no QR exists, keep the default but it simply won't render.
    void report
    return defaults
  }

  const tiffIndex = useMemo(() => {
    const isTiff = (p: string) => {
      const lower = p.toLowerCase()
      return lower.endsWith('.tif') || lower.endsWith('.tiff')
    }

    const candidates = Object.entries(fileMap)
      .filter(([path]) => isTiff(path))
      .map(([path, file]) => ({ path, file }))

    const preferred = candidates.filter((item) => item.path.toLowerCase().includes('/tiff/'))
    const list = preferred.length ? preferred : candidates

    const index = new Map<string, { path: string; file: File }>()
    for (const item of list) {
      const name = item.path.split('/').pop() || item.path
      index.set(name.toLowerCase(), { path: item.path, file: item.file })
    }
    return index
  }, [fileMap])

  const findMatchingTiffForImage = (imagePath: string) => {
    const imageName = imagePath.split('/').pop() || imagePath
    const stem = imageName.replace(/\.[^/.]+$/, '')
    const candidates = [`${imageName}.tiff`, `${imageName}.tif`, `${stem}.tiff`, `${stem}.tif`].map((n) => n.toLowerCase())
    return candidates.map((n) => tiffIndex.get(n)).find(Boolean) || null
  }

  useEffect(() => {
    if (activeAction !== 'view' || !selectedPath || fileKind !== 'image') {
      setViewMetadata(null)
      setViewMetadataStatus('idle')
      setViewMetadataError('')
      return
    }

    const tiffEntry = findMatchingTiffForImage(selectedPath)
    if (!tiffEntry) {
      setViewMetadata(null)
      setViewMetadataStatus('idle')
      setViewMetadataError('')
      return
    }

    const cached = viewMetadataCacheRef.current.get(tiffEntry.path)
    if (cached) {
      setViewMetadata(cached)
      setViewMetadataStatus('ready')
      setViewMetadataError('')

      setViewMetadataSelections((prev) => {
        const defaults = computeDefaultMetadataSelections(cached, selectedPath)
        const existing = prev[tiffEntry.path] || {}
        return { ...prev, [tiffEntry.path]: { ...defaults, ...existing } }
      })
      return
    }

    const controller = new AbortController()
    setViewMetadata(null)
    setViewMetadataStatus('loading')
    setViewMetadataError('')

    void (async () => {
      try {
        const formData = new FormData()
        formData.append('file', tiffEntry.file, tiffEntry.path)

        const response = await fetch(apiUrl('/api/metadata/probe?qr=1'), {
          method: 'POST',
          body: formData,
          signal: controller.signal,
        })

        if (!response.ok) {
          const err = await response.json().catch(() => ({}))
          throw new Error(err.detail || 'Metadata probe failed')
        }

        const report = await response.json()
        viewMetadataCacheRef.current.set(tiffEntry.path, report)
        setViewMetadata(report)
        setViewMetadataStatus('ready')

        setViewMetadataSelections((prev) => {
          const defaults = computeDefaultMetadataSelections(report, selectedPath)
          const existing = prev[tiffEntry.path] || {}
          return { ...prev, [tiffEntry.path]: { ...defaults, ...existing } }
        })
      } catch (err: any) {
        if (controller.signal.aborted) return
        setViewMetadata(null)
        setViewMetadataStatus('error')
        setViewMetadataError(err?.message || 'Metadata probe failed')
      }
    })()

    return () => controller.abort()
  }, [activeAction, selectedPath, fileKind, tiffIndex])

  useEffect(() => {
    if (viewLabelTempsDebounceRef.current !== null) {
      window.clearTimeout(viewLabelTempsDebounceRef.current)
      viewLabelTempsDebounceRef.current = null
    }
    if (viewLabelTempsAbortRef.current) {
      viewLabelTempsAbortRef.current.abort()
      viewLabelTempsAbortRef.current = null
    }

    if (activeAction !== 'view' || !selectedPath || fileKind !== 'image') {
      setViewLabelTemps(null)
      setViewLabelTempsStatus('idle')
      setViewLabelTempsError('')
      return
    }

    const tiffEntry = findMatchingTiffForImage(selectedPath)
    if (!tiffEntry) {
      setViewLabelTemps(null)
      setViewLabelTempsStatus('idle')
      setViewLabelTempsError('')
      return
    }

    // If there are no labels, keep category hidden.
    if (!selectedLabels || selectedLabels.length === 0) {
      setViewLabelTemps(null)
      setViewLabelTempsStatus('idle')
      setViewLabelTempsError('')
      return
    }

    const controller = new AbortController()
    viewLabelTempsAbortRef.current = controller
    setViewLabelTemps(null)
    setViewLabelTempsStatus('loading')
    setViewLabelTempsError('')

    // Debounce so dragging/resizing doesn't spam the backend.
    viewLabelTempsDebounceRef.current = window.setTimeout(() => {
      void (async () => {
        try {
          const payloadLabels = selectedLabels.map((l) => ({
            classId: l.classId,
            x: l.x,
            y: l.y,
            w: l.w,
            h: l.h,
            conf: l.conf,
            shape: l.shape || 'rect',
            source: l.source || 'auto',
          }))

          const formData = new FormData()
          formData.append('file', tiffEntry.file, tiffEntry.path)
          formData.append('labels', JSON.stringify(payloadLabels))

          const response = await fetch(apiUrl('/api/temperatures/labels?pad_px=3'), {
            method: 'POST',
            body: formData,
            signal: controller.signal,
          })

          if (!response.ok) {
            const err = await response.json().catch(() => ({}))
            throw new Error(err.detail || 'Label temperatures failed')
          }

          const data = await response.json()
          setViewLabelTemps(data)
          setViewLabelTempsStatus('ready')

          // Ensure these rows default to checked for this TIFF, without overriding user choices.
          const tiffKey = tiffEntry.path
          const keys: string[] = []
          const labelsOut = Array.isArray(data?.labels) ? data.labels : []
          for (const item of labelsOut) {
            const idx = Number(item?.index)
            if (!Number.isFinite(idx)) continue
            keys.push(`label_temperatures.${idx}.outside_edge_mean`)
            keys.push(`label_temperatures.${idx}.inside_mean`)
            keys.push(`label_temperatures.${idx}.inside_min`)
            keys.push(`label_temperatures.${idx}.inside_max`)
          }

          setViewMetadataSelections((prev) => {
            const existing = { ...(prev[tiffKey] || {}) }
            for (const k of keys) {
              if (existing[k] === undefined) existing[k] = true
            }
            return { ...prev, [tiffKey]: existing }
          })
        } catch (err: any) {
          if (controller.signal.aborted) return
          setViewLabelTemps(null)
          setViewLabelTempsStatus('error')
          setViewLabelTempsError(err?.message || 'Label temperatures failed')
        }
      })()
    }, 250)

    return () => {
      if (viewLabelTempsDebounceRef.current !== null) {
        window.clearTimeout(viewLabelTempsDebounceRef.current)
        viewLabelTempsDebounceRef.current = null
      }
      controller.abort()
      if (viewLabelTempsAbortRef.current === controller) {
        viewLabelTempsAbortRef.current = null
      }
    }
  }, [activeAction, selectedPath, fileKind, tiffIndex, selectedLabels])

  useEffect(() => {
    if (activeAction !== 'view' || !selectedPath || fileKind !== 'image') return
    const tiffEntry = findMatchingTiffForImage(selectedPath)
    if (!tiffEntry) return
    if (!selectedLabels || selectedLabels.length === 0) return

    const tiffKey = tiffEntry.path
    const keys: string[] = []
    for (let idx = 0; idx < selectedLabels.length; idx += 1) {
      keys.push(`fault_labels.${idx}.summary`)
    }

    setViewMetadataSelections((prev) => {
      const existing = { ...(prev[tiffKey] || {}) }
      for (const k of keys) {
        if (existing[k] === undefined) existing[k] = true
      }
      return { ...prev, [tiffKey]: existing }
    })
  }, [activeAction, selectedPath, fileKind, selectedLabels, tiffIndex])

  const parseWideTempsCsv = (csvText: string): WideTempsGrid => {
    const lines = csvText.split(/\r?\n/).filter((l) => l.trim().length > 0)
    if (lines.length < 2) throw new Error('CSV is empty')

    const header = lines[0].split(',')
    const w = Math.max(0, header.length - 1)
    const h = Math.max(0, lines.length - 1)
    if (!w || !h) throw new Error('CSV header is invalid')

    const data = new Float32Array(w * h)
    data.fill(Number.NaN)

    for (let r = 0; r < h; r += 1) {
      const parts = lines[r + 1].split(',')
      // parts[0] is the row index
      for (let c = 0; c < w; c += 1) {
        const raw = parts[c + 1] ?? ''
        if (raw === '') {
          data[r * w + c] = Number.NaN
        } else {
          const v = Number.parseFloat(raw)
          data[r * w + c] = Number.isFinite(v) ? v : Number.NaN
        }
      }
    }

    return { w, h, data }
  }

  const copyToClipboard = async (text: string) => {
    try {
      if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
        await navigator.clipboard.writeText(text)
        return true
      }
    } catch {
      // fall through to legacy path
    }

    try {
      const el = document.createElement('textarea')
      el.value = text
      el.setAttribute('readonly', 'true')
      el.style.position = 'fixed'
      el.style.left = '-9999px'
      el.style.top = '0'
      document.body.appendChild(el)
      el.select()
      const ok = document.execCommand('copy')
      document.body.removeChild(el)
      return ok
    } catch {
      return false
    }
  }

  const loadWideTempsGridFromDisk = async (csvName: string): Promise<WideTempsGrid | null> => {
    if (!folderHandle || !folderHandle.getDirectoryHandle) return null
    try {
      const tempsDir = await getWorkspaceThermalTemperaturesCsvsDir(false)
      const text = tempsDir ? await readTextFile(tempsDir, csvName) : null
      if (!text) {
        const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
        const legacyTempsDir = legacyFaultsDir?.getDirectoryHandle ? await legacyFaultsDir.getDirectoryHandle('temperatures_csvs', { create: false }) : null
        const legacyText = legacyTempsDir ? await readTextFile(legacyTempsDir, csvName) : null
        if (!legacyText) return null
        return parseWideTempsCsv(legacyText)
      }
      return parseWideTempsCsv(text)
    } catch {
      return null
    }
  }

  const ensureWideTempsGrid = async (csvName: string): Promise<WideTempsGrid | null> => {
    const cached = wideTempsCacheRef.current.get(csvName)
    if (cached) return cached

    const inFlight = wideTempsInFlightRef.current.get(csvName)
    if (inFlight) return inFlight

    const task = (async () => {
      try {
        const grid = await loadWideTempsGridFromDisk(csvName)
        if (grid) wideTempsCacheRef.current.set(csvName, grid)
        return grid
      } finally {
        wideTempsInFlightRef.current.delete(csvName)
      }
    })()

    wideTempsInFlightRef.current.set(csvName, task)
    return task
  }

  const getTempsCsvNameForSelectedImage = (path: string) => {
    const imageName = path.split('/').pop() || path
    const stem = imageName.replace(/\.[^/.]+$/, '')
    return `${stem}.pixel_temps.wide.csv`
  }

  const scheduleHoverTempUpdate = (point: { x: number; y: number }) => {
    viewHoverPointRef.current = point
    if (viewHoverRafRef.current !== null) return
    viewHoverRafRef.current = window.requestAnimationFrame(() => {
      viewHoverRafRef.current = null
      const latest = viewHoverPointRef.current
      if (!latest) return

      const csvName = viewHoverTempsCsvNameRef.current
      if (!csvName) return
      const grid = wideTempsCacheRef.current.get(csvName)
      if (!grid) return

      const px = Math.min(grid.w - 1, Math.max(0, Math.floor(latest.x * grid.w)))
      const py = Math.min(grid.h - 1, Math.max(0, Math.floor(latest.y * grid.h)))
      const v = grid.data[py * grid.w + px]

      const last = viewHoverLastRef.current
      if (last && last.px === px && last.py === py && Object.is(last.v, v)) return
      viewHoverLastRef.current = { px, py, v }
      setViewHoverPixel({ x: px, y: py })
      setViewHoverPixelTempC(v)
    })
  }

  const clearHoverTemp = () => {
    viewHoverPointRef.current = null
    viewHoverLastRef.current = null
    if (viewHoverRafRef.current !== null) {
      window.cancelAnimationFrame(viewHoverRafRef.current)
      viewHoverRafRef.current = null
    }
    setViewHoverPixel(null)
    setViewHoverPixelTempC(null)
  }

  useEffect(() => {
    clearHoverTemp()
    viewHoverTempsCsvNameRef.current = ''

    if (activeAction !== 'view' || !selectedPath || fileKind !== 'image') {
      setViewHoverTempsStatus('idle')
      return
    }

    const csvName = getTempsCsvNameForSelectedImage(selectedPath)
    viewHoverTempsCsvNameRef.current = csvName

    if (wideTempsCacheRef.current.has(csvName)) {
      setViewHoverTempsStatus('ready')
      return
    }

    setViewHoverTempsStatus('loading')
    void ensureWideTempsGrid(csvName)
      .then((grid) => {
        if (viewHoverTempsCsvNameRef.current !== csvName) return
        setViewHoverTempsStatus(grid ? 'ready' : 'missing')
      })
      .catch(() => {
        if (viewHoverTempsCsvNameRef.current !== csvName) return
        setViewHoverTempsStatus('error')
      })
  }, [activeAction, fileKind, selectedPath])

  useEffect(() => {
    const controller = new AbortController()

    fetch(apiUrl('/api/health'), { signal: controller.signal })
      .then((res) => res.json())
      .then((data) => {
        setBackendHealth(data)
        setBackendHealthError('')
        setMessage(typeof data?.message === 'string' ? data.message : 'Backend is running.')
      })
      .catch((err) => {
        if (err instanceof DOMException && err.name === 'AbortError') return
        setBackendHealth(null)
        setBackendHealthError('Backend unavailable. Start FastAPI.')
        setMessage('Backend unavailable. Start FastAPI.')
      })

    return () => controller.abort()
  }, [])

  useEffect(() => {
    if (activeMenu !== 'help' || activeHelpPage !== 'about') return

    const controller = new AbortController()
    fetch(apiUrl('/api/health'), { signal: controller.signal })
      .then((res) => res.json())
      .then((data) => {
        setBackendHealth(data)
        setBackendHealthError('')
      })
      .catch((err) => {
        if (err instanceof DOMException && err.name === 'AbortError') return
        setBackendHealth(null)
        setBackendHealthError('Backend unavailable. Start FastAPI.')
      })

    return () => controller.abort()
  }, [activeMenu, activeHelpPage])

  const buildTimeLocal = useMemo(() => {
    try {
      const d = new Date(__BUILD_TIME__)
      return Number.isFinite(d.getTime()) ? d.toLocaleString() : __BUILD_TIME__
    } catch {
      return __BUILD_TIME__
    }
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
    // Keep RGB labels independent from Thermal: store the working copy for the
    // current thermal context path while viewing RGB.
    if (activeAction !== 'view') return
    if (fileKind !== 'image') return
    if (viewVariant !== 'rgb') return
    if (!selectedPath) return
    if (!selectedViewPath || selectedViewPath === selectedPath) return
    rgbWorkingLabelsByThermalPathRef.current[selectedPath] = {
      labels: selectedLabels,
      areImageSpace: rgbLabelsAreImageSpace,
    }
  }, [activeAction, fileKind, rgbLabelsAreImageSpace, selectedLabels, selectedPath, selectedViewPath, viewVariant])

  useEffect(() => {
    return () => {
      if (realtimePersistTimerRef.current !== null) {
        window.clearTimeout(realtimePersistTimerRef.current)
        realtimePersistTimerRef.current = null
      }
      if (reportResyncTimerRef.current !== null) {
        window.clearTimeout(reportResyncTimerRef.current)
        reportResyncTimerRef.current = null
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
    loadCenterFaultFlags().catch(() => undefined)
    loadStarredFaults().catch(() => undefined)
  }, [activeAction, folderHandle])

  useEffect(() => {
    if (activeAction !== 'report' || !folderHandle) return
    loadFaultsList().catch(() => undefined)
    loadCenterFaultFlags().catch(() => undefined)
    loadStarredFaults().catch(() => undefined)
  }, [activeAction, folderHandle])

  useEffect(() => {
    if (activeAction !== 'report') return
    if (reportDefaultStarredAppliedRef.current) return
    reportDefaultStarredAppliedRef.current = true
    handleToggleOnlyStarredInReport(true)
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeAction])

  useEffect(() => {
    if (!selectedPath) {
      setLabelHistory([])
      setLabelHistoryIndex(0)
      return
    }
    setLabelHistory([selectedLabels])
    setLabelHistoryIndex(0)
  }, [selectedPath])

  useEffect(() => {
    // Cache the thermal context image dimensions so label interactions/rendering
    // can stay anchored to thermal coordinates even when viewing RGB.
    const file = selectedPath ? fileMap[selectedPath] : null
    if (!file || !file.type.startsWith('image/')) {
      setContextImageNatural(null)
      return
    }
    let cancelled = false
    void (async () => {
      const size = await getImageNaturalSizeFromFile(file)
      if (cancelled) return
      setContextImageNatural(size.width > 0 && size.height > 0 ? size : null)
    })()
    return () => {
      cancelled = true
    }
  }, [fileMap, selectedPath])

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

  const hasWritePermission = async (handle: FileSystemDirectoryHandle): Promise<boolean> => {
    if (!handle.queryPermission) return true
    try {
      const state = await handle.queryPermission({ mode: 'readwrite' } as any)
      return state === 'granted'
    } catch {
      return false
    }
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

    if ((selectedPath && !nextMap[selectedPath]) || (selectedViewPath && !nextMap[selectedViewPath])) {
      setSelectedPath('')
      setSelectedViewPath('')
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
    setSelectedViewPath('')
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
    const ok = window.confirm('Close the current folder and clear the session? (This will NOT delete any files on disk.)')
    if (!ok) return

    setFileMap({})
    setFileTree(null)
    setExpandedPaths(new Set(['']))
    setSelectedPath('')
    setSelectedViewPath('')
    setFileText('')
    if (fileUrl) URL.revokeObjectURL(fileUrl)
    setFileUrl('')
    setFileKind('')
    setActiveMenu('file')
    setActiveAction('')
    setSelectedLabels([])
    setFaultsList([])
    setCenterFaultFlags({})
    setOnlyCenterBoxFaults(false)
    setStarredFaults({})
    setOnlyStarredInReport(false)
    setScanCompleted(false)
    setFolderHandle(null)
    setLabelSaveError('')
    await clearStoredFolderHandle().catch(() => undefined)
  }

  const isImageName = (name: string) =>
    /\.(png|jpe?g|gif|webp|bmp|svg|ico)$/i.test(name)

  const isImagePath = (path: string) => isImageName(path.split('/').pop() || path)

  const parseDjiTvVariant = (name: string): { prefix: string; variant: 'T' | 'V'; ext: string } | null => {
    const m = /^(.*)_([TV])(\.[^.]+)$/i.exec(name)
    if (!m) return null
    const prefix = m[1] ?? ''
    const variant = (m[2] ?? '').toUpperCase() as 'T' | 'V'
    const ext = m[3] ?? ''
    if (!prefix || (variant !== 'T' && variant !== 'V') || !ext) return null
    return { prefix, variant, ext }
  }

  const getTocImageToken = (path: string) => {
    const base = (path.split('/').pop() || path).trim()
    if (!base) return ''

    const parsed = parseDjiTvVariant(base)
    if (parsed) {
      const token = (parsed.prefix.split('_').pop() || '').trim()
      if (token) return token
    }

    const noExt = base.replace(/\.[^.]+$/, '')
    const lastUnderscoreToken = (noExt.split('_').pop() || '').trim()
    return lastUnderscoreToken || noExt || base
  }

  // Returns a key used to pair Thermal<->RGB images even when the HHMMSS part differs.
  // Example: DJI_20250310134913_0018_T.JPG and DJI_20250310134912_0018_V.JPG
  // both map to: DJI_20250310_0018
  const getDjiDateIndexKey = (name: string) => {
    const m = /^DJI_(\d{8})\d{6}_(\d{4})_[TV](\.[^.]+)$/i.exec(name)
    if (!m) return ''
    const date = m[1] ?? ''
    const idx = m[2] ?? ''
    if (!date || !idx) return ''
    return `DJI_${date}_${idx}`
  }

  const getDjiTimestampNumber = (name: string) => {
    const m = /^DJI_(\d{14})_(\d{4})_[TV](\.[^.]+)$/i.exec(name)
    if (!m) return null
    const ts = Number(m[1])
    return Number.isFinite(ts) ? ts : null
  }

  const normalizePath = (path: string) =>
    path
      .replace(/\\/g, '/')
      .replace(/^\.\/+/, '')
      .replace(/^\/+/, '')
      .replace(/\/\.\//g, '/')

  const filePathByLowerCase = useMemo(() => {
    const out: Record<string, string> = {}
    for (const p of Object.keys(fileMap)) {
      out[p.toLowerCase()] = p
    }
    return out
  }, [fileMap])

  const dirPathLowerSet = useMemo(() => {
    const out = new Set<string>()
    const paths = Object.keys(fileMap)
    for (const p of paths) {
      const parts = p.split('/').filter(Boolean)
      if (parts.length <= 1) continue
      let acc = ''
      for (let i = 0; i < parts.length - 1; i += 1) {
        acc = acc ? `${acc}/${parts[i]}` : parts[i]
        out.add(acc.toLowerCase())
      }
    }
    return out
  }, [fileMap])

  const resolvePathCaseInsensitive = (candidatePath: string) => {
    if (!candidatePath) return ''
    if (fileMap[candidatePath]) return candidatePath
    const hit = filePathByLowerCase[candidatePath.toLowerCase()]
    return hit && fileMap[hit] ? hit : ''
  }

  const djiPairIndex = useMemo(() => {
    const thermalByKey: Record<string, string[]> = {}
    const rgbByKey: Record<string, string[]> = {}

    for (const path of Object.keys(fileMap)) {
      if (!isImagePath(path)) continue
      const base = path.split('/').pop() || ''
      const parsed = parseDjiTvVariant(base)
      if (!parsed) continue
      const key = getDjiDateIndexKey(base)
      if (!key) continue

      const bucket = parsed.variant === 'T' ? thermalByKey : rgbByKey
      if (!bucket[key]) bucket[key] = []
      bucket[key].push(path)
    }

    const sortValues = (map: Record<string, string[]>) => {
      for (const k of Object.keys(map)) {
        map[k] = (map[k] || []).slice().sort((a, b) => a.localeCompare(b))
      }
    }
    sortValues(thermalByKey)
    sortValues(rgbByKey)

    return { thermalByKey, rgbByKey }
  }, [fileMap])

  const pickBestPairedPath = (
    candidates: string[],
    preferredDir: string | null,
    sourceTimestamp: number | null
  ) => {
    if (!candidates || candidates.length === 0) return ''

    const normalizeDir = (p: string) => p.split('/').slice(0, -1).join('/').toLowerCase()
    const preferred = preferredDir ? preferredDir.toLowerCase() : ''
    const inPreferred = preferredDir
      ? candidates.filter((p) => normalizeDir(p) === preferred)
      : []
    const pool = inPreferred.length > 0 ? inPreferred : candidates

    if (sourceTimestamp !== null) {
      const scored = pool
        .map((p) => {
          const base = p.split('/').pop() || ''
          const ts = getDjiTimestampNumber(base)
          const diff = ts === null ? Number.POSITIVE_INFINITY : Math.abs(ts - sourceTimestamp)
          return { p, diff }
        })
        .sort((a, b) => (a.diff - b.diff) || a.p.localeCompare(b.p))
      return scored[0]?.p || ''
    }

    return pool.slice().sort((a, b) => a.localeCompare(b))[0] || ''
  }

  const getImageNaturalSizeFromFile = async (file: File): Promise<{ width: number; height: number }> => {
    try {
      const bitmap = await createImageBitmap(file)
      const size = { width: bitmap.width, height: bitmap.height }
      bitmap.close()
      return size
    } catch {
      return { width: 0, height: 0 }
    }
  }

  const getCenteredBoxBoundsPx = (imageW: number, imageH: number) => {
    const left = Math.max(0, Math.round((imageW - CENTER_BOX_W) / 2))
    const top = Math.max(0, Math.round((imageH - CENTER_BOX_H) / 2))
    const right = Math.min(imageW, left + CENTER_BOX_W)
    const bottom = Math.min(imageH, top + CENTER_BOX_H)
    return { left, top, right, bottom }
  }

  const isLabelCenterInsideCenterBox = (label: { x: number; y: number }, imageW: number, imageH: number) => {
    if (!(imageW > 0 && imageH > 0)) return false
    const { left, top, right, bottom } = getCenteredBoxBoundsPx(imageW, imageH)
    const cx = label.x * imageW
    const cy = label.y * imageH
    return cx >= left && cx <= right && cy >= top && cy <= bottom
  }

  const computeCenterFlagsForLabels = async (
    file: File,
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number }>
  ): Promise<{ perLabel: Array<0 | 1>; imageFlag: 0 | 1 }> => {
    const { width, height } = await getImageNaturalSizeFromFile(file)
    const perLabel = labels.map((l) => (isLabelCenterInsideCenterBox(l, width, height) ? 1 : 0) as 0 | 1)
    const imageFlag = (perLabel.some((v) => v === 1) ? 1 : 0) as 0 | 1
    return { perLabel, imageFlag }
  }

  const resolvePathInFileMap = (rawPath: string) => {
    const normalized = normalizePath(rawPath.trim())
    if (!normalized) return ''
    if (fileMap[normalized]) return normalized
    if (effectiveThermalFolderPath) {
      const candidate = normalized.startsWith(`${effectiveThermalFolderPath}/`)
        ? normalized
        : `${effectiveThermalFolderPath}/${normalized}`
      if (fileMap[candidate]) return candidate
    }
    return normalized
  }

  const findFolderPathByName = (targetName: string) => {
    if (!fileTree?.children?.length) return null
    const target = targetName.toLowerCase()

    // Collect *all* matches first (do not return early). The file tree traversal order
    // can cause a nested folder (e.g. `tiff/thermal`) to be encountered before the
    // intended top-level `thermal/`.
    const matches: string[] = []
    const stack: TreeNode[] = [...fileTree.children]
    while (stack.length > 0) {
      const node = stack.shift()!
      if (node.type === 'folder') {
        if (node.name.toLowerCase() === target) matches.push(node.path)
        if (node.children && node.children.length > 0) {
          stack.unshift(...node.children)
        }
      }
    }
    if (matches.length === 0) return null

    const parentOf = (p: string) => getParentDirPath(p) || ''
    const depthOf = (p: string) => p.split('/').filter(Boolean).length
    const hasSibling = (parent: string, name: string) => {
      const sib = parent ? `${parent}/${name}` : name
      return dirPathLowerSet.has(sib.toLowerCase())
    }

    const score = (candidate: string) => {
      const parent = parentOf(candidate)
      let s = 0
      // Prefer top-level matches.
      if (!parent) s += 2
      // Prefer the folder whose parent also contains the expected siblings.
      // This matches your contract: root contains rgb/, thermal/, tiff/.
      if (target === 'thermal') {
        if (hasSibling(parent, 'rgb')) s += 10
        if (hasSibling(parent, 'tiff') || hasSibling(parent, 'tif')) s += 5
        if (hasSibling(parent, 'workspace')) s += 1
      } else if (target === 'rgb') {
        if (hasSibling(parent, 'thermal')) s += 10
        if (hasSibling(parent, 'tiff') || hasSibling(parent, 'tif')) s += 5
        if (hasSibling(parent, 'workspace')) s += 1
      }
      return s
    }

    // Highest score wins; tie-breaker: shallowest path.
    const sorted = matches
      .slice()
      .sort((a, b) => (score(b) - score(a)) || (depthOf(a) - depthOf(b)) || a.localeCompare(b))
    return sorted[0] || null
  }

  const inferFolderPathFromFileMap = (targetName: 'thermal' | 'rgb') => {
    const target = targetName.toLowerCase()
    const paths = Object.keys(fileMap)
    for (const path of paths) {
      const parts = path.split('/').filter(Boolean)
      const idx = parts.findIndex((p) => p.toLowerCase() === target)
      if (idx >= 0) return parts.slice(0, idx + 1).join('/')
    }
    return null
  }

  const thermalFolderPath = useMemo(() => findFolderPathByName('thermal'), [fileTree])
  const rgbFolderPath = useMemo(() => findFolderPathByName('rgb'), [fileTree])

  // Use fileTree detection when available; fall back to fileMap inference so actions
  // like Scan don't fail if the tree hasn't populated yet.
  const effectiveThermalFolderPath = useMemo(
    () => thermalFolderPath ?? inferFolderPathFromFileMap('thermal'),
    [fileMap, thermalFolderPath]
  )
  const effectiveRgbFolderPath = useMemo(
    () => rgbFolderPath ?? inferFolderPathFromFileMap('rgb'),
    [fileMap, rgbFolderPath]
  )

  const openedFolderNameLower = (folderHandle?.name || '').toLowerCase()
  const openedThermalFolderDirectly = openedFolderNameLower === 'thermal' && !effectiveThermalFolderPath
  const openedRgbFolderDirectly = openedFolderNameLower === 'rgb' && !effectiveRgbFolderPath

  useEffect(() => {
    // New folder opened/restored → allow one-time ensure.
    setRgbLabelsDirEnsuredOnOpen(false)
  }, [folderHandle])

  useEffect(() => {
    // On first folder open, ensure `workspace/rgb_labels` exists (if an rgb folder exists), without re-prompting.
    if (!folderHandle || rgbLabelsDirEnsuredOnOpen) return
    const rootHandle = folderHandle

    let cancelled = false
    void (async () => {
      try {
        // If we can't locate `thermal/` within the opened folder, we don't know where
        // to place the sibling `workspace/` folder.
        if (!effectiveThermalFolderPath) return

        // Don't trigger extra permission prompts here — only create if we already have write permission.
        const canWrite = await hasWritePermission(rootHandle)
        if (!canWrite) return

        // Only ensure when an rgb folder exists in the opened tree.
        if (!effectiveRgbFolderPath) return

        // Ensure `workspace/rgb_labels`.
        await getWorkspaceRgbLabelsDir(true)
      } finally {
        if (!cancelled) setRgbLabelsDirEnsuredOnOpen(true)
      }
    })()

    return () => {
      cancelled = true
    }
  }, [effectiveRgbFolderPath, effectiveThermalFolderPath, folderHandle, rgbLabelsDirEnsuredOnOpen])

  const getRgbPathForThermal = (thermalPath: string) => {
    const name = thermalPath.split('/').pop() || ''
    const parsed = parseDjiTvVariant(name)
    if (!parsed) return null

    const unique = <T,>(values: T[]) => Array.from(new Set(values))
    const exts = unique([
      parsed.ext,
      parsed.ext.toLowerCase(),
      parsed.ext.toUpperCase(),
      '.jpg',
      '.JPG',
      '.jpeg',
      '.JPEG',
    ].filter(Boolean))
    const variants: Array<'V' | 'v'> = ['V', 'v']

    const candidates = unique(
      variants.flatMap((v) => exts.map((ext) => `${parsed.prefix}_${v}${ext}`))
    )

    if (effectiveThermalFolderPath && effectiveRgbFolderPath && thermalPath.startsWith(`${effectiveThermalFolderPath}/`)) {
      const rest = thermalPath.slice(effectiveThermalFolderPath.length + 1)
      const restParts = rest.split('/').filter(Boolean)
      restParts.pop()
      const subdir = restParts.join('/')
      const baseDir = `${effectiveRgbFolderPath}${subdir ? `/${subdir}` : ''}`
      for (const candidateName of candidates) {
        const resolved = resolvePathCaseInsensitive(`${baseDir}/${candidateName}`)
        if (resolved) return resolved
      }

      // Fallback: pair by DJI date+index even if HHMMSS differs.
      const key = getDjiDateIndexKey(name)
      const list = key ? (djiPairIndex.rgbByKey[key] || []) : []
      const ts = getDjiTimestampNumber(name)
      const picked = pickBestPairedPath(list, baseDir, ts)
      return picked || null
    }

    const dir = getParentDirPath(thermalPath)
    for (const candidateName of candidates) {
      const resolved = resolvePathCaseInsensitive(`${dir}/${candidateName}`)
      if (resolved) return resolved
    }

    // Fallback: pair by DJI date+index even if HHMMSS differs.
    const key = getDjiDateIndexKey(name)
    const list = key ? (djiPairIndex.rgbByKey[key] || []) : []
    const ts = getDjiTimestampNumber(name)
    const picked = pickBestPairedPath(list, null, ts)
    return picked || null
  }

  const getThermalPathForRgb = (rgbPath: string) => {
    const name = rgbPath.split('/').pop() || ''
    const parsed = parseDjiTvVariant(name)
    if (!parsed) return null

    const unique = <T,>(values: T[]) => Array.from(new Set(values))
    const exts = unique([
      parsed.ext,
      parsed.ext.toLowerCase(),
      parsed.ext.toUpperCase(),
      '.jpg',
      '.JPG',
      '.jpeg',
      '.JPEG',
    ].filter(Boolean))
    const variants: Array<'T' | 't'> = ['T', 't']
    const candidates = unique(
      variants.flatMap((v) => exts.map((ext) => `${parsed.prefix}_${v}${ext}`))
    )

    if (effectiveThermalFolderPath && effectiveRgbFolderPath && rgbPath.startsWith(`${effectiveRgbFolderPath}/`)) {
      const rest = rgbPath.slice(effectiveRgbFolderPath.length + 1)
      const restParts = rest.split('/').filter(Boolean)
      restParts.pop()
      const subdir = restParts.join('/')
      const baseDir = `${effectiveThermalFolderPath}${subdir ? `/${subdir}` : ''}`
      for (const candidateName of candidates) {
        const resolved = resolvePathCaseInsensitive(`${baseDir}/${candidateName}`)
        if (resolved) return resolved
      }

      // Fallback: pair by DJI date+index even if HHMMSS differs.
      const key = getDjiDateIndexKey(name)
      const list = key ? (djiPairIndex.thermalByKey[key] || []) : []
      const ts = getDjiTimestampNumber(name)
      const picked = pickBestPairedPath(list, baseDir, ts)
      return picked || null
    }

    const dir = getParentDirPath(rgbPath)
    for (const candidateName of candidates) {
      const resolved = resolvePathCaseInsensitive(`${dir}/${candidateName}`)
      if (resolved) return resolved
    }

    // Fallback: pair by DJI date+index even if HHMMSS differs.
    const key = getDjiDateIndexKey(name)
    const list = key ? (djiPairIndex.thermalByKey[key] || []) : []
    const ts = getDjiTimestampNumber(name)
    const picked = pickBestPairedPath(list, null, ts)
    return picked || null
  }

  const isRgbViewAvailableForThermal = (thermalPath: string) => {
    const rgbPath = getRgbPathForThermal(thermalPath)
    return Boolean(rgbPath && fileMap[rgbPath])
  }

  const isInThermalFolder = (path: string) => {
    if (!effectiveThermalFolderPath) return true
    return path === effectiveThermalFolderPath || path.startsWith(`${effectiveThermalFolderPath}/`)
  }

  const getParentDirPath = (path: string) => {
    const parts = path.split('/').filter(Boolean)
    parts.pop()
    return parts.join('/')
  }

  const WORKSPACE_DIR = 'workspace'
  const WORKSPACE_THERMAL_FAULTS_DIR = 'thermal_faults'
  const WORKSPACE_THERMAL_LABELS_DIR = 'thermal_labels'
  const WORKSPACE_THERMAL_LABELS_CENTER_DIR = 'thermal_labels_center'
  const WORKSPACE_THERMAL_METADATA_DIR = 'thermal_metadata'
  const WORKSPACE_THERMAL_TEMPS_CSV_DIR = 'thermal_temperatures_csvs'
  const WORKSPACE_RGB_LABELS_DIR = 'rgb_labels'

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

  const faultsListForTempTitleB = useMemo(() => {
    if (!onlyCenterBoxFaults) return faultsListImagePaths
    const hasAnyFlags = Object.keys(centerFaultFlags).length > 0
    if (!hasAnyFlags) return faultsListImagePaths
    return faultsListImagePaths.filter((p) => centerFaultFlags[p] === 1)
  }, [centerFaultFlags, faultsListImagePaths, onlyCenterBoxFaults])

  const faultsListRgbForTempTitleB = useMemo(() => {
    // Keep the same order/filtering as the thermal faults list.
    return faultsListForTempTitleB
      .map((thermalPath) => getRgbPathForThermal(thermalPath))
      .filter((p): p is string => Boolean(p && fileMap[p]))
  }, [faultsListForTempTitleB, fileMap])

  const scopedImagePaths = useMemo(() => {
    if (navScope === 'tree') return treeVisibleImagePaths
    if (navScope === 'faultsList') return faultsListForTempTitleB
    return []
  }, [faultsListForTempTitleB, navScope, treeVisibleImagePaths])

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

  const getWorkspaceProjectRootDir = async (create: boolean) => {
    if (!folderHandle) return null
    if (!folderHandle.getDirectoryHandle) return null

    // If the user opened the `thermal/` (or `rgb/`) folder directly, we cannot create
    // a sibling `workspace/` folder because we don't have access to the parent.
    if (openedThermalFolderDirectly || openedRgbFolderDirectly) return null

    // If we can't detect a thermal folder path yet (e.g. tree not populated), assume
    // the opened folder is the project root.
    if (!effectiveThermalFolderPath) return folderHandle

    const projectRootPath = getParentDirPath(effectiveThermalFolderPath)
    const parts = projectRootPath ? projectRootPath.split('/').filter(Boolean) : []
    let dir: FileSystemDirectoryHandle = folderHandle
    for (const part of parts) {
      if (!dir.getDirectoryHandle) return null
      dir = await dir.getDirectoryHandle(part, { create })
    }
    return dir
  }

  const getWorkspaceRootDir = async (create: boolean) => {
    const projectRoot = await getWorkspaceProjectRootDir(create)
    if (!projectRoot) return null
    if (!projectRoot.getDirectoryHandle) return null
    return projectRoot.getDirectoryHandle(WORKSPACE_DIR, { create })
  }

  const getWorkspaceSubdir = async (name: string, create: boolean) => {
    const workspaceRoot = await getWorkspaceRootDir(create)
    if (!workspaceRoot) return null
    if (!workspaceRoot.getDirectoryHandle) return null
    return workspaceRoot.getDirectoryHandle(name, { create })
  }

  // Resolve `workspace/<subdir>` for a specific file path.
  // This makes reads/writes stable even when the opened folder contains multiple sessions
  // (e.g. `data/<uuid>/thermal/...`), by anchoring `workspace/` next to that session's `thermal/`.
  const getWorkspaceSubdirForPath = async (targetPath: string, name: string, create: boolean) => {
    if (!folderHandle) return null
    if (!folderHandle.getDirectoryHandle) return null

    const normalizedTarget = normalizePath(String(targetPath || ''))

    const parts = normalizedTarget.split('/').filter(Boolean)
    const thermalIdx = parts.findIndex((p) => p.toLowerCase() === 'thermal')
    let projectRootParts = thermalIdx >= 0 ? parts.slice(0, thermalIdx) : []

    // If `targetPath` looks like it's rooted at `thermal/...` but the detected thermal folder is
    // nested (e.g. `thermal1_images/thermal/...`), infer the missing prefix from the detection.
    if (thermalIdx === 0 && effectiveThermalFolderPath) {
      const effParts = effectiveThermalFolderPath.split('/').filter(Boolean)
      const effThermalIdx = effParts.findIndex((p) => p.toLowerCase() === 'thermal')
      if (effThermalIdx > 0) {
        projectRootParts = effParts.slice(0, effThermalIdx)
      }
    }

    let projectRoot: FileSystemDirectoryHandle = folderHandle
    for (const part of projectRootParts) {
      if (!projectRoot.getDirectoryHandle) return null
      projectRoot = await projectRoot.getDirectoryHandle(part, { create })
    }

    if (!projectRoot.getDirectoryHandle) return null
    // Prefer the workspace folder next to the session's `thermal/`.
    try {
      const ws = await projectRoot.getDirectoryHandle(WORKSPACE_DIR, { create })
      if (ws?.getDirectoryHandle) {
        return ws.getDirectoryHandle(name, { create })
      }
    } catch {
      // ignore and fall back for reads
    }

    // Fallback for reads: if `workspace/` exists at the opened root, use it.
    // This supports older layouts where the user opened the session root directly.
    if (!create) {
      try {
        const wsAtRoot = await folderHandle.getDirectoryHandle(WORKSPACE_DIR, { create: false })
        if (wsAtRoot?.getDirectoryHandle) {
          const hit = await wsAtRoot.getDirectoryHandle(name, { create: false })
          if (hit) return hit
        }
      } catch {
        // ignore
      }
    }

    return null
  }

  const getWorkspaceThermalLabelsDirForPath = async (targetPath: string, create: boolean) =>
    getWorkspaceSubdirForPath(targetPath, WORKSPACE_THERMAL_LABELS_DIR, create)
  const getWorkspaceThermalLabelsCenterDirForPath = async (targetPath: string, create: boolean) =>
    getWorkspaceSubdirForPath(targetPath, WORKSPACE_THERMAL_LABELS_CENTER_DIR, create)
  const getWorkspaceThermalFaultsDirForPath = async (targetPath: string, create: boolean) =>
    getWorkspaceSubdirForPath(targetPath, WORKSPACE_THERMAL_FAULTS_DIR, create)

  const getWorkspaceThermalFaultsDir = async (create: boolean) => getWorkspaceSubdir(WORKSPACE_THERMAL_FAULTS_DIR, create)
  const getWorkspaceThermalLabelsDir = async (create: boolean) => getWorkspaceSubdir(WORKSPACE_THERMAL_LABELS_DIR, create)
  const getWorkspaceThermalLabelsCenterDir = async (create: boolean) => getWorkspaceSubdir(WORKSPACE_THERMAL_LABELS_CENTER_DIR, create)
  const getWorkspaceThermalMetadataDir = async (create: boolean) => getWorkspaceSubdir(WORKSPACE_THERMAL_METADATA_DIR, create)
  const getWorkspaceThermalTemperaturesCsvsDir = async (create: boolean) => getWorkspaceSubdir(WORKSPACE_THERMAL_TEMPS_CSV_DIR, create)
  const getWorkspaceRgbLabelsDir = async (create: boolean) => getWorkspaceSubdir(WORKSPACE_RGB_LABELS_DIR, create)

  const getLegacyWorkflowFaultsDir = async (create: boolean) => {
    const root = await getWorkflowRootDir(create)
    if (!root) return null
    return getFaultsDir(root, create)
  }

  const getLegacyRgbLabelsDir = async (create: boolean) => {
    const root = await getRgbRootDir(create)
    if (!root) return null
    if (!root.getDirectoryHandle) return null
    return root.getDirectoryHandle('labels', { create })
  }

  const getWorkflowRootDir = async (create: boolean) => {
    if (!folderHandle) return null
    const baseThermalPath = effectiveThermalFolderPath
    if (!baseThermalPath) return folderHandle
    if (!folderHandle.getDirectoryHandle) return null
    const parts = baseThermalPath.split('/').filter(Boolean)
    let dir: FileSystemDirectoryHandle = folderHandle
    for (const part of parts) {
      if (!dir.getDirectoryHandle) return null
      dir = await dir.getDirectoryHandle(part, { create })
    }
    return dir
  }

  const getRgbRootDir = async (create: boolean) => {
    if (!folderHandle) return null
    // Prefer the detected RGB folder path from the file tree.
    // Fallback: derive the parent folder from the currently viewed RGB image path.
    const fallbackRgbPath =
      viewVariant === 'rgb' && selectedViewPath ? selectedViewPath.split('/').slice(0, -1).join('/') : ''
    const basePath = effectiveRgbFolderPath || fallbackRgbPath
    if (!basePath) return folderHandle
    if (!folderHandle.getDirectoryHandle) return null
    const parts = basePath.split('/').filter(Boolean)
    let dir: FileSystemDirectoryHandle = folderHandle
    for (const part of parts) {
      if (!dir.getDirectoryHandle) return null
      dir = await dir.getDirectoryHandle(part, { create })
    }
    return dir
  }

  const getRgbLabelsDir = async (create: boolean) => {
    return getWorkspaceRgbLabelsDir(create)
  }

  // Ensures `workspace/rgb_labels` exists. If it doesn't exist, creates it.
  const ensureRgbLabelsDir = async () => {
    const root = await getRgbLabelsDir(true)
    if (!root) return null
    if (!root.getDirectoryHandle) return null
    return root
  }

  const getWorkflowFaultsDir = async (create: boolean) => {
    return getWorkspaceThermalFaultsDir(create)
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

  const parseYoloNumber = (raw: unknown) => {
    const s = String(raw ?? '')
      .replace(/^\uFEFF/, '')
      .trim()
    if (!s) return NaN
    // Some datasets use comma as decimal separator (e.g. "0,123").
    // Prefer interpreting comma as decimal when there's no dot present.
    if (s.includes(',') && !s.includes('.')) {
      if (/^-?\d+,\d+$/.test(s)) return Number(s.replace(',', '.'))
      if (/^-?\d{1,3}(,\d{3})+$/.test(s)) return Number(s.replace(/,/g, ''))
      return Number(s.replace(',', '.'))
    }
    return Number(s)
  }

  const inferSessionPrefixFromThermalDetection = () => {
    if (!effectiveThermalFolderPath) return ''
    const effParts = effectiveThermalFolderPath.split('/').filter(Boolean)
    const effThermalIdx = effParts.findIndex((p) => p.toLowerCase() === 'thermal')
    if (effThermalIdx > 0) return effParts.slice(0, effThermalIdx).join('/')
    return ''
  }

  // Lightweight realtime monitoring: poll only the relevant workspace folder(s) and
  // merge them into `fileMap` so label existence checks + reads stay current.
  useEffect(() => {
    if (!folderHandle) return
    if (openedThermalFolderDirectly || openedRgbFolderDirectly) return

    let cancelled = false

    const pollOnce = async () => {
      if (cancelled) return
      if (workspacePollInFlightRef.current) return

      workspacePollInFlightRef.current = (async () => {
        try {
          const projectRoot = await getWorkspaceProjectRootDir(false)
          const getProjectDir = projectRoot?.getDirectoryHandle?.bind(projectRoot)
          if (!projectRoot || !getProjectDir) return

          const sessionPrefix = inferSessionPrefixFromThermalDetection()
          const sessionPrefixWithSlash = sessionPrefix ? `${sessionPrefix}/` : ''

          const candidates: Array<{ prefix: string; dir: FileSystemDirectoryHandle }> = []
          const tryAdd = async (prefix: string, open: () => Promise<FileSystemDirectoryHandle | null>) => {
            try {
              const dir = await open()
              if (dir) candidates.push({ prefix, dir })
            } catch {
              // ignore
            }
          }

          await tryAdd(`${sessionPrefixWithSlash}${WORKSPACE_DIR}`, async () => {
            try {
              return await getProjectDir(WORKSPACE_DIR, { create: false })
            } catch {
              return null
            }
          })

          // Backward-compat: some older layouts store the workspace under `tif/` or `tiff/`.
          await tryAdd(`${sessionPrefixWithSlash}tif/${WORKSPACE_DIR}`, async () => {
            try {
              const tif = await getProjectDir('tif', { create: false })
              if (!tif?.getDirectoryHandle) return null
              return await tif.getDirectoryHandle(WORKSPACE_DIR, { create: false })
            } catch {
              return null
            }
          })

          await tryAdd(`${sessionPrefixWithSlash}tiff/${WORKSPACE_DIR}`, async () => {
            try {
              const tiff = await getProjectDir('tiff', { create: false })
              if (!tiff?.getDirectoryHandle) return null
              return await tiff.getDirectoryHandle(WORKSPACE_DIR, { create: false })
            } catch {
              return null
            }
          })

          if (candidates.length === 0) {
            if (workspaceIndexSignatureRef.current !== '') {
              workspaceIndexSignatureRef.current = ''
              setFileMap((prev) => {
                const next: Record<string, File> = {}
                for (const [p, f] of Object.entries(prev)) {
                  if (p.toLowerCase().includes(`/${WORKSPACE_DIR.toLowerCase()}/`) || p.toLowerCase().startsWith(`${WORKSPACE_DIR.toLowerCase()}/`)) {
                    continue
                  }
                  next[p] = f
                }
                return next
              })
              setWorkspaceIndexTick((t) => t + 1)
            }
            return
          }

          const workspaceEntries: Array<{ path: string; file: File }> = []
          for (const { prefix, dir } of candidates) {
            const entries = await readDirectoryEntries(dir)
            for (const e of entries) {
              const fullPath = e.path ? `${prefix}/${e.path}` : prefix
              workspaceEntries.push({ path: fullPath, file: e.file })
            }
          }

          const signature = workspaceEntries
            .map((e) => `${e.path.toLowerCase()}|${e.file.size}|${e.file.lastModified}`)
            .sort((a, b) => a.localeCompare(b))
            .join('\n')

          if (signature === workspaceIndexSignatureRef.current) return
          workspaceIndexSignatureRef.current = signature

          const prefixes = candidates.map((c) => c.prefix)
          const prefixesWithSlash = prefixes.map((p) => `${p}/`)

          setFileMap((prev) => {
            const next: Record<string, File> = {}
            for (const [p, f] of Object.entries(prev)) {
              const isUnderWorkspace = prefixes.some((prefix) => p === prefix) || prefixesWithSlash.some((prefix) => p.startsWith(prefix))
              if (isUnderWorkspace) continue
              next[p] = f
            }
            for (const entry of workspaceEntries) {
              next[entry.path] = entry.file
            }
            return next
          })

          setWorkspaceIndexTick((t) => t + 1)
        } finally {
          workspacePollInFlightRef.current = null
        }
      })()

      await workspacePollInFlightRef.current
    }

    void pollOnce()
    const id = window.setInterval(() => void pollOnce(), 2000)
    return () => {
      cancelled = true
      window.clearInterval(id)
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [folderHandle, openedThermalFolderDirectly, openedRgbFolderDirectly, effectiveThermalFolderPath])

  const getExpectedThermalLabelPaths = (imagePath: string) => {
    const normalizedImagePath = normalizePath(String(imagePath || ''))
    const baseName = normalizedImagePath.split('/').pop() || normalizedImagePath
    const stem = baseName.replace(/\.[^/.]+$/, '') || baseName

    // Many label exporters omit the trailing `_T` / `_V` in the label filename.
    // Example image: DJI_..._0011_T.JPG  -> label: DJI_..._0011.txt
    const stemNoTv = stem.replace(/_[TV]$/i, '')

    const imageParts = normalizedImagePath.split('/').filter(Boolean)
    const thermalIdx = imageParts.findIndex((p) => p.toLowerCase() === 'thermal')
    let sessionPrefix = thermalIdx >= 0 ? imageParts.slice(0, thermalIdx).join('/') : ''
    if ((!sessionPrefix || thermalIdx === 0) && effectiveThermalFolderPath) {
      sessionPrefix = inferSessionPrefixFromThermalDetection()
    }
    const sessionPrefixWithSlash = sessionPrefix ? `${sessionPrefix}/` : ''

    const relUnderThermalDir = thermalIdx >= 0 ? imageParts.slice(thermalIdx + 1, -1).join('/') : ''

    // Support alternate on-disk layouts where `workspace/` lives under `tif/` or `tiff/`.
    // (Common when TIFF assets are stored in a sibling folder.)
    const workspaceRoots = [
      `${sessionPrefixWithSlash}${WORKSPACE_DIR}/${WORKSPACE_THERMAL_LABELS_DIR}`,
      `${sessionPrefixWithSlash}tif/${WORKSPACE_DIR}/${WORKSPACE_THERMAL_LABELS_DIR}`,
      `${sessionPrefixWithSlash}tiff/${WORKSPACE_DIR}/${WORKSPACE_THERMAL_LABELS_DIR}`,
    ]

    const flatCandidates = workspaceRoots.flatMap((root) => [`${root}/${stem}.txt`, stemNoTv && `${root}/${stemNoTv}.txt`].filter(Boolean))
    const nestedCandidates = relUnderThermalDir
      ? workspaceRoots.flatMap((root) =>
          [`${root}/${relUnderThermalDir}/${stem}.txt`, stemNoTv && `${root}/${relUnderThermalDir}/${stemNoTv}.txt`].filter(Boolean)
        )
      : []

    const normalizeHit = (p: string) => resolvePathCaseInsensitive(normalizePath(p))

    const flatHit = flatCandidates.map(normalizeHit).find(Boolean) || ''
    const nestedHit = nestedCandidates.map(normalizeHit).find(Boolean) || ''

    // Prefer showing the first existing path; otherwise default to the canonical location.
    const flat = (flatHit || flatCandidates[0] || '').trim()
    const nested = relUnderThermalDir ? (nestedHit || nestedCandidates[0] || '').trim() : ''

    return {
      flat,
      nested,
      flatExists: Boolean(flatHit),
      nestedExists: Boolean(nestedHit),
    }
  }

  const parseYoloLabelsText = (text: string) => {
    const lines = String(text || '')
      .split(/\r?\n/)
      .map((l) => l.trim())
      .filter(Boolean)

    return lines
      .map((line) => {
        const parts = line.split(/\s+/)
        const classId = parseYoloNumber(parts[0])
        const x = parseYoloNumber(parts[1])
        const y = parseYoloNumber(parts[2])
        const w = parseYoloNumber(parts[3])
        const h = parseYoloNumber(parts[4])
        const conf = parts.length >= 6 ? parseYoloNumber(parts[5]) : 1
        return {
          classId: Number.isFinite(classId) ? classId : 0,
          x: Number.isFinite(x) ? x : 0,
          y: Number.isFinite(y) ? y : 0,
          w: Number.isFinite(w) ? w : 0,
          h: Number.isFinite(h) ? h : 0,
          conf: Number.isFinite(conf) ? conf : 1,
          shape: 'rect' as const,
          source: 'manual' as const,
        }
      })
      .filter((l) => Number.isFinite(l.x) && Number.isFinite(l.y) && Number.isFinite(l.w) && Number.isFinite(l.h))
  }

  const getRgbCropRectNormalized = (viewportRect: DOMRect, imgRect: DOMRect) => {
    const renderW = viewportRect.width
    const renderH = viewportRect.height
    if (!(renderW > 0 && renderH > 0)) return null
    if (!(imgRect.width > 0 && imgRect.height > 0)) return null

    // Crop is the user-visible viewport at 100% zoom. We express this crop as a
    // rectangle in full-image normalized coordinates (0..1) using DOM geometry.
    const toRgbNormRaw = (vx: number, vy: number) => {
      const pageX = viewportRect.left + vx
      const pageY = viewportRect.top + vy
      const nx = (pageX - imgRect.left) / imgRect.width
      const ny = (pageY - imgRect.top) / imgRect.height
      return { x: nx, y: ny }
    }

    const tl = toRgbNormRaw(0, 0)
    const br = toRgbNormRaw(renderW, renderH)

    const x0 = Math.min(tl.x, br.x)
    const y0 = Math.min(tl.y, br.y)
    const x1 = Math.max(tl.x, br.x)
    const y1 = Math.max(tl.y, br.y)

    const safeX0 = clamp01(x0)
    const safeY0 = clamp01(y0)
    const safeX1 = clamp01(x1)
    const safeY1 = clamp01(y1)

    const scaleX = safeX1 - safeX0
    const scaleY = safeY1 - safeY0
    if (!(scaleX > 0 && scaleY > 0)) return null

    return { x0: safeX0, y0: safeY0, scaleX, scaleY }
  }

  // Report/DOCX generation does not have access to DOMRects from the viewer.
  // This computes the default 100% zoom RGB crop rect (centered) using the
  // same viewport sizing and base scale used by the viewer.
  const getDefaultRgbCropRectNormalizedFromImageDims = (imageW: number, imageH: number) => {
    if (!(imageW > 0 && imageH > 0)) return null

    const viewportW = VIEWER_MAX_W
    const viewportH = VIEWER_MAX_H
    if (!(viewportW > 0 && viewportH > 0)) return null

    // The viewer first fits the image (contain) into the viewport, then applies
    // the base transform scale for the RGB view.
    const fitScale = Math.min(viewportW / imageW, viewportH / imageH)
    if (!(fitScale > 0)) return null

    const fittedW = imageW * fitScale
    const fittedH = imageH * fitScale
    if (!(fittedW > 0 && fittedH > 0)) return null

    const offsetX = (viewportW - fittedW) / 2
    const offsetY = (viewportH - fittedH) / 2

    const s = BASE_ZOOM * RGB_VIEW_IMAGE_SCALE
    if (!(s > 0)) return null

    const cx = viewportW / 2
    const cy = viewportH / 2
    const halfW = viewportW / (2 * s)
    const halfH = viewportH / (2 * s)

    const leftV = cx - halfW
    const topV = cy - halfH
    const rightV = cx + halfW
    const bottomV = cy + halfH

    const x0 = clamp01((leftV - offsetX) / fittedW)
    const y0 = clamp01((topV - offsetY) / fittedH)
    const x1 = clamp01((rightV - offsetX) / fittedW)
    const y1 = clamp01((bottomV - offsetY) / fittedH)

    const scaleX = x1 - x0
    const scaleY = y1 - y0
    if (!(scaleX > 0 && scaleY > 0)) return null

    return { x0, y0, scaleX, scaleY }
  }

  const convertRgbFullLabelsToCropNormalized = (
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>,
    crop: { x0: number; y0: number; scaleX: number; scaleY: number }
  ) => {
    const cropLeft = crop.x0
    const cropTop = crop.y0
    const cropRight = crop.x0 + crop.scaleX
    const cropBottom = crop.y0 + crop.scaleY

    return labels
      .map((label) => {
        const left = label.x - label.w / 2
        const top = label.y - label.h / 2
        const right = label.x + label.w / 2
        const bottom = label.y + label.h / 2
        if (![left, top, right, bottom].every(Number.isFinite)) return null

        const clippedLeft = Math.min(cropRight, Math.max(cropLeft, left))
        const clippedTop = Math.min(cropBottom, Math.max(cropTop, top))
        const clippedRight = Math.min(cropRight, Math.max(cropLeft, right))
        const clippedBottom = Math.min(cropBottom, Math.max(cropTop, bottom))
        const clippedW = clippedRight - clippedLeft
        const clippedH = clippedBottom - clippedTop
        if (!(clippedW > 0 && clippedH > 0)) return null

        const cxFull = clippedLeft + clippedW / 2
        const cyFull = clippedTop + clippedH / 2

        const x = (cxFull - crop.x0) / crop.scaleX
        const y = (cyFull - crop.y0) / crop.scaleY
        const w = clippedW / crop.scaleX
        const h = clippedH / crop.scaleY
        if (![x, y, w, h].every(Number.isFinite)) return null
        if (!(w > 0 && h > 0)) return null
        return {
          ...label,
          x: clamp01(x),
          y: clamp01(y),
          w: clamp01(w),
          h: clamp01(h),
        }
      })
      .filter(Boolean) as Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
  }

  const convertRgbCropLabelsToFullNormalized = (
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>,
    crop: { x0: number; y0: number; scaleX: number; scaleY: number }
  ) => {
    return labels
      .map((label) => {
        const x = crop.x0 + label.x * crop.scaleX
        const y = crop.y0 + label.y * crop.scaleY
        const w = label.w * crop.scaleX
        const h = label.h * crop.scaleY
        if (![x, y, w, h].every(Number.isFinite)) return null
        if (!(w > 0 && h > 0)) return null
        return {
          ...label,
          x: clamp01(x),
          y: clamp01(y),
          w: clamp01(w),
          h: clamp01(h),
        }
      })
      .filter(Boolean) as Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
  }

  const convertThermalFrameLabelsToRgbNormalized = (
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>,
    viewportRect: DOMRect,
    imgRect: DOMRect
  ) => {
    const renderW = viewportRect.width
    const renderH = viewportRect.height
    if (!(renderW > 0 && renderH > 0)) return []
    if (!(imgRect.width > 0 && imgRect.height > 0)) return []

    const frame = getRgbThermalFrameRect(renderW, renderH)
    if (!frame) return []
    const offset = getRgbOverlayOffsetPx(renderW, renderH)

    // The user-visible crop at 100% zoom is the viewport itself.
    // Anything outside this viewport is not visible and must not affect saved labels.
    const cropLeft = 0
    const cropTop = 0
    const cropRight = renderW
    const cropBottom = renderH

    const toRgbNorm = (vx: number, vy: number) => {
      const pageX = viewportRect.left + vx
      const pageY = viewportRect.top + vy
      const nx = (pageX - imgRect.left) / imgRect.width
      const ny = (pageY - imgRect.top) / imgRect.height
      return { x: clamp01(nx), y: clamp01(ny) }
    }

    const clamp = (v: number, min: number, max: number) => Math.min(max, Math.max(min, v))

    return labels
      .map((label) => {
        const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
        const centerX = isNormalized ? label.x * frame.inner.width : label.x
        const centerY = isNormalized ? label.y * frame.inner.height : label.y
        const boxW = isNormalized ? label.w * frame.inner.width : label.w
        const boxH = isNormalized ? label.h * frame.inner.height : label.h

        const leftPx = frame.inner.left + offset.x + (centerX - boxW / 2)
        const topPx = frame.inner.top + offset.y + (centerY - boxH / 2)
        const rightPx = leftPx + boxW
        const bottomPx = topPx + boxH

        // Treat the visible crop as a separate image: clip boxes to the viewport.
        const clippedLeft = clamp(leftPx, cropLeft, cropRight)
        const clippedTop = clamp(topPx, cropTop, cropBottom)
        const clippedRight = clamp(rightPx, cropLeft, cropRight)
        const clippedBottom = clamp(bottomPx, cropTop, cropBottom)
        const clippedW = Math.max(0, clippedRight - clippedLeft)
        const clippedH = Math.max(0, clippedBottom - clippedTop)
        if (!(clippedW > 0 && clippedH > 0)) return null

        const cxViewport = clippedLeft + clippedW / 2
        const cyViewport = clippedTop + clippedH / 2
        const c = toRgbNorm(cxViewport, cyViewport)

        return {
          ...label,
          x: c.x,
          y: c.y,
          w: clamp01(clippedW / imgRect.width),
          h: clamp01(clippedH / imgRect.height),
        }
      })
      .filter(Boolean) as Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
  }


  const getRgbStemForExport = () => {
    if (viewVariant !== 'rgb') return ''
    const rgbPath = selectedViewPath || ''
    const base = rgbPath.split('/').pop() || ''
    return base.replace(/\.[^/.]+$/, '')
  }

  useEffect(() => {
    // RGB export: when switching into RGB view, ensure workspace/rgb_labels/<stem>.txt exists.
    // Keep track of whether the file is still empty so Save can remain enabled until the
    // first write happens.
    if (viewVariant !== 'rgb' || activeAction !== 'view') {
      setRgbLabelsExportStatus('idle')
      setRgbLabelsExportDirty(false)
      setRgbLabelsExportFileExists(false)
      setRgbLabelsExportFileEmpty(false)
      setRgbLabelsAreImageSpace(false)
      return
    }

    const isActuallyRgbView = Boolean(selectedPath && selectedViewPath && selectedViewPath !== selectedPath)
    if (!isActuallyRgbView) {
      setRgbLabelsExportStatus('idle')
      setRgbLabelsExportDirty(false)
      setRgbLabelsExportFileExists(false)
      setRgbLabelsExportFileEmpty(false)
      setRgbLabelsAreImageSpace(false)
      return
    }

    const stem = getRgbStemForExport()
    if (!stem || !folderHandle) {
      setRgbLabelsExportStatus('idle')
      setRgbLabelsExportDirty(false)
      setRgbLabelsExportFileExists(false)
      setRgbLabelsExportFileEmpty(false)
      setRgbLabelsAreImageSpace(false)
      return
    }

    let cancelled = false
    setRgbLabelsExportDirty(false)
    setRgbLabelsExportStatus('checking')
    ;(async () => {
      // Try to create workspace/rgb_labels and an empty <stem>.txt if missing.
      // If we don't have permission, we'll just fall back to a "missing" state;
      // the Save click will request permission and try again.
      const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
      if (cancelled) return
      if (!canWrite) {
        setRgbLabelsExportFileExists(false)
        setRgbLabelsExportFileEmpty(true)
        setRgbLabelsExportStatus('missing')
        return
      }

      const labelsDir = await getRgbLabelsDir(true)
      if (cancelled) return
      if (!labelsDir) {
        setRgbLabelsExportFileExists(false)
        setRgbLabelsExportFileEmpty(true)
        setRgbLabelsExportStatus('missing')
        return
      }

      if (!labelsDir.getFileHandle) {
        setRgbLabelsExportFileExists(false)
        setRgbLabelsExportFileEmpty(true)
        setRgbLabelsExportStatus('missing')
        return
      }

      const outName = `${stem}.txt`
      const fileHandle = await labelsDir.getFileHandle(outName, { create: true })
      const file = await fileHandle.getFile()
      let text = await file.text()

      // If the workspace target is empty but a legacy file exists, migrate it.
      if (!text.trim()) {
        const legacyDir = await getLegacyRgbLabelsDir(false)
        if (legacyDir?.getFileHandle) {
          try {
            const legacyHandle = await legacyDir.getFileHandle(outName)
            const legacyFile = await legacyHandle.getFile()
            const legacyText = await legacyFile.text()
            if (legacyText.trim()) {
              await writeTextFile(labelsDir, outName, legacyText)
              text = legacyText
            }
          } catch {
            // ignore
          }
        }
      }
      if (cancelled) return
      const empty = text.trim().length === 0
      setRgbLabelsExportFileExists(true)
      setRgbLabelsExportFileEmpty(empty)
      setRgbLabelsExportStatus('exists')

      if (!empty) {
        const parsed = parseYoloLabelsText(text)
        // If we already have a cached RGB working copy for this thermal context
        // (and it's in image-space), don't overwrite it on every toggle.
        const cached = selectedPath ? rgbWorkingLabelsByThermalPathRef.current[selectedPath] : null
        if (!cached || !cached.areImageSpace) {
          // File format: crop-relative YOLO coords (visible viewport at 100%).
          // We'll convert to full-image normalized coords for stable rendering/editing.
          setRgbLabelsAreImageSpace(true)
          setRgbPendingCropFileLabels(parsed)
        }
      } else {
        setRgbLabelsAreImageSpace(false)
      }
    })().catch(() => {
      if (cancelled) return
      setRgbLabelsExportFileExists(false)
      setRgbLabelsExportFileEmpty(true)
      setRgbLabelsExportStatus('missing')
      setRgbLabelsAreImageSpace(false)
    })

    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [viewVariant, selectedViewPath, selectedPath, folderHandle, effectiveRgbFolderPath, activeAction])


  const readJsonFile = async <T,>(dir: FileSystemDirectoryHandle, name: string): Promise<T | null> => {
    const text = await readTextFile(dir, name)
    if (!text) return null
    try {
      return JSON.parse(text) as T
    } catch {
      return null
    }
  }

  const parseCenterFlagsDisk = (raw: unknown): { meta: CenterFlagsDiskV2['meta'] | null; flags: Record<string, 0 | 1> } => {
    if (!raw || typeof raw !== 'object') return { meta: null, flags: {} }

    // New format: { meta, flags }
    const asAny = raw as any
    if (asAny && typeof asAny === 'object' && asAny.meta && asAny.flags && typeof asAny.flags === 'object') {
      const metaCandidate = asAny.meta
      const metaOk =
        metaCandidate &&
        typeof metaCandidate === 'object' &&
        metaCandidate.version === 2 &&
        Number.isFinite(metaCandidate.boxW) &&
        Number.isFinite(metaCandidate.boxH)

      const flagsCandidate = asAny.flags
      const flags: Record<string, 0 | 1> = {}
      for (const [k, v] of Object.entries(flagsCandidate as Record<string, unknown>)) {
        if (v === 1 || v === 0) flags[k] = v
        else if (v === true) flags[k] = 1
        else if (v === false) flags[k] = 0
        else if (typeof v === 'number') flags[k] = v === 1 ? 1 : 0
      }

      return { meta: metaOk ? (metaCandidate as CenterFlagsDiskV2['meta']) : null, flags }
    }

    // Old format: flat map of { [path]: 0|1 }
    const flags: Record<string, 0 | 1> = {}
    for (const [k, v] of Object.entries(raw as Record<string, unknown>)) {
      if (v === 1 || v === 0) flags[k] = v
      else if (v === true) flags[k] = 1
      else if (v === false) flags[k] = 0
      else if (typeof v === 'number') flags[k] = v === 1 ? 1 : 0
    }
    return { meta: null, flags }
  }

  const readCenterFlagsDisk = async (faultsDir: FileSystemDirectoryHandle) => {
    const raw = await readJsonFile<unknown>(faultsDir, CENTER_FLAGS_FILE)
    return parseCenterFlagsDisk(raw)
  }

  const loadCenterFaultFlags = async () => {
    const faultsDir = await getWorkspaceThermalFaultsDir(false)
    const disk = faultsDir ? await readCenterFlagsDisk(faultsDir) : null
    if (!disk) {
      const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
      const legacyDisk = legacyFaultsDir ? await readCenterFlagsDisk(legacyFaultsDir) : null
      if (!legacyDisk) {
        setCenterFaultFlags({})
        return
      }
      const flags = legacyDisk.flags
      if (!flags || Object.keys(flags).length === 0) {
        setCenterFaultFlags({})
        return
      }
      const resolved: Record<string, 0 | 1> = {}
      for (const [rawPath, v] of Object.entries(flags)) {
        const candidate = resolvePathInFileMap(rawPath)
        if (!candidate) continue
        resolved[candidate] = v === 1 ? 1 : 0
      }
      setCenterFaultFlags(resolved)
      return
    }

    const flags = disk.flags
    if (!flags || Object.keys(flags).length === 0) {
      setCenterFaultFlags({})
      return
    }
    const resolved: Record<string, 0 | 1> = {}
    for (const [rawPath, v] of Object.entries(flags)) {
      const candidate = resolvePathInFileMap(rawPath)
      if (!candidate) continue
      resolved[candidate] = v === 1 ? 1 : 0
    }
    setCenterFaultFlags(resolved)
  }

  const loadStarredFaults = async () => {
    const workspaceFaultsDir = await getWorkspaceThermalFaultsDir(false)
    const text = (workspaceFaultsDir ? await readTextFile(workspaceFaultsDir, STARRED_FILE) : null)
      || (await (async () => {
        const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
        const workflowRoot = await getWorkflowRootDir(false)
        return (legacyFaultsDir ? await readTextFile(legacyFaultsDir, STARRED_FILE) : null)
          || (workflowRoot ? await readTextFile(workflowRoot, STARRED_FILE) : null)
      })())
    if (!text) {
      setStarredFaults({})
      return
    }
    const raw = normalizeFaultsText(text)
    const next: Record<string, true> = {}
    for (const item of raw) {
      const candidate = resolvePathInFileMap(item)
      if (!candidate) continue
      if (!fileMap[candidate]) continue
      next[candidate] = true
    }
    setStarredFaults(next)
  }

  const writeStarredFaultsToDisk = async (next: Record<string, true>) => {
    if (!folderHandle || !folderHandle.getDirectoryHandle) return
    const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
    if (!canWrite) throw new Error('Write permission is required to save starred images. Reopen the folder and allow write access.')

    const faultsDir = await getWorkspaceThermalFaultsDir(true)
    if (!faultsDir) {
      throw new Error(
        openedThermalFolderDirectly
          ? 'You opened the “thermal” folder directly. Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
          : openedRgbFolderDirectly
            ? 'You opened the “rgb” folder directly. Open the parent folder that contains “thermal/” (and “rgb/”) so the app can create a sibling “workspace/” folder.'
            : 'Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
      )
    }
    const lines = Object.keys(next)
      .filter((p) => Boolean(fileMap[p]))
      .sort((a, b) => a.localeCompare(b))
      .join('\n')

    await writeTextFile(faultsDir, STARRED_FILE, lines)
  }

  const toggleStarredForPath = (path: string, nextValue?: boolean) => {
    if (!path) return
    const next = { ...starredFaults }
    const willStar = typeof nextValue === 'boolean' ? nextValue : !Boolean(next[path])
    if (willStar) next[path] = true
    else delete next[path]

    setStarredFaults(next)

    if (activeAction === 'report' && onlyStarredInReport) {
      autoResyncReportDraft(undefined, undefined, undefined, next)
    }

    void (async () => {
      try {
        await writeStarredFaultsToDisk(next)
      } catch (error) {
        window.alert(error instanceof Error ? error.message : 'Failed to save starred images')
      }
    })()
  }

  const filterPathsForReport = (
    paths: string[],
    flagsOverride?: Record<string, 0 | 1>,
    onlyCenterOverride?: boolean,
    onlyStarredOverride?: boolean,
    starredOverride?: Record<string, true>
  ) => {
    const onlyStarred = onlyStarredOverride ?? onlyStarredInReport
    const onlyCenter = onlyCenterOverride ?? onlyCenterBoxFaults
    const flags = flagsOverride ?? centerFaultFlags
    const starred = starredOverride ?? starredFaults

    // Desired behavior:
    // - Neither enabled: all faults
    // - Only one enabled: only that category
    // - Both enabled: UNION (either category), with no duplicates.
    // Filtering preserves the original ordering and naturally avoids duplicates.
    const hasFlags = Object.keys(flags).length > 0
    const effectiveOnlyCenter = onlyCenter && hasFlags

    if (!onlyStarred && !effectiveOnlyCenter) return paths

    return paths.filter((p) => {
      const inStarred = onlyStarred ? Boolean(starred[p]) : false
      const inCenter = effectiveOnlyCenter ? flags[p] === 1 : false
      return inStarred || inCenter
    })
  }

  const computeAndPersistCenterFlagsForFaults = async (
    paths: string[],
    options?: { forceRecompute?: boolean }
  ): Promise<Record<string, 0 | 1>> => {
    if (!folderHandle) return { ...centerFaultFlags }
    const faultsDir = await getWorkspaceThermalFaultsDir(true)
    if (!faultsDir) {
      throw new Error(
        openedThermalFolderDirectly
          ? 'You opened the “thermal” folder directly. Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
          : openedRgbFolderDirectly
            ? 'You opened the “rgb” folder directly. Open the parent folder that contains “thermal/” (and “rgb/”) so the app can create a sibling “workspace/” folder.'
            : 'Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
      )
    }

    const forceRecompute = Boolean(options?.forceRecompute)

    // Start from persisted flags if available (most reliable across sessions),
    // unless we explicitly force a recompute (e.g. after changing center box size).
    const fromDisk = forceRecompute
      ? null
      : await (async () => {
        const ws = await readCenterFlagsDisk(faultsDir)
        if (ws && Object.keys(ws.flags).length > 0) return ws
        const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
        const legacy = legacyFaultsDir ? await readCenterFlagsDisk(legacyFaultsDir) : null
        return legacy && Object.keys(legacy.flags).length > 0 ? legacy : ws
      })()
    const nextFlags: Record<string, 0 | 1> = {}
    if (fromDisk && Object.keys(fromDisk.flags).length > 0) {
      for (const [rawPath, v] of Object.entries(fromDisk.flags)) {
        const candidate = resolvePathInFileMap(rawPath)
        if (!candidate) continue
        nextFlags[candidate] = v === 1 ? 1 : 0
      }
    } else {
      if (!forceRecompute) Object.assign(nextFlags, centerFaultFlags)
    }

    const labelsCenterDir = await getWorkspaceThermalLabelsCenterDir(true)
    if (!labelsCenterDir) return nextFlags

    const workspaceLabelsDir = await getWorkspaceThermalLabelsDir(false)
    const legacyLabelsDir = await (async () => {
      const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
      return legacyFaultsDir?.getDirectoryHandle ? await legacyFaultsDir.getDirectoryHandle('labels', { create: false }) : null
    })()

    for (const path of paths) {
      if (!forceRecompute && nextFlags[path] !== undefined) continue
      const file = fileMap[path]
      const sourceLabelsDir = workspaceLabelsDir || legacyLabelsDir
      if (!file || !sourceLabelsDir) {
        nextFlags[path] = 0
        continue
      }

      const stem = path.split('/').pop()?.replace(/\.[^/.]+$/, '') || path
      const labelText = (await readTextFile(sourceLabelsDir, `${stem}.txt`)) || ''
      const labels = labelText
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
          const [classId, x, y, w, h, conf] = line.split(/\s+/).map(parseYoloNumber)
          return {
            classId: Number.isFinite(classId) ? classId : 0,
            x: Number.isFinite(x) ? x : 0,
            y: Number.isFinite(y) ? y : 0,
            w: Number.isFinite(w) ? w : 0,
            h: Number.isFinite(h) ? h : 0,
            conf: Number.isFinite(conf) ? conf : 1,
          }
        })

      if (!labels.length) {
        nextFlags[path] = 0
        continue
      }

      const { perLabel, imageFlag } = await computeCenterFlagsForLabels(file, labels)
      const centerText = labels
        .map((l, i) => `${l.classId} ${l.x} ${l.y} ${l.w} ${l.h} ${l.conf} ${perLabel[i] ?? 0}`)
        .join('\n')
      await writeTextFile(labelsCenterDir, `${stem}.txt`, centerText)
      nextFlags[path] = imageFlag
    }

    await writeCenterFaultFlags(faultsDir, nextFlags)
    setCenterFaultFlags(nextFlags)
    return nextFlags
  }

  const ensureCenterFlagsForFaults = async (options?: { forceRecompute?: boolean }): Promise<Record<string, 0 | 1>> => {
    let forceRecompute = Boolean(options?.forceRecompute)
    if (!forceRecompute && centerFlagsEnsureInFlightRef.current) return centerFlagsEnsureInFlightRef.current

    // If the on-disk file was produced with a different box size, recompute once.
    if (!forceRecompute && folderHandle) {
      const faultsDir = await getWorkspaceThermalFaultsDir(false)
      const disk = faultsDir ? await readCenterFlagsDisk(faultsDir) : null
      const legacyDisk = !disk ? (await (async () => {
        const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
        return legacyFaultsDir ? await readCenterFlagsDisk(legacyFaultsDir) : null
      })()) : null

      const effective = disk || legacyDisk
      if (effective?.meta && (effective.meta.boxW !== CENTER_BOX_W || effective.meta.boxH !== CENTER_BOX_H)) {
        forceRecompute = true
      }
    }

    const p = (async () => {
      const paths = faultsListImagePaths
      if (!paths.length) return { ...centerFaultFlags }
      return computeAndPersistCenterFlagsForFaults(paths, { forceRecompute })
    })().finally(() => {
      centerFlagsEnsureInFlightRef.current = null
    })
    centerFlagsEnsureInFlightRef.current = p
    return p
  }

  const autoResyncReportDraft = (
    flagsOverride?: Record<string, 0 | 1>,
    onlyCenterOverride?: boolean,
    onlyStarredOverride?: boolean,
    starredOverride?: Record<string, true>
  ) => {
    if (activeAction !== 'report') return
    reportResyncPendingRef.current = { flagsOverride, onlyCenterOverride, onlyStarredOverride, starredOverride }

    // Debounce expensive draft rebuilds so rapid UI changes don't stutter.
    if (reportResyncTimerRef.current !== null) return
    reportResyncTimerRef.current = window.setTimeout(() => {
      reportResyncTimerRef.current = null
      const pending = reportResyncPendingRef.current
      reportResyncPendingRef.current = null
      if (!pending) return

      // Run on next frame to keep UI responsive.
      window.requestAnimationFrame(() => {
        if (activeAction !== 'report') return
        const rawPaths = normalizeFaultsText(reportText || faultsList.join('\n'))
        const flags = pending.flagsOverride ?? centerFaultFlags
        const onlyCenter = pending.onlyCenterOverride ?? onlyCenterBoxFaults
        const printable = filterPathsForReport(rawPaths, flags, onlyCenter, pending.onlyStarredOverride, pending.starredOverride)

        const last = lastReportPrintablePathsRef.current
        const same =
          last &&
          last.length === printable.length &&
          last.every((value, idx) => value === printable[idx])

        if (!same) {
          lastReportPrintablePathsRef.current = printable
          setDocxDraft((prev) => buildDocxDraftFromPaths(printable, prev))
        }
        setDocxPreviewStatus(`Draft auto-synced (${printable.length} item(s)).`)
        setDocxPreviewError('')
      })
    }, 80)
  }

  const handleToggleOnlyStarredInReport = (nextOnly: boolean) => {
    setOnlyStarredInReport(nextOnly)
    if (activeAction !== 'report') return
    // Pass the next checkbox value explicitly because React state updates are async.
    autoResyncReportDraft(undefined, undefined, nextOnly)
  }

  const handleToggleOnlyCenterBoxFaults = (nextOnly: boolean) => {
    setOnlyCenterBoxFaults(nextOnly)

    if (!nextOnly) {
      autoResyncReportDraft(undefined, false)
      return
    }

    // Turned ON: ensure flags exist, then re-filter the faults list and auto-sync Report.
    if (activeAction === 'report') {
      setDocxPreviewStatus('Computing center-box flags…')
      setDocxPreviewError('')
    }

    void (async () => {
      const flags = await ensureCenterFlagsForFaults()
      autoResyncReportDraft(flags, true)
    })()
  }

  const loadFaultsList = async () => {
    if (!folderHandle) return
    const faultsDir = await getWorkspaceThermalFaultsDir(false)
    const text = faultsDir ? await readTextFile(faultsDir, 'faults.txt') : null
    if (!text) {
      const workflowRoot = await getWorkflowRootDir(false)
      const legacy = workflowRoot ? await readTextFile(workflowRoot, 'faults.txt') : null
      const fallback = legacy || (await readTextFile(folderHandle, 'faults.txt'))
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

  const fileStemFromPath = (value: string) => {
    const normalized = (value || '').replace(/\\/g, '/').trim()
    const base = normalized.split('/').pop() || normalized
    // Remove only the last extension (.JPG, .jpeg, .tif, etc.).
    return base.replace(/\.[^.]+$/, '')
  }

  const normalizeImageCaption = (caption: string | undefined | null, path: string) => {
    const c = (caption || '').trim()
    if (!c) return fileStemFromPath(path)

    // If caption looks like a file path or filename-with-extension, normalize it to a stem.
    const looksLikePath = /[\\/]/.test(c)
    const looksLikeFilenameWithExt = /\.[a-z0-9]{1,6}$/i.test(c)
    if (looksLikePath || looksLikeFilenameWithExt) return fileStemFromPath(c)

    return c
  }

  const buildDocxDraftFromPaths = (paths: string[], prev: DocxDraftItem[]) => {
    const prevByPath = new Map(prev.map((item) => [item.path, item]))
    return paths.map((path, idx) => {
      const existing = prevByPath.get(path) as (DocxDraftItem & { notes?: string; fields?: DocxDraftTableRow[]; tables?: DocxDraftTable[] }) | undefined

      const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`
      const isPlaceholderTables = (tables: DocxDraftTable[] | null) => {
        if (!Array.isArray(tables) || tables.length !== 1) return false
        const t = tables[0]
        const titleEmpty = !(t?.title || '').trim()
        const rows = Array.isArray(t?.rows) ? t.rows : []
        if (!titleEmpty || rows.length !== 1) return false
        const r = rows[0]
        return !((r?.name || '').trim() || (r?.description || '').trim())
      }

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

      const cleanedExistingTables = existingTables && isPlaceholderTables(existingTables) ? [] : existingTables

      const tables = cleanedExistingTables && cleanedExistingTables.length
        ? cleanedExistingTables
        : migratedTables && migratedTables.length
          ? migratedTables
          : []
      return {
        id: existing?.id ?? `${Date.now()}-${Math.random().toString(16).slice(2)}`,
        path,
        include: existing?.include ?? true,
        caption: normalizeImageCaption(existing?.caption ?? '', path),
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
    setDocxDraft((prev) =>
      prev.map((item) => {
        if (item.id !== draftId) return item
        const nextTables = (item.tables || []).filter((t) => t.id !== tableId)
        return {
          ...item,
          tables: nextTables,
        }
      })
    )
  }

  const updateDocxDraftTableTitleForAllImages = (draftId: string, tableId: string, title: string) => {
    setDocxDraft((prev) => {
      const source = prev.find((i) => i.id === draftId)
      const sourceTables = Array.isArray(source?.tables) ? source!.tables : []
      const tableIndex = sourceTables.findIndex((t) => t.id === tableId)
      if (tableIndex < 0) {
        // Fallback: update only the current item.
        return prev.map((item) => {
          if (item.id !== draftId) return item
          const nextTables = (item.tables || []).map((t) => (t.id === tableId ? { ...t, title } : t))
          return { ...item, tables: nextTables }
        })
      }

      return prev.map((item) => {
        const tables = Array.isArray(item.tables) ? item.tables : []
        if (tableIndex >= tables.length) return item
        const nextTables = tables.map((t, idx) => (idx === tableIndex ? { ...t, title } : t))
        return { ...item, tables: nextTables }
      })
    })
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

  const updateDocxDraftRowDescriptionForAllImagesByTableTitleAndRowName = (tableTitle: string, rowName: string, description: string) => {
    const titleKey = (tableTitle || '').trim().toLowerCase()
    const rowKey = (rowName || '').trim().toLowerCase()
    if (!titleKey || !rowKey) return

    setDocxDraft((prev) =>
      prev.map((item) => {
        const tables = Array.isArray(item.tables) ? item.tables : []
        if (!tables.length) return item

        let anyChanged = false
        const nextTables = tables.map((t) => {
          if (((t?.title || '').trim().toLowerCase() !== titleKey)) return t

          const rows = Array.isArray(t?.rows) ? t.rows : []
          if (!rows.length) return t

          let changedRows = false
          const nextRows = rows.map((r) => {
            if (((r?.name || '').trim().toLowerCase() !== rowKey)) return r
            if (String(r?.description ?? '') === description) return r
            changedRows = true
            return { ...r, description }
          })

          if (!changedRows) return t
          anyChanged = true
          return { ...t, rows: nextRows }
        })

        return anyChanged ? { ...item, tables: nextTables } : item
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
    const rawPaths = normalizeFaultsText(reportText || faultsList.join('\n'))
    const printable = filterPathsForReport(rawPaths)
    setDocxDraft((prev) => buildDocxDraftFromPaths(printable, prev))
    setDocxPreviewStatus(`Draft synced (${printable.length} item(s)).`)
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

  // const loadReportFromDisk = async () => {
  //   setReportError('')
  //   setReportStatus('Loading…')

  //   if (!folderHandle) {
  //     setReportError('Choose a folder with File → Open Folder…')
  //     setReportStatus('')
  //     return
  //   }
  //   const workflowRoot = await getWorkflowRootDir(false)
  //   const text = (workflowRoot ? await readTextFile(workflowRoot, 'faults.txt') : null) || (await readTextFile(folderHandle, 'faults.txt')) || ''
  //   const normalized = normalizeFaultsText(text)
  //   setReportText(normalized.join('\n'))
  //   const printable = filterPathsForReport(normalized)
  //   setDocxDraft((prev) => buildDocxDraftFromPaths(printable, prev))
    
  // }

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

      const faultsDir = await getWorkspaceThermalFaultsDirForPath(selectedPath, true)
      if (!faultsDir) {
        throw new Error(
          openedThermalFolderDirectly
            ? 'You opened the “thermal” folder directly. Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
            : openedRgbFolderDirectly
              ? 'You opened the “rgb” folder directly. Open the parent folder that contains “thermal/” (and “rgb/”) so the app can create a sibling “workspace/” folder.'
              : 'Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
        )
      }
      await writeTextFile(faultsDir, 'faults.txt', content)

      setFaultsList(normalized)
      setReportText(content)
      setDocxDraft((prev) => buildDocxDraftFromPaths(normalized, prev))
      setReportStatus(`Saved ${normalized.length} item(s).`)
    } catch (error) {
      setReportError(error instanceof Error ? error.message : 'Failed to save faults.txt')
      setReportStatus('')
    }
  }

  const buildWordReportBlobFromDraft = async (draft: DocxDraftItem[], options?: { includeRgbPairs?: boolean }) => {
    const includeRgbPairs = options?.includeRgbPairs ?? true

    const paragraphRunsFromMultiline = (value: string) => {
      const parts = value.replace(/\r\n/g, '\n').split('\n')
      return parts.map((part, idx) => new TextRun({ text: part, break: idx === 0 ? 0 : 1 }))
    }

    const paragraphRunsFromMultilineSized = (value: string, size: number) => {
      const parts = value.replace(/\r\n/g, '\n').split('\n')
      return parts.map((part, idx) => new TextRun({ text: part, break: idx === 0 ? 0 : 1, size }))
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

    // Images section sizing: keep photos large, but conservative enough so
    // docx-preview pagination + PDF capture doesn't clip the footer/page number.
    // Tables will start on the next page when present.
    const maxW = 540
    const maxH = 610
    const children: Array<Paragraph | Table> = []

    const items = draft.filter((d) => d.include)

    const parseMetadataTimestampToDate = (raw: unknown): Date | null => {
      if (raw === null || raw === undefined) return null
      if (typeof raw === 'number' && Number.isFinite(raw)) {
        const d = new Date(raw)
        return Number.isFinite(d.getTime()) ? d : null
      }

      const s = String(raw).trim()
      if (!s) return null

      // Common EXIF/ExifTool format: "YYYY:MM:DD HH:MM:SS" (sometimes with timezone).
      const m = s.match(
        /^(\d{4})[:\-/](\d{2})[:\-/](\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?(?:\s*(Z|[+-]\d{2}:?\d{2})\s*)?$/
      )
      if (m) {
        const year = Number(m[1])
        const month = Number(m[2])
        const day = Number(m[3])
        const hour = m[4] ? Number(m[4]) : 0
        const minute = m[5] ? Number(m[5]) : 0
        const second = m[6] ? Number(m[6]) : 0
        const tz = m[7] || ''

        if (tz === 'Z' || /^[+-]\d{2}:?\d{2}$/.test(tz)) {
          const mm = String(month).padStart(2, '0')
          const dd = String(day).padStart(2, '0')
          const hh = String(hour).padStart(2, '0')
          const mi = String(minute).padStart(2, '0')
          const ss = String(second).padStart(2, '0')
          const tzNorm = tz === 'Z' ? 'Z' : tz.includes(':') ? tz : `${tz.slice(0, 3)}:${tz.slice(3)}`
          const iso = `${year}-${mm}-${dd}T${hh}:${mi}:${ss}${tzNorm}`
          const d = new Date(iso)
          return Number.isFinite(d.getTime()) ? d : null
        }

        const d = new Date(year, month - 1, day, hour, minute, second)
        return Number.isFinite(d.getTime()) ? d : null
      }

      const d = new Date(s)
      return Number.isFinite(d.getTime()) ? d : null
    }

    const getCoverDateText = async (): Promise<string> => {
      try {
        const candidates = items.length ? items : draft
        for (const item of candidates) {
          const meta = await readMetadataForImagePath(item.path)
          if (!meta) continue
          const categories: any = meta?.categories && typeof meta.categories === 'object' ? meta.categories : {}
          const timestamps = categories?.timestamps || {}
          const ts = timestamps?.CreateDate ?? timestamps?.DateTimeOriginal ?? meta?.summary?.DateTimeOriginal ?? meta?.summary?.DateTime
          const d = parseMetadataTimestampToDate(ts)
          if (!d) continue
          return new Intl.DateTimeFormat('en-GB', { day: 'numeric', month: 'long', year: 'numeric' }).format(d)
        }
      } catch {
        // best-effort only
      }
      return ''
    }

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
      // Make report annotations easier to read by slightly expanding the
      // drawn boxes/ellipses (model outputs can be tight around the fault).
      const REPORT_LABEL_SCALE = 1.8 * 1.5
      const REPORT_LABEL_STROKE_SCALE = 1.3

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
        // so labels remain visible on any background.
        const baseLineWidth = Math.max(1, Math.round(Math.min(canvas.width, canvas.height) / 320))
        const lineWidth = Math.max(1, Math.round(baseLineWidth * REPORT_LABEL_STROKE_SCALE))
        const green = 'rgba(0, 255, 0, 0.98)'
        const black = 'rgba(0, 0, 0, 0.75)'
        ctx.lineJoin = 'round'
        ctx.lineCap = 'round'

        // Keep the fault number smaller than the border thickness.
        const fontPx = Math.max(10, Math.round(baseLineWidth * 6))
        ctx.font = `bold ${fontPx}px sans-serif`
        ctx.textBaseline = 'top'
        ctx.textAlign = 'left'

        type Rect = { x: number; y: number; w: number; h: number }
        const overlaps = (a: Rect, b: Rect) => a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y

        // Precompute the scaled label rectangles so we can place numbers without overlapping other labels.
        const computed = labels.map((label) => {
          const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
          const cx = (isNormalized ? label.x : label.x / canvas.width) * canvas.width
          const cy = (isNormalized ? label.y : label.y / canvas.height) * canvas.height
          const bwRaw = (isNormalized ? label.w : label.w / canvas.width) * canvas.width
          const bhRaw = (isNormalized ? label.h : label.h / canvas.height) * canvas.height
          const bw = Math.min(canvas.width, Math.max(1, bwRaw * REPORT_LABEL_SCALE))
          const bh = Math.min(canvas.height, Math.max(1, bhRaw * REPORT_LABEL_SCALE))

          const width = bw
          const height = bh
          const left = Math.max(0, Math.min(canvas.width - width, cx - width / 2))
          const top = Math.max(0, Math.min(canvas.height - height, cy - height / 2))

          return { label, cx, cy, left, top, width, height }
        })

        const labelRects: Rect[] = computed.map((c) => ({ x: c.left, y: c.top, w: c.width, h: c.height }))
        const placedNumberRects: Rect[] = []

        const clampRectToCanvas = (r: Rect): Rect => {
          const x = Math.max(0, Math.min(canvas.width - r.w, Math.round(r.x)))
          const y = Math.max(0, Math.min(canvas.height - r.h, Math.round(r.y)))
          return { x, y, w: r.w, h: r.h }
        }

        const canPlaceNumberRect = (idx: number, rect: Rect) => {
          for (let j = 0; j < labelRects.length; j += 1) {
            if (j === idx) continue
            if (overlaps(rect, labelRects[j])) return false
          }
          for (const p of placedNumberRects) {
            if (overlaps(rect, p)) return false
          }
          return true
        }

        const drawFaultNumber = (idx: number) => {
          const n = idx + 1
          const text = String(n)
          const m = ctx.measureText(text)
          const textW = Math.max(1, Math.ceil(m.width))
          const textH = fontPx
          const r = labelRects[idx]

          const gap = Math.max(2, Math.round(lineWidth))

          const candidates: Rect[] = [
            // Outside: above
            { x: r.x, y: r.y - textH - gap, w: textW, h: textH },
            { x: r.x + r.w - textW, y: r.y - textH - gap, w: textW, h: textH },
            // Outside: right
            { x: r.x + r.w + gap, y: r.y, w: textW, h: textH },
            { x: r.x + r.w + gap, y: r.y + r.h - textH, w: textW, h: textH },
            // Outside: left
            { x: r.x - textW - gap, y: r.y, w: textW, h: textH },
            { x: r.x - textW - gap, y: r.y + r.h - textH, w: textW, h: textH },
            // Outside: below
            { x: r.x, y: r.y + r.h + gap, w: textW, h: textH },
            { x: r.x + r.w - textW, y: r.y + r.h + gap, w: textW, h: textH },
            // Inside fallback (last resort)
            { x: r.x + gap, y: r.y + gap, w: textW, h: textH },
            { x: r.x + r.w - textW - gap, y: r.y + gap, w: textW, h: textH },
          ]

          let placed: Rect | null = null
          for (const cand of candidates) {
            const c = clampRectToCanvas(cand)
            if (canPlaceNumberRect(idx, c)) {
              placed = c
              break
            }
          }

          if (!placed) {
            // Worst-case: clamp above-left.
            placed = clampRectToCanvas({ x: r.x, y: r.y - textH - gap, w: textW, h: textH })
          }

          placedNumberRects.push(placed)

          // No background; small green number with a subtle black outline for contrast.
          ctx.lineWidth = Math.max(1, Math.round(baseLineWidth))
          ctx.strokeStyle = black
          ctx.strokeText(text, placed.x, placed.y)
          ctx.fillStyle = green
          ctx.fillText(text, placed.x, placed.y)
        }

        for (let i = 0; i < computed.length; i += 1) {
          const { label, cx, cy, left, top, width, height } = computed[i]

          if (label.shape === 'ellipse') {
            const rxRaw = Math.max(1, width / 2)
            const ryRaw = Math.max(1, height / 2)
            // Clamp radii to stay within canvas even after scaling.
            const rx = Math.max(1, Math.min(rxRaw, cx, canvas.width - cx))
            const ry = Math.max(1, Math.min(ryRaw, cy, canvas.height - cy))

            ctx.beginPath()
            ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2)
            ctx.lineWidth = lineWidth + 1
            ctx.strokeStyle = black
            ctx.stroke()

            ctx.beginPath()
            ctx.ellipse(cx, cy, rx, ry, 0, 0, Math.PI * 2)
            ctx.lineWidth = lineWidth
            ctx.strokeStyle = green
            ctx.stroke()
          } else {
            ctx.lineWidth = lineWidth + 1
            ctx.strokeStyle = black
            ctx.strokeRect(left, top, width, height)

            ctx.lineWidth = lineWidth
            ctx.strokeStyle = green
            ctx.strokeRect(left, top, width, height)
          }

          // Fault number next to the label.
          drawFaultNumber(i)
        }
      }

      const blob = await new Promise<Blob>((resolve, reject) => {
        canvas.toBlob((b) => (b ? resolve(b) : reject(new Error('Failed to encode PNG'))), 'image/png')
      })

      return blobToUint8Array(blob)
    }

    // Cover page (always first page)
    // Use a borderless table with fixed row heights so placement is stable in Word and docx-preview.
    const coverDateText = await getCoverDateText()
    const coverContentHeightTwip = Math.round(convertInchesToTwip(11.69 - 1.0)) // A4 height minus 0.5" top+bottom margins
    const coverHeights = {
      topOffset: Math.round(convertInchesToTwip(0.25)),
      topLogo: Math.round(convertInchesToTwip(1.25)),
      spacer1: Math.round(convertInchesToTwip(1.25)),
      title: Math.round(convertInchesToTwip(0.8)),
      date: Math.round(convertInchesToTwip(0.35)),
      bottomLogo: Math.round(convertInchesToTwip(4.4)),
    }

    // Prevent the cover top logo from being clipped by keeping its pixel height
    // safely within the fixed row height.
    const twipPerIn = convertInchesToTwip(1)
    const topLogoRowIn = coverHeights.topLogo / twipPerIn
    const topLogoMaxHeightPx = Math.max(80, Math.floor(topLogoRowIn * 96 * 0.9))

    const toposolLogo = await fetchPublicImageRun('toposol-logo.png', 'png', 660, topLogoMaxHeightPx, true)
    const thermalLogo = await fetchPublicImageRun('thermal-logo.jpg', 'jpg', 580, 320, true)
    const used =
      coverHeights.topOffset +
      coverHeights.topLogo +
      coverHeights.spacer1 +
      coverHeights.title +
      coverHeights.date +
      coverHeights.bottomLogo
    // Pull the bottom image slightly upward on the cover page.
    const coverPullUpTwip = Math.round(convertInchesToTwip(0.65))
    const minSpacer2 = Math.round(convertInchesToTwip(0.35))
    const maxSpacer2 = Math.round(convertInchesToTwip(1.2))
    const spacer2Raw = coverContentHeightTwip - used - coverPullUpTwip
    const spacer2 = Math.min(maxSpacer2, Math.max(minSpacer2, spacer2Raw))

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

    const spacerParagraph = () => new Paragraph({ spacing: { before: 0, after: 0 }, children: [new TextRun({ text: '' })] })

    const coverCenterParagraph = (runs: TextRun[] | ImageRun[]) =>
      new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: runs })

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
                coverCenterParagraph(
                  toposolLogo ? [toposolLogo] : [new TextRun({ text: '[Missing toposol-logo.png]', color: 'B91C1C' })]
                ),
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
              coverCenterParagraph([new TextRun({ text: 'THERMOGRAPHY REPORT', bold: true, size: 56 })]),
            ]),
          ],
        }),
        new TableRow({
          height: { value: coverHeights.date, rule: HeightRule.EXACT },
          children: [
            borderlessCell([
              coverCenterParagraph([new TextRun({ text: coverDateText || '', size: 22 })]),
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
              coverCenterParagraph(
                thermalLogo ? [thermalLogo] : [new TextRun({ text: '[Missing thermal-logo.jpg]', color: 'B91C1C' })]
              ),
            ]),
          ],
        }),
      ],
    })

    children.push(coverTable)

    const BOOKMARK_DESCRIPTION = 'section_description'
    const BOOKMARK_EQUIPMENT = 'section_equipment'
    const BOOKMARK_REPORTS = 'section_reports'
    const BOOKMARK_CUSTOM_PREFIX = 'section_custom_'
    const BOOKMARK_IMAGE_PREFIX = 'section_report_image_'

    // Used to avoid auto-rendering an RGB pair when that RGB image is already explicitly included.
    const includedPathSet = new Set(items.map((it) => it.path))

    const cleanCustomChapters = reportChapters
      .map((c) => ({
        id: c.id,
        chapterTitle: (c.chapterTitle || '').trim(),
        sections: (Array.isArray(c.sections) ? c.sections : [])
          .map((s) => ({
            ...s,
            title: (s.title || '').trim(),
            text: (s.text || '').trim(),
          }))
          .filter((s) => s.title || s.text || s.imageFile),
      }))
      .filter((c) => c.chapterTitle)

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
    const reportsPage = 4 + cleanCustomChapters.length

    const tocRightStop = Math.round(convertInchesToTwip(7.1))
    const tocTitleStop = Math.round(convertInchesToTwip(0.7))
    const tocNumberStop = Math.round(convertInchesToTwip(0.55))
    const tocImageIndent = Math.round(convertInchesToTwip(0.3))

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

    const tocImageLine = (index: number, token: string, bookmarkId: string, page: number) =>
      new Paragraph({
        alignment: AlignmentType.LEFT,
        tabStops: [
          { type: TabStopType.LEFT, position: tocImageIndent },
          { type: TabStopType.RIGHT, position: tocNumberStop + tocImageIndent },
          { type: TabStopType.LEFT, position: tocTitleStop + tocImageIndent },
          { type: TabStopType.RIGHT, position: tocRightStop, leader: LeaderType.DOT },
        ],
        spacing: { before: 40, after: 40 },
        children: [
          new TextRun({ text: '\t' }),
          new TextRun({ text: `${index}.` }),
          new TextRun({ text: '\t\t' }),
          new InternalHyperlink({
            anchor: bookmarkId,
            children: [new TextRun({ text: token, color: '000000' })],
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
      ...cleanCustomChapters.map((c, idx) => {
        const page = 4 + idx
        const chapterIndex = 3 + idx
        const bookmarkId = `${BOOKMARK_CUSTOM_PREFIX}${c.id}`
        return tocLine(chapterIndex, c.chapterTitle, bookmarkId, page)
      }),
      tocLine(3 + cleanCustomChapters.length, 'Reports (anomaly detected)', BOOKMARK_REPORTS, reportsPage),
      ...items.map((it, idx) => {
        const token = getTocImageToken(it.path) || (it.path.split('/').pop() || it.path)
        // Best-effort: assume 1 page per image item (tables may add extra pages).
        const page = reportsPage + idx
        return tocImageLine(idx + 1, token, `${BOOKMARK_IMAGE_PREFIX}${idx + 1}`, page)
      })
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

    // Custom user chapters (each on its own page).
    for (const chapter of cleanCustomChapters) {
      children.push(new Paragraph({ children: [new PageBreak()] }))
      children.push(sectionTitle(chapter.chapterTitle, `${BOOKMARK_CUSTOM_PREFIX}${chapter.id}`))

      if (!chapter.sections.length) {
        children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: '' })] }))
      } else {
        for (const section of chapter.sections) {
          if (section.title) {
            children.push(
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 120, after: 80 },
                children: [new TextRun({ text: section.title, bold: true, size: 24 })],
              })
            )
          }

          if (section.text) {
            children.push(
              new Paragraph({
                alignment: AlignmentType.CENTER,
                spacing: { before: 0, after: 120 },
                children: paragraphRunsFromMultiline(section.text),
              })
            )
          }

          if (section.imageFile) {
            const run = await fileToImageRun(section.imageFile, 520, 420)
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
                  children: [new TextRun({ text: `Unsupported chapter image: ${section.imageFile.type || 'unknown type'}`, color: 'B91C1C' })],
                })
              )
            }
          }
        }
      }
    }

    // Reports section: start a new page, show the title, then start the first image on the same page.
    children.push(new Paragraph({ children: [new PageBreak()] }))
    children.push(sectionTitle('Reports (anomaly detected)', BOOKMARK_REPORTS))
    children.push(new Paragraph({ children: [new TextRun({ text: '' })] }))

    for (let idx = 0; idx < items.length; idx += 1) {
      const item = items[idx]
      const file = fileMap[item.path]

      // Guard pagination: render each included image item on its own A4 page.
      // This prevents docx-preview from trying to lay out multiple large blocks
      // on the same page (which can lead to odd scaling/overflow in preview/PDF).
      if (idx > 0) {
        children.push(new Paragraph({ children: [new PageBreak()] }))
      }

      const tocToken = getTocImageToken(item.path)
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 60 },
          children: [
            new Bookmark({
              id: `${BOOKMARK_IMAGE_PREFIX}${idx + 1}`,
              children: [new TextRun({ text: `${idx + 1}. ${tocToken || ''}`.trim(), bold: true, size: 28 })],
            }),
          ],
        })
      )

      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: normalizeImageCaption(item.caption, item.path), bold: true })],
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

      // Pagination/fit bookkeeping (inches). We use this only to decide whether tables can start
      // under the thermal image when there is NO RGB pair.
      const A4_HEIGHT_IN = 11.69
      const PAGE_MARGINS_IN = 1.0 // 0.5" top + 0.5" bottom
      const HEADER_RESERVE_IN = 0.20
      const FOOTER_RESERVE_IN = 0.35
      const CONTENT_HEIGHT_IN = A4_HEIGHT_IN - PAGE_MARGINS_IN - HEADER_RESERVE_IN - FOOTER_RESERVE_IN

      const THERMAL_CAPTION_IN = 0.30
      const CAPTION_GAP_IN = 0.12
      const RGB_CAPTION_SPACING_IN = (140 + 60) / 1440
      const IMAGE_GAP_AFTER_IN = 0.14

      // Determine if we will render the paired RGB on the SAME page.
      // If yes, we will scale both images so the combined height fits.
      let rgbRender:
        | null
        | {
            path: string
            file: File
            imageType: 'png' | 'jpg' | 'gif' | 'bmp'
            width: number
            height: number
            buf: ArrayBuffer
          } = null

      if (includeRgbPairs) {
        try {
          const base = item.path.split('/').pop() || ''
          const parsed = parseDjiTvVariant(base)
          const isThermal = Boolean(parsed && parsed.variant === 'T')
          const rgbPath = isThermal ? getRgbPathForThermal(item.path) : null

          if (rgbPath && fileMap[rgbPath] && !includedPathSet.has(rgbPath)) {
            const rgbFile = fileMap[rgbPath]
            if (rgbFile && rgbFile.type.startsWith('image/')) {
              const rgbImageType = (() => {
                const t = (rgbFile.type || '').toLowerCase()
                if (t === 'image/png') return 'png'
                if (t === 'image/jpeg' || t === 'image/jpg') return 'jpg'
                if (t === 'image/gif') return 'gif'
                if (t === 'image/bmp') return 'bmp'
                return null
              })()

              if (rgbImageType) {
                const rgbBuf = await rgbFile.arrayBuffer()
                const rgbSize = await getImageSize(rgbFile)
                rgbRender = {
                  path: rgbPath,
                  file: rgbFile,
                  imageType: rgbImageType,
                  width: rgbSize.width,
                  height: rgbSize.height,
                  buf: rgbBuf,
                }
              }
            }
          }
        } catch {
          // best-effort only
        }
      }

      // Compute final image sizes.
      let scale = Math.min(maxW / width, maxH / height, 1)
      let w = Math.max(1, Math.round(width * scale))
      let h = Math.max(1, Math.round(height * scale))

      // If RGB will be rendered on the same page, downscale BOTH images to fit within the page.
      // This avoids preview/PDF weirdness and keeps the layout stable.
      let rgbW = 0
      let rgbH = 0
      if (rgbRender) {
        const scaleRgbWidth = Math.min(maxW / Math.max(1, rgbRender.width), maxH / Math.max(1, rgbRender.height), 1)

        const thermalHpx = height * scale
        const rgbHpx = rgbRender.height * scaleRgbWidth

        // Total non-image vertical space (captions + gaps + reserve).
        const nonImageIn = THERMAL_CAPTION_IN + CAPTION_GAP_IN + (THERMAL_CAPTION_IN + RGB_CAPTION_SPACING_IN) + CAPTION_GAP_IN + IMAGE_GAP_AFTER_IN
        const maxCombinedImagesPx = Math.max(1, Math.floor(Math.max(0.5, CONTENT_HEIGHT_IN - nonImageIn) * 96))

        const sumPx = Math.max(1, thermalHpx + rgbHpx)
        // Safety factor: Word/LibreOffice can differ slightly in layout; keep a bit of headroom.
        const pairedSafety = 0.92
        const down = Math.min(1, (maxCombinedImagesPx * pairedSafety) / sumPx)

        const finalThermalScale = scale * down
        const finalRgbScale = scaleRgbWidth * down

        w = Math.max(1, Math.round(width * finalThermalScale))
        h = Math.max(1, Math.round(height * finalThermalScale))
        rgbW = Math.max(1, Math.round(rgbRender.width * finalRgbScale))
        rgbH = Math.max(1, Math.round(rgbRender.height * finalRgbScale))
      }

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

      // Optional paired RGB image (same page).
      const didRenderRgb = Boolean(rgbRender)
      if (rgbRender) {
        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 140, after: 60 },
            children: [new TextRun({ text: fileStemFromPath(rgbRender.path), bold: true })],
          })
        )

        const rgbLabels = await readRgbLabelsForRgbPath(rgbRender.path)

        let usedRgbLabeled = false
        if (rgbLabels.length > 0) {
          try {
            const labeledBytes = await renderLabeledPngBytes(rgbRender.file, rgbLabels, rgbW, rgbH)
            pushImageParagraph(
              new ImageRun({
                type: 'png',
                data: labeledBytes,
                transformation: { width: rgbW, height: rgbH },
              })
            )
            usedRgbLabeled = true
          } catch {
            usedRgbLabeled = false
          }
        }

        if (!usedRgbLabeled) {
          pushImageParagraph(
            new ImageRun({
              type: rgbRender.imageType,
              data: new Uint8Array(rgbRender.buf),
              transformation: { width: rgbW, height: rgbH },
            })
          )
        }
      }

      // Pre-compute whether we will render a QR block.
      let qrBlock: null | { bytes: Uint8Array; lat: any; lon: any } = null
      try {
        const meta = await readMetadataForImagePath(item.path)
        if (meta) {
          const defaults = computeDefaultMetadataSelections(meta, item.path)
          const selTiffEntry = findMatchingTiffForImage(item.path)
          const selTiffKey = selTiffEntry?.path || ''
          const existingSel = (selTiffKey && viewMetadataSelections[selTiffKey]) || {}
          const sel = { ...defaults, ...existingSel }

          if (sel['geolocation.qr_code'] === true) {
            const qrName = meta?.qr_png || meta?.categories?.maps?.qr_png
            let bytes: Uint8Array | null = null

            const categories: any = meta?.categories && typeof meta.categories === 'object' ? meta.categories : {}
            const geo = categories?.geolocation || {}
            const lat = geo?.latitude ?? meta?.summary?.Latitude
            const lon = geo?.longitude ?? meta?.summary?.Longitude

            if (typeof qrName === 'string' && qrName.trim()) {
              const metadataDir = await getWorkspaceThermalMetadataDir(false)
              const legacyFaultsDir = !metadataDir ? await getLegacyWorkflowFaultsDir(false) : null
              const legacyMetadataDir = legacyFaultsDir?.getDirectoryHandle
                ? await legacyFaultsDir.getDirectoryHandle('metadata', { create: false })
                : null

              const dir = metadataDir || legacyMetadataDir
              if (dir && dir.getFileHandle) {
                const fh = await dir.getFileHandle(qrName)
                const f = await fh.getFile()
                const buf = await f.arrayBuffer()
                bytes = new Uint8Array(buf)
              }
            } else if (typeof meta?.qr_png_base64 === 'string' && meta.qr_png_base64.length > 0) {
              bytes = Uint8Array.from(atob(meta.qr_png_base64), (c) => c.charCodeAt(0))
            }

            if (bytes) {
              qrBlock = { bytes, lat, lon }
            }
          }
        }
      } catch {
        // best-effort only
      }

      // One or more user-defined tables under the image (centered). No column header row.
      // IMPORTANT: Let Word handle pagination naturally.
      // If we pre-split into chunks and inject PageBreaks, Word can leave unused space on a page.
      // Also, ensure rows are allowed to break across pages (cantSplit=false).
      const cellBorders = {
        top: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
        left: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
        right: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
      }

      const tableCellMargins = {
        top: 80,
        bottom: 80,
        left: 140,
        right: 140,
      }

      const TABLE_TEXT_SIZE = 22 // ~11pt

      const makeDraftTable = (rowsToRender: Array<Pick<DocxDraftTableRow, 'name' | 'description'>>) =>
        new Table({
          width: { size: 86, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          alignment: AlignmentType.CENTER,
          borders: {
            top: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
            bottom: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
            left: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
            right: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
            insideVertical: { style: BorderStyle.SINGLE, size: 6, color: 'D1D5DB' },
          },
          rows: rowsToRender.map((r, rowIdx) => {
            const zebraFill = rowIdx % 2 === 0 ? 'FFFFFF' : 'F9FAFB'
            const nameFill = rowIdx % 2 === 0 ? 'F3F4F6' : 'EEF2F7'

            const nameText = (r.name || '').trim()
            const descText = (r.description || '').trim()

            return new TableRow({
              // Allow rows to break across pages (Word setting: "Allow row to break across pages").
              // This is key for preventing entire tables from being pushed to the next page.
              cantSplit: false,
              children: [
                new TableCell({
                  width: { size: 33, type: WidthType.PERCENTAGE },
                  borders: cellBorders,
                  margins: tableCellMargins,
                  verticalAlign: VerticalAlignTable.CENTER,
                  shading: { type: ShadingType.CLEAR, color: 'auto', fill: nameFill },
                  children: [
                    new Paragraph({
                      keepNext: false,
                      keepLines: false,
                      spacing: { before: 40, after: 40 },
                      children: [new TextRun({ text: nameText, bold: true, size: TABLE_TEXT_SIZE })],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 67, type: WidthType.PERCENTAGE },
                  borders: cellBorders,
                  margins: tableCellMargins,
                  verticalAlign: VerticalAlignTable.CENTER,
                  shading: { type: ShadingType.CLEAR, color: 'auto', fill: zebraFill },
                  children: [
                    new Paragraph({
                      keepNext: false,
                      keepLines: false,
                      spacing: { before: 40, after: 40 },
                      children: paragraphRunsFromMultilineSized(descText, TABLE_TEXT_SIZE),
                    }),
                  ],
                }),
              ],
            })
          }),
        })

      const pushTableTitle = (title: string) => {
        const t = (title || '').trim()
        if (!t) return

        children.push(
          new Paragraph({
            alignment: AlignmentType.CENTER,
            spacing: { before: 160, after: 80 },
            // Prevent orphan titles at the bottom of a page.
            // Keep the title with the next block (the table).
            keepNext: true,
            keepLines: true,
            children: [new TextRun({ text: t, bold: true })],
          })
        )
      }

      const tablesToRender = Array.isArray(item.tables) && item.tables.length ? item.tables : []

      // Table placement rules:
      // - If RGB was rendered, always start tables/QR on the next page (page is dedicated to the images).
      // - If NO RGB, allow tables/QR to start under the thermal image on the same page when there is space.
      if ((tablesToRender.length > 0 || qrBlock) && didRenderRgb) {
        children.push(new Paragraph({ children: [new PageBreak()] }))
      } else if (tablesToRender.length > 0 || qrBlock) {
        // Estimate remaining space after the thermal block.
        const thermalImageIn = h / 96
        const remainingIn = Math.max(0, CONTENT_HEIGHT_IN - (THERMAL_CAPTION_IN + CAPTION_GAP_IN + thermalImageIn + IMAGE_GAP_AFTER_IN))
        // If there's very little room left, don't strand a table at the bottom.
        if (remainingIn < 1.05) {
          children.push(new Paragraph({ children: [new PageBreak()] }))
        }
      }

      for (const t of tablesToRender) {
        const title = (t.title || '').trim()
        const rawRows = (t.rows || []).filter((r) => (r.name || '').trim() || (r.description || '').trim())
        const rows: Array<Pick<DocxDraftTableRow, 'name' | 'description'>> = rawRows.length ? rawRows : [{ name: '', description: '' }]

        pushTableTitle(title)
        children.push(makeDraftTable(rows))
        // Small gap after each table.
        children.push(new Paragraph({ children: [new TextRun({ text: '' })], spacing: { after: 140 } }))
      }

      // Optional QR code block (based on the View metadata checkbox state).
      if (qrBlock) {
        const coordText = (v: any) => {
          if (v === null || v === undefined || v === '') return 'N/A'
          if (typeof v === 'number') return Number.isFinite(v) ? String(v) : 'N/A'
          if (typeof v === 'string') return v.trim() || 'N/A'
          return String(v)
        }

        const QR_PX = 210

        children.push(
          new Paragraph({
            alignment: AlignmentType.LEFT,
            spacing: { before: 160, after: 80 },
            keepNext: true,
            keepLines: true,
            children: [new TextRun({ text: 'Geolocation', bold: true })],
          })
        )

        const geoCellBorders = {
          top: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
          bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
          left: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
          right: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
        }

        const geoTable = new Table({
          width: { size: 80, type: WidthType.PERCENTAGE },
          layout: TableLayoutType.FIXED,
          alignment: AlignmentType.CENTER,
          borders: {
            top: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
            bottom: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
            left: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
            right: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
            insideHorizontal: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
            insideVertical: { style: BorderStyle.SINGLE, size: 8, color: '000000' },
          },
          rows: [
            new TableRow({
              cantSplit: false,
              children: [
                new TableCell({
                  width: { size: 66, type: WidthType.PERCENTAGE },
                  borders: geoCellBorders,
                  margins: { top: 120, bottom: 120, left: 120, right: 120 },
                  verticalAlign: VerticalAlignTable.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      keepNext: false,
                      keepLines: false,
                      children: [
                        new ImageRun({
                          type: 'png',
                          data: qrBlock.bytes,
                          transformation: { width: QR_PX, height: QR_PX },
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 34, type: WidthType.PERCENTAGE },
                  borders: geoCellBorders,
                  margins: { top: 180, bottom: 180, left: 220, right: 220 },
                  verticalAlign: VerticalAlignTable.TOP,
                  children: [
                    new Paragraph({
                      keepNext: false,
                      keepLines: false,
                      spacing: { after: 120 },
                      children: [new TextRun({ text: 'Latitude', bold: true })],
                    }),
                    new Paragraph({
                      keepNext: false,
                      keepLines: false,
                      spacing: { after: 240 },
                      children: [new TextRun({ text: coordText(qrBlock.lat) })],
                    }),
                    new Paragraph({
                      keepNext: false,
                      keepLines: false,
                      spacing: { after: 120 },
                      children: [new TextRun({ text: 'Longitude', bold: true })],
                    }),
                    new Paragraph({
                      keepNext: false,
                      keepLines: false,
                      children: [new TextRun({ text: coordText(qrBlock.lon) })],
                    }),
                  ],
                }),
              ],
            }),
          ],
        })

        children.push(geoTable)
      }
    }

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

        ctx.globalAlpha = 0.18
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
    const tinyToposolFooterLogo = await fetchPublicImageRun('toposol-logo.png', 'png', 36, 36, true)

    const headerChildren: Paragraph[] = []

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

    const footerBordersNone = {
      top: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      bottom: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      left: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
      right: { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' },
    }

    const defaultFooter = new Footer({
      children: [
        new Table({
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
              children: [
                new TableCell({
                  width: { size: 30, type: WidthType.PERCENTAGE },
                  borders: footerBordersNone,
                  margins: { top: 40, bottom: 40, left: 0, right: 0 },
                  verticalAlign: VerticalAlignTable.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.LEFT,
                      children: tinyToposolFooterLogo ? [tinyToposolFooterLogo] : [],
                    }),
                  ],
                }),
                new TableCell({
                  width: { size: 70, type: WidthType.PERCENTAGE },
                  borders: footerBordersNone,
                  margins: { top: 40, bottom: 40, left: 0, right: 0 },
                  verticalAlign: VerticalAlignTable.CENTER,
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.RIGHT,
                      children: [new TextRun({ children: ['Page | ', PageNumber.CURRENT] })],
                    }),
                  ],
                }),
              ],
            }),
          ],
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
            default: defaultFooter,
          },
          children,
        },
      ],
    })

    return await Packer.toBlob(doc)
  }

  const previewWordReport = async (): Promise<Blob | null> => {
    // Ensure the latest in-progress draft edits are committed.
    try {
      const el = document.activeElement as HTMLElement | null
      el?.blur?.()
      await new Promise((r) => window.setTimeout(r, 0))
    } catch {
      // best-effort only
    }

    docxPreviewLastErrorRef.current = ''
    setDocxPreviewError('')
    setDocxPreviewStatus('Building preview…')

    const missingChapterTitles = reportChapters.filter((c) => !(c.chapterTitle || '').trim()).length
    if (missingChapterTitles > 0) {
      const msg = 'Please set a title for each added chapter.'
      docxPreviewLastErrorRef.current = msg
      setDocxPreviewError(msg)
      setDocxPreviewStatus('')
      return null
    }

    const host = docxPreviewHostRef.current
    if (!host) {
      const msg = 'Preview host is not available.'
      docxPreviewLastErrorRef.current = msg
      setDocxPreviewError(msg)
      setDocxPreviewStatus('')
      return null
    }

    try {
      const draftToUse = docxDraft.length ? docxDraft : buildDocxDraftFromPaths(normalizeFaultsText(reportText || faultsList.join('\n')), [])

      setDocxPreviewStatus('Generating DOCX…')
      const blob = await buildWordReportBlobFromDraft(draftToUse, { includeRgbPairs: includeRgbInDocx })

      setDocxPreviewStatus('Converting to PDF (LibreOffice)…')

      const form = new FormData()
      form.append('file', blob, 'report.docx')
      const resp = await fetch(apiUrl('/api/convert/docx-to-pdf'), {
        method: 'POST',
        body: form,
      })

      if (!resp.ok) {
        let detail = `${resp.status} ${resp.statusText}`
        try {
          const j = (await resp.json()) as any
          if (j && typeof j.detail === 'string') detail = j.detail
        } catch {
          try {
            const t = (await resp.text()).trim()
            if (t) detail = t
          } catch {
            // ignore
          }
        }
        throw new Error(detail)
      }

      const pdfBlob = await resp.blob()
      docxPreviewPdfBlobRef.current = pdfBlob
      const url = URL.createObjectURL(pdfBlob)
      if (docxPreviewPdfUrlRef.current) URL.revokeObjectURL(docxPreviewPdfUrlRef.current)
      docxPreviewPdfUrlRef.current = url

      host.innerHTML = ''
      const iframe = document.createElement('iframe')
      iframe.src = url
      iframe.title = 'PDF preview'
      iframe.style.width = '100%'
      iframe.style.height = '100%'
      iframe.style.border = 'none'
      host.appendChild(iframe)

      setDocxPreviewStatus('Preview updated.')
      docxPreviewLastErrorRef.current = ''
      return pdfBlob
    } catch (error) {
      const msg = error instanceof Error ? error.message : 'Failed to render DOCX preview'
      docxPreviewLastErrorRef.current = msg
      setDocxPreviewError(msg)
      setDocxPreviewStatus('')
      return null
    }
  }

  const downloadWordReport = async () => {
    // Ensure the latest in-progress draft edits are committed.
    try {
      const el = document.activeElement as HTMLElement | null
      el?.blur?.()
      await new Promise((r) => window.setTimeout(r, 0))
    } catch {
      // best-effort only
    }

    setReportError('')
    setReportStatus('Generating Word report…')

    const missingChapterTitles = reportChapters.filter((c) => !(c.chapterTitle || '').trim()).length
    if (missingChapterTitles > 0) {
      setReportError('Please set a title for each added chapter.')
      setReportStatus('')
      return
    }

    try {
      const draftToUse = docxDraft.length ? docxDraft : buildDocxDraftFromPaths(normalizeFaultsText(reportText || faultsList.join('\n')), [])
      const includedCount = draftToUse.filter((d) => d.include).length

      const blob = await buildWordReportBlobFromDraft(draftToUse, { includeRgbPairs: includeRgbInDocx })
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
    // Ensure the latest in-progress draft edits are committed.
    try {
      const el = document.activeElement as HTMLElement | null
      el?.blur?.()
      await new Promise((r) => window.setTimeout(r, 0))
    } catch {
      // best-effort only
    }

    setReportError('')
    setReportStatus('Generating PDF…')

    const missingChapterTitles = reportChapters.filter((c) => !(c.chapterTitle || '').trim()).length
    if (missingChapterTitles > 0) {
      setReportError('Please set a title for each added chapter.')
      setReportStatus('')
      return
    }

    try {
      const pdfBlob = await previewWordReport()
      if (!pdfBlob) throw new Error(docxPreviewLastErrorRef.current || 'Failed to generate PDF preview.')

      const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19)
      const filename = `faults-report-${stamp}.pdf`

      setReportStatus('Downloading PDF…')
      const url = URL.createObjectURL(pdfBlob)

      const ua = navigator.userAgent || ''
      const isIOS = /iPad|iPhone|iPod/.test(ua)
      const isSafari = /Safari\//.test(ua) && !/(Chrome|Chromium|Edg|OPR)\//.test(ua)

      // Safari/iOS commonly ignores programmatic downloads; open the PDF viewer instead.
      if (isIOS || isSafari) {
        const w = window.open(url, '_blank', 'noopener,noreferrer')
        if (!w) {
          setReportError('Popup blocked. Allow popups to open the generated PDF.')
        }
        window.setTimeout(() => URL.revokeObjectURL(url), 1500)
      } else {
        const a = document.createElement('a')
        a.href = url
        a.download = filename
        document.body.appendChild(a)
        a.click()
        a.remove()
        window.setTimeout(() => URL.revokeObjectURL(url), 1500)
      }
      setReportStatus('PDF downloaded.')
    } catch (error) {
      const msg = error instanceof Error ? error.message : 'Failed to generate PDF'
      setReportError(msg)
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

  const addReportChapter = () => {
    const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`
    const sectionId = `${Date.now()}-${Math.random().toString(16).slice(2)}`
    setReportChapters((prev) => [
      ...prev,
      {
        id,
        chapterTitle: '',
        sections: [{ id: sectionId, title: '', text: '', imageFile: null, imagePreviewUrl: null }],
      },
    ])
    setReportLeftTab(`chapter:${id}`)
  }

  const removeReportChapter = (id: string) => {
    const chapter = reportChapters.find((c) => c.id === id)
    const label = (chapter?.chapterTitle || '').trim() || 'this chapter'
    const ok = window.confirm(`Delete “${label}”?`)
    if (!ok) return

    setReportChapters((prev) => prev.filter((c) => c.id !== id))
    setReportLeftTab((prev) => (prev === `chapter:${id}` ? 'description' : prev))
  }

  const updateReportChapter = (id: string, patch: Partial<ReportCustomChapter>) => {
    setReportChapters((prev) => prev.map((c) => (c.id === id ? { ...c, ...patch } : c)))
  }

  const addReportChapterSection = (chapterId: string) => {
    const sectionId = `${Date.now()}-${Math.random().toString(16).slice(2)}`
    setReportChapters((prev) =>
      prev.map((c) =>
        c.id !== chapterId
          ? c
          : {
              ...c,
              sections: [...(Array.isArray(c.sections) ? c.sections : []), { id: sectionId, title: '', text: '', imageFile: null, imagePreviewUrl: null }],
            }
      )
    )
  }

  const removeReportChapterSection = (chapterId: string, sectionId: string) => {
    const ok = window.confirm('Delete this subchapter?')
    if (!ok) return

    setReportChapters((prev) =>
      prev.map((c) => {
        if (c.id !== chapterId) return c
        const nextSections = (Array.isArray(c.sections) ? c.sections : []).filter((s) => s.id !== sectionId)
        return { ...c, sections: nextSections.length ? nextSections : [{ id: `${Date.now()}-${Math.random().toString(16).slice(2)}`, title: '', text: '', imageFile: null, imagePreviewUrl: null }] }
      })
    )
  }

  const updateReportChapterSection = (chapterId: string, sectionId: string, patch: Partial<ReportCustomChapterSection>) => {
    setReportChapters((prev) =>
      prev.map((c) => {
        if (c.id !== chapterId) return c
        const nextSections = (Array.isArray(c.sections) ? c.sections : []).map((s) => (s.id === sectionId ? { ...s, ...patch } : s))
        return { ...c, sections: nextSections }
      })
    )
  }

  const setReportChapterSectionImage = async (chapterId: string, sectionId: string, file: File | null) => {
    if (!file) {
      updateReportChapterSection(chapterId, sectionId, { imageFile: null, imagePreviewUrl: null })
      return
    }
    const preview = await readFileAsDataUrl(file)
    updateReportChapterSection(chapterId, sectionId, { imageFile: file, imagePreviewUrl: preview })
  }

  const readLabelsForPath = async (path: string) => {
    try {
      const normalizedImagePath = normalizePath(String(path || ''))
      const baseName = normalizedImagePath.split('/').pop() || normalizedImagePath
      const stem = baseName.replace(/\.[^/.]+$/, '') || baseName

      // Backward-compat: some sessions write label files without the DJI `_T` / `_V` suffix.
      const baseNameNoTv = baseName.replace(/_[TV](\.[^.]+)$/i, '$1')
      const stemNoTv = stem.replace(/_[TV]$/i, '')

      // Backward-compat: some older sessions used '<imageNameWithExt>.txt' (e.g. DJI_0001.JPG.txt)
      // while newer writes use '<stem>.txt' (e.g. DJI_0001.txt).
      // Additionally, some datasets omit the trailing `_T` / `_V` in label filenames.
      const candidates = Array.from(
        new Set(
          [`${stem}.txt`, `${baseName}.txt`, stemNoTv && `${stemNoTv}.txt`, baseNameNoTv && `${baseNameNoTv}.txt`]
            .filter(Boolean)
            .map(String)
        )
      )

      // Fast-path: if labels exist in the opened tree, read them directly from `fileMap`.
      // This avoids brittle directory-handle traversal and guarantees rendering when the files are present.
      let labelText: string | null = null
      let matchedPath: string | null = null
      const tryReadTextFromFileMap = async (candidatePath: string) => {
        const resolved = resolvePathCaseInsensitive(normalizePath(candidatePath))
        const file = resolved ? fileMap[resolved] : null
        if (!file) return null
        try {
          return await file.text()
        } catch {
          return null
        }
      }

      const imageParts = normalizedImagePath.split('/').filter(Boolean)
      const thermalIdx = imageParts.findIndex((p) => p.toLowerCase() === 'thermal')
      let sessionPrefix = thermalIdx >= 0 ? imageParts.slice(0, thermalIdx).join('/') : ''
      if ((!sessionPrefix || thermalIdx === 0) && effectiveThermalFolderPath) {
        const effParts = effectiveThermalFolderPath.split('/').filter(Boolean)
        const effThermalIdx = effParts.findIndex((p) => p.toLowerCase() === 'thermal')
        if (effThermalIdx > 0) {
          sessionPrefix = effParts.slice(0, effThermalIdx).join('/')
        }
      }
      const sessionPrefixWithSlash = sessionPrefix ? `${sessionPrefix}/` : ''

      const relUnderThermalDir = thermalIdx >= 0 ? imageParts.slice(thermalIdx + 1, -1).join('/') : ''
      const imageParentPath = getParentDirPath(normalizedImagePath)

      // Support alternate layouts where `workspace/` is stored under `tif/` or `tiff/`.
      const workspaceRoots = [
        `${sessionPrefixWithSlash}${WORKSPACE_DIR}/${WORKSPACE_THERMAL_LABELS_DIR}`,
        `${sessionPrefixWithSlash}tif/${WORKSPACE_DIR}/${WORKSPACE_THERMAL_LABELS_DIR}`,
        `${sessionPrefixWithSlash}tiff/${WORKSPACE_DIR}/${WORKSPACE_THERMAL_LABELS_DIR}`,
      ]
      const legacyRoot = `${sessionPrefixWithSlash}thermal/faults/labels`

      const fileMapCandidates: string[] = []
      for (const name of candidates) {
        for (const root of workspaceRoots) {
          fileMapCandidates.push(`${root}/${name}`)
          if (relUnderThermalDir) fileMapCandidates.push(`${root}/${relUnderThermalDir}/${name}`)
        }
        fileMapCandidates.push(`${legacyRoot}/${name}`)
        fileMapCandidates.push(`${imageParentPath}/${name}`)
        if (relUnderThermalDir) {
          fileMapCandidates.push(`${legacyRoot}/${relUnderThermalDir}/${name}`)
        }
      }

      for (const candidatePath of fileMapCandidates) {
        const text = await tryReadTextFromFileMap(candidatePath)
        if (text && text.trim()) {
          labelText = text
          matchedPath = resolvePathCaseInsensitive(normalizePath(candidatePath)) || normalizePath(candidatePath)
          break
        }
      }

      // Only attempt directory-handle traversal if the file isn't already present in `fileMap`.
      // (Some environments/sessions may not allow resolving sibling workspace directories reliably.)
      const workspaceLabelsDir = !labelText
        ? await getWorkspaceThermalLabelsDirForPath(normalizedImagePath, false).catch(() => null)
        : null

      const legacyLabelsDirByPath = !labelText
        ? await (async () => {
            if (!folderHandle || !folderHandle.getDirectoryHandle) return null
            const parts = normalizedImagePath.split('/').filter(Boolean)
            const thermalIdxLocal = parts.findIndex((p) => p.toLowerCase() === 'thermal')
            if (thermalIdxLocal < 0) return null

            let thermalDir: FileSystemDirectoryHandle = folderHandle
            for (const part of parts.slice(0, thermalIdxLocal + 1)) {
              if (!thermalDir.getDirectoryHandle) return null
              thermalDir = await thermalDir.getDirectoryHandle(part, { create: false })
            }

            if (!thermalDir.getDirectoryHandle) return null
            const faults = await thermalDir.getDirectoryHandle('faults', { create: false })
            if (!faults?.getDirectoryHandle) return null
            return faults.getDirectoryHandle('labels', { create: false })
          })().catch(() => null)
        : null

      const legacyLabelsDirGlobal = !labelText
        ? await (async () => {
            const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
            if (!legacyFaultsDir?.getDirectoryHandle) return null
            try {
              return await legacyFaultsDir.getDirectoryHandle('labels', { create: false })
            } catch {
              return null
            }
          })().catch(() => null)
        : null

      const legacyLabelsDir = !labelText ? (legacyLabelsDirByPath || legacyLabelsDirGlobal) : null

      const readTextFileFlexible = async (dir: FileSystemDirectoryHandle, name: string) => {
        const direct = await readTextFile(dir, name)
        if (direct !== null) return direct

        // Case-insensitive fallback: scan directory entries for a matching filename.
        try {
          if (!dir.getFileHandle) return null
          const target = name.toLowerCase()
          const entriesFn = (dir as any).entries as undefined | (() => AsyncIterableIterator<[string, any]>)
          if (!entriesFn) return null
          for await (const [entryName, entry] of entriesFn.call(dir)) {
            if (!entry || entry.kind !== 'file') continue
            if (String(entryName).toLowerCase() !== target) continue
            const fh = await dir.getFileHandle(entryName)
            const f = await fh.getFile()
            return await f.text()
          }
        } catch {
          // ignore
        }

        return null
      }

      const tryGetSubdir = async (root: FileSystemDirectoryHandle, relPath: string) => {
        if (!relPath) return root
        if (!root.getDirectoryHandle) return null
        const parts = relPath.split('/').filter(Boolean)
        let dir: FileSystemDirectoryHandle = root
        for (const part of parts) {
          if (!dir.getDirectoryHandle) return null
          try {
            dir = await dir.getDirectoryHandle(part, { create: false })
          } catch {
            return null
          }
        }
        return dir
      }

      const thermalRoot = effectiveThermalFolderPath
      const parentPath = getParentDirPath(path)
      const inferRelUnderThermal = (p: string) => {
        const normalized = String(p || '')
        if (!normalized) return ''

        if (thermalRoot && (normalized === thermalRoot || normalized.startsWith(`${thermalRoot}/`))) {
          return normalized.slice(thermalRoot.length).replace(/^\//, '')
        }

        // If we don't have a detected thermal root yet, but paths are still prefixed
        // with a top-level `thermal/`, strip it so nested label folders match.
        const parts = normalized.split('/').filter(Boolean)
        if (parts.length > 0 && parts[0].toLowerCase() === 'thermal') {
          return parts.slice(1).join('/')
        }

        return normalized
      }

      const relDir = parentPath ? inferRelUnderThermal(parentPath) : ''

      if (!labelText) {
        const labelRoots = [workspaceLabelsDir, legacyLabelsDir].filter(Boolean) as FileSystemDirectoryHandle[]
        for (const root of labelRoots) {
          // 1) Try flat storage at the root.
          for (const name of candidates) {
            labelText = await readTextFileFlexible(root, name)
            if (labelText) break
          }
          if (labelText) break

          // 2) Try nested storage mirroring the thermal folder structure.
          if (relDir) {
            const nested = await tryGetSubdir(root, relDir)
            if (nested) {
              for (const name of candidates) {
                labelText = await readTextFileFlexible(nested, name)
                if (labelText) break
              }
            }
          }
          if (labelText) break
        }

        // Final fallback: some datasets store labels alongside the image file itself.
        // (Read-only; we do not write into `thermal/`.)
        if (!labelText) {
          const workflowRoot = await getWorkflowRootDir(false)
          if (workflowRoot) {
            const imageRelPath = parentPath ? inferRelUnderThermal(parentPath) : ''
            const imageDir = await tryGetSubdir(workflowRoot, imageRelPath)
            if (imageDir) {
              for (const name of candidates) {
                labelText = await readTextFileFlexible(imageDir, name)
                if (labelText) break
              }
            }
          }
        }
      }

      if (!labelText) {
        // Diagnostics: helps confirm we are looking in the correct nested workspace.
        // (Open DevTools console to view.)
        console.warn('[labels] not found', {
          imagePath: normalizedImagePath,
          effectiveThermalFolderPath,
          sessionPrefix,
          candidates,
          fileMapCandidateCount: fileMapCandidates.length,
          fileMapCandidateSample: fileMapCandidates.slice(0, 10),
          hasWorkspaceLabelsDir: Boolean(workspaceLabelsDir),
          hasLegacyLabelsDir: Boolean(legacyLabelsDir),
          relDir,
        })
        return []
      }

      if (matchedPath) {
        console.debug('[labels] loaded', { imagePath: normalizedImagePath, labelPath: matchedPath })
      }
      const raw = labelText
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean)
        .map((line) => {
          const [classId, x, y, w, h, conf] = line.split(/\s+/).map(parseYoloNumber)
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

      // If labels are stored in absolute pixel coords (legacy), normalize them
      // to YOLO-style [0..1] using the context image dimensions so they render
      // correctly on both Thermal and RGB views.
      const needsNormalize = raw.some((l) => l.x > 1 || l.y > 1 || l.w > 1 || l.h > 1)
      if (!needsNormalize) return raw

      const file = fileMap[path]
      if (!file) return raw

      const { width, height } = await getImageNaturalSizeFromFile(file)
      if (!(width > 0 && height > 0)) return raw

      const clamp01 = (v: number) => Math.min(1, Math.max(0, v))
      return raw.map((l) => ({
        ...l,
        x: clamp01(l.x / width),
        y: clamp01(l.y / height),
        w: clamp01(l.w / width),
        h: clamp01(l.h / height),
      }))
    } catch {
      return []
    }
  }

  // When labels are generated externally (scan/backend) after the image is already open,
  // re-check and load them automatically. To avoid overriding manual edits, only do
  // this when the user hasn't edited labels yet (history still has the initial empty snapshot).
  useEffect(() => {
    if (!selectedPath) return
    if (!folderHandle) return
    if (fileKind !== 'image') return
    if (selectedLabels.length > 0) return
    if (labelHistoryIndex !== 0) return
    if (labelHistory.length !== 1) return
    if ((labelHistory[0] || []).length !== 0) return

    let cancelled = false
    void (async () => {
      const labels = await readLabelsForPath(selectedPath)
      if (cancelled) return
      if (!labels || labels.length === 0) return
      setSelectedLabels(labels)
      setSelectedLabelIndex(null)
      setLabelHistory([labels])
      setLabelHistoryIndex(0)
    })()

    return () => {
      cancelled = true
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [fileKind, folderHandle, labelHistory, labelHistoryIndex, selectedLabels.length, selectedPath, workspaceIndexTick])

  const readRgbLabelsForRgbPath = async (rgbPath: string) => {
    try {
      const labelsDir = (await getRgbLabelsDir(false)) || (await getLegacyRgbLabelsDir(false))
      if (!labelsDir) return []

      const stem = rgbPath.split('/').pop()?.replace(/\.[^/.]+$/, '') || rgbPath
      const labelText = await readTextFile(labelsDir, `${stem}.txt`)
      if (!labelText) return []

      const raw = parseYoloLabelsText(labelText)
      if (raw.length === 0) return []

      // If labels are stored in absolute px coords, normalize first.
      const needsNormalize = raw.some((l) => l.x > 1 || l.y > 1 || l.w > 1 || l.h > 1)
      const rgbFile = fileMap[rgbPath]
      if (!rgbFile) return raw

      const { width, height } = await getImageNaturalSizeFromFile(rgbFile)
      if (!(width > 0 && height > 0)) return raw

      const clamp01Local = (v: number) => Math.min(1, Math.max(0, v))
      const normalized = needsNormalize
        ? raw.map((l) => ({
            ...l,
            x: clamp01Local(l.x / width),
            y: clamp01Local(l.y / height),
            w: clamp01Local(l.w / width),
            h: clamp01Local(l.h / height),
          }))
        : raw

      // Convert crop-relative coords (current RGB label file format) into full-image
      // coords for correct overlay when rendering the full RGB image in the report.
      const crop = getDefaultRgbCropRectNormalizedFromImageDims(width, height)
      if (!crop) return normalized

      // Backward-compat heuristic: older RGB label files may already be full-image normalized.
      const x0 = crop.x0
      const y0 = crop.y0
      const x1 = crop.x0 + crop.scaleX
      const y1 = crop.y0 + crop.scaleY
      const candidates = normalized.filter((l) => Number.isFinite(l.x) && Number.isFinite(l.y))
      const inCrop = candidates.filter((l) => l.x >= x0 && l.x <= x1 && l.y >= y0 && l.y <= y1).length
      const ratio = candidates.length ? inCrop / candidates.length : 0
      const treatAsFullImage = ratio >= 0.9

      return treatAsFullImage ? normalized : convertRgbCropLabelsToFullNormalized(normalized, crop)
    } catch {
      return []
    }
  }

  const readMetadataForImagePath = async (imagePath: string) => {
    try {
      const metadataDir = await getWorkspaceThermalMetadataDir(false)
      const legacyFaultsDir = !metadataDir ? await getLegacyWorkflowFaultsDir(false) : null
      const legacyMetadataDir = legacyFaultsDir?.getDirectoryHandle
        ? await legacyFaultsDir.getDirectoryHandle('metadata', { create: false })
        : null

      const dir = metadataDir || legacyMetadataDir
      if (!dir) return null
      const imageName = imagePath.split('/').pop() || imagePath
      const jsonName = `${imageName}.json`
      return await readJsonFile<any>(dir, jsonName)
    } catch {
      return null
    }
  }

  const isPlaceholderEquipmentItems = (items: ReportEquipmentItem[]) => {
    if (!Array.isArray(items) || items.length !== 1) return false
    const it = items[0]
    const titleEmpty = !(it.title || '').trim()
    const textEmpty = !(it.text || '').trim()
    return titleEmpty && textEmpty && !it.imageFile
  }

  const equipmentDefaultsAppliedRef = useRef(false)

  useEffect(() => {
    if (activeAction !== 'report') return
    if (equipmentDefaultsAppliedRef.current) return
    if (!docxDraft.length) return
    if (!isPlaceholderEquipmentItems(equipmentItems)) return

    equipmentDefaultsAppliedRef.current = true

    void (async () => {
      try {
        // If any image metadata reports M3T camera model, prefill equipment.
        let hasM3T = false
        for (const item of docxDraft) {
          if (item.include === false) continue
          const meta = await readMetadataForImagePath(item.path)
          if (!meta) continue

          const categories = meta?.categories && typeof meta.categories === 'object' ? meta.categories : {}
          const device = categories?.device || {}
          const modelRaw = (device?.Model ?? meta?.summary?.Model ?? '').toString().trim()
          if (modelRaw.toUpperCase() === 'M3T') {
            hasM3T = true
            break
          }
        }

        if (!hasM3T) return

        // Load the default equipment image from public assets.
        const base = import.meta.env.BASE_URL || '/'
        const url = `${base}m3t.png`
        const res = await fetch(url)
        if (!res.ok) return
        const blob = await res.blob()
        const file = new File([blob], 'm3t.png', { type: blob.type || 'image/png' })
        const preview = URL.createObjectURL(blob)

        const id = `${Date.now()}-${Math.random().toString(16).slice(2)}`
        setEquipmentItems([
          {
            id,
            title: 'Mavic 3T',
            text: '',
            imageFile: file,
            imagePreviewUrl: preview,
          },
        ])
      } catch {
        // Best-effort only.
      }
    })()
  }, [activeAction, docxDraft, equipmentItems])

  type MetadataValueKind = 'default' | 'temp' | 'distance' | 'percent' | 'speed' | 'irradiance'

  const formatMetadataValue = (v: any, kind: MetadataValueKind = 'default') => {
    if (v === null || v === undefined || v === '') return 'N/A'

    const num = () => {
      if (typeof v === 'number') return Number.isFinite(v) ? v : null
      if (typeof v === 'string') {
        const m = v.replace(',', '.').match(/-?\d+(?:\.\d+)?/)
        if (!m) return null
        const n = Number(m[0])
        return Number.isFinite(n) ? n : null
      }
      return null
    }

    if (kind === 'temp') {
      const n = num()
      return n === null ? 'N/A' : `${n.toFixed(2)} °C`
    }
    if (kind === 'distance') {
      const n = num()
      return n === null ? 'N/A' : `${n} m`
    }
    if (kind === 'percent') {
      const n = num()
      return n === null ? 'N/A' : `${n} %`
    }
    if (kind === 'speed') {
      const n = num()
      return n === null ? 'N/A' : `${n} m/s`
    }
    if (kind === 'irradiance') {
      const n = num()
      return n === null ? 'N/A' : `${n} W/m²`
    }

    if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v)
    try {
      return JSON.stringify(v)
    } catch {
      return String(v)
    }
  }

  const buildMetadataTablesForReport = async (
    imagePath: string,
    options?: { selectionOverride?: Record<string, boolean>; overridesOverride?: Record<string, any> }
  ): Promise<DocxDraftTable[] | null> => {
    const report = await readMetadataForImagePath(imagePath)
    if (!report) return null

    const categories: any = report?.categories && typeof report.categories === 'object' ? report.categories : {}

    const tiffEntry = findMatchingTiffForImage(imagePath)
    const tiffKey = tiffEntry?.path || ''

    const defaults = computeDefaultMetadataSelections(report, imagePath)
    const existingSelection = options?.selectionOverride ?? ((tiffKey && viewMetadataSelections[tiffKey]) || {})
    const selection: Record<string, boolean> = { ...defaults, ...existingSelection }
    const overrides = options?.overridesOverride ?? ((tiffKey && viewMetadataOverrides[tiffKey]) || {})

    const getEffective = (id: string, fallback: any) => (overrides && overrides[id] !== undefined ? overrides[id] : fallback)
    const isChecked = (id: string) => selection[id] === true

    const imageName = imagePath.split('/').pop() || imagePath
    const device = categories?.device || {}
    const flight = categories?.flight || {}
    const imageInfo = categories?.image_info || {}
    const timestamps = categories?.timestamps || {}
    const pixelStats = categories?.measurement_temperatures?.pixel_stats || {}
    const params = categories?.measurement_params || {}
    const geo = categories?.geolocation || {}

    const lat = geo?.latitude ?? report?.summary?.Latitude
    const lon = geo?.longitude ?? report?.summary?.Longitude

    const derivedImageInfo: Record<string, any> = {
      file_name: imageName,
      tiff_file: report?.file,
      camera_model: device?.Model ?? report?.summary?.Model,
      serial_number:
        report?.exiftool_meta?.['ExifIFD:SerialNumber'] ??
        report?.pillow_exif?.['ExifIFD:SerialNumber'] ??
        device?.SerialNumber ??
        report?.summary?.CameraSerialNumber,
      focal_length: device?.FocalLength ?? report?.summary?.FocalLength,
      f_number: device?.FNumber ?? report?.summary?.FNumber,
      width: imageInfo?.ImageWidth ?? report?.summary?.ImageWidth,
      height: imageInfo?.ImageHeight ?? report?.summary?.ImageHeight,
      timestamp_created: timestamps?.CreateDate ?? timestamps?.DateTimeOriginal ?? report?.summary?.DateTimeOriginal ?? report?.summary?.DateTime,
      latitude: lat,
      longitude: lon,
    }

    type RowModel = { id: string; label: string; value: any; kind: MetadataValueKind }
    const row = (id: string, label: string, value: any, kind: MetadataValueKind = 'default'): RowModel => ({ id, label, value, kind })

    const newId = () => `${Date.now()}-${Math.random().toString(16).slice(2)}`

    const faultLabels = await readLabelsForPath(imagePath)

    const titleCaseKey = (s: string) => String(s).replace(/_/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase())

    // Best-effort: compute per-label temperature stats once, and enrich the Fault labels table.
    const labelTempsByIndex = new Map<number, any>()
    if (tiffEntry && faultLabels.length) {
      try {
        const formData = new FormData()
        formData.append('file', tiffEntry.file, tiffEntry.path)
        formData.append(
          'labels',
          JSON.stringify(
            faultLabels.map((l, idx) => ({
              index: idx,
              classId: l.classId,
              x: l.x,
              y: l.y,
              w: l.w,
              h: l.h,
              conf: l.conf,
              shape: 'rect',
              source: 'auto',
            }))
          )
        )
        const response = await fetch(apiUrl('/api/temperatures/labels?pad_px=3'), { method: 'POST', body: formData })
        if (response.ok) {
          const data = await response.json()
          const labelsOut = Array.isArray(data?.labels) ? data.labels : []
          for (const item of labelsOut) {
            const idx = Number(item?.index)
            if (!Number.isFinite(idx)) continue
            labelTempsByIndex.set(idx, item)
          }
        }
      } catch {
        // best-effort only
      }
    }

    const faultLabelRows: RowModel[] = faultLabels.map((l, idx) => {
      const title = `Fault ${idx + 1}`
      const f = (n: any) => (typeof n === 'number' && Number.isFinite(n) ? n.toFixed(4) : '0')
      const conf = Number.isFinite(Number(l?.conf))
        ? (Number(l.conf) >= 0 && Number(l.conf) <= 1 ? ` — ${Math.round(Number(l.conf) * 100)}%` : ` — ${String(l.conf)}`)
        : ''

      const temps = labelTempsByIndex.get(idx)
      const tempSummary = temps
        ? ` | edge=${formatMetadataValue(temps?.outside_edge_mean, 'temp')}, avg=${formatMetadataValue(temps?.inside_mean, 'temp')}, min=${formatMetadataValue(temps?.inside_min, 'temp')}, max=${formatMetadataValue(temps?.inside_max, 'temp')}`
        : ''

      const summary = `x=${f(l?.x)}, y=${f(l?.y)}, w=${f(l?.w)}, h=${f(l?.h)}${conf}${tempSummary}`
      return row(`fault_labels.${idx}.summary`, title, summary)
    })

    // Default-check fault label rows (new category) unless the user explicitly changed them.
    for (let idx = 0; idx < faultLabelRows.length; idx += 1) {
      const k = faultLabelRows[idx].id
      if (selection[k] === undefined) selection[k] = true
    }

    // Build the requested structure:
    // Table = "Faults" (summary), Subtable = "Fault 1" ... with fields for avg outside/avg inside/min/max inside.
    // Default-check these temperature fields when we have values, without overriding user choices.
    const faultsTables: DocxDraftTable[] = []
    if (faultLabels.length) {
      const includedFaultIdx: number[] = []
      const faultSummaryRows: DocxDraftTableRow[] = []

      for (let idx = 0; idx < faultLabels.length; idx += 1) {
        const faultKey = `fault_labels.${idx}.summary`
        if (selection[faultKey] === undefined) selection[faultKey] = true
        if (!isChecked(faultKey)) continue

        includedFaultIdx.push(idx)
        faultSummaryRows.push({
          id: newId(),
          name: `Fault ${idx + 1}`,
          description: `Possible ${getFaultTypeLabel(faultLabels[idx]?.classId)}`,
        })
      }

      if (faultSummaryRows.length) {
        faultsTables.push({ id: newId(), title: 'Faults', rows: faultSummaryRows })
      }

      for (const idx of includedFaultIdx) {
        const temps = labelTempsByIndex.get(idx)

        const title = `Fault ${idx + 1}`
        const fields: Array<{ key: string; label: string; value: any }> = [
          { key: `label_temperatures.${idx}.outside_edge_mean`, label: `${title} — Outside edge avg`, value: temps?.outside_edge_mean },
          { key: `label_temperatures.${idx}.inside_mean`, label: `${title} — Inside avg`, value: temps?.inside_mean },
          { key: `label_temperatures.${idx}.inside_min`, label: `${title} — Inside min`, value: temps?.inside_min },
          { key: `label_temperatures.${idx}.inside_max`, label: `${title} — Inside max`, value: temps?.inside_max },
        ]

        for (const f of fields) {
          if (selection[f.key] === undefined && temps) selection[f.key] = true
        }

        const rows: DocxDraftTableRow[] = []
        for (const f of fields) {
          if (!isChecked(f.key)) continue
          rows.push({ id: newId(), name: f.label, description: formatMetadataValue(f.value, 'temp') })
        }

        if (rows.length) {
          faultsTables.push({ id: newId(), title, rows })
        }
      }
    }

    const categoryModels: Array<{ title: string; rows: RowModel[] }> = [
      {
        title: 'Temperature measurements',
        rows: [
          row('measurement_temperatures.pixel_stats.min', 'Minimum', pixelStats?.min, 'temp'),
          row('measurement_temperatures.pixel_stats.mean', 'Average', pixelStats?.mean, 'temp'),
          row('measurement_temperatures.pixel_stats.max', 'Maximum', pixelStats?.max, 'temp'),
          row('measurement_temperatures.pixel_stats.looks_like_celsius_guess', 'Looks like °C', pixelStats?.looks_like_celsius_guess),
        ].filter((r) => r.value !== undefined),
      },
      {
        title: 'Thermal parameters',
        rows: [
          row('measurement_params.Distance', 'Distance', params?.Distance, 'distance'),
          row('measurement_params.RelativeHumidity', 'Relative humidity', params?.RelativeHumidity, 'percent'),
          row('measurement_params.Emissivity', 'Emissivity', params?.Emissivity),
          row('measurement_params.AmbientTemperature', 'Ambient temperature', params?.AmbientTemperature, 'temp'),
          row('measurement_params.WindSpeed', 'Wind speed', params?.WindSpeed, 'speed'),
          row('measurement_params.Irradiance', 'Irradiance', params?.Irradiance, 'irradiance'),
        ].filter((r) => r.value !== undefined),
      },
      {
        title: 'Flight & gimbal',
        rows: Object.entries(flight && typeof flight === 'object' ? flight : {}).map(([k, v]) =>
          row(`flight.${k}`, titleCaseKey(k), v)
        ),
      },
      {
        title: 'Image info',
        rows: [
          row('image.file_name', 'File name', derivedImageInfo.file_name),
          row('image.tiff_file', 'TIFF file', derivedImageInfo.tiff_file),
          row('image.camera_model', 'Camera model', derivedImageInfo.camera_model),
          row('image.serial_number', 'Serial number', derivedImageInfo.serial_number),
          row('image.focal_length', 'Focal length', derivedImageInfo.focal_length),
          row('image.f_number', 'F-number', derivedImageInfo.f_number),
          row('image.width', 'Width', derivedImageInfo.width),
          row('image.height', 'Height', derivedImageInfo.height),
          row('image.timestamp_created', 'Timestamp created', derivedImageInfo.timestamp_created),
          row('image.latitude', 'Latitude', derivedImageInfo.latitude),
          row('image.longitude', 'Longitude', derivedImageInfo.longitude),
        ].filter((r) => r.value !== undefined),
      },
    ]

    const tables: DocxDraftTable[] = []
    if (faultsTables.length) tables.push(...faultsTables)
    for (const cat of categoryModels) {
      const checkedRows = cat.rows
        .filter((r) => isChecked(r.id))
        .map((r) => {
          const effective = getEffective(r.id, r.value)
          return {
            id: newId(),
            name: r.label,
            description: formatMetadataValue(effective, r.kind),
          }
        })

      if (!checkedRows.length) continue
      tables.push({ id: newId(), title: cat.title, rows: checkedRows })
    }

    return tables.length ? tables : null
  }

  const reportMetadataSyncTimersRef = useRef<Map<string, number>>(new Map())

  const isSystemReportTableTitle = (title: string) => {
    const t = (title || '').trim()
    if (!t) return false
    if (t === 'Faults') return true
    if (/^Fault\s+\d+$/i.test(t)) return true
    if (t === 'Temperature measurements') return true
    if (t === 'Thermal parameters') return true
    if (t === 'Flight & gimbal') return true
    if (t === 'Image info') return true
    return false
  }

  const requestReportMetadataSyncForImage = (
    imagePath: string,
    tiffKey: string,
    selectionOverride?: Record<string, boolean>,
    overridesOverride?: Record<string, any>
  ) => {
    if (!imagePath || !tiffKey) return

    const existing = reportMetadataSyncTimersRef.current.get(tiffKey)
    if (existing !== undefined) window.clearTimeout(existing)

    const timer = window.setTimeout(() => {
      reportMetadataSyncTimersRef.current.delete(tiffKey)

      void (async () => {
        const tables = await buildMetadataTablesForReport(imagePath, { selectionOverride, overridesOverride })
        reportMetadataDefaultsAppliedRef.current.add(imagePath)
        reportFaultTablesAppliedRef.current.add(imagePath)

        setDocxDraft((prev) => {
          if (!prev.length) return prev

          const signatureForTables = (tables: DocxDraftTable[]) => {
            const parts: string[] = []
            for (const t of tables) {
              parts.push(`T:${(t?.title || '').trim()}`)
              const rows = Array.isArray(t?.rows) ? t.rows : []
              for (const r of rows) {
                parts.push(`R:${(r?.name || '').trim()}=${(r?.description || '').trim()}`)
              }
            }
            return parts.join('\n')
          }

          let anyChanged = false
          const next = prev.map((item) => {
            if (item.path !== imagePath) return item

            const existingTables = Array.isArray(item.tables) ? item.tables : []
            const kept = existingTables.filter((t) => !isSystemReportTableTitle(t?.title || ''))
            const systemTables = Array.isArray(tables) ? tables : []
            const nextTables = [...systemTables, ...kept]

            const same = signatureForTables(nextTables) === signatureForTables(existingTables)
            if (same) return item
            anyChanged = true
            return { ...item, tables: nextTables }
          })

          return anyChanged ? next : prev
        })
      })().catch(() => undefined)
    }, 180)

    reportMetadataSyncTimersRef.current.set(tiffKey, timer)
  }

  const refreshReportMetadataTables = async () => {
    if (activeAction !== 'report') return
    if (!docxDraft.length) return

    setReportError('')

    const included = docxDraft.filter((d) => d.include !== false)
    if (!included.length) {
      setReportStatus('No included items to refresh.')
      return
    }

    setReportStatus('Refreshing metadata…')

    try {
      const updates = new Map<string, DocxDraftTable[]>()

      for (let i = 0; i < included.length; i += 1) {
        const item = included[i]
        setReportStatus(`Refreshing metadata… ${i + 1} / ${included.length}`)

        const entry = findMatchingTiffForImage(item.path)
        const tiffKey = entry?.path || ''
        if (!tiffKey) continue

        const selection = viewMetadataSelections[tiffKey]
        const overrides = viewMetadataOverrides[tiffKey]

        const tables = await buildMetadataTablesForReport(item.path, {
          selectionOverride: selection,
          overridesOverride: overrides,
        })
        if (!tables || !tables.length) continue
        updates.set(item.id, tables)

        // Prevent the placeholder-injection effect from fighting us.
        reportMetadataDefaultsAppliedRef.current.add(item.path)
        reportFaultTablesAppliedRef.current.add(item.path)
      }

      if (!updates.size) {
        setReportStatus('Metadata refreshed (no tables changed).')
        return
      }

      setDocxDraft((prev) =>
        prev.map((item) => {
          const systemTables = updates.get(item.id)
          if (!systemTables) return item

          const existing = Array.isArray(item.tables) ? item.tables : []
          const kept = existing.filter((t) => !isSystemReportTableTitle(t?.title || ''))
          return { ...item, tables: [...systemTables, ...kept] }
        })
      )

      setReportStatus('Metadata refreshed.')
    } catch (error) {
      setReportError(error instanceof Error ? error.message : 'Failed to refresh metadata')
      setReportStatus('')
    }
  }

  const reportAutoRefreshSigRef = useRef<string>('')

  const reportAutoRefreshSig = useMemo(() => {
    if (!docxDraft.length) return ''
    return docxDraft.map((d) => `${d.id}:${d.include ? 1 : 0}`).join('|')
  }, [docxDraft])

  useEffect(() => {
    if (activeAction !== 'report') return
    if (!docxDraft.length) return

    if (!reportAutoRefreshSig) return
    if (reportAutoRefreshSigRef.current === reportAutoRefreshSig) return
    reportAutoRefreshSigRef.current = reportAutoRefreshSig

    // Auto-refresh system metadata tables when entering Report or when a new draft is created.
    void refreshReportMetadataTables()
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeAction, reportAutoRefreshSig])

  const isPlaceholderDraftTables = (tables: DocxDraftItem['tables']) => {
    if (!Array.isArray(tables) || tables.length !== 1) return false
    const t = tables[0]
    const titleEmpty = !(t.title || '').trim()
    const rows = Array.isArray(t.rows) ? t.rows : []
    if (!titleEmpty || rows.length !== 1) return false
    const r = rows[0]
    return !((r?.name || '').trim() || (r?.description || '').trim())
  }

  const isEmptyOrPlaceholderDraftTables = (tables: DocxDraftItem['tables']) => {
    if (!Array.isArray(tables) || !tables.length) return true
    return isPlaceholderDraftTables(tables)
  }

  const hasFaultsTables = (tables: DocxDraftItem['tables']) => {
    if (!Array.isArray(tables) || !tables.length) return false
    const re = /^Fault\s+\d+$/i
    return tables.some((t) => {
      const title = (t?.title || '').trim()
      return title === 'Faults' || re.test(title)
    })
  }

  const reportMetadataDefaultsAppliedRef = useRef<Set<string>>(new Set())
  const reportFaultTablesAppliedRef = useRef<Set<string>>(new Set())
  const reportMetadataDefaultsInFlightRef = useRef<Promise<void> | null>(null)

  useEffect(() => {
    if (activeAction !== 'report') return
    if (!docxDraft.length) return
    if (reportMetadataDefaultsInFlightRef.current) return

    const pending = docxDraft.filter((item) => {
      if (item.include === false) return false

      const needsDefaults =
        !reportMetadataDefaultsAppliedRef.current.has(item.path) &&
        isEmptyOrPlaceholderDraftTables(item.tables)

      const needsFaults =
        !reportFaultTablesAppliedRef.current.has(item.path) &&
        !hasFaultsTables(item.tables)

      return needsDefaults || needsFaults
    })

    if (!pending.length) return

    reportMetadataDefaultsInFlightRef.current = (async () => {
      const replaceUpdates = new Map<string, DocxDraftTable[]>()
      const prependFaultsUpdates = new Map<string, DocxDraftTable[]>()
      for (const item of pending) {
        const needsDefaults =
          !reportMetadataDefaultsAppliedRef.current.has(item.path) &&
          isEmptyOrPlaceholderDraftTables(item.tables)

        const needsFaults =
          !reportFaultTablesAppliedRef.current.has(item.path) &&
          !hasFaultsTables(item.tables)

        // Mark as applied even if we can't load anything, to avoid retry loops.
        if (needsDefaults) reportMetadataDefaultsAppliedRef.current.add(item.path)

        const tables = await buildMetadataTablesForReport(item.path)
        if (!tables || !tables.length) continue

        if (needsDefaults) {
          replaceUpdates.set(item.id, tables)
          continue
        }

        if (needsFaults) {
          const re = /^Fault\s+\d+$/i
          const faultTables = tables.filter((t) => {
            const title = (t?.title || '').trim()
            return title === 'Faults' || re.test(title)
          })
          if (faultTables.length) {
            prependFaultsUpdates.set(item.id, faultTables)
            // Only mark as applied once we actually have something to inject.
            reportFaultTablesAppliedRef.current.add(item.path)
          }
        }
      }

      if (!replaceUpdates.size && !prependFaultsUpdates.size) return
      setDocxDraft((prev) =>
        prev.map((item) => {
          const nextTables = replaceUpdates.get(item.id)
          if (nextTables) return { ...item, tables: nextTables }

          const faultTables = prependFaultsUpdates.get(item.id)
          if (faultTables && faultTables.length) {
            const existing = Array.isArray(item.tables) ? item.tables : []
            return { ...item, tables: [...faultTables, ...existing] }
          }

          return item
        })
      )
    })().finally(() => {
      reportMetadataDefaultsInFlightRef.current = null
    })
  }, [activeAction, docxDraft])

  const writeTextFile = async (dir: FileSystemDirectoryHandle, name: string, content: string) => {
    if (!dir.getFileHandle) throw new Error('Folder write is not available in this browser.')
    const fileHandle = await dir.getFileHandle(name, { create: true })
    if (!fileHandle.createWritable) throw new Error('Unable to write files: File System Access API is unavailable or permission is missing.')
    const writable = await fileHandle.createWritable()
    await writable.write(content)
    await writable.close()
  }

  const writeBinaryFile = async (dir: FileSystemDirectoryHandle, name: string, content: Blob | ArrayBuffer) => {
    if (!dir.getFileHandle) throw new Error('Folder write is not available in this browser.')
    const fileHandle = await dir.getFileHandle(name, { create: true })
    if (!fileHandle.createWritable) throw new Error('Unable to write files: File System Access API is unavailable or permission is missing.')
    const writable = await fileHandle.createWritable()
    await writable.write(content)
    await writable.close()
  }

  const writeCenterFaultFlags = async (faultsDir: FileSystemDirectoryHandle, flags: Record<string, 0 | 1>) => {
    const payload: CenterFlagsDiskV2 = { meta: CURRENT_CENTER_FLAGS_META, flags }
    await writeTextFile(faultsDir, CENTER_FLAGS_FILE, JSON.stringify(payload, null, 2))
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
      if (openedThermalFolderDirectly) {
        throw new Error('You opened the “thermal” folder directly. Open the parent folder that contains “thermal/” (and optionally “rgb/”) so the app can create a sibling “workspace/” folder.')
      }
      if (openedRgbFolderDirectly) {
        throw new Error('You opened the “rgb” folder directly. Open the parent folder that contains “thermal/” and “rgb/” so the app can create a sibling “workspace/” folder.')
      }

      const thermalPath = effectiveThermalFolderPath
      if (!thermalPath) {
        throw new Error('Unable to locate a “thermal” folder inside the opened folder. Open the parent folder that contains “thermal/”.')
      }

      const workflowRoot = await getWorkflowRootDir(false)
      if (!workflowRoot) throw new Error('Unable to access thermal folder')

      const faultsDir = await getWorkspaceThermalFaultsDir(true)
      if (!faultsDir) throw new Error('Unable to access workspace/thermal_faults folder')

      const labelsDir = await getWorkspaceThermalLabelsDir(true)
      if (!labelsDir) throw new Error('Unable to access workspace/thermal_labels folder')

      const labelsCenterDir = await getWorkspaceThermalLabelsCenterDir(true)
      if (!labelsCenterDir) throw new Error('Unable to access workspace/thermal_labels_center folder')

      const metadataDir = await getWorkspaceThermalMetadataDir(true)
      if (!metadataDir) throw new Error('Unable to access workspace/thermal_metadata folder')

      const temperaturesCsvDir = await getWorkspaceThermalTemperaturesCsvsDir(true)
      if (!temperaturesCsvDir) throw new Error('Unable to access workspace/thermal_temperatures_csvs folder')

      if (overwrite) {
        await clearDirectory(faultsDir)
        await clearDirectory(labelsDir)
        await clearDirectory(labelsCenterDir)
        await clearDirectory(metadataDir)
        await clearDirectory(temperaturesCsvDir)
        setStarredFaults({})
        setOnlyStarredInReport(false)
        await writeTextFile(faultsDir, STARRED_FILE, '')
      }

      const buildTiffIndex = () => {
        const isTiff = (p: string) => {
          const lower = p.toLowerCase()
          return lower.endsWith('.tif') || lower.endsWith('.tiff')
        }
        const candidates = Object.entries(fileMap)
          .filter(([path]) => isTiff(path))
          .map(([path, file]) => ({ path, file }))

        const preferred = candidates.filter((item) => item.path.toLowerCase().includes('/tiff/'))
        const list = preferred.length ? preferred : candidates

        const index = new Map<string, { path: string; file: File }>()
        for (const item of list) {
          const name = item.path.split('/').pop() || item.path
          index.set(name.toLowerCase(), { path: item.path, file: item.file })
        }
        return index
      }

      const tiffIndex = buildTiffIndex()
      const pendingMetadata: Array<Promise<void>> = []
      const pendingTempsCsvs: Array<Promise<void>> = []

      const scheduleMetadataProbe = (imagePath: string) => {
        const imageName = imagePath.split('/').pop() || imagePath
        const candidates = [`${imageName}.tiff`, `${imageName}.tif`].map((n) => n.toLowerCase())
        const tiffEntry = candidates.map((n) => tiffIndex.get(n)).find(Boolean)
        if (!tiffEntry) return

        const task = (async () => {
          const formData = new FormData()
          formData.append('file', tiffEntry.file, tiffEntry.path)

          const response = await fetch(apiUrl('/api/metadata/tiff?qr=1'), {
            method: 'POST',
            body: formData,
          })

          if (!response.ok) {
            const err = await response.json().catch(() => ({}))
            throw new Error(err.detail || 'Metadata probe failed')
          }

          const report = (await response.json()) as any

          const qrBase64 = report?.qr_png_base64
          if (typeof qrBase64 === 'string' && qrBase64.length > 0) {
            const pngName = `${imageName}.maps.qr.png`
            const binary = Uint8Array.from(atob(qrBase64), (c) => c.charCodeAt(0))
            await writeBinaryFile(metadataDir, pngName, new Blob([binary], { type: 'image/png' }))

            delete report.qr_png_base64
            report.qr_png = pngName
            report.categories = report.categories && typeof report.categories === 'object' ? report.categories : {}
            report.categories.maps = report.categories.maps && typeof report.categories.maps === 'object' ? report.categories.maps : {}
            report.categories.maps.qr_png = pngName
          }
          const outName = `${imageName}.json`
          await writeTextFile(metadataDir, outName, JSON.stringify(report, null, 2))
        })()

        pendingMetadata.push(
          task.catch((err) => {
            console.warn(`Metadata probe failed for ${imagePath}:`, err)
          })
        )
      }

      const scheduleTemperaturesCsv = (imagePath: string) => {
        const imageName = imagePath.split('/').pop() || imagePath
        const stem = imageName.replace(/\.[^/.]+$/, '')
        const candidates = [`${imageName}.tiff`, `${imageName}.tif`].map((n) => n.toLowerCase())
        const tiffEntry = candidates.map((n) => tiffIndex.get(n)).find(Boolean)
        if (!tiffEntry) return

        const task = (async () => {
          const formData = new FormData()
          formData.append('file', tiffEntry.file, tiffEntry.path)

          const response = await fetch(apiUrl('/api/temperatures/csv?mode=wide&sample=1&nan_empty=1'), {
            method: 'POST',
            body: formData,
          })

          if (!response.ok) {
            const err = await response.json().catch(() => ({}))
            throw new Error(err.detail || 'Temperature CSV export failed')
          }

          const blob = await response.blob()
          const outName = `${stem}.pixel_temps.wide.csv`
          await writeBinaryFile(temperaturesCsvDir, outName, blob)
        })()

        pendingTempsCsvs.push(
          task.catch((err) => {
            console.warn(`Temperature CSV export failed for ${imagePath}:`, err)
          })
        )
      }

      const existingText = overwrite
        ? ''
        : (await readTextFile(faultsDir, 'faults.txt'))
          || (await (async () => {
            const legacyFaultsDir = await getLegacyWorkflowFaultsDir(false)
            const legacy = legacyFaultsDir ? await readTextFile(legacyFaultsDir, 'faults.txt') : null
            if (legacy) return legacy
            const legacyRoot = await getWorkflowRootDir(false)
            return legacyRoot ? await readTextFile(legacyRoot, 'faults.txt') : null
          })())
          || ''
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
      const nextCenterFlags: Record<string, 0 | 1> = overwrite ? {} : { ...centerFaultFlags }

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
            const labels = data.labels.map((label: any) => ({
              classId: Number(label.classId) || 0,
              x: Number(label.x) || 0,
              y: Number(label.y) || 0,
              w: Number(label.w) || 0,
              h: Number(label.h) || 0,
              conf: Number.isFinite(Number(label.conf)) ? Number(label.conf) : 1,
            })) as Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number }>

            const labelText = labels
              .map((l) => `${l.classId} ${l.x} ${l.y} ${l.w} ${l.h} ${l.conf}`)
              .join('\n')
            await writeTextFile(labelsDir, `${stem}.txt`, labelText)

            const { perLabel, imageFlag } = await computeCenterFlagsForLabels(file, labels)
            const centerText = labels
              .map((l, i) => `${l.classId} ${l.x} ${l.y} ${l.w} ${l.h} ${l.conf} ${perLabel[i] ?? 0}`)
              .join('\n')
            await writeTextFile(labelsCenterDir, `${stem}.txt`, centerText)
            nextCenterFlags[path] = imageFlag

            detected.push(path)
            existingSet.add(path)

            // In parallel with scanning, probe the corresponding TIFF and save a JSON report.
            scheduleMetadataProbe(path)

            // In parallel with scanning, export per-pixel temperatures CSV for the corresponding TIFF.
            scheduleTemperaturesCsv(path)
          }
        }
      }

      const pendingAll = [...pendingMetadata, ...pendingTempsCsvs]
      if (pendingAll.length > 0) {
        setScanStatus(`Finalizing sidecars (${pendingAll.length} item(s))…`)
        await Promise.allSettled(pendingAll)
      }

      // Backfill center flags for any existing faults that were already present
      // (e.g. previous scan run) so the filter works consistently.
      for (const path of existingSet) {
        if (nextCenterFlags[path] !== undefined) continue
        const file = fileMap[path]
        if (!file) {
          nextCenterFlags[path] = 0
          continue
        }
        const stem = path.split('/').pop()?.replace(/\.[^/.]+$/, '') || path
        const labelText = (await readTextFile(labelsDir, `${stem}.txt`)) || ''
        const labels = labelText
          .split(/\r?\n/)
          .map((line) => line.trim())
          .filter(Boolean)
          .map((line) => {
            const [classId, x, y, w, h, conf] = line.split(/\s+/).map(parseYoloNumber)
            return {
              classId: Number.isFinite(classId) ? classId : 0,
              x: Number.isFinite(x) ? x : 0,
              y: Number.isFinite(y) ? y : 0,
              w: Number.isFinite(w) ? w : 0,
              h: Number.isFinite(h) ? h : 0,
              conf: Number.isFinite(conf) ? conf : 1,
            }
          })
        if (labels.length === 0) {
          nextCenterFlags[path] = 0
          continue
        }
        const { perLabel, imageFlag } = await computeCenterFlagsForLabels(file, labels)
        const centerText = labels
          .map((l, i) => `${l.classId} ${l.x} ${l.y} ${l.w} ${l.h} ${l.conf} ${perLabel[i] ?? 0}`)
          .join('\n')
        await writeTextFile(labelsCenterDir, `${stem}.txt`, centerText)
        nextCenterFlags[path] = imageFlag
      }

      const merged = [...existingSet]
      await writeTextFile(faultsDir, 'faults.txt', merged.join('\n'))

      await writeCenterFaultFlags(faultsDir, nextCenterFlags)
      setCenterFaultFlags(nextCenterFlags)

      setFaultsList(merged)
      setScanCompleted(true)
      setScanStatus(`Scan completed . Faults found: ${detected.length}`)

      // Refresh the in-memory file tree so newly created workspace files become visible
      // without requiring the user to reopen the folder.
      await refreshFolderEntries(folderHandle)

      // If the user currently has an image open with no loaded labels, try reloading
      // now that scan outputs exist on disk.
      if (selectedPath && fileKind === 'image' && selectedLabelsRef.current.length === 0) {
        const next = await readLabelsForPath(selectedPath)
        if (next.length > 0) {
          selectedLabelsRef.current = next
          setSelectedLabels(next)
          setSelectedLabelIndex(null)
          setLabelHistory([next])
          setLabelHistoryIndex(0)
        }
      }
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

  const handleSelectFile = async (path: string, variantOverride?: 'thermal' | 'rgb') => {
    const clickedFile = fileMap[path]
    if (!clickedFile) return

    const clickedIsImage = clickedFile.type.startsWith('image/')

    // Context (thermal) path is used for labels/metadata.
    // View (thermal or rgb) path is used for display.
    let contextPath = path
    let viewPath = path

    if (clickedIsImage) {
      const clickedName = path.split('/').pop() || ''
      const parsed = parseDjiTvVariant(clickedName)
      const clickedVariant = parsed?.variant ?? null

      if (clickedVariant === 'V') {
        const thermalCandidate = getThermalPathForRgb(path)
        if (thermalCandidate && fileMap[thermalCandidate]) {
          contextPath = thermalCandidate
        }
        // If user clicked an RGB image, force RGB view.
        setViewVariant('rgb')
        viewPath = path
      } else {
        // Thermal (or non-matching) image: follow current viewVariant preference.
        const effectiveVariant = variantOverride ?? viewVariant
        if (effectiveVariant === 'rgb') {
          if (variantOverride) setViewVariant('rgb')
          const rgbCandidate = getRgbPathForThermal(contextPath)
          if (rgbCandidate && fileMap[rgbCandidate]) {
            viewPath = rgbCandidate
          } else {
            viewPath = contextPath
          }
        } else {
          if (variantOverride) setViewVariant('thermal')
          viewPath = contextPath
        }
      }
    }

    const viewFile = clickedIsImage ? fileMap[viewPath] : clickedFile
    if (!viewFile) return

    setSelectedPath(contextPath)
    setSelectedViewPath(clickedIsImage ? viewPath : contextPath)
    if (fileUrl) URL.revokeObjectURL(fileUrl)
    setFileUrl('')
    setFileText('')
    setSelectedLabels([])
    setSelectedLabelIndex(null)
    setShowLabels(true)
    setZoom(1)
    setPan({ x: 0, y: 0 })
    setDrawMode('select')

    if (viewFile.type.startsWith('image/')) {
      setFileKind('image')
      setFileUrl(URL.createObjectURL(viewFile))
      const isViewingRgb = Boolean(clickedIsImage && viewPath && contextPath && viewPath !== contextPath)
      const cached = isViewingRgb ? rgbWorkingLabelsByThermalPathRef.current[contextPath] : null
      const labels = cached ? cached.labels : await readLabelsForPath(contextPath)
      if (isViewingRgb) {
        setRgbLabelsAreImageSpace(Boolean(cached?.areImageSpace))
        if (!cached) {
          rgbWorkingLabelsByThermalPathRef.current[contextPath] = { labels, areImageSpace: false }
        }
      }
      setSelectedLabels(labels)
      setSelectedLabelIndex(null)
      setShowLabels(true)
      setZoom(1)
      setPan({ x: 0, y: 0 })
      setDrawMode('select')
      setLabelHistory([labels])
      setLabelHistoryIndex(0)
      if (isViewingRgb) setRgbLabelsExportDirty(false)
      return
    }

    if (viewFile.type.startsWith('text/') || viewFile.name.match(/\.(md|txt|json|csv|ts|tsx|js|jsx|py|html|css)$/i)) {
      setFileKind('text')
      const text = await viewFile.text()
      setFileText(text)
      return
    }

    setFileKind('other')
  }

  const displayedRootPath = useMemo(() => {
    if (effectiveThermalFolderPath) return getParentDirPath(effectiveThermalFolderPath)
    if (!fileTree?.children?.length) return ''
    const stack: TreeNode[] = [...fileTree.children]
    while (stack.length > 0) {
      const node = stack.shift()!
      if (node.type === 'folder' && node.name.toLowerCase() === 'thermal') {
        return getParentDirPath(node.path)
      }
      if (node.type === 'folder' && node.children && node.children.length > 0) {
        stack.unshift(...node.children)
      }
    }
    return ''
  }, [effectiveThermalFolderPath, fileTree])

  const displayedRootNode = useMemo(() => {
    if (!fileTree?.children?.length) return null
    const target = displayedRootPath
    if (!target) return fileTree
    const stack: TreeNode[] = [...fileTree.children]
    while (stack.length > 0) {
      const node = stack.shift()!
      if (node.type === 'folder' && node.path === target) return node
      if (node.type === 'folder' && node.children && node.children.length > 0) {
        stack.unshift(...node.children)
      }
    }
    return fileTree
  }, [displayedRootPath, fileTree])

  const rootLabel = useMemo(() => {
    if (!fileTree?.children?.length) return 'No folder selected'
    if (!displayedRootPath) return folderHandle?.name || 'project'
    return displayedRootPath.split('/').filter(Boolean).pop() || folderHandle?.name || 'project'
  }, [displayedRootPath, fileTree, folderHandle?.name])

  const rootChildren = useMemo(() => {
    if (!fileTree?.children?.length) return []
    if (!displayedRootNode) return []
    return displayedRootNode.children ?? []
  }, [displayedRootNode, fileTree])

  useEffect(() => {
    if (!fileTree?.children?.length) return
    const workspacePath = displayedRootPath ? `${displayedRootPath}/workspace` : 'workspace'
    setExpandedPaths((prev) => {
      const next = new Set(prev)
      next.add('')
      next.add(workspacePath)
      return next
    })
  }, [displayedRootPath, fileTree])

  const hasFolderSelected = useMemo(
    () => Boolean(fileTree && (rootChildren.length > 0 || (fileTree.children && fileTree.children.length > 0))),
    [fileTree, rootChildren.length]
  )

  const showLeftExplorer = hasFolderSelected && (activeMenu === 'file' || (activeMenu === 'actions' && activeAction === 'view'))
  const showRightExplorer = hasFolderSelected && activeMenu === 'actions' && activeAction === 'view'
  const isViewExplorer = activeMenu === 'actions' && activeAction === 'view'
  const isHomePage = activeMenu === 'file'

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

    if (starredFaults[path]) {
      const nextStarred = { ...starredFaults }
      delete nextStarred[path]
      setStarredFaults(nextStarred)
      void (async () => {
        try {
          await writeStarredFaultsToDisk(nextStarred)
        } catch {
          // ignore; removal still succeeded
        }
      })()
    }

    try {
      const faultsDir = await getWorkspaceThermalFaultsDir(true)
      if (!faultsDir) {
        throw new Error(
          openedThermalFolderDirectly
            ? 'You opened the “thermal” folder directly. Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
            : openedRgbFolderDirectly
              ? 'You opened the “rgb” folder directly. Open the parent folder that contains “thermal/” (and “rgb/”) so the app can create a sibling “workspace/” folder.'
              : 'Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
        )
      }

      await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))

      // Keep center_flags.json in sync.
      const { flags } = await readCenterFlagsDisk(faultsDir)
      if (flags[path] !== undefined) {
        const nextFlags = { ...flags }
        delete nextFlags[path]
        await writeCenterFaultFlags(faultsDir, nextFlags)
        setCenterFaultFlags(nextFlags)
      }
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
        const faultsDir = await getWorkspaceThermalFaultsDirForPath(node.path, true)
        if (faultsDir) {
          await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))
        }
      }

      if (!isFolder) {
        const labelsDir = await getWorkspaceThermalLabelsDirForPath(node.path, false)
        if (labelsDir?.removeEntry) {
          await labelsDir.removeEntry(getLabelFileName(node.path)).catch(() => undefined)
        }

        const labelsCenterDir = await getWorkspaceThermalLabelsCenterDirForPath(node.path, false)
        if (labelsCenterDir?.removeEntry) {
          await labelsCenterDir.removeEntry(getLabelFileName(node.path)).catch(() => undefined)
        }

        const faultsDir = await getWorkspaceThermalFaultsDirForPath(node.path, false)
        if (faultsDir) {
          const { flags } = await readCenterFlagsDisk(faultsDir)
          if (flags[node.path] !== undefined) {
            const nextFlags = { ...flags }
            delete nextFlags[node.path]
            await writeCenterFaultFlags(faultsDir, nextFlags)
            setCenterFaultFlags(nextFlags)
          }
        }

        if (starredFaults[node.path]) {
          const nextStarred = { ...starredFaults }
          delete nextStarred[node.path]
          setStarredFaults(nextStarred)
          void (async () => {
            try {
              await writeStarredFaultsToDisk(nextStarred)
            } catch {
              // ignore
            }
          })()
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
    if (viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath) {
      setRgbLabelsExportDirty(true)
      return
    }
    await persistLabels(snapshot)
  }

  const persistLabels = async (labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; source?: 'auto' | 'manual' }>) => {
    try {
      // Thermal labels are persisted to `workspace/thermal_labels/...`.
      // RGB labels must remain independent and are saved only via the RGB export button.
      if (viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath) return
      if (!folderHandle || !folderHandle.getDirectoryHandle || !selectedPath) return

      const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
      if (!canWrite) {
        setLabelSaveError('Write permission denied. Reopen the folder and allow write access to save labels.')
        return
      }

      const faultsDir = await getWorkspaceThermalFaultsDir(true)
      if (!faultsDir) {
        throw new Error(
          openedThermalFolderDirectly
            ? 'You opened the “thermal” folder directly. Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
            : openedRgbFolderDirectly
              ? 'You opened the “rgb” folder directly. Open the parent folder that contains “thermal/” (and “rgb/”) so the app can create a sibling “workspace/” folder.'
              : 'Open the parent folder that contains “thermal/” so the app can create a sibling “workspace/” folder.'
        )
      }

      const labelsDir = await getWorkspaceThermalLabelsDirForPath(selectedPath, true)
      if (!labelsDir) throw new Error('Unable to access workspace/thermal_labels folder')

      const labelsCenterDir = await getWorkspaceThermalLabelsCenterDirForPath(selectedPath, true)
      if (!labelsCenterDir) throw new Error('Unable to access workspace/thermal_labels_center folder')

      const lines = labels.map((label) =>
        `${label.classId} ${label.x} ${label.y} ${label.w} ${label.h} ${Number.isFinite(label.conf) ? label.conf : 1}`
      )
      await writeTextFile(labelsDir, getLabelFileName(selectedPath), lines.join('\n'))

      const file = fileMap[selectedPath]
      if (file) {
        const normalizedLabels = labels.map((l) => ({
          classId: l.classId,
          x: l.x,
          y: l.y,
          w: l.w,
          h: l.h,
          conf: Number.isFinite(l.conf) ? l.conf : 1,
        }))
        const { perLabel, imageFlag } = await computeCenterFlagsForLabels(file, normalizedLabels)
        const centerLines = normalizedLabels.map((l, i) => `${l.classId} ${l.x} ${l.y} ${l.w} ${l.h} ${l.conf} ${perLabel[i] ?? 0}`)
        await writeTextFile(labelsCenterDir, getLabelFileName(selectedPath), centerLines.join('\n'))

        const disk = await readCenterFlagsDisk(faultsDir)
        const next = { ...disk.flags, [selectedPath]: labels.length > 0 ? imageFlag : 0 }
        await writeCenterFaultFlags(faultsDir, next)
        setCenterFaultFlags(next)
      }

      if (labels.length > 0) {
        if (!faultsList.includes(selectedPath)) {
          const nextFaults = [...faultsList, selectedPath]
          setFaultsList(nextFaults)
          await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))
        }
      } else {
        if (faultsList.includes(selectedPath)) {
          const nextFaults = faultsList.filter((path) => path !== selectedPath)
          setFaultsList(nextFaults)
          await writeTextFile(faultsDir, 'faults.txt', nextFaults.join('\n'))
        }
      }

      setLabelSaveError('')
    } catch (error) {
      setLabelSaveError(error instanceof Error ? error.message : 'Failed to save labels')
    }
  }

  const syncReportFaultDescriptionsForPath = (
    path: string,
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; source?: 'auto' | 'manual' }>
  ) => {
    if (!path) return
    const possibleByIndex = labels.map((l) => formatPossibleFaultType(l.classId))

    setDocxDraft((prev) => {
      if (!prev.length) return prev

      let anyChanged = false
      const next = prev.map((item) => {
        if (item.path !== path) return item
        const tables = Array.isArray(item.tables) ? item.tables : []
        if (!tables.length) return item

        let changed = false
        const nextTables = tables.map((t) => {
          const title = (t?.title || '').trim()
          if (title !== 'Faults') return t

          const rows = Array.isArray(t.rows) ? t.rows : []
          if (!rows.length) return t

          let rowsChanged = false
          const nextRows = rows.map((r) => {
            const name = (r?.name || '').trim()
            const m = /^Fault\s+(\d+)$/i.exec(name)
            if (!m) return r
            const idx = Number(m[1]) - 1
            if (!Number.isFinite(idx) || idx < 0 || idx >= possibleByIndex.length) return r

            const nextDesc = possibleByIndex[idx]
            const currentDesc = (r?.description || '').trim()
            const isAuto = currentDesc === '' || /^Possible\s+/i.test(currentDesc)
            if (!isAuto) return r
            if (r.description === nextDesc) return r

            rowsChanged = true
            return { ...r, description: nextDesc }
          })

          if (!rowsChanged) return t
          changed = true
          return { ...t, rows: nextRows }
        })

        if (!changed) return item
        anyChanged = true
        return { ...item, tables: nextTables }
      })

      return anyChanged ? next : prev
    })
  }

  const saveRgbLabelsExport = async () => {
    try {
      if (rgbLabelsExportSaving) return
      if (viewVariant !== 'rgb' || activeAction !== 'view' || fileKind !== 'image') return
      if (!folderHandle || !folderHandle.getDirectoryHandle || !selectedViewPath) return
      if (!selectedPath || selectedViewPath === selectedPath) return
      if (zoom !== 1) {
        setLabelSaveError('Set zoom to 100% (1.0) before saving RGB labels.')
        return
      }

      const stem = getRgbStemForExport()
      if (!stem) return

      // Only allow saving when user has modified labels while viewing RGB.
      if (!rgbLabelsExportDirty) return

      setRgbLabelsExportSaving(true)
      setLabelSaveError('')

      const canWrite = await ensureHandlePermission(folderHandle, 'readwrite')
      if (!canWrite) {
        setLabelSaveError('Write permission denied. Reopen the folder and allow write access to save RGB labels.')
        return
      }

      const labelsDir = await ensureRgbLabelsDir()
      if (!labelsDir) throw new Error('Unable to access workspace/rgb_labels folder')

      const outName = `${stem}.txt`

      const labels = selectedLabelsRef.current
      let rgbFullNormalized = labels

      const viewport = imageContainerRef.current
      const img = imageRef.current
      if (!viewport) throw new Error('Viewer viewport is unavailable')
      if (!img) throw new Error('RGB image is unavailable')

      const viewportRect = viewport.getBoundingClientRect()
      const imgRect = img.getBoundingClientRect()
      const crop = getRgbCropRectNormalized(viewportRect, imgRect)
      if (!crop) throw new Error('Unable to compute RGB crop (100% viewport)')

      if (!rgbLabelsAreImageSpace) {
        const converted = convertThermalFrameLabelsToRgbNormalized(labels, viewportRect, imgRect)
        if (labels.length > 0 && !converted.length) throw new Error('Unable to map labels into RGB frame')
        rgbFullNormalized = converted
      }

      // Persist as crop-relative YOLO coords (what user sees at 100% zoom).
      const rgbCropNormalized = convertRgbFullLabelsToCropNormalized(rgbFullNormalized, crop)

      const lines = rgbCropNormalized.map((label) => {
        const cx = clamp01(label.x)
        const cy = clamp01(label.y)
        const w = clamp01(label.w)
        const h = clamp01(label.h)
        const conf = Number.isFinite(label.conf) ? label.conf : 1
        return `${Number(label.classId) || 0} ${cx.toFixed(6)} ${cy.toFixed(6)} ${w.toFixed(6)} ${h.toFixed(6)} ${Number.isFinite(conf) ? conf.toFixed(6) : '1.000000'}`
      })

      const outText = lines.join('\n')
      await writeTextFile(labelsDir, outName, outText)
      setRgbLabelsExportFileExists(true)
      setRgbLabelsExportFileEmpty(outText.trim().length === 0)
      setRgbLabelsExportStatus('exists')
      setRgbLabelsExportDirty(false)
      setLabelSaveError('')

      setRgbLabelsAreImageSpace(true)
      setSelectedLabels(rgbFullNormalized)
      setSelectedLabelIndex(null)
      setLabelHistory([rgbFullNormalized])
      setLabelHistoryIndex(0)
    } catch (error) {
      setLabelSaveError(error instanceof Error ? error.message : 'Failed to save RGB labels')
    } finally {
      setRgbLabelsExportSaving(false)
    }
  }

  const scheduleRealtimePersist = (
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
  ) => {
    if (viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath) return
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

  const flushRealtimePersist = (
    labels: Array<{ classId: number; x: number; y: number; w: number; h: number; conf: number; shape?: 'rect' | 'ellipse'; source?: 'auto' | 'manual' }>
  ) => {
    if (viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath) return
    if (!folderHandle || !selectedPath) return
    realtimePersistLabelsRef.current = labels
    if (realtimePersistTimerRef.current !== null) {
      window.clearTimeout(realtimePersistTimerRef.current)
      realtimePersistTimerRef.current = null
    }
    void persistLabels(labels)
  }

  const clamp01 = (value: number) => Math.min(Math.max(value, 0), 1)

  useEffect(() => {
    try {
      window.localStorage.setItem('rgb-align-offset-v1', JSON.stringify(rgbAlignOffset))
    } catch {
      // ignore
    }
  }, [rgbAlignOffset])

  useEffect(() => {
    try {
      const safe = Number.isFinite(rgbAlignScale) ? rgbAlignScale : 1
      window.localStorage.setItem('rgb-align-scale-v1', String(safe))
    } catch {
      // ignore
    }
  }, [rgbAlignScale])

  const getRgbThermalFrameRect = (renderW: number, renderH: number) => {
    if (viewVariant !== 'rgb') return null
    const thermalW = contextImageNatural?.width ?? 0
    const thermalH = contextImageNatural?.height ?? 0

    if (!(renderW > 0 && renderH > 0 && thermalW > 0 && thermalH > 0)) return null

    // Map thermal labels into a centered, locked 1536×1229 frame (or smaller if
    // the viewport is smaller). This is independent of the RGB image size.
    // Allow a small manual scale tweak so RGB labels can match the thermal view.
    const safeAlignScale = Number.isFinite(rgbAlignScaleEffective) ? Math.min(2, Math.max(0.5, rgbAlignScaleEffective)) : 1

    const lockedW = Math.min(VIEWER_MAX_W, renderW)
    const lockedH = Math.min(VIEWER_MAX_H, renderH)
    const lockedLeft = (renderW - lockedW) / 2
    const lockedTop = (renderH - lockedH) / 2

    const baseScale = Math.min(lockedW / thermalW, lockedH / thermalH) * safeAlignScale
    const innerW = thermalW * baseScale
    const innerH = thermalH * baseScale

    const baseLeft = lockedLeft + (lockedW - innerW) / 2
    const baseTop = lockedTop + (lockedH - innerH) / 2

    const innerLeft = baseLeft
    const innerTop = baseTop

    return {
      // Keep `outer` for compatibility with existing rendering code.
      // It is intentionally identical to `inner` now.
      outer: {
        left: innerLeft,
        top: innerTop,
        width: innerW,
        height: innerH,
      },
      inner: {
        left: innerLeft,
        top: innerTop,
        width: innerW,
        height: innerH,
      },
      thermalW,
      thermalH,
    }
  }

  const getRgbOverlayOffsetPx = (renderW: number, renderH: number) => {
    // Per-bucket scale (pitch/altitude) for more accurate alignment
    const scale = Number.isFinite(rgbAlignScaleEffective) ? rgbAlignScaleEffective : 1
    const userX = (Number.isFinite(rgbAlignOffset.x) ? rgbAlignOffset.x : 0) * renderW * scale
    const userY = (Number.isFinite(rgbAlignOffset.y) ? rgbAlignOffset.y : 0) * renderH * scale

    const base = getYawInterpolatedOffsetPx(viewAlignYawDegree)
    const fgAdj = getFlightGimbalYawAdjustmentPx(viewFlightYawDegree, viewGimbalYawDegree)
    return {
      x: (base.x + fgAdj.x) * scale + userX,
      y: (base.y + fgAdj.y) * scale + userY,
    }
  }

  // Maps a pointer event to normalized image coordinates (0..1) based on the
  // actual rendered <img> rect. This stays correct under zoom/pan transforms.
  const toImageCoords = (event: React.MouseEvent) => {
    // In RGB mode, pointer mapping is locked to the viewer viewport (1536×1229 frame)
    // and must NOT depend on the transformed <img> rect.
    if (viewVariant === 'rgb') {
      if (rgbLabelsAreImageSpace && selectedViewPath && selectedPath && selectedViewPath !== selectedPath) {
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

      const viewport = imageContainerRef.current
      if (!viewport) return null
      const rect = viewport.getBoundingClientRect()
      if (!rect.width || !rect.height) return null
      const frame = getRgbThermalFrameRect(rect.width, rect.height)
      if (!frame) return null
      const rx = event.clientX - rect.left
      const ry = event.clientY - rect.top

      // Apply the same pixel offset as the overlay rendering, but inverted,
      // so clicks/drags stay aligned with the shifted labels.
      const offset = getRgbOverlayOffsetPx(rect.width, rect.height)
      const adjX = rx - offset.x
      const adjY = ry - offset.y
      return {
        x: clamp01((adjX - frame.inner.left) / frame.inner.width),
        y: clamp01((adjY - frame.inner.top) / frame.inner.height),
      }
    }

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

  useEffect(() => {
    // If the RGB labels file exists but is empty, initialize RGB labels based on
    // what the user currently sees (default 100% zoom/crop) by converting the
    // thermal-frame overlay into real RGB image-space coords.
    if (activeAction !== 'view') return
    if (viewVariant !== 'rgb') return
    if (!selectedPath || !selectedViewPath || selectedViewPath === selectedPath) return
    if (fileKind !== 'image') return
    if (!rgbLabelsExportFileExists || !rgbLabelsExportFileEmpty) return
    if (rgbLabelsAreImageSpace) return
    if (zoom !== 1) return

    const viewport = imageContainerRef.current
    const img = imageRef.current
    if (!viewport || !img) return

    const viewportRect = viewport.getBoundingClientRect()
    const imgRect = img.getBoundingClientRect()
    const renderW = viewportRect.width
    const renderH = viewportRect.height
    if (!(renderW > 0 && renderH > 0)) return
    if (!(imgRect.width > 0 && imgRect.height > 0)) return

    const labels = selectedLabelsRef.current
    const rgbNormalized = convertThermalFrameLabelsToRgbNormalized(labels, viewportRect, imgRect)
    if (!rgbNormalized.length) return

    setRgbLabelsAreImageSpace(true)
    setSelectedLabels(rgbNormalized)
    setSelectedLabelIndex(null)
    setLabelHistory([rgbNormalized])
    setLabelHistoryIndex(0)
  }, [
    activeAction,
    fileKind,
    imageMetrics.height,
    imageMetrics.naturalHeight,
    imageMetrics.naturalWidth,
    imageMetrics.width,
    rgbLabelsAreImageSpace,
    rgbLabelsExportFileEmpty,
    rgbLabelsExportFileExists,
    selectedPath,
    selectedViewPath,
    viewVariant,
    zoom,
  ])

  useEffect(() => {
    // When the RGB labels file contains crop-relative coords, convert them to
    // full-image normalized coords once we have DOM geometry, so labels render
    // and edit stably (no drift on zoom/pan).
    if (!rgbPendingCropFileLabels) return
    if (activeAction !== 'view') return
    if (viewVariant !== 'rgb') return
    if (!selectedPath || !selectedViewPath || selectedViewPath === selectedPath) return
    if (fileKind !== 'image') return
    if (zoom !== 1) return

    const viewport = imageContainerRef.current
    const img = imageRef.current
    if (!viewport || !img) return

    const viewportRect = viewport.getBoundingClientRect()
    const imgRect = img.getBoundingClientRect()
    const crop = getRgbCropRectNormalized(viewportRect, imgRect)
    if (!crop) return

    // Backward-compat: older RGB label files may already be full-image normalized.
    // Heuristic: if almost all centers already fall inside the crop rect in full-image
    // coords, treat the file as full-image to avoid shrink/shift.
    const x0 = crop.x0
    const y0 = crop.y0
    const x1 = crop.x0 + crop.scaleX
    const y1 = crop.y0 + crop.scaleY
    const candidates = rgbPendingCropFileLabels.filter((l) => Number.isFinite(l.x) && Number.isFinite(l.y))
    const inCrop = candidates.filter((l) => l.x >= x0 && l.x <= x1 && l.y >= y0 && l.y <= y1).length
    const ratio = candidates.length ? inCrop / candidates.length : 0
    const treatAsFullImage = ratio >= 0.9

    const full = treatAsFullImage ? rgbPendingCropFileLabels : convertRgbCropLabelsToFullNormalized(rgbPendingCropFileLabels, crop)
    setSelectedLabels(full)
    setSelectedLabelIndex(null)
    setLabelHistory([full])
    setLabelHistoryIndex(0)
    setRgbPendingCropFileLabels(null)
  }, [
    activeAction,
    fileKind,
    imageMetrics.height,
    imageMetrics.naturalHeight,
    imageMetrics.naturalWidth,
    imageMetrics.width,
    rgbPendingCropFileLabels,
    selectedPath,
    selectedViewPath,
    viewVariant,
    zoom,
  ])

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

    const viewScale = viewVariant === 'rgb' ? RGB_VIEW_IMAGE_SCALE : 1
    const scale = zoom * BASE_ZOOM * viewScale
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

  const handleDrawSurfaceMouseDown = (event: React.MouseEvent) => {
    if (!fileUrl) return

    // RGB alignment adjustment: Alt+drag shifts the thermal-to-RGB mapping.
    // Alt+double-click resets.
    if (viewVariant === 'rgb' && drawMode === 'select' && event.altKey) {
      event.preventDefault()
      event.stopPropagation()

      if (event.detail >= 2) {
        setRgbAlignOffset({ x: 0, y: 0 })
        setRgbAlignScale(1)
        if (rgbAlignScaleContextKey) {
          try {
            window.localStorage.removeItem(getRgbAlignScaleStorageKey(rgbAlignScaleContextKey))
          } catch {
            // ignore
          }
          setRgbAlignScaleContextTick((v) => v + 1)
        }
        setIsRgbAlignDragging(false)
        setRgbAlignDragStart(null)
        setRgbAlignDragOrigin(null)
        return
      }

      setIsRgbAlignDragging(true)
      setRgbAlignDragStart({ x: event.clientX, y: event.clientY })
      setRgbAlignDragOrigin({ ...rgbAlignOffset })
      return
    }

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
      return
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
  }

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
                    Scanner
                  </button>
                  <button
                    className={`dropdown-item ${!hasFolderSelected ? 'disabled' : ''}`}
                    disabled={!hasFolderSelected}
                    onClick={() => {
                      setActiveMenu('actions')
                      setActiveAction('report')
                      // void loadReportFromDisk()
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
                  <button
                    className="dropdown-item"
                    onClick={() => {
                      setActiveHelpPage('documentation')
                      setActiveMenu('help')
                    }}
                  >
                    Documentation
                  </button>
                  <button
                    className="dropdown-item"
                    onClick={() => {
                      setActiveHelpPage('about')
                      setActiveMenu('help')
                    }}
                  >
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
                  <div className={`explorer-pane explorer-pane-tree ${isHomePage ? 'explorer-pane-tree-full' : ''}`}>
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
                  {isViewExplorer && faultsListImagePaths.length > 0 && (
                    <>
                      <div className="explorer-pane">
                        <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 8 }}>
                          <div className="explorer-pane-title">Thermal</div>
                          {activeAction === 'view' && (
                            <label
                              style={{ display: 'flex', alignItems: 'center', gap: 6, fontSize: 12, color: '#6b7a90' }}
                              title={`Show only images where at least one label center is inside the ${CENTER_BOX_W}×${CENTER_BOX_H} center box (centered faults)`}
                            >
                              <input
                                type="checkbox"
                                checked={onlyCenterBoxFaults}
                                onChange={(e) => handleToggleOnlyCenterBoxFaults(e.target.checked)}
                              />
                              Centered faults
                            </label>
                          )}
                        </div>
                        <div className="tree">
                          {activeAction === 'view' ? (
                            faultsListForTempTitleB.length > 0 ? (
                              faultsListForTempTitleB.map((path) => {
                                const name = path.split('/').pop() || path
                                const isMissing = !fileMap[path]
                                return (
                                  <button
                                    key={path}
                                    className={`tree-item file ${selectedPath === path ? 'active' : ''}`}
                                    style={{ paddingLeft: 2 }}
                                    onClick={() => {
                                      setNavScope('faultsList')
                                      void handleSelectFile(path, 'thermal')
                                    }}
                                    onContextMenu={(event) => openExplorerContextMenu(event, { name, path, type: 'file' }, 'faultsList')}
                                    disabled={isMissing}
                                    title={isMissing ? 'File not found in folder' : path}
                                  >
                                    <span className="file-label">
                                      <span className="file-icon">📷</span>
                                      <span className="file-name">{starredFaults[path] ? `★ ${name}` : name}</span>
                                    </span>
                                  </button>
                                )
                              })
                            ) : (
                              <div className="explorer-empty">{onlyCenterBoxFaults ? 'No center-box faults.' : 'No faults listed yet.'}</div>
                            )
                          ) : null}
                        </div>
                        {activeAction === 'view' && (
                          <div className="explorer-pane-footer">Images: {faultsListForTempTitleB.length}</div>
                        )}
                      </div>

                      <div className="explorer-pane">
                        <div className="explorer-pane-title">RGB</div>
                        <div className="tree">
                          {activeAction === 'view' ? (
                            faultsListRgbForTempTitleB.length > 0 ? (
                              faultsListRgbForTempTitleB.map((path) => {
                                const name = path.split('/').pop() || path
                                const isMissing = !fileMap[path]
                                const isActive = Boolean(selectedViewPath && selectedViewPath === path)
                                return (
                                  <button
                                    key={path}
                                    className={`tree-item file ${isActive ? 'active' : ''}`}
                                    style={{ paddingLeft: 2 }}
                                    onClick={() => {
                                      setNavScope('faultsList')
                                      void handleSelectFile(path, 'rgb')
                                    }}
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
                              <div className="explorer-empty">No RGB matches.</div>
                            )
                          ) : null}
                        </div>
                        {activeAction === 'view' && (
                          <div className="explorer-pane-footer">Images: {faultsListRgbForTempTitleB.length}</div>
                        )}
                      </div>
                    </>
                  )}
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
          {activeMenu === 'file' && (
            <section id="welcome-section" className="card home-card">
              <div className="home-hero">
                <div className="home-hero-text">
                  <span className="home-badge">Workspace</span>
                  <h1 id="welcome-header">Welcome</h1>
                  <p>Open a folder to browse files and preview their contents.</p>
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
                    <li>Choose a folder to get started → This folder must contain three folders with exact names: rgb, thermal and tiff</li>
                    <li>Then click "Scan for faults" to proceed to the scan page or choose Actions→Scanner</li>
                    <li>You can always use Actions→View page. Without scanning, you won't see any faults e.t.c.</li>
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
                  {activeAction === 'scan' && 'Scanner'}
                  {activeAction === 'report' && 'Report'}
                  {!activeAction && 'Actions'}
                </h2>
                {activeAction === 'view' && (
                  <div className="page-header-right">
                    {selectedPath && (
                      <div
                        className="view-filename"
                        title={selectedViewPath || selectedPath}
                      >
                        <span className="view-filename-name">{(selectedViewPath || selectedPath).split('/').pop() || (selectedViewPath || selectedPath)}</span>
                        {fileKind === 'image' && selectedViewPath === selectedPath && viewHoverPixel && viewHoverTempsStatus === 'ready' && (
                          <span className="view-filename-temp" title={`x=${viewHoverPixel.x}, y=${viewHoverPixel.y}`}>
                            {Number.isFinite(viewHoverPixelTempC as number) ? `${(viewHoverPixelTempC as number).toFixed(2)} °C` : 'N/A'}
                          </span>
                        )}
                      </div>
                    )}
                    {selectedPath && fileKind === 'image' && fileUrl && (
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
                        {(() => {
                          const isViewingRgb = Boolean(selectedViewPath && selectedPath && selectedViewPath !== selectedPath)
                          const hasRgb = Boolean(selectedPath && isRgbViewAvailableForThermal(selectedPath))
                          const nextVariant: 'thermal' | 'rgb' = isViewingRgb ? 'thermal' : 'rgb'
                          const label = isViewingRgb ? 'Thermal' : 'RGB'

                          return (
                            <button
                              className="link-button"
                              onClick={() => {
                                if (!selectedPath) return
                                if (nextVariant === 'rgb' && !hasRgb) return

                                setViewVariant(nextVariant)
                                const nextViewPath = nextVariant === 'rgb' ? getRgbPathForThermal(selectedPath) : selectedPath
                                if (!nextViewPath || !fileMap[nextViewPath]) return
                                setSelectedViewPath(nextViewPath)
                                if (fileUrl) URL.revokeObjectURL(fileUrl)
                                setFileUrl(URL.createObjectURL(fileMap[nextViewPath]))

                                void (async () => {
                                  const isNextViewingRgb = nextVariant === 'rgb' && nextViewPath !== selectedPath
                                  const cached = isNextViewingRgb ? rgbWorkingLabelsByThermalPathRef.current[selectedPath] : null
                                  const labels = cached ? cached.labels : await readLabelsForPath(selectedPath)
                                  if (isNextViewingRgb) {
                                    setRgbLabelsAreImageSpace(Boolean(cached?.areImageSpace))
                                    if (!cached) {
                                      rgbWorkingLabelsByThermalPathRef.current[selectedPath] = { labels, areImageSpace: false }
                                    }
                                  }
                                  setSelectedLabels(labels)
                                  setSelectedLabelIndex(null)
                                  setLabelHistory([labels])
                                  setLabelHistoryIndex(0)
                                  if (isNextViewingRgb) setRgbLabelsExportDirty(false)
                                })()
                              }}
                              disabled={nextVariant === 'rgb' && !hasRgb}
                              title={nextVariant === 'rgb' && !hasRgb ? 'No matching RGB image found' : `Switch to ${label} view`}
                            >
                              {label}
                            </button>
                          )
                        })()}
                        <button
                          className="link-button"
                          onClick={() => toggleStarredForPath(selectedPath)}
                          title={starredFaults[selectedPath] ? 'Unstar this image' : 'Star this image'}
                        >
                          {starredFaults[selectedPath] ? '★ Starred' : '☆ Star'}
                        </button>
                      </div>
                    )}
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
                          className={`image-viewport ${viewVariant === 'rgb' ? 'rgb-max-frame' : ''} ${isPanning ? 'panning' : zoom > 1 ? 'can-pan' : ''}`}
                          ref={imageContainerRef}
                          style={size ? { width: `${size.width}px`, height: `${size.height}px` } : undefined}
                        onWheel={(event) => {
                          // RGB alignment: Alt+wheel scales the thermal-to-RGB mapping.
                          if (!(viewVariant === 'rgb' && drawMode === 'select' && event.altKey)) return
                          event.preventDefault()
                          event.stopPropagation()

                          const delta = Math.sign(event.deltaY)
                          const step = event.shiftKey ? 0.05 : 0.01
                          const current = Number.isFinite(rgbAlignScaleEffective) ? rgbAlignScaleEffective : 1
                          const next = current + (delta > 0 ? -step : step)
                          const clamped = Math.min(2, Math.max(0.5, Number(next.toFixed(3))))

                          // If we have pitch/alt context, store an override for this bucket.
                          // Otherwise, keep the original global behavior.
                          if (rgbAlignScaleContextKey) {
                            try {
                              window.localStorage.setItem(getRgbAlignScaleStorageKey(rgbAlignScaleContextKey), String(clamped))
                            } catch {
                              // ignore
                            }
                            setRgbAlignScaleContextTick((v) => v + 1)
                          } else {
                            setRgbAlignScale(clamped)
                          }
                        }}
                        onContextMenu={(event) => {
                          if (activeAction !== 'view' || fileKind !== 'image') return
                          event.preventDefault()
                          event.stopPropagation()

                          const point = toImageCoords(event)
                          if (!point) return

                          const csvName = viewHoverTempsCsvNameRef.current
                          if (!csvName) return

                          void (async () => {
                            let grid = wideTempsCacheRef.current.get(csvName) || null
                            if (!grid) {
                              setViewHoverTempsStatus('loading')
                              grid = await ensureWideTempsGrid(csvName)
                              setViewHoverTempsStatus(grid ? 'ready' : 'missing')
                            }
                            if (!grid) return

                            const px = Math.min(grid.w - 1, Math.max(0, Math.floor(point.x * grid.w)))
                            const py = Math.min(grid.h - 1, Math.max(0, Math.floor(point.y * grid.h)))
                            const v = grid.data[py * grid.w + px]

                            setViewHoverPixel({ x: px, y: py })
                            setViewHoverPixelTempC(v)

                            const text = Number.isFinite(v) ? `${v.toFixed(2)} °C` : 'N/A'
                            const ok = await copyToClipboard(text)
                            if (!ok) console.warn('Failed to copy temperature to clipboard')
                          })()
                        }}
                        onMouseMove={(event) => {
                          if (isRgbAlignDragging && viewVariant === 'rgb' && rgbAlignDragStart && rgbAlignDragOrigin) {
                            const img = imageRef.current
                            if (!img) return
                            const rect = img.getBoundingClientRect()
                            if (!rect.width || !rect.height) return
                            const dx = (event.clientX - rgbAlignDragStart.x) / rect.width
                            const dy = (event.clientY - rgbAlignDragStart.y) / rect.height
                            setRgbAlignOffset({
                              x: rgbAlignDragOrigin.x + dx,
                              y: rgbAlignDragOrigin.y + dy,
                            })
                            return
                          }

                          const point = toImageCoords(event)
                          if (!point) return

                          // Hover temperature (from wide CSV) – throttled via rAF.
                          if (viewHoverTempsStatus === 'ready') {
                            scheduleHoverTempUpdate(point)
                          }

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
                            if (viewVariant === 'rgb') setRgbLabelsExportDirty(true)
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
                            if (viewVariant === 'rgb') setRgbLabelsExportDirty(true)
                          }
                        }}
                        onMouseUp={() => {
                          if (isDrawing && draftLabel) {
                            const minSize = 0.01
                            if (draftLabel.w > minSize && draftLabel.h > minSize) {
                              const next = [
                                ...selectedLabels,
                                {
                                  classId: 9,
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
                              if (!(viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath)) {
                                persistLabels(next)
                              }
                              if (viewVariant === 'rgb') setRgbLabelsExportDirty(true)
                            }
                          }
                          if (isResizingLabel) {
                            const current = selectedLabelsRef.current
                            pushLabelHistory(current)
                            if (!(viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath)) {
                              persistLabels(current)
                            }
                          }
                          if (isDraggingLabel) {
                            const current = selectedLabelsRef.current
                            pushLabelHistory(current)
                            if (!(viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath)) {
                              persistLabels(current)
                            }
                          }
                          if (isPanning) {
                            setIsPanning(false)
                            setPanStart(null)
                            setPanOrigin(null)
                          }
                          if (isRgbAlignDragging) {
                            setIsRgbAlignDragging(false)
                            setRgbAlignDragStart(null)
                            setRgbAlignDragOrigin(null)
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
                          clearHoverTemp()
                          if (isRgbAlignDragging) {
                            setIsRgbAlignDragging(false)
                            setRgbAlignDragStart(null)
                            setRgbAlignDragOrigin(null)
                          }
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
                          style={{
                            transform: `translate(${pan.x}px, ${pan.y}px) scale(${zoom * BASE_ZOOM * (viewVariant === 'rgb' ? RGB_VIEW_IMAGE_SCALE : 1)})`,
                          }}
                        >
                          <img
                            ref={imageRef}
                            src={fileUrl}
                            alt={selectedViewPath || selectedPath}
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
                            onMouseDown={handleDrawSurfaceMouseDown}
                          />
                          {viewVariant === 'thermal' && onlyCenterBoxFaults && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div
                              className="center-box-overlay"
                              style={{ width: imageMetrics.width, height: imageMetrics.height }}
                            >
                              {(() => {
                                const { width: renderW, height: renderH, naturalWidth, naturalHeight } = imageMetrics

                                const safeNaturalW = naturalWidth || renderW
                                const safeNaturalH = naturalHeight || renderH
                                const scale = Math.min(renderW / safeNaturalW, renderH / safeNaturalH)
                                const contentW = safeNaturalW * scale
                                const contentH = safeNaturalH * scale
                                const offsetX = (renderW - contentW) / 2
                                const offsetY = (renderH - contentH) / 2

                                const bounds = getCenteredBoxBoundsPx(safeNaturalW, safeNaturalH)
                                const left = offsetX + bounds.left * scale
                                const top = offsetY + bounds.top * scale
                                const width = Math.max(0, (bounds.right - bounds.left) * scale)
                                const height = Math.max(0, (bounds.bottom - bounds.top) * scale)

                                return <div className="center-box-rect" style={{ left, top, width, height }} />
                              })()}
                            </div>
                          )}
                          {viewVariant === 'thermal' && showLabels && selectedLabels.length > 0 && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div
                              className="label-overlay"
                              style={{ width: imageMetrics.width, height: imageMetrics.height }}
                            >
                          {selectedLabels.map((label, index) => {
                            const { width: renderW, height: renderH, naturalWidth, naturalHeight } = imageMetrics

                            const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1

                            const safeNaturalW = naturalWidth || renderW
                            const safeNaturalH = naturalHeight || renderH
                            const scale = Math.min(renderW / safeNaturalW, renderH / safeNaturalH)
                            const contentW = safeNaturalW * scale
                            const contentH = safeNaturalH * scale
                            const offsetX = (renderW - contentW) / 2
                            const offsetY = (renderH - contentH) / 2

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
                          {viewVariant === 'thermal' && showLabels && drawMode === 'select' && selectedLabelIndex !== null && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div
                              className="label-handles"
                              style={{ width: imageMetrics.width, height: imageMetrics.height }}
                            >
                          {(() => {
                            const label = selectedLabels[selectedLabelIndex]
                            if (!label) return null

                            const { width: renderW, height: renderH, naturalWidth, naturalHeight } = imageMetrics
                            const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1

                            const safeNaturalW = naturalWidth || renderW
                            const safeNaturalH = naturalHeight || renderH
                            const scale = Math.min(renderW / safeNaturalW, renderH / safeNaturalH)
                            const contentW = safeNaturalW * scale
                            const contentH = safeNaturalH * scale
                            const offsetX = (renderW - contentW) / 2
                            const offsetY = (renderH - contentH) / 2

                            const centerX = isNormalized ? label.x * safeNaturalW : label.x
                            const centerY = isNormalized ? label.y * safeNaturalH : label.y
                            const boxW = isNormalized ? label.w * safeNaturalW : label.w
                            const boxH = isNormalized ? label.h * safeNaturalH : label.h

                            const left = offsetX + (centerX - boxW / 2) * scale
                            const top = offsetY + (centerY - boxH / 2) * scale
                            const width = boxW * scale
                            const height = boxH * scale
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
                          {viewVariant === 'thermal' && showLabels && draftLabel && imageMetrics.width > 0 && imageMetrics.height > 0 && (
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

                          {viewVariant === 'rgb' && rgbLabelsAreImageSpace && showLabels && selectedLabels.length > 0 && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div className="label-overlay" style={{ width: imageMetrics.width, height: imageMetrics.height }}>
                              {selectedLabels.map((label, index) => {
                                const { width: renderW, height: renderH } = imageMetrics
                                const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
                                const centerX = isNormalized ? label.x * renderW : label.x
                                const centerY = isNormalized ? label.y * renderH : label.y
                                const boxW = isNormalized ? label.w * renderW : label.w
                                const boxH = isNormalized ? label.h * renderH : label.h
                                const left = centerX - boxW / 2
                                const top = centerY - boxH / 2

                                return (
                                  <div
                                    key={`${label.classId}-${index}`}
                                    className={`label-box ${label.shape === 'ellipse' ? 'label-box-ellipse' : ''} ${selectedLabelIndex === index ? 'label-box-selected' : ''}`}
                                    style={{ left, top, width: boxW, height: boxH }}
                                  />
                                )
                              })}
                            </div>
                          )}

                          {viewVariant === 'rgb' && rgbLabelsAreImageSpace && showLabels && drawMode === 'select' && selectedLabelIndex !== null && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div className="label-handles" style={{ width: imageMetrics.width, height: imageMetrics.height }}>
                              {(() => {
                                const label = selectedLabels[selectedLabelIndex]
                                if (!label) return null

                                const { width: renderW, height: renderH } = imageMetrics
                                const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
                                const centerX = isNormalized ? label.x * renderW : label.x
                                const centerY = isNormalized ? label.y * renderH : label.y
                                const boxW = isNormalized ? label.w * renderW : label.w
                                const boxH = isNormalized ? label.h * renderH : label.h
                                const left = centerX - boxW / 2
                                const top = centerY - boxH / 2
                                const right = left + boxW
                                const bottom = top + boxH
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

                          {viewVariant === 'rgb' && rgbLabelsAreImageSpace && showLabels && draftLabel && imageMetrics.width > 0 && imageMetrics.height > 0 && (
                            <div className="manual-overlay" style={{ width: imageMetrics.width, height: imageMetrics.height }}>
                              {(() => {
                                const { width: renderW, height: renderH } = imageMetrics
                                const left = (draftLabel.x - draftLabel.w / 2) * renderW
                                const top = (draftLabel.y - draftLabel.h / 2) * renderH
                                const width = draftLabel.w * renderW
                                const height = draftLabel.h * renderH
                                return (
                                  <div
                                    className={`manual-label ${draftLabel.shape === 'ellipse' ? 'ellipse' : ''}`}
                                    style={{ left, top, width, height }}
                                  />
                                )
                              })()}
                            </div>
                          )}
                        </div>

                        {/* RGB overlays must be in viewport space (locked 1536×1229 frame), not inside the transformed image-zoom. */}
                        {viewVariant === 'rgb' && !rgbLabelsAreImageSpace && (
                          <>
                            <div className={`draw-surface ${drawMode}`} onMouseDown={handleDrawSurfaceMouseDown} />

                            {onlyCenterBoxFaults && (size?.width ?? 0) > 0 && (size?.height ?? 0) > 0 && (
                              <div className="center-box-overlay" style={{ width: size?.width, height: size?.height }}>
                                {(() => {
                                  const renderW = size?.width ?? 0
                                  const renderH = size?.height ?? 0
                                  const frame = getRgbThermalFrameRect(renderW, renderH)
                                  if (!frame) return null
                                  const offset = getRgbOverlayOffsetPx(renderW, renderH)
                                  const bounds = getCenteredBoxBoundsPx(frame.thermalW, frame.thermalH)
                                  const scale = frame.inner.width / frame.thermalW
                                  const left = frame.inner.left + offset.x + bounds.left * scale
                                  const top = frame.inner.top + offset.y + bounds.top * scale
                                  const width = Math.max(0, (bounds.right - bounds.left) * scale)
                                  const height = Math.max(0, (bounds.bottom - bounds.top) * scale)
                                  return <div className="center-box-rect" style={{ left, top, width, height }} />
                                })()}
                              </div>
                            )}

                            {showLabels && selectedLabels.length > 0 && (size?.width ?? 0) > 0 && (size?.height ?? 0) > 0 && (
                              <div className="label-overlay" style={{ width: size?.width, height: size?.height }}>
                                {selectedLabels.map((label, index) => {
                                  const renderW = size?.width ?? 0
                                  const renderH = size?.height ?? 0
                                  const frame = getRgbThermalFrameRect(renderW, renderH)
                                  if (!frame) return null
                                  const offset = getRgbOverlayOffsetPx(renderW, renderH)

                                  const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
                                  const centerX = isNormalized ? label.x * frame.inner.width : label.x
                                  const centerY = isNormalized ? label.y * frame.inner.height : label.y
                                  const boxW = isNormalized ? label.w * frame.inner.width : label.w
                                  const boxH = isNormalized ? label.h * frame.inner.height : label.h

                                  const left = frame.inner.left + offset.x + (centerX - boxW / 2)
                                  const top = frame.inner.top + offset.y + (centerY - boxH / 2)
                                  const width = boxW
                                  const height = boxH

                                  return (
                                    <div
                                      key={`${label.classId}-${index}`}
                                      className={`label-box ${label.shape === 'ellipse' ? 'label-box-ellipse' : ''} ${selectedLabelIndex === index ? 'label-box-selected' : ''}`}
                                      style={{ left, top, width, height }}
                                    />
                                  )
                                })}
                              </div>
                            )}

                            {showLabels && drawMode === 'select' && selectedLabelIndex !== null && (size?.width ?? 0) > 0 && (size?.height ?? 0) > 0 && (
                              <div className="label-handles" style={{ width: size?.width, height: size?.height }}>
                                {(() => {
                                  const label = selectedLabels[selectedLabelIndex]
                                  if (!label) return null

                                  const renderW = size?.width ?? 0
                                  const renderH = size?.height ?? 0
                                  const frame = getRgbThermalFrameRect(renderW, renderH)
                                  if (!frame) return null
                                  const offset = getRgbOverlayOffsetPx(renderW, renderH)

                                  const isNormalized = label.x <= 1 && label.y <= 1 && label.w <= 1 && label.h <= 1
                                  const centerX = isNormalized ? label.x * frame.inner.width : label.x
                                  const centerY = isNormalized ? label.y * frame.inner.height : label.y
                                  const boxW = isNormalized ? label.w * frame.inner.width : label.w
                                  const boxH = isNormalized ? label.h * frame.inner.height : label.h
                                  const left = frame.inner.left + offset.x + (centerX - boxW / 2)
                                  const top = frame.inner.top + offset.y + (centerY - boxH / 2)
                                  const width = boxW
                                  const height = boxH
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

                            {showLabels && draftLabel && (size?.width ?? 0) > 0 && (size?.height ?? 0) > 0 && (
                              <div className="manual-overlay" style={{ width: size?.width, height: size?.height }}>
                                {(() => {
                                  const renderW = size?.width ?? 0
                                  const renderH = size?.height ?? 0
                                  const frame = getRgbThermalFrameRect(renderW, renderH)
                                  if (!frame) return null
                                  const offset = getRgbOverlayOffsetPx(renderW, renderH)

                                  const left = frame.inner.left + offset.x + (draftLabel.x - draftLabel.w / 2) * frame.inner.width
                                  const top = frame.inner.top + offset.y + (draftLabel.y - draftLabel.h / 2) * frame.inner.height
                                  const width = draftLabel.w * frame.inner.width
                                  const height = draftLabel.h * frame.inner.height

                                  return (
                                    <div
                                      className={`manual-label ${draftLabel.shape === 'ellipse' ? 'ellipse' : ''}`}
                                      style={{ left, top, width, height }}
                                    />
                                  )
                                })()}
                              </div>
                            )}
                          </>
                        )}

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
                  {/* <p style={{ marginTop: 0 }}>
                    Edit <strong>faults.txt</strong> (one image path per line). This does not delete images; it only updates the list.
                  </p> */}
                  <div className="report-toolbar">
                    {/* <button className="link-button" onClick={() => loadReportFromDisk()}>
                      Reload
                    </button> */}

                    <button
                      className="link-button"
                      onClick={() => setReportAdvancedOpen((prev) => !prev)}
                      title={reportAdvancedOpen ? 'Hide advanced report actions' : 'Show advanced report actions'}
                    >
                      Advanced
                    </button>

                    {reportAdvancedOpen && (
                      <>
                        <button
                          className="link-button"
                          onClick={() => {
                            const normalized = normalizeFaultsText((docxDraft.length ? docxDraft.map((d) => d.path) : normalizeFaultsText(reportText)).join('\n'))
                            setReportText(normalized.join('\n'))
                            const printable = filterPathsForReport(normalized)
                            setDocxDraft((prev) => buildDocxDraftFromPaths(printable, prev))
                          }}
                          title="Normalize and de-duplicate the report image list"
                        >
                          Normalize list
                        </button>
                        <button className="link-button" onClick={() => saveReportToDisk()} title="Save the current list to workspace/thermal_faults/faults.txt">
                          Save list
                        </button>
                        <button className="link-button" onClick={() => syncDocxDraftFromEditor()} title="Rebuild the editable DOCX draft from the current list">
                          Sync draft
                        </button>
                        <button className="link-button" onClick={() => void refreshReportMetadataTables()} title="Recompute and refresh metadata tables in the draft">
                          Refresh metadata
                        </button>
                      </>
                    )}

                    <button className="link-button" onClick={() => previewWordReport()} title="Generate an in-app DOCX preview">
                      Preview
                    </button>
                    <button className="link-button" onClick={() => downloadWordReport()} title="Download the report as a .docx file">
                      Download DOCX
                    </button>
                    <button className="link-button" onClick={() => downloadPdfFromPreview()} title="Download the report as a PDF (uses the preview renderer)">
                      Download PDF
                    </button>

                    <label
                      style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 8 }}
                      title={`Include only images with at least one label inside the ${CENTER_BOX_W}×${CENTER_BOX_H} center box`}
                    >
                      <input
                        type="checkbox"
                        checked={onlyCenterBoxFaults}
                        onChange={(e) => handleToggleOnlyCenterBoxFaults(e.target.checked)}
                      />
                      Centered
                    </label>
                    <label
                      style={{ display: 'inline-flex', alignItems: 'center', gap: 6, marginLeft: 8 }}
                      title="Include only starred images in the report"
                    >
                      <input
                        type="checkbox"
                        checked={onlyStarredInReport}
                        onChange={(e) => handleToggleOnlyStarredInReport(e.target.checked)}
                      />
                      Starred
                    </label>
                  </div>
                  <div ref={reportSplitRef} className="report-split">
                    <div className="report-left">
                      <div className="report-left-window report-left-window-combined">
                        <div className="help-tabs" style={{ marginBottom: 10, flexWrap: 'wrap' }}>
                          <button
                            className={`link-button help-lang-button ${reportLeftTab === 'description' ? 'active' : ''}`}
                            onClick={() => setReportLeftTab('description')}
                          >
                            Description
                          </button>
                          <button
                            className={`link-button help-lang-button ${reportLeftTab === 'equipment' ? 'active' : ''}`}
                            onClick={() => setReportLeftTab('equipment')}
                          >
                            Equipment
                          </button>

                          {reportChapters.map((c, idx) => {
                            const id = `chapter:${c.id}` as const
                            const label = (c.chapterTitle || '').trim() || `New ${idx + 1}`
                            return (
                              <span key={c.id} style={{ display: 'inline-flex', alignItems: 'center', gap: 4 }}>
                                <button
                                  className={`link-button help-lang-button ${reportLeftTab === id ? 'active' : ''}`}
                                  onClick={() => setReportLeftTab(id)}
                                  title={label}
                                  type="button"
                                >
                                  {label}
                                </button>
                                <button
                                  className="tiny-button danger"
                                  onClick={() => removeReportChapter(c.id)}
                                  title="Delete tab/chapter"
                                  type="button"
                                >
                                  ✕
                                </button>
                              </span>
                            )
                          })}

                          <button className="link-button help-lang-button" onClick={() => addReportChapter()} title="Add a new chapter">
                            Add
                          </button>
                        </div>

                        {reportLeftTab === 'description' ? (
                          <div className="report-left-section">
                            <div className="report-sidebar-title explorer-pane-title">Description</div>
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
                                    <button
                                      className="tiny-button danger"
                                      onClick={() => {
                                        const ok = window.confirm('Delete this description block?')
                                        if (!ok) return
                                        removeDescriptionRow(row.id)
                                      }}
                                      title="Remove"
                                      type="button"
                                    >
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
                        ) : null}

                        {reportLeftTab === 'equipment' ? (
                          <div className="report-left-section">
                            <div className="report-sidebar-title explorer-pane-title">Equipment</div>
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
                                      onClick={() => {
                                        const ok = window.confirm('Delete this equipment item?')
                                        if (!ok) return
                                        removeEquipmentItem(item.id)
                                      }}
                                      title="Remove"
                                      type="button"
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
                                          onClick={() => {
                                            const ok = window.confirm('Remove this image?')
                                            if (!ok) return
                                            void setEquipmentItemImage(item.id, null)
                                          }}
                                          title="Remove image"
                                          type="button"
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
                        ) : null}

                        {reportLeftTab.startsWith('chapter:') ? (() => {
                          const chapterId = reportLeftTab.slice('chapter:'.length)
                          const chapter = reportChapters.find((c) => c.id === chapterId) || null
                          if (!chapter) return null

                          const missingTitle = !(chapter.chapterTitle || '').trim()
                          return (
                            <div className="report-left-section">
                              <div className="report-sidebar-title explorer-pane-title">Chapter</div>
                              <div className="report-sidebar-subtitle">This becomes a new chapter in the PDF. Image is auto-scaled to fit.</div>

                              {missingTitle ? (
                                <div className="report-error" style={{ marginBottom: 10 }}>
                                  Chapter/Tab title is required.
                                </div>
                              ) : null}

                              <label className="description-label">
                                Chapter / Tab title (required)
                                <input
                                  className="description-input"
                                  value={chapter.chapterTitle}
                                  onChange={(e) => updateReportChapter(chapter.id, { chapterTitle: e.target.value })}
                                  placeholder="e.g. Findings / Notes / Appendix"
                                />
                              </label>

                              <div style={{ display: 'flex', gap: 8, alignItems: 'center', marginBottom: 10 }}>
                                <button className="link-button" onClick={() => addReportChapterSection(chapter.id)} type="button">
                                  + Add subchapter
                                </button>
                                <div className="docx-draft-meta">{(chapter.sections || []).length} block(s)</div>
                              </div>

                              <div className="description-rows">
                                {(chapter.sections || []).map((section, idx) => (
                                  <div key={section.id} className="description-row">
                                    <div className="description-row-header">
                                      <div className="description-row-title">{idx + 1}.</div>
                                      <button
                                        className="tiny-button danger"
                                        onClick={() => removeReportChapterSection(chapter.id, section.id)}
                                        title="Remove"
                                        type="button"
                                      >
                                        ✕
                                      </button>
                                    </div>

                                    <label className="description-label">
                                      Title (optional)
                                      <input
                                        className="description-input"
                                        value={section.title}
                                        onChange={(e) => updateReportChapterSection(chapter.id, section.id, { title: e.target.value })}
                                        placeholder="Optional title shown inside the PDF"
                                      />
                                    </label>

                                    <label className="description-label">
                                      Text (optional)
                                      <textarea
                                        className="description-textarea"
                                        value={section.text}
                                        onChange={(e) => updateReportChapterSection(chapter.id, section.id, { text: e.target.value })}
                                        placeholder="Optional text…"
                                        rows={6}
                                      />
                                    </label>

                                    <div className="equipment-row-image">
                                      <input
                                        className="equipment-row-file"
                                        type="file"
                                        accept="image/*"
                                        onChange={(e) => {
                                          const f = e.target.files?.[0] || null
                                          void setReportChapterSectionImage(chapter.id, section.id, f)
                                        }}
                                      />
                                      {section.imagePreviewUrl ? (
                                        <div className="equipment-row-image-preview">
                                          <img src={section.imagePreviewUrl} alt="subchapter" />
                                          <button
                                            className="equipment-row-clear-image"
                                            onClick={() => {
                                              const ok = window.confirm('Remove this image?')
                                              if (!ok) return
                                              void setReportChapterSectionImage(chapter.id, section.id, null)
                                            }}
                                            title="Remove image"
                                            type="button"
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
                          )
                        })() : null}
                      </div>
                    </div>

                    <div className="report-main">
                      <div className="docx-preview">
                        <div className="docx-preview-header">
                          <strong>PDF preview</strong>
                          <span className="docx-preview-hint">(generated via LibreOffice; matches downloaded PDF)</span>
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
                        <div className="report-sidebar-title explorer-pane-title">DOCX draft</div>
                        <div className="report-sidebar-subtitle">Edit captions/fields/order before export</div>
                        <label
                          style={{ display: 'inline-flex', alignItems: 'center', gap: 6, margin: '4px 0 10px 0' }}
                          title="When enabled, each Thermal image will include its paired RGB image under it (when available)."
                        >
                          <input type="checkbox" checked={includeRgbInDocx} onChange={(e) => setIncludeRgbInDocx(e.target.checked)} />
                          Include RGBs
                        </label>
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
                                    {idx + 1}. {fileStemFromPath(item.path)}
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
                                    <button
                                      className="tiny-button danger"
                                      onClick={() => {
                                        const label = fileStemFromPath(item.path)
                                        const ok = window.confirm(`Remove “${label}” from the DOCX draft?`)
                                        if (!ok) return
                                        removeDocxDraftItem(item.id)
                                      }}
                                      type="button"
                                    >
                                      Remove
                                    </button>
                                  </div>
                                </div>
                                <div className="docx-draft-fields">
                                  <label className="docx-field">
                                    <div className="docx-field-label">Caption</div>
                                    <DraftTextInput
                                      className="docx-field-input"
                                      value={item.caption}
                                      onCommit={(next) => updateDocxDraftItem(item.id, { caption: next })}
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
                                              <DraftTextInput
                                                className="docx-field-input"
                                                value={table.title}
                                                onCommit={(next) => updateDocxDraftTableTitleForAllImages(item.id, table.id, next)}
                                                placeholder={`Table ${tableIdx + 1} title (optional)`}
                                              />
                                            </div>
                                            <button
                                              className="tiny-button danger"
                                              onClick={() => {
                                                const ok = window.confirm('Remove this table?')
                                                if (!ok) return
                                                removeDocxDraftTable(item.id, table.id)
                                              }}
                                              type="button"
                                              title="Remove table"
                                            >
                                              Remove table
                                            </button>
                                          </div>

                                          {(table.rows || []).map((row) => (
                                            <div key={row.id} className="docx-kv-row">
                                              <DraftTextInput
                                                className="docx-field-input docx-kv-input"
                                                value={row.name}
                                                onCommit={(next) => updateDocxDraftTableRow(item.id, table.id, row.id, { name: next })}
                                                placeholder="Field name"
                                              />
                                              <DraftTextInput
                                                className="docx-field-input docx-kv-input"
                                                value={row.description}
                                                onCommit={(next) => {
                                                  const isDistance =
                                                    (table.title || '').trim().toLowerCase() === 'thermal parameters' &&
                                                    (row.name || '').trim().toLowerCase() === 'distance'

                                                  if (isDistance) {
                                                    updateDocxDraftRowDescriptionForAllImagesByTableTitleAndRowName('Thermal parameters', 'Distance', next)
                                                    return
                                                  }

                                                  updateDocxDraftTableRow(item.id, table.id, row.id, { description: next })
                                                }}
                                                placeholder="Description"
                                              />
                                              <button
                                                className="tiny-button danger"
                                                onClick={() => {
                                                  const ok = window.confirm('Remove this row?')
                                                  if (!ok) return
                                                  removeDocxDraftTableRow(item.id, table.id, row.id)
                                                }}
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

          {activeMenu === 'help' && (
            <section className="card card-help">
              <div className="page-header-row">
                <h2>
                  {activeHelpPage === 'about'
                    ? helpLang === 'en'
                      ? 'About'
                      : 'About'
                    : helpLang === 'en'
                      ? 'Documentation'
                      : 'Documentation'}
                </h2>
                <div className="page-header-right">
                  <div className="help-tabs">
                    {activeHelpPage === 'documentation' ? (
                      <button
                        className="link-button"
                        onClick={() => setActiveHelpPage('about')}
                        type="button"
                        title={helpLang === 'en' ? 'Application information' : 'Πληροφορίες εφαρμογής'}
                      >
                        About
                      </button>
                    ) : (
                      <button
                        className="link-button"
                        onClick={() => setActiveHelpPage('documentation')}
                        type="button"
                        title={helpLang === 'en' ? 'How to use the application' : 'Οδηγός χρήσης'}
                      >
                        Documentation
                      </button>
                    )}

                    <div className="help-lang">
                      {helpLang === 'en' ? (
                        <button
                          className="link-button help-lang-button"
                          onClick={() => setHelpLang('el')}
                          type="button"
                          title="Ελληνικά"
                        >
                          EL
                        </button>
                      ) : (
                        <button
                          className="link-button help-lang-button"
                          onClick={() => setHelpLang('en')}
                          type="button"
                          title="English"
                        >
                          EN
                        </button>
                      )}
                    </div>
                  </div>
                </div>
              </div>

              {activeHelpPage === 'documentation' ? (
                <div className="help-doc-scroll">
                  <div className="help-doc">
                    <p className="help-lead">
                      {helpLang === 'en'
                        ? 'This guide explains, in simple steps, how to use the application to view Thermal/RGB images, manage labels, and create reports.'
                        : 'Αυτός ο οδηγός εξηγεί με απλά βήματα πώς να χρησιμοποιείτε την εφαρμογή για προβολή θερμικών/RGB εικόνων, επισήμανση (labels) και δημιουργία report.'}
                    </p>

                    <h3 id="doc-start">
                      {helpLang === 'en' ? '1) Getting started (Choose folder)' : '1) Ξεκίνημα (Επιλογή φακέλου)'}
                    </h3>
                    <ol>
                      <li>
                        {helpLang === 'en' ? (
                          <>
                            Click <b>Home → Choose Folder...</b> and select the <b>parent</b> folder that contains <b>exactly</b> these subfolders:
                            <b> rgb</b>, <b>thermal</b>, <b>tiff</b>.
                          </>
                        ) : (
                          <>
                            Πατήστε <b>Home → Choose Folder...</b> και επιλέξτε τον <b>φάκελο parent</b> που περιέχει <b>ακριβώς</b> τους υποφακέλους:
                            <b> rgb</b>, <b>thermal</b>, <b>tiff</b>.
                          </>
                        )}
                      </li>
                      <li>
                        {helpLang === 'en'
                          ? 'When asked for folder access (Windows/Browser prompt), choose Allow.'
                          : 'Όταν σας ζητηθεί άδεια πρόσβασης (Windows/Browser prompt), επιλέξτε Allow.'}
                      </li>
                      <li>
                        {helpLang === 'en'
                          ? 'The application uses/updates a workspace folder for generated files (labels, lists, etc.).'
                          : 'Η εφαρμογή χρησιμοποιεί/ενημερώνει έναν φάκελο workspace για τα παραγόμενα αρχεία (labels, λίστες κ.λπ.).'}
                      </li>
                    </ol>

                    <h3 id="doc-explorer">{helpLang === 'en' ? '2) Explorer (Left)' : '2) Explorer (Αριστερά)'}</h3>
                    <ul>
                      <li>{helpLang === 'en' ? 'Select images/files from here.' : 'Από εδώ επιλέγετε εικόνες/αρχεία.'}</li>
                      <li>
                        {helpLang === 'en' ? (
                          <>
                            On the <b>View</b> page you will also see the <b>Fault labels</b> list (right side) for the selected image.
                          </>
                        ) : (
                          <>
                            Στη σελίδα <b>View</b> θα δείτε και τη λίστα <b>Fault labels</b> (δεξιά) για τη συγκεκριμένη εικόνα.
                          </>
                        )}
                      </li>
                    </ul>

                    <h3 id="doc-view">{helpLang === 'en' ? '3) View (Preview & labels)' : '3) View (Προβολή & Labels)'}</h3>
                    <ul>
                      <li>
                        {helpLang === 'en'
                          ? 'Go to Actions → View and select a thermal image from the Explorer.'
                          : 'Πηγαίνετε Actions → View και επιλέξτε μία θερμική εικόνα από τον Explorer.'}
                      </li>
                      <li>
                        {helpLang === 'en' ? 'Using the toolbar above the image, you can:' : 'Με τα κουμπιά πάνω από την εικόνα μπορείτε να:'}
                        <ul>
                          <li>{helpLang === 'en' ? <><b>Zoom</b> (+/−) and reset to 100%.</> : <><b>Zoom</b> (+/−) και επαναφορά στο 100%.</>}</li>
                          <li>
                            {helpLang === 'en'
                              ? <><b>Show/Hide labels</b> to toggle overlays.</>
                              : <><b>Show/Hide labels</b> για να εμφανίζονται/κρύβονται τα overlays.</>}
                          </li>
                          <li>
                            {helpLang === 'en'
                              ? <>Switch <b>Thermal ↔ RGB</b> (when a matching RGB image exists).</>
                              : <>Εναλλαγή <b>Thermal ↔ RGB</b> (όπου υπάρχει αντίστοιχη RGB εικόνα).</>}
                          </li>
                        </ul>
                      </li>
                      <li>
                        {helpLang === 'en' ? (
                          <>
                            Labels are saved to a <b>.txt</b> file inside <b>workspace</b> and appear automatically in the right panel.
                          </>
                        ) : (
                          <>
                            Τα labels αποθηκεύονται σε αρχείο <b>.txt</b> μέσα στο <b>workspace</b> και εμφανίζονται αυτόματα στο δεξί panel.
                          </>
                        )}
                      </li>
                    </ul>

                    <h3 id="doc-scan">{helpLang === 'en' ? '4) Scanner (Find faults)' : '4) Scanner (Εντοπισμός faults)'}</h3>
                    <ol>
                      <li>{helpLang === 'en' ? <>Go to <b>Actions → Scanner</b>.</> : <>Πηγαίνετε <b>Actions → Scanner</b>.</>}</li>
                      <li>{helpLang === 'en' ? <>Click <b>Scan</b> and wait for it to finish.</> : <>Πατήστε <b>Scan</b> και περιμένετε να ολοκληρωθεί.</>}</li>
                      <li>
                        {helpLang === 'en'
                          ? <>After scanning, faults/labels become available and you can review them in <b>View</b>.</>
                          : <>Μετά το scan, εμφανίζονται τα faults/labels και μπορείτε να τα δείτε στο <b>View</b>.</>}
                      </li>
                    </ol>

                    <h3 id="doc-report">{helpLang === 'en' ? '5) Report (Compose & export)' : '5) Report (Σύνταξη & Export)'}</h3>
                    <ul>
                      <li>{helpLang === 'en' ? <>Go to <b>Actions → Report</b>.</> : <>Πηγαίνετε <b>Actions → Report</b>.</>}</li>
                      <li>
                        {helpLang === 'en'
                          ? <>Use the filters (e.g. <b>Starred</b>) to build the list that goes into the report.</>
                          : <>Χρησιμοποιήστε τα φίλτρα (π.χ. <b>Starred</b>) για να φτιάξετε τη λίστα που θα μπει στο report.</>}
                      </li>
                      <li>
                        {helpLang === 'en'
                          ? <>In Report you can add notes/descriptions and export to a file (DOCX/PDF depending on the available buttons).</>
                          : <>Στο report μπορείτε να κρατάτε σημειώσεις/περιγραφές και να κάνετε export σε αρχείο (DOCX/PDF ανάλογα με τα κουμπιά).</>}
                      </li>
                    </ul>

                    <h3 id="doc-tips">{helpLang === 'en' ? '6) Common issues / tips' : '6) Συχνά θέματα / Tips'}</h3>
                    <ul>
                      <li>
                        {helpLang === 'en' ? (
                          <>
                            <b>I cannot see labels:</b> make sure you opened the parent folder (the one containing <b>rgb/thermal/tiff</b> and <b>workspace</b>).
                          </>
                        ) : (
                          <>
                            <b>Δεν βλέπω labels:</b> ελέγξτε ότι ανοίξατε τον φάκελο parent (αυτόν που περιέχει τα <b>rgb/thermal/tiff</b> και το <b>workspace</b>).
                          </>
                        )}
                      </li>
                      <li>
                        {helpLang === 'en' ? (
                          <>
                            <b>It does not ask/keep permissions:</b> prefer <b>Microsoft Edge</b> or <b>Google Chrome</b>.
                          </>
                        ) : (
                          <>
                            <b>Δεν ζητάει/δεν κρατάει άδεια:</b> προτιμήστε <b>Microsoft Edge</b> ή <b>Google Chrome</b>.
                          </>
                        )}
                      </li>
                      <li>
                        {helpLang === 'en' ? (
                          <>
                            <b>Updates feel slow:</b> the app watches <b>workspace</b> in near real-time (polling) — usually within a few seconds.
                          </>
                        ) : (
                          <>
                            <b>Αργεί να ενημερωθεί:</b> η εφαρμογή διαβάζει το <b>workspace</b> σε πραγματικό χρόνο (polling) — συνήθως σε λίγα δευτερόλεπτα.
                          </>
                        )}
                      </li>
                    </ul>
                  </div>
                </div>
              ) : (
                <div className="help-doc-scroll">
                  <div className="help-doc">
                    <p className="help-lead">
                      {helpLang === 'en'
                        ? 'Application for viewing thermal/RGB imagery, managing labels/faults, and creating reports.'
                        : 'Εφαρμογή για προβολή θερμικών/RGB εικόνων, διαχείριση labels/faults και δημιουργία report.'}
                    </p>

                    <h3>{helpLang === 'en' ? 'Versions' : 'Versions'}</h3>
                    <div className="help-kv">
                      <div className="help-kv-row">
                        <div className="help-kv-key">Frontend</div>
                        <div className="help-kv-value">
                          {helpLang === 'en' ? (
                            <>
                              App: <b>{__APP_VERSION__}</b> · React: <b>{__REACT_VERSION__}</b> · Vite: <b>{__VITE_VERSION__}</b> · TypeScript:{' '}
                              <b>{__TS_VERSION__}</b> · Build: <b>{buildTimeLocal}</b>
                            </>
                          ) : (
                            <>
                              App: <b>{__APP_VERSION__}</b> · React: <b>{__REACT_VERSION__}</b> · Vite: <b>{__VITE_VERSION__}</b> · TypeScript:{' '}
                              <b>{__TS_VERSION__}</b> · Build: <b>{buildTimeLocal}</b>
                            </>
                          )}
                        </div>
                      </div>

                      <div className="help-kv-row">
                        <div className="help-kv-key">Backend</div>
                        <div className="help-kv-value">
                          {backendHealth?.versions ? (
                            <>
                              {Object.entries(backendHealth.versions)
                                .sort(([a], [b]) => a.localeCompare(b))
                                .map(([k, v], idx, arr) => (
                                  <span key={k}>
                                    {k}: <b>{v}</b>
                                    {idx < arr.length - 1 ? ' · ' : ''}
                                  </span>
                                ))}
                            </>
                          ) : (
                            <>
                              {helpLang === 'en'
                                ? backendHealthError || 'Not available (start FastAPI).'
                                : backendHealthError || 'Δεν είναι διαθέσιμο (ξεκινήστε το FastAPI).'}
                            </>
                          )}
                        </div>
                      </div>
                    </div>

                    <h3>{helpLang === 'en' ? 'Developed by' : 'Αναπτύχθηκε από'}</h3>
                    <div className="help-kv">
                      <div className="help-kv-row">
                        <div className="help-kv-key">Name</div>
                        <div className="help-kv-value">{helpLang === 'en' ? <b>Gavriil Ilikidis</b> : <b>Γαβριήλ Ηλικίδης</b>}</div>
                      </div>
                      <div className="help-kv-row">
                        <div className="help-kv-key">Company</div>
                        <div className="help-kv-value">
                          {helpLang === 'en' ? (
                            <>
                              Built for <b>TOPOSOL</b>
                            </>
                          ) : (
                            <>
                              Δημιουργήθηκε για την <b>TOPOSOL</b>
                            </>
                          )}
                        </div>
                      </div>
                      <div className="help-kv-row">
                        <div className="help-kv-key">Contact</div>
                        <div className="help-kv-value">
                          {helpLang === 'en' ? (
                            <>
                              Support: <b>gabrielilikidis@gmail.com</b>
                            </>
                          ) : (
                            <>
                              Υποστήριξη: <b>gabrielilikidis@gmail.com</b>
                            </>
                          )}
                        </div>
                      </div>
                    </div>

                    <div className="help-footnote">
                      {helpLang === 'en' ? (
                        <>
                          For usage instructions: <b>Help → Documentation</b>. For folder access issues: try <b>Edge</b>/<b>Chrome</b>.
                        </>
                      ) : (
                        <>
                          Για χρήση: <b>Help → Documentation</b>. Για θέματα πρόσβασης φακέλων: δοκιμάστε <b>Edge</b>/<b>Chrome</b>.
                        </>
                      )}
                    </div>
                  </div>
                </div>
              )}
            </section>
          )}

          {activeMenu !== 'file' && activeMenu !== 'actions' && activeMenu !== 'help' && (
            <section className="card">
              <h2>Status</h2>
              <p>{message}</p>
            </section>
          )}

          

          {activeMenu !== 'actions' && activeMenu !== 'file' && activeMenu !== 'help' && (
            <section className="card viewer">
              <h3>Preview</h3>
              {!selectedPath && <p>Select a file from the explorer.</p>}
              {selectedPath && fileKind === 'image' && fileUrl && (
                <img src={fileUrl} alt={selectedViewPath || selectedPath} className="preview-image" />
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
                    <div className="explorer-pane-title">Fault labels</div>
                    <div className="tree">
                      {selectedLabels.length === 0 && (
                        <div className="explorer-empty">
                          {openedThermalFolderDirectly
                            ? 'You opened the “thermal” folder directly. Open the parent folder that contains “thermal/” and “workspace/” so the app can read workspace/thermal_labels.'
                            : openedRgbFolderDirectly
                              ? 'You opened the “rgb” folder directly. Open the parent folder that contains “thermal/” (and “workspace/”) so the app can read workspace/thermal_labels.'
                              : (() => {
                                  if (!selectedPath) return 'No faults for this image.'
                                  const expected = getExpectedThermalLabelPaths(selectedPath)
                                  const lines = [
                                    'No faults for this image.',
                                    `Expected: ${expected.flat} (${expected.flatExists ? 'loaded' : 'missing'})`,
                                  ]
                                  if (expected.nested) {
                                    lines.push(`Nested: ${expected.nested} (${expected.nestedExists ? 'loaded' : 'missing'})`)
                                  }
                                  return lines.join('\n')
                                })()}
                        </div>
                      )}
                      {selectedLabels.map((label, index) => (
                        <button
                          key={`fault-label-${index}`}
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
                              <span
                                className="fault-type-wrap"
                                onClick={(e) => {
                                  e.preventDefault()
                                  e.stopPropagation()
                                }}
                              >
                                <select
                                  value={String(normalizeFaultTypeId(label.classId))}
                                  title="Defect type"
                                  className="fault-type-select"
                                  onChange={(e) => {
                                    e.preventDefault()
                                    e.stopPropagation()
                                    const nextId = normalizeFaultTypeId(Number(e.target.value))
                                    const next = selectedLabels.map((l, i) => (i === index ? { ...l, classId: nextId } : l))
                                    selectedLabelsRef.current = next
                                    setSelectedLabels(next)
                                    pushLabelHistory(next)
                                    if (selectedPath) syncReportFaultDescriptionsForPath(selectedPath, next)
                                    if (!(viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath)) {
                                      flushRealtimePersist(next)
                                    }
                                    if (viewVariant === 'rgb') setRgbLabelsExportDirty(true)
                                  }}
                                >
                                  {FAULT_TYPE_OPTIONS.map((opt) => (
                                    <option key={opt.id} value={opt.id}>
                                      {opt.label}
                                    </option>
                                  ))}
                                </select>
                              </span>
                            </span>
                          </span>
                        </button>
                      ))}
                    </div>
                    <div className="explorer-pane-footer">Labels: {selectedLabels.length}</div>
                  </div>
                  <div className="explorer-pane metadata-pane">
                    <div className="explorer-pane-title">Metadata</div>
                    <div className="metadata-panel">
                      {(!selectedPath || fileKind !== 'image') && <div className="explorer-empty">Select an image to view metadata.</div>}

                      {selectedPath && fileKind === 'image' && viewMetadataStatus === 'idle' && (
                        <div className="explorer-empty">No matching TIFF metadata found for this image.</div>
                      )}

                      {selectedPath && fileKind === 'image' && viewMetadataStatus === 'loading' && (
                        <div className="explorer-empty">Loading metadata…</div>
                      )}

                      {selectedPath && fileKind === 'image' && viewMetadataStatus === 'error' && (
                        <div className="explorer-empty">{viewMetadataError || 'Failed to load metadata.'}</div>
                      )}

                      {selectedPath && fileKind === 'image' && viewMetadataStatus === 'ready' && viewMetadata && (() => {
                        const report: any = viewMetadata
                        const categories: any = report?.categories && typeof report.categories === 'object' ? report.categories : {}

                        const titleCase = (s: string) => s.replace(/_/g, ' ').replace(/\b\w/g, (c) => c.toUpperCase())
                        const activeTiffEntry = selectedPath ? findMatchingTiffForImage(selectedPath) : null
                        const tiffKey = activeTiffEntry?.path || ''
                        const selection = (tiffKey && viewMetadataSelections[tiffKey]) || {}
                        const overrides = (tiffKey && viewMetadataOverrides[tiffKey]) || {}

                        const setChecked = (id: string, checked: boolean) => {
                          if (!tiffKey) return
                          const nextForFile = { ...(selection || {}), [id]: checked }
                          setViewMetadataSelections((prev) => {
                            return { ...prev, [tiffKey]: nextForFile }
                          })

                          // If this image is already in the report draft, keep its auto metadata tables in sync.
                          if (selectedPath && fileKind === 'image') {
                            requestReportMetadataSyncForImage(selectedPath, tiffKey, nextForFile, overrides)
                          }
                        }

                        const getEffectiveValue = (id: string, fallback: any) => {
                          if (overrides && overrides[id] !== undefined) return overrides[id]
                          return fallback
                        }

                        const beginEdit = (id: string, currentValue: any) => {
                          if (!tiffKey) return
                          const existing = overrides && overrides[id] !== undefined ? overrides[id] : currentValue
                          setViewMetadataEditing({ tiffKey, id, draft: existing === null || existing === undefined ? '' : String(existing) })
                        }

                        const commitEdit = () => {
                          const e = viewMetadataEditing
                          if (!e || !e.tiffKey) return
                          const nextText = (e.draft ?? '').trim()

                          const isThermalDistance = e.id === 'measurement_params.Distance'

                          // Compute next overrides for this file so we can sync the report immediately.
                          const nextOverridesForFile: Record<string, any> = { ...(viewMetadataOverrides[e.tiffKey] || {}) }
                          if (nextText === '') {
                            delete nextOverridesForFile[e.id]
                          } else {
                            nextOverridesForFile[e.id] = nextText
                          }

                          setViewMetadataOverrides((prev) => {
                            const nextForFile = { ...(prev[e.tiffKey] || {}) }
                            if (nextText === '') {
                              delete nextForFile[e.id]
                            } else {
                              nextForFile[e.id] = nextText
                            }

                            // Report UX: if the user edits Thermal parameters -> Distance once,
                            // propagate to all report images so they don't need to repeat it.
                            if (isThermalDistance && docxDraft.length) {
                              const nextAll: Record<string, Record<string, any>> = { ...prev }

                              const apply = (tk: string) => {
                                if (!tk) return
                                const per = { ...(nextAll[tk] || {}) }
                                if (nextText === '') delete per[e.id]
                                else per[e.id] = nextText
                                nextAll[tk] = per
                              }

                              apply(e.tiffKey)
                              for (const d of docxDraft) {
                                if (d.include === false) continue
                                const entry = findMatchingTiffForImage(d.path)
                                apply(entry?.path || '')
                              }

                              return nextAll
                            }

                            return { ...prev, [e.tiffKey]: nextForFile }
                          })
                          setViewMetadataEditing(null)

                          if (isThermalDistance && docxDraft.length) {
                            // Keep the current draft tables consistent immediately.
                            if (nextText !== '') {
                              updateDocxDraftRowDescriptionForAllImagesByTableTitleAndRowName(
                                'Thermal parameters',
                                'Distance',
                                formatMetadataValue(nextText, 'distance')
                              )
                            }

                            // Sync metadata tables for all included report images.
                            for (const d of docxDraft) {
                              if (d.include === false) continue
                              const entry = findMatchingTiffForImage(d.path)
                              const tk = entry?.path || ''
                              if (!tk) continue
                              const sel = viewMetadataSelections[tk]

                              const nextOverridesForTk: Record<string, any> = { ...(viewMetadataOverrides[tk] || {}) }
                              if (nextText === '') delete nextOverridesForTk[e.id]
                              else nextOverridesForTk[e.id] = nextText

                              requestReportMetadataSyncForImage(d.path, tk, sel, nextOverridesForTk)
                            }
                          }

                          if (selectedPath && fileKind === 'image') {
                            const entry = findMatchingTiffForImage(selectedPath)
                            if (entry?.path === e.tiffKey) {
                              requestReportMetadataSyncForImage(selectedPath, e.tiffKey, selection, nextOverridesForFile)
                            }
                          }
                        }

                        const cancelEdit = () => setViewMetadataEditing(null)

                        type ValueKind = 'default' | 'temp' | 'distance' | 'percent' | 'speed' | 'irradiance'

                        const formatValue = (v: any, kind: ValueKind = 'default') => {
                          if (v === null || v === undefined || v === '') return 'N/A'

                          const num = () => {
                            if (typeof v === 'number') return Number.isFinite(v) ? v : null
                            if (typeof v === 'string') {
                              const m = v.replace(',', '.').match(/-?\d+(?:\.\d+)?/)
                              if (!m) return null
                              const n = Number(m[0])
                              return Number.isFinite(n) ? n : null
                            }
                            return null
                          }

                          if (kind === 'temp') {
                            const n = num()
                            return n === null ? 'N/A' : `${n.toFixed(2)} °C`
                          }
                          if (kind === 'distance') {
                            const n = num()
                            return n === null ? 'N/A' : `${n} m`
                          }
                          if (kind === 'percent') {
                            const n = num()
                            return n === null ? 'N/A' : `${n} %`
                          }
                          if (kind === 'speed') {
                            const n = num()
                            return n === null ? 'N/A' : `${n} m/s`
                          }
                          if (kind === 'irradiance') {
                            const n = num()
                            return n === null ? 'N/A' : `${n} W/m²`
                          }

                          if (typeof v === 'string' || typeof v === 'number' || typeof v === 'boolean') return String(v)
                          try {
                            return JSON.stringify(v)
                          } catch {
                            return String(v)
                          }
                        }

                        const row = (id: string, label: string, value: any, pinned = false, kind: ValueKind = 'default') => ({ id, label, value, pinned, kind })

                        const mapsLink = report?.maps_link || categories?.maps?.google_maps_link
                        const qrB64 = report?.qr_png_base64

                        const geo = categories?.geolocation || {}
                        const lat = geo?.latitude ?? report?.summary?.Latitude
                        const lon = geo?.longitude ?? report?.summary?.Longitude

                        const device = categories?.device || {}
                        const imageInfo = categories?.image_info || {}
                        const timestamps = categories?.timestamps || {}
                        const pixelStats = categories?.measurement_temperatures?.pixel_stats || {}
                        const params = categories?.measurement_params || {}

                        const labelTempsItems: Array<any> = Array.isArray(viewLabelTemps?.labels) ? viewLabelTemps.labels : []
                        const labelTempsByIndex = (() => {
                          const m = new Map<number, any>()
                          for (const item of labelTempsItems) {
                            const idx = Number(item?.index)
                            if (!Number.isFinite(idx)) continue
                            m.set(idx, item)
                          }
                          return m
                        })()

                        const selectedImageName = selectedPath.split('/').pop() || selectedPath

                        const derivedImageInfo: Record<string, any> = {
                          file_name: selectedImageName,
                          tiff_file: report?.file,
                          camera_model: device?.Model ?? report?.summary?.Model,
                          serial_number:
                            report?.exiftool_meta?.['ExifIFD:SerialNumber'] ??
                            report?.pillow_exif?.['ExifIFD:SerialNumber'] ??
                            device?.SerialNumber ??
                            report?.summary?.CameraSerialNumber,
                          focal_length: device?.FocalLength ?? report?.summary?.FocalLength,
                          f_number: device?.FNumber ?? report?.summary?.FNumber,
                          width: imageInfo?.ImageWidth ?? report?.summary?.ImageWidth,
                          height: imageInfo?.ImageHeight ?? report?.summary?.ImageHeight,
                          timestamp_created: timestamps?.CreateDate ?? timestamps?.DateTimeOriginal ?? report?.summary?.DateTimeOriginal ?? report?.summary?.DateTime,
                          latitude: lat,
                          longitude: lon,
                        }

                        const categoryModels: Array<{
                          key: string
                          title: string
                          open?: boolean
                          rows: Array<{ id: string; label: string; value: any; pinned: boolean; kind: ValueKind }>
                          extra?: React.ReactNode
                        }> = [
                          {
                            key: 'measurement_temperatures',
                            title: 'Temperature measurements',
                            open: true,
                            rows: [
                              row('measurement_temperatures.pixel_stats.min', 'Minimum', pixelStats?.min, true, 'temp'),
                              row('measurement_temperatures.pixel_stats.mean', 'Average', pixelStats?.mean, true, 'temp'),
                              row('measurement_temperatures.pixel_stats.max', 'Maximum', pixelStats?.max, true, 'temp'),
                              row('measurement_temperatures.pixel_stats.looks_like_celsius_guess', 'Looks like °C', pixelStats?.looks_like_celsius_guess, false),
                            ].filter((r) => r.value !== undefined),
                          },
                          {
                            key: 'fault_labels',
                            title: 'Fault labels',
                            open: true,
                            rows: selectedLabels
                              .map((label, idx) => {
                                const title = label?.source === 'manual' ? `Manual ${idx + 1}` : `Fault ${idx + 1}`
                                const f = (n: any) => (typeof n === 'number' && Number.isFinite(n) ? n.toFixed(4) : '0')
                                const conf = Number.isFinite(Number(label?.conf))
                                  ? (Number(label.conf) >= 0 && Number(label.conf) <= 1 ? ` — ${Math.round(Number(label.conf) * 100)}%` : ` — ${String(label.conf)}`)
                                  : ''
                                const shape = label?.shape ? ` — ${label.shape}` : ''
                                const temps = labelTempsByIndex.get(idx)
                                const tempSummary = temps
                                  ? ` | edge=${formatValue(temps?.outside_edge_mean, 'temp')}, avg=${formatValue(temps?.inside_mean, 'temp')}, min=${formatValue(temps?.inside_min, 'temp')}, max=${formatValue(temps?.inside_max, 'temp')}`
                                  : ''
                                const summary = `x=${f(label?.x)}, y=${f(label?.y)}, w=${f(label?.w)}, h=${f(label?.h)}${conf}${shape}${tempSummary}`
                                return row(`fault_labels.${idx}.summary`, title, summary, true)
                              })
                              .filter((r) => r && r.value !== undefined),
                          },
                          {
                            key: 'label_temperatures',
                            title: 'Label temperatures',
                            open: true,
                            rows: labelTempsItems
                              .map((item) => {
                                const idx = Number(item?.index)
                                if (!Number.isFinite(idx)) return []

                                const label = selectedLabels[idx]
                                const labelTitle = label?.source === 'manual' ? `Manual ${idx + 1}` : `Fault ${idx + 1}`
                                return [
                                  row(`label_temperatures.${idx}.outside_edge_mean`, `${labelTitle} — Outside edge avg`, item?.outside_edge_mean, true, 'temp'),
                                  row(`label_temperatures.${idx}.inside_mean`, `${labelTitle} — Inside avg`, item?.inside_mean, true, 'temp'),
                                  row(`label_temperatures.${idx}.inside_min`, `${labelTitle} — Inside min`, item?.inside_min, true, 'temp'),
                                  row(`label_temperatures.${idx}.inside_max`, `${labelTitle} — Inside max`, item?.inside_max, true, 'temp'),
                                ]
                              })
                              .flat()
                              .filter((r) => r && r.value !== undefined),
                            extra:
                              selectedLabels.length > 0 && viewLabelTempsStatus === 'loading'
                                ? (
                                    <div style={{ fontSize: 12, color: '#6b7a90' }}>Computing label temperatures…</div>
                                  )
                                : selectedLabels.length > 0 && viewLabelTempsStatus === 'error'
                                  ? (
                                      <div style={{ fontSize: 12, color: '#c23' }}>{viewLabelTempsError || 'Label temperatures failed.'}</div>
                                    )
                                  : undefined,
                          },
                          {
                            key: 'measurement_params',
                            title: 'Thermal parameters',
                            open: true,
                            rows: [
                              row('measurement_params.Distance', 'Distance', params?.Distance, true, 'distance'),
                              row('measurement_params.RelativeHumidity', 'Relative humidity', params?.RelativeHumidity, true, 'percent'),
                              row('measurement_params.Emissivity', 'Emissivity', params?.Emissivity, true),
                              row('measurement_params.AmbientTemperature', 'Ambient temperature', params?.AmbientTemperature, true, 'temp'),
                              row('measurement_params.WindSpeed', 'Wind speed', params?.WindSpeed, true, 'speed'),
                              row('measurement_params.Irradiance', 'Irradiance', params?.Irradiance, true, 'irradiance'),
                            ].filter((r) => r.value !== undefined),
                          },
                          {
                            key: 'image',
                            title: 'Image info',
                            open: true,
                            rows: [
                              row('image.file_name', 'File name', derivedImageInfo.file_name, true),
                              row('image.tiff_file', 'TIFF file', derivedImageInfo.tiff_file, true),
                              row('image.camera_model', 'Camera model', derivedImageInfo.camera_model, true),
                              row('image.serial_number', 'Serial number', derivedImageInfo.serial_number, true),
                              row('image.focal_length', 'Focal length', derivedImageInfo.focal_length, true),
                              row('image.f_number', 'F-number', derivedImageInfo.f_number, true),
                              row('image.width', 'Width', derivedImageInfo.width, true),
                              row('image.height', 'Height', derivedImageInfo.height, true),
                              row('image.timestamp_created', 'Timestamp created', derivedImageInfo.timestamp_created, true),
                              row('image.latitude', 'Latitude', derivedImageInfo.latitude, true),
                              row('image.longitude', 'Longitude', derivedImageInfo.longitude, true),
                            ].filter((r) => r.value !== undefined),
                          },
                          {
                            key: 'geolocation',
                            title: 'Geolocation',
                            open: true,
                            rows: [
                              row('geolocation.latitude', 'Latitude', lat, true),
                              row('geolocation.longitude', 'Longitude', lon, true),
                              row('geolocation.qr_code', 'QR code', typeof qrB64 === 'string' && qrB64.length > 0 ? 'Available' : null, true),
                            ].filter((r) => r.value !== undefined),
                            extra: (
                              <>
                                {mapsLink && (
                                  <a className="metadata-link" href={String(mapsLink)} target="_blank" rel="noreferrer">
                                    Open in Google Maps
                                  </a>
                                )}
                                {typeof qrB64 === 'string' && qrB64.length > 0 && (
                                  <img className="metadata-qr" src={`data:image/png;base64,${qrB64}`} alt="Maps QR" />
                                )}
                              </>
                            ),
                          },
                          {
                            key: 'flight',
                            title: 'Flight & gimbal',
                            rows: Object.entries((categories?.flight && typeof categories.flight === 'object') ? categories.flight : {}).map(([k, v]) =>
                              row(`flight.${k}`, titleCase(k), v, false)
                            ),
                          },
                          {
                            key: 'pixel_data',
                            title: 'Pixel data',
                            rows: Object.entries((categories?.pixel_data && typeof categories.pixel_data === 'object') ? categories.pixel_data : {}).map(([k, v]) =>
                              row(`pixel_data.${k}`, titleCase(k), v, false)
                            ),
                          },
                          {
                            key: 'device',
                            title: 'Device',
                            rows: Object.entries((categories?.device && typeof categories.device === 'object') ? categories.device : {}).map(([k, v]) =>
                              row(`device.${k}`, titleCase(k), v, false)
                            ),
                          },
                          {
                            key: 'timestamps',
                            title: 'Timestamps',
                            rows: Object.entries((categories?.timestamps && typeof categories.timestamps === 'object') ? categories.timestamps : {}).map(([k, v]) =>
                              row(`timestamps.${k}`, titleCase(k), v, false)
                            ),
                          },
                          {
                            key: 'tiff_info',
                            title: 'TIFF info',
                            rows: Object.entries((categories?.image_info && typeof categories.image_info === 'object') ? categories.image_info : {}).map(([k, v]) =>
                              row(`image_info.${k}`, titleCase(k), v, false)
                            ),
                          },
                        ].filter((c) => c.rows.length > 0 || c.extra)

                        const isChecked = (id: string) => {
                          if (selection[id] !== undefined) return Boolean(selection[id])
                          // Default to unchecked unless the default initializer already ran.
                          return false
                        }

                        const sortedRows = (rows: Array<{ id: string; label: string; value: any; pinned: boolean; kind: ValueKind }>) => {
                          const pinned = rows.filter((r) => r.pinned)
                          const rest = rows.filter((r) => !r.pinned)
                          return [...pinned, ...rest]
                        }

                        return (
                          <>
                            {categoryModels.map((cat) => (
                              <details key={cat.key} open={cat.open} className="metadata-block">
                                <summary className="metadata-summary">{cat.title}</summary>
                                <div className="metadata-block-body">
                                  <div className="metadata-rows">
                                    {sortedRows(cat.rows).map((r) => (
                                      <div className="metadata-row" key={r.id}>
                                        <label className="metadata-check">
                                          <input
                                            type="checkbox"
                                            checked={isChecked(r.id)}
                                            onChange={(e) => setChecked(r.id, e.target.checked)}
                                          />
                                        </label>
                                        <div className={`metadata-key ${r.pinned ? 'pinned' : ''}`}>{r.label}</div>
                                        {viewMetadataEditing && viewMetadataEditing.tiffKey === tiffKey && viewMetadataEditing.id === r.id ? (
                                          <input
                                            className="metadata-edit-input"
                                            value={viewMetadataEditing.draft}
                                            autoFocus
                                            onChange={(e) => setViewMetadataEditing((prev) => (prev ? { ...prev, draft: e.target.value } : prev))}
                                            onBlur={() => commitEdit()}
                                            onKeyDown={(e) => {
                                              if (e.key === 'Enter') commitEdit()
                                              if (e.key === 'Escape') cancelEdit()
                                            }}
                                          />
                                        ) : (
                                          <div className="metadata-value" title={formatValue(getEffectiveValue(r.id, r.value), r.kind)}>
                                            <button
                                              type="button"
                                              className="metadata-value-edit"
                                              title="Click to edit"
                                              onClick={() => beginEdit(r.id, r.value)}
                                              disabled={!tiffKey}
                                            >
                                              {formatValue(getEffectiveValue(r.id, r.value), r.kind)}
                                            </button>
                                          </div>
                                        )}
                                      </div>
                                    ))}
                                  </div>
                                  {cat.extra && <div className="metadata-maps">{cat.extra}</div>}
                                </div>
                              </details>
                            ))}
                          </>
                        )
                      })()}
                    </div>
                  </div>
                  <div className="explorer-pane">
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
                            const ok = window.confirm('Delete the selected label?')
                            if (!ok) return
                            const next = selectedLabels.filter((_, idx) => idx !== selectedLabelIndex)
                            setSelectedLabelIndex(null)
                            setSelectedLabels(next)
                            pushLabelHistory(next)
                            if (!(viewVariant === 'rgb' && selectedViewPath && selectedPath && selectedViewPath !== selectedPath)) {
                              persistLabels(next)
                            }
                            if (viewVariant === 'rgb') setRgbLabelsExportDirty(true)
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
                        {viewVariant === 'rgb' && activeAction === 'view' && selectedPath && fileKind === 'image' && (
                          <button
                            type="button"
                            className="tool-button"
                            onClick={() => {
                              void saveRgbLabelsExport()
                            }}
                            disabled={
                              rgbLabelsExportSaving ||
                              rgbLabelsExportStatus === 'checking' ||
                              !rgbLabelsExportDirty
                            }
                            title={
                              rgbLabelsExportStatus === 'checking'
                                ? 'Checking workspace/rgb_labels…'
                                : !rgbLabelsExportDirty
                                  ? 'Edit/move/resize RGB labels to enable saving'
                                  : rgbLabelsExportFileExists
                                    ? rgbLabelsExportFileEmpty
                                      ? 'Save RGB labels to fill the empty workspace/rgb_labels file'
                                      : 'Save (overwrite) RGB labels in workspace/rgb_labels'
                                    : 'Save RGB labels to workspace/rgb_labels'
                            }
                          >
                            Save
                          </button>
                        )}
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

      <footer className="app-footer">© 2026 TOPOSOL. Software by Gavriil Ilikidis.</footer>

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
