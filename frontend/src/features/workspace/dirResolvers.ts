import type { FileSystemDirectoryHandle } from './fileSystemAccess'
import {
  WORKSPACE_DIR,
  WORKSPACE_RGB_LABELS_DIR,
  WORKSPACE_THERMAL_FAULTS_DIR,
  WORKSPACE_THERMAL_LABELS_CENTER_DIR,
  WORKSPACE_THERMAL_LABELS_DIR,
  WORKSPACE_THERMAL_METADATA_DIR,
  WORKSPACE_THERMAL_TEMPS_CSV_DIR,
} from './constants'
import { getParentDirPath, normalizePath } from './pathUtils'

export const getFaultsDir = async (handle: FileSystemDirectoryHandle, create: boolean) => {
  if (!handle.getDirectoryHandle) return null
  return handle.getDirectoryHandle('faults', { create })
}

export const getWorkspaceProjectRootDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) => {
  const { folderHandle, openedThermalFolderDirectly, openedRgbFolderDirectly, effectiveThermalFolderPath, create } = args

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

export const getWorkspaceRootDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) => {
  const projectRoot = await getWorkspaceProjectRootDir(args)
  if (!projectRoot) return null
  if (!projectRoot.getDirectoryHandle) return null
  return projectRoot.getDirectoryHandle(WORKSPACE_DIR, { create: args.create })
}

export const getWorkspaceSubdir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  name: string
  create: boolean
}) => {
  const workspaceRoot = await getWorkspaceRootDir(args)
  if (!workspaceRoot) return null
  if (!workspaceRoot.getDirectoryHandle) return null
  return workspaceRoot.getDirectoryHandle(args.name, { create: args.create })
}

// Resolve `workspace/<subdir>` for a specific file path.
// This makes reads/writes stable even when the opened folder contains multiple sessions
// (e.g. `data/<uuid>/thermal/...`), by anchoring `workspace/` next to that session's `thermal/`.
export const getWorkspaceSubdirForPath = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  effectiveThermalFolderPath: string | null
  targetPath: string
  name: string
  create: boolean
}) => {
  const { folderHandle, effectiveThermalFolderPath, targetPath, name, create } = args

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

export const getWorkspaceThermalLabelsDirForPath = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  effectiveThermalFolderPath: string | null
  targetPath: string
  create: boolean
}) =>
  getWorkspaceSubdirForPath({
    folderHandle: args.folderHandle,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    targetPath: args.targetPath,
    name: WORKSPACE_THERMAL_LABELS_DIR,
    create: args.create,
  })

export const getWorkspaceThermalLabelsCenterDirForPath = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  effectiveThermalFolderPath: string | null
  targetPath: string
  create: boolean
}) =>
  getWorkspaceSubdirForPath({
    folderHandle: args.folderHandle,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    targetPath: args.targetPath,
    name: WORKSPACE_THERMAL_LABELS_CENTER_DIR,
    create: args.create,
  })

export const getWorkspaceThermalFaultsDirForPath = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  effectiveThermalFolderPath: string | null
  targetPath: string
  create: boolean
}) =>
  getWorkspaceSubdirForPath({
    folderHandle: args.folderHandle,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    targetPath: args.targetPath,
    name: WORKSPACE_THERMAL_FAULTS_DIR,
    create: args.create,
  })

export const getWorkspaceThermalFaultsDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) =>
  getWorkspaceSubdir({
    folderHandle: args.folderHandle,
    openedThermalFolderDirectly: args.openedThermalFolderDirectly,
    openedRgbFolderDirectly: args.openedRgbFolderDirectly,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    name: WORKSPACE_THERMAL_FAULTS_DIR,
    create: args.create,
  })

export const getWorkspaceThermalLabelsDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) =>
  getWorkspaceSubdir({
    folderHandle: args.folderHandle,
    openedThermalFolderDirectly: args.openedThermalFolderDirectly,
    openedRgbFolderDirectly: args.openedRgbFolderDirectly,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    name: WORKSPACE_THERMAL_LABELS_DIR,
    create: args.create,
  })

export const getWorkspaceThermalLabelsCenterDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) =>
  getWorkspaceSubdir({
    folderHandle: args.folderHandle,
    openedThermalFolderDirectly: args.openedThermalFolderDirectly,
    openedRgbFolderDirectly: args.openedRgbFolderDirectly,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    name: WORKSPACE_THERMAL_LABELS_CENTER_DIR,
    create: args.create,
  })

export const getWorkspaceThermalMetadataDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) =>
  getWorkspaceSubdir({
    folderHandle: args.folderHandle,
    openedThermalFolderDirectly: args.openedThermalFolderDirectly,
    openedRgbFolderDirectly: args.openedRgbFolderDirectly,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    name: WORKSPACE_THERMAL_METADATA_DIR,
    create: args.create,
  })

export const getWorkspaceThermalTemperaturesCsvsDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) =>
  getWorkspaceSubdir({
    folderHandle: args.folderHandle,
    openedThermalFolderDirectly: args.openedThermalFolderDirectly,
    openedRgbFolderDirectly: args.openedRgbFolderDirectly,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    name: WORKSPACE_THERMAL_TEMPS_CSV_DIR,
    create: args.create,
  })

export const getWorkspaceRgbLabelsDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  openedThermalFolderDirectly: boolean
  openedRgbFolderDirectly: boolean
  effectiveThermalFolderPath: string | null
  create: boolean
}) =>
  getWorkspaceSubdir({
    folderHandle: args.folderHandle,
    openedThermalFolderDirectly: args.openedThermalFolderDirectly,
    openedRgbFolderDirectly: args.openedRgbFolderDirectly,
    effectiveThermalFolderPath: args.effectiveThermalFolderPath,
    name: WORKSPACE_RGB_LABELS_DIR,
    create: args.create,
  })

export const getWorkflowRootDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  effectiveThermalFolderPath: string | null
  create: boolean
}) => {
  const { folderHandle, effectiveThermalFolderPath, create } = args
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

export const getRgbRootDir = async (args: {
  folderHandle: FileSystemDirectoryHandle | null
  effectiveRgbFolderPath: string | null
  viewVariant: 'thermal' | 'rgb'
  selectedViewPath: string
  create: boolean
}) => {
  const { folderHandle, effectiveRgbFolderPath, viewVariant, selectedViewPath, create } = args

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
