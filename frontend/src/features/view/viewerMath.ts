export const clampPanToBounds = (args: {
  zoom: number
  viewportW: number
  viewportH: number
  baseW: number
  baseH: number
  baseZoom: number
  viewScale: number
  pan: { x: number; y: number }
}) => {
  const { zoom, viewportW, viewportH, baseW, baseH, baseZoom, viewScale, pan } = args

  if (zoom <= 1) return { x: 0, y: 0 }
  if (!viewportW || !viewportH || !baseW || !baseH) return pan

  const scale = zoom * baseZoom * viewScale
  const scaledW = baseW * scale
  const scaledH = baseH * scale

  const maxX = Math.max(0, (scaledW - viewportW) / 2)
  const maxY = Math.max(0, (scaledH - viewportH) / 2)

  return {
    x: Math.min(maxX, Math.max(-maxX, pan.x)),
    y: Math.min(maxY, Math.max(-maxY, pan.y)),
  }
}

export const calcViewerSize = (args: {
  naturalW: number
  naturalH: number
  hostW: number
  hostH: number
  maxW: number
  maxH: number
}) => {
  const { naturalW, naturalH, hostW, hostH, maxW, maxH } = args

  if (!naturalW || !naturalH) return null

  const availableW = Math.min(maxW, hostW || maxW)
  const availableH = Math.min(maxH, hostH || maxH)
  const scale = Math.min(availableW / naturalW, availableH / naturalH)

  return {
    width: Math.round(naturalW * scale),
    height: Math.round(naturalH * scale),
  }
}
