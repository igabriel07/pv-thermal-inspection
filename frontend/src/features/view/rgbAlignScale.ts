import { RGB_ALIGN_ALT_BUCKET_M, RGB_ALIGN_PITCH_BUCKET_DEG } from './constants'

export const bucketRound = (value: number, step: number) => {
  if (!Number.isFinite(value) || !Number.isFinite(step) || step <= 0) return value
  return Math.round(value / step) * step
}

export const makeRgbAlignScaleContextKey = (gimbalPitchDeg: number | null, relativeAltM: number | null) => {
  if (gimbalPitchDeg === null || relativeAltM === null) return null
  const p = bucketRound(gimbalPitchDeg, RGB_ALIGN_PITCH_BUCKET_DEG)
  const a = bucketRound(relativeAltM, RGB_ALIGN_ALT_BUCKET_M)
  if (!Number.isFinite(p) || !Number.isFinite(a)) return null
  return `p${p.toFixed(0)}_a${a.toFixed(0)}`
}

export const getRgbAlignScaleStorageKey = (contextKey: string | null) => {
  if (!contextKey) return 'rgb-align-scale-v1'
  return `rgb-align-scale-v1:${contextKey}`
}

export const parseStoredScale = (raw: string | null) => {
  const n = raw ? Number(raw) : NaN
  return Number.isFinite(n) ? Math.min(2, Math.max(0.5, n)) : null
}
