import { DEFAULT_RGB_YAW_OFFSET_PX, RGB_YAW_ANCHORS, RGB_YAW_SNAP_DEG } from './constants'

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

export const extractFlightYawDegree = (report: any | null): number | null => {
  if (!report) return null

  // Preferred structured location from backend/check_for_metadata_tiff.py
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

export const extractGimbalYawDegree = (report: any | null): number | null => {
  if (!report) return null

  // Preferred structured location from backend/check_for_metadata_tiff.py
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

export const extractGimbalPitchDegree = (report: any | null): number | null => {
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

export const extractRelativeAltitude = (report: any | null): number | null => {
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
// flight yaw and gimbal yaw.
export const extractYawForAlignment = (report: any | null): number | null => {
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
const RGB_FLIGHT_GIMBAL_YAW_ADJUSTMENTS: Array<{
  flightYawDeg: number
  gimbalYawDeg: number
  tolDeg: number
  deltaPx: { x: number; y: number }
}> = [
  { flightYawDeg: 270, gimbalYawDeg: 268.7, tolDeg: 0.6, deltaPx: { x: -5, y: -5 } },
  { flightYawDeg: 183, gimbalYawDeg: 9.9, tolDeg: 0.6, deltaPx: { x: 0, y: -40 } },
  { flightYawDeg: 90.8, gimbalYawDeg: 269.2, tolDeg: 0.6, deltaPx: { x: -20, y: -30 } },
  { flightYawDeg: 261.5, gimbalYawDeg: 85.2, tolDeg: 0.6, deltaPx: { x: 0, y: -20 } },
  { flightYawDeg: 269.9, gimbalYawDeg: 270.2, tolDeg: 0.6, deltaPx: { x: 20, y: 3 } },
  { flightYawDeg: 270, gimbalYawDeg: 270.6, tolDeg: 0.6, deltaPx: { x: -20, y: 0 } },
]

export const getFlightGimbalYawAdjustmentPx = (flightYaw: number | null, gimbalYaw: number | null) => {
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

export const getYawInterpolatedOffsetPx = (flightYawDegree: number | null) => {
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

  return { ...DEFAULT_RGB_YAW_OFFSET_PX }
}
