// View feature constants (kept separate from App.tsx for maintainability).

// Multiplier applied on top of the user-controlled zoom so the default view
// renders larger without changing the zoom UI value.
export const BASE_ZOOM = 1.44

// When viewing an RGB image, render it slightly smaller for better framing.
export const RGB_VIEW_IMAGE_SCALE = 0.98

// Locked viewer frame (used for pointer mapping & RGB overlays).
export const VIEWER_MAX_W = 1536
export const VIEWER_MAX_H = 1229

// Bucket sizes for pitch/altitude-conditioned RGB alignment.
export const RGB_ALIGN_PITCH_BUCKET_DEG = 5
export const RGB_ALIGN_ALT_BUCKET_M = 5

// Default RGB yaw offset calibration.
export const DEFAULT_RGB_YAW_OFFSET_PX = { x: -15, y: 5 }

// If the flight yaw is within this many degrees of an anchor, use the anchor offset exactly.
export const RGB_YAW_SNAP_DEG = 10

export const RGB_YAW_ANCHORS = [
  { angleDeg: 0, offsetPx: { ...DEFAULT_RGB_YAW_OFFSET_PX } },
  // +90° bucket: move labels 12px up and 10px left.
  { angleDeg: 90, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 10, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 12 } },
  // +90.80° bucket: move labels 5px further up (relative to the +90° bucket).
  { angleDeg: 90.8, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 10, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 20 } },
  // +93.70° bucket: move labels 2px down (relative to the +90° bucket).
  { angleDeg: 93.7, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 10, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 13 } },
  { angleDeg: 133.4, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x + 5, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 22 } },
  { angleDeg: 167.2, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 1, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 10 } },
  { angleDeg: 174.4, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x + 9, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 25 } },
  { angleDeg: 180, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 35 } },
  { angleDeg: 269.9, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x + 2, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 3 } },
  { angleDeg: 270, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 12, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 7 } },
  { angleDeg: 270.1, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 7, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 5 } },
  { angleDeg: 261.5, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x - 5, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 20 } },
  { angleDeg: 266.1, offsetPx: { x: DEFAULT_RGB_YAW_OFFSET_PX.x, y: DEFAULT_RGB_YAW_OFFSET_PX.y - 35 } },
  { angleDeg: 267.7, offsetPx: { ...DEFAULT_RGB_YAW_OFFSET_PX } },
] as const
