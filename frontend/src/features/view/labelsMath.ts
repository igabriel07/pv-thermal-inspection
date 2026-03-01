export type LabelLike = { x: number; y: number; w: number; h: number }

export const hitTestLabel = (labels: LabelLike[], point: { x: number; y: number }) => {
  for (let i = labels.length - 1; i >= 0; i -= 1) {
    const label = labels[i]
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

export const resizeLabel = (
  origin: { x: number; y: number; w: number; h: number },
  handle: 'nw' | 'ne' | 'sw' | 'se',
  point: { x: number; y: number }
) => {
  const clamp = (value: number, min: number, max: number) => Math.min(Math.max(value, min), max)
  const minSize = 0.01

  let left = origin.x - origin.w / 2
  let right = origin.x + origin.w / 2
  let top = origin.y - origin.h / 2
  let bottom = origin.y + origin.h / 2

  left = clamp(left, 0, 1)
  right = clamp(right, 0, 1)
  top = clamp(top, 0, 1)
  bottom = clamp(bottom, 0, 1)

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
