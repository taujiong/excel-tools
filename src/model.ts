export const columnNameMap = {
  uid: '学员UID',
  course: '学科',
  teacher: '教师',
} as const

export type Meta = Record<keyof typeof columnNameMap, string>
