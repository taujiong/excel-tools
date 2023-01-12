import { loadExcelData } from './file.js'

interface TeacherInfo {
  name: string
  old: {
    xu: number
    total: number
  }
  new: {
    xu: number
    total: number
  }
}

const springMetas = loadExcelData('./assets/春2-0110.xlsx')
const autumnMetas = loadExcelData('./assets/秋2-0110.xlsx')
const winterMetas = loadExcelData('./assets/寒2-0110.xlsx')

const teacherMap = new Map<string, TeacherInfo>()

for (const { id: key, teacher: value } of winterMetas) {
  const teacher = getTeacher(value)
  if (isOld(key)) {
    teacher.old.total++
    if (isXu(key)) teacher.old.xu++
  } else {
    teacher.new.total++
    if (isXu(key)) teacher.new.xu++
  }
}

for (const [name, teacher] of teacherMap.entries()) {
  console.log(
    `${name}\t\t${teacher.old.xu}\t${teacher.old.total}\t${teacher.new.xu}\t${teacher.new.total}\t`
  )
}

function isOld(id: string) {
  return autumnMetas.some((meta) => meta.id === id)
}

function isXu(id: string) {
  return springMetas.some((meta) => meta.id === id)
}

function getTeacher(name: string) {
  let teacher = teacherMap.get(name)
  if (!teacher) {
    teacher = {
      name,
      old: {
        xu: 0,
        total: 0,
      },
      new: {
        xu: 0,
        total: 0,
      },
    }
    teacherMap.set(name, teacher)
  }

  return teacher
}
