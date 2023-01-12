import fs from 'fs'
import type { Sheet } from 'xlsx'
import { readFile, set_fs } from 'xlsx'
import { DATA_SHEET_NAME, HEAD_CROSS_ROWS } from './constant.js'
import type { Meta } from './model.js'
import { columnNameMap } from './model.js'

set_fs(fs)

export function loadExcelData(fileName: string) {
  console.log(`正在读取 excel：${fileName}`)

  const wb = readFile(fileName)
  const sheet = wb.Sheets[DATA_SHEET_NAME]

  const { startCol, startRow, endCol, endRow } = getShape(sheet)
  const targetHeadNameMap = findTargetColumnForKey(sheet, startCol, endCol)
  const metaMap = extractMeta(sheet, targetHeadNameMap, startRow, endRow)
  console.log(metaMap.length)

  return metaMap
}

interface ComputedMeta {
  id: string
  teacher: string
}

// eslint-disable-next-line max-params
function extractMeta(
  sheet: Sheet,
  metaDef: Meta,
  startRow: number,
  endRow: number
): ComputedMeta[] {
  const metas: ComputedMeta[] = []
  for (let i = startRow; i <= endRow; i++) {
    let ignore = false
    const meta: Record<string, string> = {}
    for (const [key, char] of Object.entries(metaDef)) {
      meta[key] = sheet[`${char}${i}`]?.v
      if (!meta[key]) {
        ignore = true
        continue
      }
    }

    if (ignore) continue

    const mapKey = `${meta['uid']}-${meta['course']}`
    // if (map.has(mapKey)) {
    //   console.log(`数据已存在：${i}-${mapKey}`)
    // }
    metas.push({
      id: mapKey,
      teacher: meta['teacher'],
    })
  }

  return metas
}

function getShape(sheet: Sheet) {
  const [leftTop, rightBottom] = sheet['!ref']!.split(':')
  const [startCol, startRow] = getCellPos(leftTop)
  const [endCol, endRow] = getCellPos(rightBottom)
  const colNum = endCol - startCol + 1
  const headCrossRows = sheet['!margins'] === undefined ? 1 : HEAD_CROSS_ROWS
  let rowNum = endRow - startRow - headCrossRows + 1
  console.log(`数据表共 ${colNum} 列，${rowNum} 行`)

  return {
    startRow: startRow + headCrossRows,
    startCol,
    endCol,
    endRow,
  }
}

function getCellPos(cellDef: string) {
  const col = cellDef.substring(0, 1).charCodeAt(0)
  const rowStr = cellDef.substring(1)
  const row = parseInt(rowStr, 10)

  return [col, row]
}

function findTargetColumnForKey(sheet: Sheet, startCol: number, endCol: number) {
  const map = {
    uid: '',
    course: '',
    teacher: '',
  } satisfies Meta
  const headNames = getHeadNames(sheet, startCol, endCol)

  for (const [key, targetHeadName] of Object.entries(columnNameMap)) {
    for (const headName of headNames) {
      if (targetHeadName === headName) {
        const index = headNames.indexOf(headName)
        const charCode = startCol + index
        map[key as keyof Meta] = String.fromCharCode(charCode)
      }
    }
  }

  validateSheetHead(map)

  return map as Meta
}

function getHeadNames(sheet: Sheet, startCol: number, endCol: number) {
  const headNames: string[] = []

  for (let i = startCol; i <= endCol; i++) {
    const char = String.fromCharCode(i)
    const headName = sheet[`${char}1`]?.v

    if (!headName) continue

    headNames.push(headName)
  }

  console.log(`表头数据：${headNames.join('，')}`)

  return headNames
}

function validateSheetHead(map: Record<string, string>) {
  const missedHeads = Object.keys(columnNameMap).filter((headName) => !map[headName])
  if (missedHeads.length !== 0) {
    // @ts-ignore
    const missedHeadNames = missedHeads.map((head) => columnNameMap[head]).join('和')
    throw new Error(`数据表结构不正确，缺少表头为${missedHeadNames}的列`)
  }
}
