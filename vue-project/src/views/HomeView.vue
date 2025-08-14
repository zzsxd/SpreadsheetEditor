<template>
  <div class="home">
    <SheetTabs
      :sheets="sheets"
      :active-sheet="activeSheet"
      @select-sheet="selectSheet"
      @add-sheet="addSheet"
      @rename-sheet="renameSheet"
      @delete-sheet="deleteSheet"
    />

    <Toolbar
      :current-styles="getCurrentCellStyles()"
      @apply-styles="applyStylesToSelectedCells"
      @import-data="handleImport"
      @export-excel="exportExcel"
    />

    <SearchPanel
      :columns="columns"
      @search="handleSearch"
    />

    <ExcelGrid
      :data="currentData"
      :columns="columns"
      :selected-cells="selectedCells"
      :search-results="searchResults"
      :row-height="ROW_HEIGHT"
      @update:selected-cells="selectedCells = $event"
      @update-cell="updateCell"
      @resize-column="resizeColumn"
    />
  </div>
</template>

<script setup>
import { ref, computed } from 'vue'
import ExcelJS from 'exceljs'
import SheetTabs from '@/components/SheetTabs.vue'
import Toolbar from '@/components/Toolbar.vue'
import ExcelGrid from '@/components/ExcelGrid.vue'
import SearchPanel from '@/components/SearchPanel.vue'

/** ---- Константы ---- */
const ROWS_COUNT = 100
const COLS_COUNT = 26
const DEFAULT_COL_WIDTH = 120
const ROW_HEIGHT = 30 // для виртуализации
defineExpose({}) // (безопасная заглушка, если нужен доступ снаружи)

/** ---- Состояние ---- */
const nextSheetId = ref(2)
const sheets = ref([])
const activeSheet = ref(null)
const columns = ref([])
const selectedCells = ref([])
const searchResults = ref([])

columns.value = Array.from({ length: COLS_COUNT }, (_, i) => ({
  id: (i + 1).toString(),
  name: String.fromCharCode(65 + i),
  width: DEFAULT_COL_WIDTH
}))

function createNewSheet(name) {
  const sheetId = nextSheetId.value++
  return {
    id: sheetId,
    name: name || `Лист${sheetId}`,
    columns: [...columns.value],
    data: Array(ROWS_COUNT).fill().map(() => {
      const row = {}
      columns.value.forEach(col => {
        row[col.id] = { value: '', style: {}, meta: {} }
      })
      return row
    })
  }
}

sheets.value = [createNewSheet('Лист1')]
activeSheet.value = sheets.value[0].id

const currentSheet = computed(() => sheets.value.find(s => s.id === activeSheet.value) || null)

const currentData = computed(() => currentSheet.value ? currentSheet.value.data : [])
const selectSheet = (id) => { activeSheet.value = id; selectedCells.value = []; searchResults.value = []; columns.value = [...(currentSheet.value?.columns || [])] }

const addSheet = () => {
  const newSheet = createNewSheet()
  sheets.value.push(newSheet)
  selectSheet(newSheet.id)
}

const renameSheet = ({ id, newName }) => {
  const sheet = sheets.value.find(s => s.id === id)
  if (sheet && newName.trim()) sheet.name = newName.trim()
}

const deleteSheet = (sheetId) => {
  if (sheets.value.length <= 1) return
  if (confirm('Вы уверены, что хотите удалить этот лист?')) {
    const idx = sheets.value.findIndex(s => s.id === sheetId)
    if (idx !== -1) {
      sheets.value.splice(idx, 1)
      if (activeSheet.value === sheetId) selectSheet(sheets.value[0].id)
    }
  }
}

/** ---- Текущие стили выделенной ячейки ---- */
const getCurrentCellStyles = () => {
  if (selectedCells.value.length === 0) return {
    fontWeight: 'normal', textAlign: 'left', color: '#000000',
    backgroundColor: '#ffffff', fontSize: '14px'
  }
  const first = selectedCells.value[0]
  const sheet = currentSheet.value
  return sheet?.data[first.rowIndex]?.[first.colId]?.style || {
    fontWeight: 'normal', textAlign: 'left', color: '#000000',
    backgroundColor: '#ffffff', fontSize: '14px'
  }
}

const updateCell = ({ rowIndex, colId, value }) => {
  const sheet = currentSheet.value
  if (sheet && sheet.data[rowIndex]) sheet.data[rowIndex][colId].value = value
}

const applyStylesToSelectedCells = (styles) => {
  if (!selectedCells.value.length) return
  const sheet = currentSheet.value
  if (!sheet) return
  selectedCells.value.forEach(({ rowIndex, colId }) => {
    const cell = sheet.data[rowIndex]?.[colId]
    if (cell) cell.style = { ...cell.style, ...styles }
  })
}

const resizeColumn = ({ colId, width }) => {
  const column = columns.value.find(c => c.id === colId)
  if (column) {
    column.width = Math.max(50, width)
    // синхронизируем с листом
    const sheet = currentSheet.value
    if (sheet) {
      const idx = sheet.columns.findIndex(c => c.id === colId)
      if (idx !== -1) sheet.columns[idx].width = column.width
    }
  }
}

/** ---- Поиск ---- */
const handleSearch = ({ queries, matchCase, columns: searchColumns }) => {
  if (!queries || queries.length === 0) { searchResults.value = []; return }
  const sheet = currentSheet.value
  if (!sheet) return
  const results = []
  sheet.data.forEach((row, rowIndex) => {
    for (const [colId, cell] of Object.entries(row)) {
      if (searchColumns !== 'all' && !searchColumns.includes(colId)) continue
      const text = matchCase ? (cell.value ?? '') : (cell.value ?? '').toLowerCase()
      const ok = queries.some(q => (matchCase ? q : q.toLowerCase()))
        && queries.some(q => text.includes(matchCase ? q : q.toLowerCase()))
      if (ok) results.push({ rowIndex, colId })
    }
  })
  searchResults.value = results
  if (results.length) selectedCells.value = results
}

/** ---------------------------------------------------
 *     Импорт Excel — гиперссылки + цвета + скорость
 *  --------------------------------------------------*/

/** Базовая Office тема (приближенно), для конвертации theme->rgb */
const BASE_THEME_RGB = [
  '#000000', '#FFFFFF', '#EEECE1', '#1F497D', '#4F81BD',
  '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646'
]

function applyTint(hex, tint = 0) {
  // Excel tint: >0 lighten, <0 darken
  const clamp = v => Math.max(0, Math.min(255, v))
  const to = tint > 0 ? 255 : 0
  const amt = Math.abs(tint)
  const r = parseInt(hex.slice(1, 3), 16)
  const g = parseInt(hex.slice(3, 5), 16)
  const b = parseInt(hex.slice(5, 7), 16)
  const nr = clamp(Math.round(r + (to - r) * amt))
  const ng = clamp(Math.round(g + (to - g) * amt))
  const nb = clamp(Math.round(b + (to - b) * amt))
  return `#${nr.toString(16).padStart(2,'0')}${ng.toString(16).padStart(2,'0')}${nb.toString(16).padStart(2,'0')}`
}

function colorToHex(color) {
  if (!color) return null
  if (color.argb) {
    // ABGR — берём RGB (последние 6)
    return `#${color.argb.slice(2)}`
  }
  if (typeof color.theme === 'number') {
    const base = BASE_THEME_RGB[color.theme % BASE_THEME_RGB.length] || '#000000'
    return applyTint(base, color.tint || 0)
  }
  return null
}

function extractFillColor(fill) {
  if (!fill) return null
  if (fill.type === 'pattern') {
    const fg = colorToHex(fill.fgColor)
    const bg = colorToHex(fill.bgColor)
    return fg || bg || null
  }
  // градиенты отображаем как первый цвет
  if (fill.type === 'gradient') {
    const stop = fill.stops?.[0]
    return colorToHex(stop?.color) || null
  }
  return null
}

function richTextToPlain(rich) {
  try {
    return (rich?.richText || [])
      .map(r => r.text ?? '')
      .join('')
  } catch { return '' }
}

function hyperlinkToMeta(value) {
  // value: { text, hyperlink, tooltip? }
  return {
    text: value?.text ?? value?.hyperlink ?? '',
    hyperlink: value?.hyperlink ?? null,
  }
}

function getCellDisplayAndMeta(cell) {
  // Возвращаем { text, meta } для значения разных типов (в т.ч. формулы/гиперссылки/riched)
  let text = ''
  let meta = {}

  const v = cell.value

  // Формула
  if (v && typeof v === 'object' && 'formula' in v) {
    text = (v.result != null) ? String(v.result) : ''
  }
  // Гиперссылка
  else if (v && typeof v === 'object' && 'hyperlink' in v) {
    const h = hyperlinkToMeta(v)
    text = String(h.text || '')
    meta.hyperlink = h.hyperlink
  }
  // RichText
  else if (v && typeof v === 'object' && 'richText' in v) {
    text = richTextToPlain(v)
  }
  // Дата/число/строка/булево
  else if (v != null) {
    text = String(v)
  }

  return { text, meta }
}

const handleImport = async (event) => {
  const file = event.target.files[0]
  if (!file) return

  try {
    const buffer = await file.arrayBuffer()
    const workbook = new ExcelJS.Workbook()
    await workbook.xlsx.load(buffer)

    // очищаем листы
    sheets.value = []

    workbook.eachSheet((worksheet, idx) => {
      const sheetColumns = []
      const colCount = Math.max(1, worksheet.actualColumnCount || worksheet.columnCount || COLS_COUNT)

      for (let i = 1; i <= colCount; i++) {
        // забираем ширину колонки из xlsx (примерная конвертация в px)
        const xlsxWidth = worksheet.getColumn(i)?.width // в "символах"
        const px = xlsxWidth ? Math.round(xlsxWidth * 7) : DEFAULT_COL_WIDTH
        sheetColumns.push({
          id: i.toString(),
          name: String.fromCharCode(64 + i),
          width: Math.max(50, px)
        })
      }

      const sheetData = []
      // ВАЖНО: формируем строки один раз — без лишних .push в глубине циклов
      const totalRows = worksheet.actualRowCount || worksheet.rowCount || 0
      for (let r = 1; r <= totalRows; r++) {
        const row = worksheet.getRow(r)
        const rowObj = {}
        for (let c = 1; c <= colCount; c++) {
          const cell = row.getCell(c)
          const { text, meta } = getCellDisplayAndMeta(cell)
          const font = cell.font || {}
          const fill = extractFillColor(cell.fill)

          let colorHex = '#000000' // дефолт чёрный

          if (font.color) {
            if (font.color.argb) {
              const candidate = `#${font.color.argb.slice(2)}`.toUpperCase()
              if (candidate !== '#FFFFFF') { // игнорируем белый
                colorHex = candidate
              }
            } else {
              const converted = colorToHex(font.color)?.toUpperCase()
              if (converted && converted !== '#FFFFFF') { // игнорируем белый
                colorHex = converted
              }
            }
          }

          const bgHex = (fill || '#FFFFFF').toUpperCase()

          rowObj[c.toString()] = {
            value: text,
            meta, // meta.hyperlink если был
            style: {
              fontWeight: font.bold ? 'bold' : 'normal',
              textAlign: cell.alignment?.horizontal || 'left',
              color: colorHex,
              backgroundColor: bgHex,
              fontSize: font.size ? `${font.size}px` : '14px'
            }
          }
        }
        sheetData.push(rowObj)
      }

      const newSheet = {
        id: idx, // ExcelJS нумерует с 1
        name: worksheet.name || `Лист${idx}`,
        columns: sheetColumns,
        data: sheetData
      }
      sheets.value.push(newSheet)
    })

    if (sheets.value.length) {
      activeSheet.value = sheets.value[0].id
      columns.value = [...sheets.value[0].columns]
    }
  } catch (e) {
    console.error('Ошибка импорта:', e)
    alert('Не удалось загрузить файл')
  } finally {
    event.target.value = ''
  }
}

/** ---- Экспорт ---- */
const exportExcel = async () => {
  const workbook = new ExcelJS.Workbook()
  sheets.value.forEach(sheet => {
    const ws = workbook.addWorksheet(sheet.name)

    // ширины колонок (примерный перевод px -> "символы")
    ws.columns = sheet.columns.map(c => ({ width: Math.min(50, Math.max(5, Math.round(c.width / 7))) }))

    sheet.data.forEach(row => {
      const values = sheet.columns.map(c => row[c.id]?.value || '')
      const excelRow = ws.addRow(values)

      sheet.columns.forEach((c, idx) => {
        const cellData = row[c.id]
        if (!cellData) return
        const cell = excelRow.getCell(idx + 1)

        // гиперссылка обратно
        if (cellData.meta?.hyperlink) {
          cell.value = { text: String(cellData.value || ''), hyperlink: String(cellData.meta.hyperlink) }
        }

        const st = cellData.style || {}
        cell.font = {
          bold: st.fontWeight === 'bold',
          color: st.color ? { argb: st.color.replace('#', '').padStart(6, '0') } : undefined,
          size: st.fontSize ? parseInt(st.fontSize) : undefined
        }
        if (st.textAlign) cell.alignment = { horizontal: st.textAlign, vertical: 'middle' }
        if (st.backgroundColor) {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: st.backgroundColor.replace('#', '') } }
        }
        cell.border = {
          top: { style: 'thin', color: { argb: '000000' } },
          bottom: { style: 'thin', color: { argb: '000000' } },
          left: { style: 'thin', color: { argb: '000000' } },
          right: { style: 'thin', color: { argb: '000000' } }
        }
      })
    })
  })

  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const link = document.createElement('a')
  link.href = URL.createObjectURL(blob)
  link.download = `${sheets.value[0]?.name || 'таблица'}.xlsx`
  link.click()
}
</script>

<style scoped>
.home {
  display: flex;
  flex-direction: column;
  height: 100vh;
  padding: 0;
}
</style>
