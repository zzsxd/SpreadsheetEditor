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
      :search-results="searchResults"
      @search="handleSearch"
      @navigate="goToSearchResult"
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
import { ref, computed, onMounted, onUnmounted } from 'vue'
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

const copySelectedCells = async () => {
  if (!selectedCells.value.length) return
  const sheet = currentSheet.value
  if (!sheet) return

  // Определяем диапазон выделения
  const minRow = Math.min(...selectedCells.value.map(c => c.rowIndex))
  const maxRow = Math.max(...selectedCells.value.map(c => c.rowIndex))
  const minCol = Math.min(...selectedCells.value.map(c => parseInt(c.colId)))
  const maxCol = Math.max(...selectedCells.value.map(c => parseInt(c.colId)))

  let rows = []
  for (let r = minRow; r <= maxRow; r++) {
    let row = []
    for (let c = minCol; c <= maxCol; c++) {
      row.push(sheet.data[r][c]?.value || '')
    }
    rows.push(row.join('\t'))
  }

  const text = rows.join('\n')
  await navigator.clipboard.writeText(text)
}


const deleteSelectedCells = () => {
  if (!selectedCells.value.length) return
  const sheet = currentSheet.value
  selectedCells.value.forEach(({ rowIndex, colId }) => {
    sheet.data[rowIndex][colId].value = ''
  })
}

const applyStylesToSelectedCells = (styles) => {
  if (!selectedCells.value.length) return
  const sheet = currentSheet.value
  if (!sheet) return
  
  selectedCells.value.forEach(({ rowIndex, colId }) => {
    const cell = sheet.data[rowIndex]?.[colId]
    if (cell) {
      // Сохраняем существующие стили и применяем новые
      cell.style = { ...cell.style, ...styles }
      
      // Принудительно обновляем отображение
      if (styles.color) {
        cell.style.color = styles.color
      }
      if (styles.backgroundColor) {
        cell.style.backgroundColor = styles.backgroundColor
      }
    }
  })
  
  // Принудительное обновление реактивности
  sheet.data = [...sheet.data]
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

const goToSearchResult = (cell) => {
  selectedCells.value = [cell]
  // можно добавить прокрутку до ячейки
  const container = document.querySelector('.excel-grid')
  if (container) {
    const y = cell.rowIndex * ROW_HEIGHT
    container.scrollTop = y
  }
}


/** ---- Поиск ---- */
const handleSearch = ({ queries, columns: searchColumns }) => {
  if (!queries || queries.length === 0) { 
    searchResults.value = [] 
    return 
  }
  const sheet = currentSheet.value
  if (!sheet) return
  const results = []

  sheet.data.forEach((row, rowIndex) => {
    for (const [colId, cell] of Object.entries(row)) {
      if (searchColumns !== 'all' && !searchColumns.includes(colId)) continue
      const text = (cell.value ?? '').toString().toLowerCase()
      
      // ✅ теперь проверяем, что ВСЕ запросы присутствуют в тексте
      const ok = queries.every(q => text.includes(q.toLowerCase()))
      
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
// Обновленная базовая тема Excel с более точными цветами
// Полная палитра Excel theme colors
// Обновленная палитра цветов темы Excel
const EXCEL_THEME_COLORS = [
  '#000000', // 0: Dark 1
  '#FFFFFF', // 1: Light 1 (базовый белый)
  '#E7E6E6', // 2: Light 2 (светло-серый)
  '#44546A', // 3: Dark 2 (темно-синий)
  '#5B9BD5', // 4: Accent 1 (синий)
  '#ED7D31', // 5: Accent 2 (оранжевый)
  '#A5A5A5', // 6: Accent 3 (серый)
  '#FFC000', // 7: Accent 4 (желтый)
  '#4472C4', // 8: Accent 5 (синий)
  '#70AD47', // 9: Accent 6 (зеленый)
  '#0563C1', // 10: Hyperlink (синий)
  '#954F72', // 11: Followed Hyperlink (фиолетовый)
  '#808080', // 12: Style #1 (Gray)
  '#C0C0C0', // 13: Style #2 (Silver)
  '#FF9900', // 14: Style #3 (Orange)
  '#FFFFCC'  // 15: Style #4 (Light Yellow)
]


function colorToHexSafe(color) {
  if (!color) return null
  
  try {
    // 1. Проверяем RGB (прямое значение)
    if (color.rgb) {
      let rgb = color.rgb.replace(/^#/, '')
      if (rgb.length === 8) rgb = rgb.slice(2) // убираем альфу ARGB
      if (rgb.length === 6) return `#${rgb.toUpperCase()}`
    }
    
    // 2. Проверяем ARGB
    if (color.argb) {
      let argb = color.argb.replace(/^#/, '').toUpperCase()
      if (argb.length === 8) return `#${argb.slice(2)}` // убираем альфу
      if (argb.length === 6) return `#${argb}`
    }
    
    // 3. Проверяем THEME + tint (самая важная часть)
    if (typeof color.theme === 'number') {
      const themeIndex = color.theme
      const tint = color.tint || 0
      
      if (themeIndex >= 0 && themeIndex < EXCEL_THEME_COLORS.length) {
        const baseColor = EXCEL_THEME_COLORS[themeIndex]
        const resultColor = applyTintCorrected(baseColor, tint)
        console.log('Theme color conversion:', { themeIndex, tint, baseColor, resultColor })
        return resultColor
      }
    }
    
    return null
  } catch (e) {
    console.error("Ошибка обработки цвета:", color, e)
    return null
  }
}




// Функция для извлечения цвета шрифта
function extractFontColor(font) {
  if (!font || !font.color) {
    return '#000000' // по умолчанию всегда черный
  }
  const colorHex = colorToHexSafe(font.color)
  // Excel иногда возвращает белый (#FFFFFF) вместо "по умолчанию"
  if (!colorHex || colorHex.toUpperCase() === '#FFFFFF') {
    return '#000000'
  }
  return colorHex
}


function applyTintCorrected(hex, tint) {
  if (!hex || tint === 0 || tint === undefined) return hex
  
  // Конвертируем HEX в RGB
  let r = parseInt(hex.slice(1, 3), 16)
  let g = parseInt(hex.slice(3, 5), 16)
  let b = parseInt(hex.slice(5, 7), 16)
  
  // Правильная формула для Excel tint согласно официальной документации
  if (tint < 0) {
    // Затемнение: color = base * (1 + tint)
    // tint отрицательный, поэтому 1 + tint < 1
    r = Math.round(r * (1 + tint))
    g = Math.round(g * (1 + tint))
    b = Math.round(b * (1 + tint))
  } else {
    // Осветление: color = base * (1 - tint) + 255 * tint
    r = Math.round(r * (1 - tint) + 255 * tint)
    g = Math.round(g * (1 - tint) + 255 * tint)
    b = Math.round(b * (1 - tint) + 255 * tint)
  }
  
  // Ограничиваем значения 0-255
  r = Math.max(0, Math.min(255, r))
  g = Math.max(0, Math.min(255, g))
  b = Math.max(0, Math.min(255, b))
  
  const result = `#${r.toString(16).padStart(2, '0')}${g.toString(16).padStart(2, '0')}${b.toString(16).padStart(2, '0')}`
  return result
}

// Функция для извлечения цвета заливки
function extractFillColor(fill) {
  if (!fill) return '#FFFFFF'
  
  let colorHex = null
  let colorObj = null

  if (fill.type === 'pattern') {
    // Сначала проверяем fgColor
    if (fill.fgColor) {
      colorObj = fill.fgColor
      colorHex = colorToHexSafe(fill.fgColor)
    }
    
    // Если fgColor не найден или белый, проверяем bgColor
    if (!colorHex || colorHex === '#FFFFFF') {
      if (fill.bgColor) {
        colorObj = fill.bgColor
        colorHex = colorToHexSafe(fill.bgColor)
      }
    }
  } 
  else if (fill.type === 'gradient' && fill.stops?.length > 0) {
    colorObj = fill.stops[0].color
    colorHex = colorToHexSafe(fill.stops[0].color)
  }

  // Применяем специальную обработку для серых цветов
  if (colorHex && colorObj) {
    colorHex = handleExcelGrayColors(colorObj, colorHex)
  }

  return colorHex || '#FFFFFF'
}

function handleExcelGrayColors(colorObj, calculatedHex) {
  if (!colorObj || !calculatedHex) return calculatedHex
  
  // Определяем известные серые цвета Excel
  const excelGrayColors = {
    '#D9D9D9': '#D9D9D9', // Standard gray 1
    '#E7E6E6': '#E7E6E6', // Standard gray 2
    '#F2F2F2': '#F2F2F2', // Standard gray 3
    '#BFBFBF': '#BFBFBF', // Standard gray 4
    '#A6A6A6': '#A6A6A6', // Standard gray 5
    '#808080': '#808080', // Standard gray 6
  }
  
  // Если это известный серый цвет, возвращаем его напрямую
  if (excelGrayColors[calculatedHex.toUpperCase()]) {
    return excelGrayColors[calculatedHex.toUpperCase()]
  }
  
  // Особый случай: Light 1 (белый) с tint для создания серого
  if (colorObj.theme === 1 && colorObj.tint < 0) {
    const tint = colorObj.tint
    // Эмпирическая корректировка для популярных серых оттенков
    if (tint >= -0.15) return '#F2F2F2'
    if (tint >= -0.25) return '#E7E6E6'
    if (tint >= -0.35) return '#D9D9D9'
    if (tint >= -0.5) return '#BFBFBF'
    return '#A6A6A6'
  }
  
  return calculatedHex
}


// Функция для определения, является ли цвет серым из Excel theme
function isGrayFromExcelTheme(colorObj, hexColor) {
  // Если это theme color с индексом 1 (белый) и отрицательным tint
  if (colorObj.theme === 1 && colorObj.tint < 0) {
    return true
  }
  
  // Если это стандартный серый цвет Excel
  const grayColors = ['#808080', '#A0A0A0', '#C0C0C0', '#D3D3D3', '#D9D9D9', '#E6E6E6']
  if (grayColors.includes(hexColor.toUpperCase())) {
    return true
  }
  
  // Проверяем, является ли цвет серым по компонентам RGB
  if (hexColor.startsWith('#')) {
    const r = parseInt(hexColor.slice(1, 3), 16)
    const g = parseInt(hexColor.slice(3, 5), 16)
    const b = parseInt(hexColor.slice(5, 7), 16)
    
    // Цвет считается серым, если все компоненты близки по значению
    const threshold = 40
    return Math.abs(r - g) < threshold && 
           Math.abs(g - b) < threshold && 
           Math.abs(r - b) < threshold
  }
  
  return false
}


const tableColors = ref([]);

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
      const colCount = worksheet.columnCount || COLS_COUNT

      // Создаем колонки
      for (let i = 1; i <= colCount; i++) {
        const xlsxWidth = worksheet.getColumn(i)?.width
        const px = xlsxWidth ? Math.round(xlsxWidth * 7) : DEFAULT_COL_WIDTH
        sheetColumns.push({
          id: i.toString(),
          name: String.fromCharCode(64 + i),
          width: Math.max(50, px)
        })
      }

      const sheetData = []
      const rowCount = worksheet.rowCount || 0
      
      // Обрабатываем все строки
      for (let r = 1; r <= rowCount; r++) {
        const row = worksheet.getRow(r)
        const rowObj = {}
        
        // Обрабатываем все ячейки в строке
        for (let c = 1; c <= colCount; c++) {
          const cell = row.getCell(c)
          const { text, meta } = getCellDisplayAndMeta(cell)
          const font = cell.font || {}
          
          // Получаем цвета
          const colorHex = extractFontColor(font)
          const bgHex = extractFillColor(cell.fill)

          // Отладочная информация для каждой ячейки
          console.log('Cell colors:', {
            row: r,
            col: c,
            textColor: colorHex,
            bgColor: bgHex,
            font: font,
            fill: cell.fill
          })

          rowObj[c.toString()] = {
            value: text,
            meta,
            style: {
              fontWeight: font.bold ? 'bold' : 'normal',
              textAlign: cell.alignment?.horizontal || 'left',
              color: colorHex,
              backgroundColor: bgHex,
              fontSize: font.size ? `${font.size}px` : '14px',
              fontStyle: font.italic ? 'italic' : 'normal',
              textDecoration: font.underline ? 'underline' : 'none'
            }
          }
        }
        sheetData.push(rowObj)
      }

      // Добавляем недостающие строки до ROWS_COUNT
      while (sheetData.length < ROWS_COUNT) {
        const emptyRow = {}
        sheetColumns.forEach(col => {
          emptyRow[col.id] = { value: '', style: {}, meta: {} }
        })
        sheetData.push(emptyRow)
      }

      const newSheet = {
        id: idx + 1,
        name: worksheet.name || `Лист${idx + 1}`,
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


const pasteIntoSelectedCells = async () => {
  if (!selectedCells.value.length) return
  const sheet = currentSheet.value
  if (!sheet) return

  try {
    const text = await navigator.clipboard.readText()
    if (!text) return

    // Разбиваем на строки и колонки
    let rows = text
      .trim()
      .split(/\r?\n/)
      .map(r => r.split('\t'))

    const start = selectedCells.value[0]

    const rowCount = rows.length
    const colCount = Math.max(...rows.map(r => r.length))

    // 🚀 Если это "вертикальное копирование" (несколько строк, но всегда по 1 колонке)
    if (colCount === 1 && rowCount > 1) {
      rows.forEach((row, rIdx) => {
        const rowIndex = start.rowIndex + rIdx
        const colId = start.colId
        if (sheet.data[rowIndex] && sheet.data[rowIndex][colId]) {
          sheet.data[rowIndex][colId].value = row[0]
        }
      })
    }
    // 🚀 Если это "горизонтальное копирование" (одна строка, несколько колонок)
    else if (rowCount === 1 && colCount > 1) {
      rows[0].forEach((val, cIdx) => {
        const rowIndex = start.rowIndex
        const colId = (parseInt(start.colId) + cIdx).toString()
        if (sheet.data[rowIndex] && sheet.data[rowIndex][colId]) {
          sheet.data[rowIndex][colId].value = val
        }
      })
    }
    // 🚀 Если это прямоугольный блок (таблица NxM)
    else {
      rows.forEach((row, rIdx) => {
        row.forEach((val, cIdx) => {
          const rowIndex = start.rowIndex + rIdx
          const colId = (parseInt(start.colId) + cIdx).toString()
          if (sheet.data[rowIndex] && sheet.data[rowIndex][colId]) {
            sheet.data[rowIndex][colId].value = val
          }
        })
      })
    }

    // Обновляем Vue
    sheet.data = [...sheet.data]
  } catch (e) {
    console.error("Ошибка вставки:", e)
  }
}




// Обработка горячих клавиш
const handleKeyDown = (e) => {
  if ((e.ctrlKey || e.metaKey) && !e.altKey) {
    switch (e.key) {
      case 'c':
        e.preventDefault()
        copySelectedCells()
        break
      case 'x':
        e.preventDefault()
        copySelectedCells()
        deleteSelectedCells()
        break
      case 'z':
        e.preventDefault()
        // Можно добавить отмену действия при необходимости
        break
      case 'y':
        e.preventDefault()
        // Можно добавить повтор действия при необходимости
        break
      case 'f':
        e.preventDefault()
        // Фокус на поиск
        document.querySelector('.search-panel input')?.focus()
        break
      case 'o':
        if (e.ctrlKey || e.metaKey) {
          e.preventDefault()
          document.querySelector('input[type="file"]')?.click()
        }
        break
      case 's':
        if (e.ctrlKey || e.metaKey) {
          e.preventDefault()
          exportExcel()
        }
        break
        case 'v':
          e.preventDefault()
          pasteIntoSelectedCells()
          break
    }
  }
}

onMounted(() => {
  document.addEventListener('keydown', handleKeyDown)
})

onUnmounted(() => {
  document.removeEventListener('keydown', handleKeyDown)
})

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
          color: st.color ? { argb: normalizeToARGB(st.color) } : undefined,
          size: st.fontSize ? parseInt(st.fontSize) : undefined
        }
        if (st.textAlign) cell.alignment = { horizontal: st.textAlign, vertical: 'middle' }
      if (st.backgroundColor) {
        cell.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: st.backgroundColor ? { argb: normalizeToARGB(st.backgroundColor) } : undefined,
        }
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
function normalizeToARGB(hex) {
  // hex = "#RRGGBB"
  const clean = hex.replace('#', '').toUpperCase()
  return clean.length === 6 ? 'FF' + clean : clean
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
