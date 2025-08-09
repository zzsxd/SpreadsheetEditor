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
      @update:selected-cells="selectedCells = $event"
      @update-cell="updateCell"
      @resize-column="resizeColumn"
    />
  </div>
</template>

<script setup>
import { ref, computed } from 'vue'
import * as XLSX from 'xlsx'
import SheetTabs from '@/components/SheetTabs.vue'
import Toolbar from '@/components/Toolbar.vue'
import ExcelGrid from '@/components/ExcelGrid.vue'
import SearchPanel from '@/components/SearchPanel.vue'

// Константы для настройки таблицы
const ROWS_COUNT = 100
const COLS_COUNT = 26
const DEFAULT_COL_WIDTH = 120

// Сначала объявляем все переменные
const nextSheetId = ref(2)
const sheets = ref([])
const activeSheet = ref(null)
const columns = ref([])
const selectedCells = ref([])
const searchResults = ref([])

// Инициализируем columns до создания листов
columns.value = Array.from({ length: COLS_COUNT }, (_, i) => ({
  id: (i + 1).toString(),
  name: String.fromCharCode(65 + i),
  width: DEFAULT_COL_WIDTH
}))

// Функция создания нового листа (должна быть объявлена после columns)
function createNewSheet(name) {
  const sheetId = nextSheetId.value++
  return {
    id: sheetId,
    name: name || `Лист${sheetId}`,
    data: Array(ROWS_COUNT).fill().map(() => {
      const row = {}
      columns.value.forEach(col => {
        row[col.id] = { value: '', style: {} }
      })
      return row
    })
  }
}

// Теперь инициализируем sheets
sheets.value = [createNewSheet('Лист1')]
activeSheet.value = sheets.value[0].id

// Вычисляемые свойства
const currentData = computed(() => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet ? sheet.data : []
})

// Методы для работы с листами
const selectSheet = (sheetId) => {
  activeSheet.value = sheetId
  selectedCells.value = []
  searchResults.value = []
}

const addSheet = () => {
  const newSheet = createNewSheet()
  sheets.value.push(newSheet)
  selectSheet(newSheet.id)
}

const renameSheet = ({ id, newName }) => {
  const sheet = sheets.value.find(s => s.id === id)
  if (sheet && newName.trim()) {
    sheet.name = newName.trim()
  }
}

const deleteSheet = (sheetId) => {
  if (sheets.value.length <= 1) return
  
  if (confirm('Вы уверены, что хотите удалить этот лист?')) {
    const index = sheets.value.findIndex(s => s.id === sheetId)
    if (index !== -1) {
      sheets.value.splice(index, 1)
      if (activeSheet.value === sheetId) {
        selectSheet(sheets.value[0].id)
      }
    }
  }
}

// Остальные методы остаются без изменений
const getCurrentCellStyles = () => {
  if (selectedCells.value.length === 0) return {
    fontWeight: 'normal',
    textAlign: 'left',
    color: '#000000',
    backgroundColor: '#ffffff',
    fontSize: '14px'
  }
  
  const firstCell = selectedCells.value[0]
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet?.data[firstCell.rowIndex]?.[firstCell.colId]?.style || {
    fontWeight: 'normal',
    textAlign: 'left',
    color: '#000000',
    backgroundColor: '#ffffff',
    fontSize: '14px'
  }
}

const updateCell = ({ rowIndex, colId, value }) => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet && sheet.data[rowIndex]) {
    sheet.data[rowIndex][colId].value = value
  }
}

const applyStylesToSelectedCells = (styles) => {
  if (selectedCells.value.length === 0) return
  
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (!sheet) return
  
  selectedCells.value.forEach(({ rowIndex, colId }) => {
    if (sheet.data[rowIndex]?.[colId]) {
      sheet.data[rowIndex][colId].style = {
        ...sheet.data[rowIndex][colId].style,
        ...styles
      }
    }
  })
}

const resizeColumn = ({ colId, width }) => {
  const column = columns.value.find(c => c.id === colId)
  if (column) {
    column.width = Math.max(50, width)
  }
}

const handleSearch = ({ queries, matchCase, columns: searchColumns }) => {
  if (!queries || queries.length === 0) {
    searchResults.value = []
    return
  }

  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (!sheet) return

  const results = []

  sheet.data.forEach((row, rowIndex) => {
    Object.entries(row).forEach(([colId, cell]) => {
      if (searchColumns !== 'all' && !searchColumns.includes(colId)) return
      
      const cellText = matchCase 
        ? cell.value 
        : cell.value?.toLowerCase()
      
      const anyMatch = queries.some(query => {
        const searchWord = matchCase ? query : query.toLowerCase()
        return cellText?.includes(searchWord)
      })
      
      if (anyMatch) {
        results.push({ rowIndex, colId })
      }
    })
  })

  searchResults.value = results
  if (results.length > 0) {
    selectedCells.value = results
  }
}

const handleImport = (jsonData) => {
  if (!jsonData || !jsonData.length) return

  const maxColumns = Math.max(...jsonData.map(row => row.length))
  
  const newColumns = Array.from({ length: maxColumns }, (_, i) => ({
    id: (i + 1).toString(),
    name: String.fromCharCode(65 + i),
    width: DEFAULT_COL_WIDTH
  }))

  const newData = jsonData.map(row => {
    const cellData = {}
    newColumns.forEach((col, colIndex) => {
      cellData[col.id] = {
        value: row[colIndex] !== undefined ? String(row[colIndex]) : '',
        style: {}
      }
    })
    return cellData
  })

  const newSheet = createNewSheet('Импортированный лист')
  newSheet.data = newData
  sheets.value.push(newSheet)
  columns.value = newColumns
  selectSheet(newSheet.id)
}

const exportExcel = () => {
  const workbook = XLSX.utils.book_new()
  
  // Экспортируем каждый лист
  sheets.value.forEach(sheet => {
    const exportData = sheet.data.map(row => {
      return columns.value.map(col => row[col.id]?.value || '')
    })
    
    const worksheet = XLSX.utils.aoa_to_sheet(exportData)
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name)
  })
  
  // Сохраняем файл с именем первого листа
  const fileName = sheets.value[0]?.name || 'workbook'
  XLSX.writeFile(workbook, `${fileName}.xlsx`)
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