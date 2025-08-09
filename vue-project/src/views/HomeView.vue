<template>
  <div class="home">
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
import Toolbar from '@/components/Toolbar.vue'
import ExcelGrid from '@/components/ExcelGrid.vue'
import SearchPanel from '@/components/SearchPanel.vue'

const sheets = ref([
  {
    id: 1,
    name: 'Лист1',
    data: Array(20).fill().map((_, row) => ({
      '1': { value: `Ячейка A${row + 1}`, style: {} },
      '2': { value: `Ячейка B${row + 1}`, style: {} },
      '3': { value: `Ячейка C${row + 1}`, style: {} }
    }))
  }
])

const activeSheet = ref(1)
const columns = ref([
  { id: '1', name: 'A', width: 120 },
  { id: '2', name: 'B', width: 150 },
  { id: '3', name: 'C', width: 180 }
])

const selectedCells = ref([])
const searchResults = ref([])

const currentData = computed(() => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet ? sheet.data : []
})

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

// Обновленный поиск по нескольким словам
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
      
      // Проверяем совпадение с любым из запросов
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
    selectedCells.value = results // Выделяем все найденные ячейки
  }
}

const handleImport = (jsonData) => {
  if (!jsonData || !jsonData.length) return

  const maxColumns = Math.max(...jsonData.map(row => row.length))
  
  const newColumns = Array.from({ length: maxColumns }, (_, i) => ({
    id: (i + 1).toString(),
    name: String.fromCharCode(65 + i),
    width: 120
  }))

  const newData = jsonData.map(row => {
    const cellData = {}
    newColumns.forEach((col, colIndex) => {
      cellData[col.id] = {
        value: row[colIndex] !== undefined ? String(row[colIndex]) : '',
        style: {
          fontWeight: 'normal',
          textAlign: 'left',
          color: '#000000',
          backgroundColor: '#ffffff',
          fontSize: '14px'
        }
      }
    })
    return cellData
  })

  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet) {
    sheet.data = newData
    columns.value = newColumns
    selectedCells.value = []
    searchResults.value = []
  }
}

const exportExcel = () => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (!sheet) return
  
  const exportData = sheet.data.map(row => {
    return columns.value.map(col => row[col.id]?.value || '')
  })
  
  const worksheet = XLSX.utils.aoa_to_sheet(exportData)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name)
  XLSX.writeFile(workbook, `${sheet.name}.xlsx`)
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