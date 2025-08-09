
<template>
  <div class="home">
    <Toolbar 
      :current-styles="getCurrentCellStyles()"
      @apply-styles="applyStylesToSelectedCells"
      @import-data="handleImport"
      @export-excel="exportExcel"
    />
    <ExcelGrid
      :data="currentData"
      :columns="columns"
      :selected-cells="selectedCells"
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

// Инициализация данных
const sheets = ref([
  {
    id: 1,
    name: 'Лист1',
    data: Array(20).fill().map((_, row) => ({
      1: { value: `Ячейка A${row + 1}`, style: {} },
      2: { value: `Ячейка B${row + 1}`, style: {} },
      3: { value: `Ячейка C${row + 1}`, style: {} }
    }))
  }
])

const activeSheet = ref(1)
const columns = ref([
  { id: 1, name: 'A', width: 120 },
  { id: 2, name: 'B', width: 150 },
  { id: 3, name: 'C', width: 180 }
])


const selectedCells = ref([])

// Вычисляемое свойство для текущих данных
const currentData = computed(() => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet ? sheet.data : []
})
// Получение стилей для выделенных ячеек
const getCurrentCellStyles = () => {
  if (selectedCells.value.length === 0) return {
    fontWeight: 'normal',
    textAlign: 'left',
    color: '#000000',
    backgroundColor: 'transparent',
    fontSize: '14px'
  }
  
  // Берем стили первой выделенной ячейки
  const firstCell = selectedCells.value[0]
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet?.data[firstCell.rowIndex]?.[firstCell.colId]?.style || {
    fontWeight: 'normal',
    textAlign: 'left',
    color: '#000000',
    backgroundColor: 'transparent',
    fontSize: '14px'
  }
}

// Обновление выбранной ячейки
const updateSelectedCells = (cells) => {
  selectedCells.value = cells
}

// Обновление содержимого ячейки
const updateCell = ({ rowIndex, colId, value }) => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet && sheet.data[rowIndex]) {
    sheet.data[rowIndex][colId].value = value
  }
}



// Применение стилей к выделенным ячейкам
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

// Изменение размера столбца
const resizeColumn = ({ colId, width }) => {
  const column = columns.value.find(c => c.id === colId)
  if (column) {
    column.width = Math.max(50, width)
  }
}

// Импорт данных из Excel
const handleImport = (jsonData) => {
  if (!jsonData || !jsonData.length) return

  // Определяем количество колонок
  const maxColumns = Math.max(...jsonData.map(row => row.length))
  
  // Создаем новые колонки
  const newColumns = Array.from({ length: maxColumns }, (_, i) => ({
    id: i + 1,
    name: String.fromCharCode(65 + i),
    width: 120
  }))

  // Преобразуем данные
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

  // Обновляем текущий лист
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet) {
    sheet.data = newData
    columns.value = newColumns
  }
}

// Экспорт в Excel
const exportExcel = () => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (!sheet) return
  
  // Преобразуем данные
  const exportData = sheet.data.map(row => {
    return columns.value.map(col => row[col.id]?.value || '')
  })
  
  // Создаем книгу Excel
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