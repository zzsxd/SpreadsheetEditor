<template>
  <div class="home">
    <Toolbar 
      @import="handleImport"
      @export="handleExport"
      @format-cell="formatCell"
      @add-column="addColumn"
    />
    <SearchPanel 
      :columns="columns"
      @search="handleSearch"
    />
    <SheetTabs 
      :sheets="sheets"
      :active-sheet="activeSheet"
      @add-sheet="addSheet"
      @rename-sheet="renameSheet"
      @select-sheet="selectSheet"
      @delete-sheet="deleteSheet"
    />
    <ExcelGrid 
      :data="currentData"
      :columns="columns"
      @update-cell="updateCell"
      @copy="handleCopy"
      @cut="handleCut"
      @resize-column="resizeColumn"
    />
  </div>
</template>

<script setup>
import { ref, computed } from 'vue'
import * as XLSX from 'xlsx'
import Toolbar from '@/components/Toolbar.vue'
import SearchPanel from '@/components/SearchPanel.vue'
import SheetTabs from '@/components/SheetTabs.vue'
import ExcelGrid from '@/components/ExcelGrid.vue'

// Инициализация данных
const sheets = ref([
  { id: 1, name: 'Лист1', data: [] }
])
const activeSheet = ref(1)
const columns = ref([
  { id: 1, name: 'A', width: 120 },
  { id: 2, name: 'B', width: 150 },
  { id: 3, name: 'C', width: 180 }
])

// Инициализация начальных данных
initializeData()

function initializeData() {
  sheets.value.forEach(sheet => {
    sheet.data = Array(20).fill().map((_, row) => 
      columns.value.reduce((acc, col) => {
        acc[col.id] = { 
          value: `Ячейка ${col.name}${row + 1}`,
          style: {}
        }
        return acc
      }, {})
    )
  })
}

// Computed свойства
const currentData = computed(() => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet ? sheet.data : []
})

// Методы для работы с данными
function addColumn() {
  const newId = columns.value.length > 0 
    ? Math.max(...columns.value.map(c => c.id)) + 1 
    : 1
  const newName = String.fromCharCode(65 + columns.value.length)
  
  columns.value.push({
    id: newId,
    name: newName,
    width: 120
  })
  
  // Добавляем столбец во все строки всех листов
  sheets.value.forEach(sheet => {
    sheet.data.forEach(row => {
      row[newId] = { value: '', style: {} }
    })
  })
}

function resizeColumn({ colId, width }) {
  const column = columns.value.find(c => c.id === colId)
  if (column) {
    column.width = Math.max(50, width) // Минимальная ширина 50px
  }
}

const handleImport = (file) => {
  const reader = new FileReader()
  reader.onload = (e) => {
    const data = new Uint8Array(e.target.result)
    const workbook = XLSX.read(data, { type: 'array' })
    
    // Clear existing sheets
    sheets.value = []
    
    workbook.SheetNames.forEach((name, index) => {
      const worksheet = workbook.Sheets[name]
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 })
      
      // Transform to our format
      const formattedData = jsonData.map(row => {
        return columns.value.reduce((acc, col, colIndex) => {
          acc[col.id] = { 
            value: row[colIndex] || '',
            style: {}
          }
          return acc
        }, {})
      })
      
      sheets.value.push({
        id: index + 1,
        name: name,
        data: formattedData
      })
    })
    
    if (sheets.value.length > 0) {
      activeSheet.value = sheets.value[0].id
    }
  }
  reader.readAsArrayBuffer(file)
}

const handleExport = () => {
  const workbook = XLSX.utils.book_new()
  
  sheets.value.forEach(sheet => {
    // Transform our data to worksheet format
    const wsData = sheet.data.map(row => {
      return columns.value.map(col => row[col.id]?.value || '')
    })
    
    const worksheet = XLSX.utils.aoa_to_sheet(wsData)
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name)
  })
  
  XLSX.writeFile(workbook, 'excel-export.xlsx')
}

const updateCell = ({ rowIndex, colId, value }) => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet && sheet.data[rowIndex]) {
    sheet.data[rowIndex][colId].value = value
  }
}

const formatCell = ({ rowIndex, colId, style }) => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet && sheet.data[rowIndex]) {
    sheet.data[rowIndex][colId].style = { 
      ...sheet.data[rowIndex][colId].style,
      ...style 
    }
  }
}

const handleCopy = (data) => {
  navigator.clipboard.writeText(data)
}

const handleCut = ({ rowIndex, colId }) => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet && sheet.data[rowIndex]) {
    const value = sheet.data[rowIndex][colId].value
    navigator.clipboard.writeText(value)
    sheet.data[rowIndex][colId].value = ''
  }
}

const handleSearch = ({ terms, searchInColumns }) => {
  // Implement search logic
  console.log('Searching for:', terms, 'in columns:', searchInColumns)
}

const addSheet = () => {
  const newId = Math.max(...sheets.value.map(s => s.id), 0) + 1
  sheets.value.push({
    id: newId,
    name: `Лист${newId}`,
    data: Array(20).fill().map((_, row) => 
      columns.value.reduce((acc, col) => {
        acc[col.id] = { 
          value: `Ячейка ${col.name}${row + 1}`,
          style: {}
        }
        return acc
      }, {})
    )
  })
  activeSheet.value = newId
}

const renameSheet = ({ id, newName }) => {
  const sheet = sheets.value.find(s => s.id === id)
  if (sheet) {
    sheet.name = newName
  }
}

const selectSheet = (id) => {
  activeSheet.value = id
}

const deleteSheet = (id) => {
  if (sheets.value.length > 1) {
    sheets.value = sheets.value.filter(s => s.id !== id)
    if (activeSheet.value === id) {
      activeSheet.value = sheets.value[0].id
    }
  }
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