import { defineStore } from 'pinia'
import { ref } from 'vue'

export const useExcelStore = defineStore('excel', () => {
  const sheets = ref([
    { id: 1, name: 'Лист1', data: [] },
    { id: 2, name: 'Лист2', data: [] }
  ])
  
  const activeSheet = ref(1)
  const columns = ref([
    { id: 1, name: 'A', width: 100 },
    { id: 2, name: 'B', width: 150 },
    { id: 3, name: 'C', width: 200 }
  ])
  
  // Initialize sample data
  const initializeData = () => {
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
  
  initializeData()
  
  return { sheets, activeSheet, columns }
})