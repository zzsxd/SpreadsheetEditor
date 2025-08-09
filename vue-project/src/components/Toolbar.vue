<template>
  <div class="toolbar">
    <div class="file-actions">
      <button @click="triggerImport">
        <i class="material-icons">file_upload</i> Импорт
      </button>
      <input 
        type="file" 
        ref="fileInput" 
        @change="handleImport" 
        accept=".xlsx,.xls" 
        style="display: none"
      >
      <button @click="$emit('export-excel')">
        <i class="material-icons">file_download</i> Экспорт
      </button>
    </div>
    
    <div class="format-section">
      <button 
        @click="toggleBold" 
        :class="{ active: currentStyles.fontWeight === 'bold' }"
      >
        <span class="icon-bold">B</span>
      </button>
      
      <select v-model="currentStyles.textAlign" @change="updateStyles">
        <option value="left">По левому краю</option>
        <option value="center">По центру</option>
        <option value="right">По правому краю</option>
      </select>
      
      <input 
        type="color" 
        v-model="currentStyles.color" 
        @change="updateStyles" 
        title="Цвет текста"
      >
      
      <input 
        type="color" 
        v-model="currentStyles.backgroundColor" 
        @change="updateStyles" 
        title="Цвет фона"
      >
      
      <select v-model="currentStyles.fontSize" @change="updateStyles">
        <option value="12px">12px</option>
        <option value="14px">14px</option>
        <option value="16px">16px</option>
        <option value="18px">18px</option>
        <option value="20px">20px</option>
      </select>
    </div>
  </div>
</template>

<script setup>
import { ref, watch } from 'vue'
import * as XLSX from 'xlsx'

const props = defineProps({
  currentStyles: {
    type: Object,
    default: () => ({
      fontWeight: 'normal',
      textAlign: 'left',
      color: '#000000',
      backgroundColor: '',
      fontSize: '14px'
    })
  }
})

const emit = defineEmits([
  'apply-styles',
  'import-data',
  'export-excel'
])

const fileInput = ref(null)
const currentStyles = ref({ ...props.currentStyles })

watch(() => props.currentStyles, (newStyles) => {
  currentStyles.value = { ...newStyles }
}, { deep: true })

const triggerImport = () => {
  fileInput.value.click()
}

const handleImport = async (e) => {
  const file = e.target.files[0]
  if (!file) return
  
  try {
    const data = await file.arrayBuffer()
    const workbook = XLSX.read(data)
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
    const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 })
    emit('import-data', jsonData)
  } catch (error) {
    console.error('Ошибка импорта:', error)
    alert('Не удалось загрузить файл')
  }
}

const toggleBold = () => {
  currentStyles.value.fontWeight = currentStyles.value.fontWeight === 'bold' ? 'normal' : 'bold'
  updateStyles()
}

const updateStyles = () => {
  emit('apply-styles', { ...currentStyles.value })
}
</script>

<style scoped>
.toolbar {
  padding: 10px;
  background: #f5f5f5;
  border-bottom: 1px solid #ddd;
}

.format-section {
  display: flex;
  gap: 10px;
  align-items: center;
}

button {
  padding: 5px 10px;
  cursor: pointer;
}

button.active {
  background: #42b983;
  color: white;
}

.icon-bold {
  font-weight: bold;
}

select, input[type="color"] {
  padding: 5px;
  border-radius: 4px;
  border: 1px solid #ddd;
}

input[type="color"] {
  width: 30px;
  height: 30px;
  padding: 2px;
}
</style>