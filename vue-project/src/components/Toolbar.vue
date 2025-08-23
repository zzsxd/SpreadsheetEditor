<template>
  <div class="toolbar">
    <div class="file-actions">
      <button @click="triggerImport">
        Импорт
      </button>
      <input 
        type="file" 
        ref="fileInput" 
        @change="handleImport" 
        accept=".xlsx,.xls" 
        style="display: none"
      >
      <button @click="$emit('export-excel')">
        Экспорт
      </button>
    </div>
    
    <div class="format-section">
      <button 
        @click="toggleBold" 
        :class="{ active: currentStyles.fontWeight === 'bold' }"
        title="Жирный текст"
      >
        Ж
      </button>
      
      <select v-model="currentStyles.textAlign" @change="updateStyles" title="Выравнивание текста">
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
      
      <select v-model="currentStyles.fontSize" @change="updateStyles" title="Размер шрифта">
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

const props = defineProps({
  currentStyles: {
    type: Object,
    default: () => ({
      fontWeight: 'normal',
      textAlign: 'left',
      color: '#000000',
      backgroundColor: '#ffffff',
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

const handleImport = (e) => {
  emit('import-data', e)
  e.target.value = ''
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
  display: flex;
  flex-wrap: wrap;
  gap: 15px;
  align-items: center;
}

.file-actions {
  display: flex;
  gap: 10px;
}

.format-section {
  display: flex;
  gap: 10px;
  align-items: center;
  flex-wrap: wrap;
}

button {
  padding: 5px 10px;
  cursor: pointer;
  border: 1px solid #ddd;
  border-radius: 4px;
  background: white;
  min-width: 80px;
}

button:hover {
  background: #f0f0f0;
}

button.active {
  background: #42b983;
  color: white;
}

select, input[type="color"] {
  padding: 5px;
  border-radius: 4px;
  border: 1px solid #ddd;
  height: 32px;
}

input[type="color"] {
  width: 32px;
  padding: 2px;
  cursor: pointer;
}
</style>