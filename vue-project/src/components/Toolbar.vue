<template>
  <div class="toolbar">
    <button class="toolbar-btn" @click="triggerImport">
      <i class="material-icons">file_upload</i> Импорт
    </button>
    <input 
      type="file" 
      ref="fileInput" 
      @change="handleFileChange" 
      accept=".xlsx,.xls" 
      style="display: none"
    >
    
    <button class="toolbar-btn" @click="$emit('export')">
      <i class="material-icons">file_download</i> Экспорт
    </button>
    
    <button class="toolbar-btn" @click="$emit('add-column')">
      <i class="material-icons">add</i> Столбец
    </button>
    
    <div class="format-actions">
      <button class="toolbar-btn" @click="applyFormat({ fontWeight: fontWeight === 'bold' ? 'normal' : 'bold' })">
        <i class="material-icons">format_bold</i>
      </button>
      
      <button class="toolbar-btn" @click="applyFormat({ textAlign: 'left' })">
        <i class="material-icons">format_align_left</i>
      </button>
      
      <button class="toolbar-btn" @click="applyFormat({ textAlign: 'center' })">
        <i class="material-icons">format_align_center</i>
      </button>
      
      <button class="toolbar-btn" @click="applyFormat({ textAlign: 'right' })">
        <i class="material-icons">format_align_right</i>
      </button>
      
      <input type="color" v-model="colorValue" @change="applyFormat({ color: colorValue })">
      
      <select v-model="fontSize" @change="applyFormat({ fontSize: fontSize + 'px' })">
        <option v-for="size in [8, 10, 12, 14, 16, 18, 20, 24]" :value="size">{{ size }}px</option>
      </select>
    </div>
  </div>
</template>

<script setup>
import { ref } from 'vue'

const emit = defineEmits(['import', 'export', 'format-cell', 'add-column'])
const fileInput = ref(null)
const fontWeight = ref('normal')
const colorValue = ref('#000000')
const fontSize = ref(12)

const triggerImport = () => {
  fileInput.value.click()
}

const handleFileChange = (e) => {
  const file = e.target.files[0]
  if (file) {
    emit('import', file)
  }
}

const applyFormat = (style) => {
  if (style.fontWeight) {
    fontWeight.value = style.fontWeight
  }
  emit('format-cell', { style })
}
</script>

<style scoped>
.toolbar {
  display: flex;
  gap: 10px;
  padding: 10px;
  background: #2c3e50;
  border-bottom: 1px solid #ddd;
  align-items: center;
}

.toolbar-btn {
  background: #42b983;
  color: white;
  border: none;
  border-radius: 4px;
  padding: 8px 12px;
  display: flex;
  align-items: center;
  gap: 5px;
  cursor: pointer;
  transition: background 0.2s;
}

.toolbar-btn:hover {
  background: #369f6e;
}

.format-actions {
  display: flex;
  gap: 8px;
  align-items: center;
  margin-left: auto;
}

.material-icons {
  font-size: 18px;
}

input[type="color"] {
  width: 30px;
  height: 30px;
  border: none;
  cursor: pointer;
}

select {
  padding: 5px;
  border-radius: 4px;
  border: 1px solid #ddd;
}
</style>