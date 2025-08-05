<template>
  <div class="sheet-tabs">
    <div 
      v-for="sheet in sheets" 
      :key="sheet.id" 
      class="sheet-tab"
      :class="{ active: sheet.id === activeSheet }"
      @click="$emit('select-sheet', sheet.id)"
    >
      <span>{{ sheet.name }}</span>
      <button 
        @click.stop="$emit('delete-sheet', sheet.id)" 
        class="delete-tab"
        v-if="sheets.length > 1"
      >
        ×
      </button>
    </div>
    <button @click="$emit('add-sheet')" class="add-tab">+</button>
    
    <div v-if="renamingSheet" class="rename-dialog">
      <input 
        type="text" 
        v-model="newSheetName" 
        @keyup.enter="confirmRename"
        ref="renameInput"
      >
      <button @click="confirmRename">OK</button>
      <button @click="cancelRename">Отмена</button>
    </div>
  </div>
</template>

<script setup>
import { ref, nextTick, watch } from 'vue'

const props = defineProps({
  sheets: {
    type: Array,
    required: true
  },
  activeSheet: {
    type: Number,
    required: true
  }
})

const emit = defineEmits(['add-sheet', 'rename-sheet', 'select-sheet', 'delete-sheet'])

const renamingSheet = ref(null)
const newSheetName = ref('')
const renameInput = ref(null)

const startRenaming = (sheet) => {
  renamingSheet.value = sheet.id
  newSheetName.value = sheet.name
  nextTick(() => {
    renameInput.value.focus()
  })
}

const confirmRename = () => {
  if (newSheetName.value.trim()) {
    emit('rename-sheet', { 
      id: renamingSheet.value, 
      newName: newSheetName.value 
    })
  }
  cancelRename()
}

const cancelRename = () => {
  renamingSheet.value = null
  newSheetName.value = ''
}

// Double click to rename
const handleDoubleClick = (sheet) => {
  startRenaming(sheet)
}

watch(() => props.sheets, () => {
  cancelRename()
}, { deep: true })
</script>

<style scoped>
.sheet-tabs {
  display: flex;
  border-bottom: 1px solid #ddd;
  padding: 0 10px;
  background: #f0f0f0;
}

.sheet-tab {
  padding: 8px 15px;
  cursor: pointer;
  border: 1px solid #ddd;
  border-bottom: none;
  margin-right: 5px;
  background: white;
  display: flex;
  align-items: center;
  border-radius: 5px 5px 0 0;
  position: relative;
  bottom: -1px;
}

.sheet-tab.active {
  background: #42b983;
  color: white;
  border-color: #42b983;
}

.delete-tab {
  margin-left: 5px;
  background: none;
  border: none;
  color: #999;
  cursor: pointer;
  padding: 0 5px;
}

.delete-tab:hover {
  color: #ff0000;
}

.add-tab {
  padding: 0 10px;
  background: none;
  border: 1px dashed #999;
  margin-left: 5px;
  cursor: pointer;
}

.add-tab:hover {
  border-color: #42b983;
  color: #42b983;
}

.rename-dialog {
  position: absolute;
  background: white;
  border: 1px solid #ddd;
  padding: 10px;
  z-index: 100;
  box-shadow: 0 2px 10px rgba(0,0,0,0.1);
}
</style>