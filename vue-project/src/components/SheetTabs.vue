<template>
  <div class="sheet-tabs-container">
    <div class="sheet-tabs-scroll">
      <div class="sheet-tabs">
        <div 
          v-for="sheet in sheets" 
          :key="sheet.id"
          class="sheet-tab"
          :class="{ active: sheet.id === activeSheet }"
          @click="selectSheet(sheet.id)"
          @dblclick="startRenaming(sheet)"
        >
          <span v-if="!isRenaming(sheet.id)">{{ sheet.name }}</span>
          <input
            v-else
            type="text"
            v-model="renameValue"
            @blur="confirmRename"
            @keyup.enter="confirmRename"
            @keyup.esc="cancelRename"
            @click.stop
            ref="renameInput"
          >
          <button 
            class="delete-tab"
            @click.stop="deleteSheet(sheet.id)"
            v-if="sheets.length > 1"
          >
            ×
          </button>
        </div>
        <button class="add-tab" @click="addSheet">
          +
        </button>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, nextTick } from 'vue'

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

const emit = defineEmits([
  'select-sheet',
  'add-sheet',
  'rename-sheet',
  'delete-sheet'
])

const renamingSheetId = ref(null)
const renameValue = ref('')
const renameInput = ref(null)

const isRenaming = (sheetId) => {
  return renamingSheetId.value === sheetId
}

const selectSheet = (sheetId) => {
  if (!isRenaming(sheetId)) {
    emit('select-sheet', sheetId)
  }
}

const startRenaming = async (sheet) => {
  renamingSheetId.value = sheet.id
  renameValue.value = sheet.name
  await nextTick()
  renameInput.value[0].focus()
  renameInput.value[0].select()
}

const confirmRename = () => {
  if (renameValue.value.trim()) {
    emit('rename-sheet', {
      id: renamingSheetId.value,
      newName: renameValue.value.trim()
    })
  }
  cancelRename()
}

const cancelRename = () => {
  renamingSheetId.value = null
  renameValue.value = ''
}

const addSheet = () => {
  emit('add-sheet')
}

const deleteSheet = (sheetId) => {
  emit('delete-sheet', sheetId)
}
</script>

<style scoped>
.sheet-tabs-container {
  background: #f0f0f0;
  border-bottom: 1px solid #ddd;
  overflow: hidden;
}

.sheet-tabs-scroll {
  overflow-x: auto;
  overflow-y: hidden;
  white-space: nowrap;
  padding: 0 10px;
}

.sheet-tabs {
  display: inline-flex;
  height: 30px;
}

.sheet-tab {
  position: relative;
  padding: 0 25px 0 15px;
  margin-right: 2px;
  background: #e0e0e0;
  border: 1px solid #ddd;
  border-bottom: none;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  min-width: 80px;
  max-width: 150px;
  height: 100%;
  user-select: none;
}

.sheet-tab.active {
  background: white;
  border-top: 2px solid #42b983;
  padding-top: 0;
}

.sheet-tab span {
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.sheet-tab input {
  width: 100%;
  border: none;
  outline: none;
  background: transparent;
  font-size: inherit;
  font-family: inherit;
}

.delete-tab {
  position: absolute;
  right: 5px;
  top: 50%;
  transform: translateY(-50%);
  background: none;
  border: none;
  cursor: pointer;
  color: #999;
  font-size: 16px;
  line-height: 1;
  padding: 0;
  width: 16px;
  height: 16px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 50%;
}

.delete-tab:hover {
  background: #e74c3c;
  color: white;
}

.add-tab {
  padding: 0 15px;
  background: none;
  border: 1px dashed #999;
  margin-left: 5px;
  cursor: pointer;
  color: #555;
  font-weight: bold;
  min-width: 30px;
}

.add-tab:hover {
  border-color: #42b983;
  color: #42b983;
}
</style>