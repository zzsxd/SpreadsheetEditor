<template>
  <div class="excel-grid" ref="gridContainer">
    <div class="grid-header" ref="header">
      <div 
        v-for="col in columns" 
        :key="col.id" 
        class="header-cell"
        :style="{ width: col.width + 'px' }"
      >
        <span>{{ col.name }}</span>
        <div 
          class="resize-handle"
          @mousedown="startResize(col.id, $event)"
        ></div>
      </div>
    </div>
    
    <div class="grid-body" @scroll="handleScroll" ref="gridBody">
      <div 
        v-for="(row, rowIndex) in visibleRows" 
        :key="rowIndex" 
        class="grid-row"
      >
        <div 
          v-for="col in columns" 
          :key="col.id" 
          class="grid-cell"
          :style="{ 
            width: col.width+17 + 'px',
            ...row[col.id]?.style || {}
          }"
          @dblclick="startEditing(rowIndex, col.id)"
        >
          <span 
            v-if="!isEditing(rowIndex, col.id)"
            class="cell-content"
            :title="row[col.id]?.value"
          >
            {{ row[col.id]?.value }}
          </span>
          
          <input
            v-else
            type="text"
            v-model="editingValue"
            @blur="stopEditing"
            @keyup.enter="stopEditing"
            @keyup.esc="cancelEditing"
            ref="editingInput"
          >
        </div>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, nextTick, onMounted } from 'vue'

const props = defineProps({
  data: {
    type: Array,
    required: true
  },
  columns: {
    type: Array,
    required: true
  }
})

const emit = defineEmits(['update-cell', 'copy', 'cut', 'resize-column'])

const gridContainer = ref(null)
const editingInput = ref(null)
const visibleStart = ref(0)
const visibleEnd = ref(50)
const rowHeight = 30

const editingCell = ref(null)
const editingValue = ref('')

const visibleRows = computed(() => {
  return props.data.slice(visibleStart.value, visibleEnd.value)
})

const isEditing = (rowIndex, colId) => {
  return editingCell.value?.rowIndex === rowIndex && 
         editingCell.value?.colId === colId
}

const startEditing = (rowIndex, colId) => {
  editingCell.value = { rowIndex, colId }
  editingValue.value = props.data[rowIndex][colId]?.value || ''
  nextTick(() => {
    editingInput.value?.focus()
  })
}

const stopEditing = () => {
  if (editingCell.value) {
    const { rowIndex, colId } = editingCell.value
    emit('update-cell', { 
      rowIndex, 
      colId, 
      value: editingValue.value 
    })
    editingCell.value = null
  }
}

const cancelEditing = () => {
  editingCell.value = null
}

const handleScroll = (e) => {
  const scrollTop = e.target.scrollTop
  visibleStart.value = Math.floor(scrollTop / rowHeight)
  visibleEnd.value = visibleStart.value + Math.ceil(gridContainer.value?.clientHeight / rowHeight) + 10
}

const handleCopy = (text) => {
  emit('copy', text)
}

const handleCut = (rowIndex, colId) => {
  emit('cut', { rowIndex, colId })
}

onMounted(() => {
  if (gridContainer.value) {
    visibleEnd.value = Math.ceil(gridContainer.value.clientHeight / rowHeight) + 10
  }
})
const startResize = (colId, e) => {
  e.preventDefault()
  const startX = e.clientX
  const column = props.columns.find(c => c.id === colId)
  const startWidth = column.width
  
  const doResize = (moveEvent) => {
    const newWidth = startWidth + (moveEvent.clientX - startX)
    emit('resize-column', { colId, width: newWidth })
  }
  
  const stopResize = () => {
    document.removeEventListener('mousemove', doResize)
    document.removeEventListener('mouseup', stopResize)
  }
  
  document.addEventListener('mousemove', doResize)
  document.addEventListener('mouseup', stopResize)
}
</script>

<style scoped>
.excel-grid {
  display: flex;
  flex-direction: column;
  height: 100%;
  overflow: hidden;
}

.grid-header {
  display: flex;
  background: #f0f0f0;
  border-bottom: 1px solid #ddd;
  position: sticky;
  top: 0;
  z-index: 10;
}

.header-cell {
  position: relative;
  padding: 8px;
  border-right: 1px solid #ddd;
  font-weight: bold;
  text-align: center;
  user-select: none;
  flex-shrink: 0;
}

.grid-body {
  flex: 1;
  overflow-y: auto;
  position: relative;
}

.grid-row {
  display: flex;
  min-height: 30px;
  border-bottom: 1px solid #eee;
}

.grid-cell {
  padding: 5px;
  border-right: 1px solid #eee;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  flex-shrink: 0;
  box-sizing: border-box;
}

.cell-content {
  display: inline-block;
  width: 100%;
  height: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
}

.resize-handle {
  position: absolute;
  top: 0;
  right: 0;
  width: 5px;
  height: 100%;
  cursor: col-resize;
  background: transparent;
}

.resize-handle:hover {
  background: #42b983;
}

.grid-cell input {
  width: calc(100% - 10px);
  height: calc(100% - 10px);
  border: 1px solid #42b983;
  padding: 0 5px;
  box-sizing: border-box;
}
</style>