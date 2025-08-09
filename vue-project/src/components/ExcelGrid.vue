<template>
  <div 
    class="excel-grid" 
    @mousedown="startSelection"
    @mouseup="endSelection"
    @mousemove="handleSelection"
  >
    <table class="excel-table">
      <thead>
        <tr>
          <th 
            v-for="col in columns" 
            :key="col.id"
            class="header-cell"
            :style="{ width: col.width + 'px' }"
          >
            <div class="header-content">
              {{ col.name }}
              <div 
                class="resize-handle"
                @mousedown.stop="startResize(col, $event)"
              ></div>
            </div>
          </th>
        </tr>
      </thead>
      <tbody>
        <tr 
          v-for="(row, rowIndex) in data" 
          :key="rowIndex"
          class="data-row"
          :data-row-index="rowIndex"
        >
          <td 
            v-for="col in columns" 
            :key="col.id"
            class="data-cell"
            :class="{
              selected: isSelected(rowIndex, col.id),
              'selection-area': isInSelectionArea(rowIndex, col.id)
            }"
            :style="{
              width: col.width + 'px',
              ...getCellStyle(row[col.id])
            }"
            :data-col-id="col.id"
            @click.stop="handleCellClick(rowIndex, col.id, $event)"
            @dblclick.stop="startEditing(rowIndex, col.id)"
          >
            <div class="cell-content">
              <span v-if="!isEditing(rowIndex, col.id)">
                {{ row[col.id]?.value || '' }}
              </span>
              <input
                v-else
                ref="editorInput"
                type="text"
                v-model="editingValue"
                @blur="stopEditing"
                @keyup.enter="stopEditing"
                @keyup.esc="cancelEditing"
                :style="getInputStyle(row[col.id])"
                class="cell-input"
              >
            </div>
          </td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script setup>
import { ref, nextTick } from 'vue'

const props = defineProps({
  data: {
    type: Array,
    required: true
  },
  columns: {
    type: Array,
    required: true
  },
  selectedCells: {
    type: Array,
    default: () => []
  }
})

const emit = defineEmits([
  'update:selected-cells',
  'update-cell',
  'resize-column'
])

const editorInput = ref(null)
const editingCell = ref(null)
const editingValue = ref('')
const isSelecting = ref(false)
const selectionStart = ref(null)

// Редактирование ячеек
const startEditing = async (rowIndex, colId) => {
  editingCell.value = { rowIndex, colId }
  editingValue.value = props.data[rowIndex][colId]?.value || ''
  await nextTick()
  editorInput.value?.focus()
}

const stopEditing = () => {
  if (editingCell.value) {
    emit('update-cell', {
      rowIndex: editingCell.value.rowIndex,
      colId: editingCell.value.colId,
      value: editingValue.value
    })
    editingCell.value = null
  }
}

const cancelEditing = () => {
  editingCell.value = null
}

const isEditing = (rowIndex, colId) => {
  return editingCell.value?.rowIndex === rowIndex && 
         editingCell.value?.colId === colId
}

// Выделение ячеек
const startSelection = (e) => {
  const cell = getCellFromEvent(e)
  if (!cell) return

  isSelecting.value = true
  selectionStart.value = cell

  if (e.ctrlKey || e.metaKey) {
    // Добавляем к существующему выделению
    const existingIndex = props.selectedCells.findIndex(
      c => c.rowIndex === cell.rowIndex && c.colId === cell.colId
    )
    
    if (existingIndex >= 0) {
      // Удаляем если уже выделена
      const newSelection = [...props.selectedCells]
      newSelection.splice(existingIndex, 1)
      emit('update:selected-cells', newSelection)
    } else {
      // Добавляем если не выделена
      emit('update:selected-cells', [...props.selectedCells, cell])
    }
  } else if (e.shiftKey && props.selectedCells.length > 0) {
    // Выделяем диапазон
    const lastSelected = props.selectedCells[props.selectedCells.length - 1]
    updateSelectionArea(lastSelected, cell)
  } else {
    // Новое выделение
    emit('update:selected-cells', [cell])
  }
}

const handleSelection = (e) => {
  if (!isSelecting.value || !selectionStart.value) return
  
  const cell = getCellFromEvent(e)
  if (cell) {
    updateSelectionArea(selectionStart.value, cell)
  }
}

const endSelection = () => {
  isSelecting.value = false
}

const handleCellClick = (rowIndex, colId, event) => {
  // Обработка кликов уже происходит в startSelection
  // Эта функция нужна только для stopPropagation
}

const getCellFromEvent = (e) => {
  const cellElement = e.target.closest('.data-cell')
  if (!cellElement) return null
  
  const rowElement = cellElement.parentElement
  const rowIndex = parseInt(rowElement.dataset.rowIndex)
  const colId = parseInt(cellElement.dataset.colId)
  
  return { rowIndex, colId }
}

const updateSelectionArea = (startCell, endCell) => {
  const startRow = Math.min(startCell.rowIndex, endCell.rowIndex)
  const endRow = Math.max(startCell.rowIndex, endCell.rowIndex)
  const startCol = Math.min(startCell.colId, endCell.colId)
  const endCol = Math.max(startCell.colId, endCell.colId)

  const selected = []
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      selected.push({ rowIndex: row, colId: col })
    }
  }
  
  emit('update:selected-cells', selected)
}

const isSelected = (rowIndex, colId) => {
  return props.selectedCells.some(
    cell => cell.rowIndex === rowIndex && cell.colId === colId
  )
}

const isInSelectionArea = (rowIndex, colId) => {
  if (!isSelecting.value || !selectionStart.value) return false
  
  const currentCell = getCurrentHoveredCell()
  if (!currentCell) return false
  
  const startRow = Math.min(selectionStart.value.rowIndex, currentCell.rowIndex)
  const endRow = Math.max(selectionStart.value.rowIndex, currentCell.rowIndex)
  const startCol = Math.min(selectionStart.value.colId, currentCell.colId)
  const endCol = Math.max(selectionStart.value.colId, currentCell.colId)
  
  return rowIndex >= startRow && rowIndex <= endRow &&
         colId >= startCol && colId <= endCol
}

const getCurrentHoveredCell = () => {
  // В реальном проекте нужно отслеживать текущую позицию мыши
  // Здесь упрощенная реализация
  return selectionStart.value
}

// Стили ячеек
const getCellStyle = (cellData) => {
  return {
    fontWeight: cellData?.style?.fontWeight || 'normal',
    textAlign: cellData?.style?.textAlign || 'left',
    color: cellData?.style?.color || '#000000',
    backgroundColor: cellData?.style?.backgroundColor || 'transparent',
    fontSize: cellData?.style?.fontSize || '14px'
  }
}

const getInputStyle = (cellData) => {
  return getCellStyle(cellData)
}

// Изменение размера столбцов
const startResize = (column, e) => {
  e.preventDefault()
  const startX = e.clientX
  const startWidth = column.width
  
  const doResize = (moveEvent) => {
    const newWidth = startWidth + (moveEvent.clientX - startX)
    emit('resize-column', { 
      colId: column.id, 
      width: Math.max(30, newWidth) 
    })
  }
  
  const stopResize = () => {
    document.removeEventListener('mousemove', doResize)
    document.removeEventListener('mouseup', stopResize)
  }
  
  document.addEventListener('mousemove', doResize)
  document.addEventListener('mouseup', stopResize, { once: true })
}
</script>

<style scoped>
.excel-grid {
  width: 100%;
  height: 100%;
  overflow: auto;
  font-family: Arial, sans-serif;
}

.excel-table {
  border-collapse: collapse;
  width: max-content;
  min-width: 100%;
}

.header-cell {
  position: relative;
  padding: 8px;
  border: 1px solid #ddd;
  background-color: #f0f0f0;
  font-weight: bold;
  text-align: center;
  user-select: none;
}

.header-content {
  display: flex;
  justify-content: center;
  align-items: center;
  position: relative;
}

.resize-handle {
  position: absolute;
  right: 0;
  top: 0;
  width: 5px;
  height: 100%;
  cursor: col-resize;
  background-color: rgba(0, 0, 0, 0.1);
}

.resize-handle:hover {
  background-color: #42b983;
}

.data-row {
  height: 30px;
}

.data-cell {
  position: relative;
  padding: 4px 8px;
  border: 1px solid #ddd;
  vertical-align: middle;
  cursor: cell;
}

.data-cell.selected {
  background-color: #e6f3ff;
  outline: 2px solid #42b983;
  outline-offset: -2px;
}

.data-cell.selection-area {
  background-color: #f0f7ff;
}

.cell-content {
  width: 100%;
  height: 100%;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.cell-input {
  width: 100%;
  height: 100%;
  border: none;
  outline: none;
  padding: 0;
  margin: 0;
  font: inherit;
  color: inherit;
  background: transparent;
}
</style>