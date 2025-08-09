<template>
  <div class="excel-container">
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
                'selection-area': isInSelectionArea(rowIndex, col.id),
                'search-match': isSearchMatch(rowIndex, col.id)
              }"
              :style="{
                width: col.width + 'px',
                maxWidth: col.width + 'px',
                ...getCellStyle(row[col.id])
              }"
              :data-col-id="col.id"
              @click.stop="handleCellClick(rowIndex, col.id, $event)"
              @dblclick.stop="startEditing(rowIndex, col.id, $event)"
              @mouseenter="showTooltipIfNeeded($event, row[col.id]?.value)"
              @mouseleave="hideTooltip"
            >
              <div class="cell-content" :class="{ 'text-truncated': shouldTruncate(row[col.id]?.value) }">
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

    <div 
      v-if="tooltip.visible" 
      class="tooltip" 
      :style="{
        left: `${tooltip.x}px`,
        top: `${tooltip.y}px`
      }"
    >
      {{ tooltip.text }}
    </div>
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
  },
  searchResults: {
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
const tooltip = ref({
  visible: false,
  text: '',
  x: 0,
  y: 0
})

// Редактирование ячеек
const startEditing = async (rowIndex, colId, event) => {
  editingCell.value = { rowIndex, colId }
  editingValue.value = props.data[rowIndex][colId]?.value || ''
  await nextTick()
  if (editorInput.value && editorInput.value[0]) {
    editorInput.value[0].focus()
  }
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

  // Ctrl/Cmd + клик - добавить/удалить из выделения
  if (e.ctrlKey || e.metaKey) {
    const existingIndex = props.selectedCells.findIndex(
      c => c.rowIndex === cell.rowIndex && c.colId === cell.colId
    )
    
    if (existingIndex >= 0) {
      const newSelection = [...props.selectedCells]
      newSelection.splice(existingIndex, 1)
      emit('update:selected-cells', newSelection)
    } else {
      emit('update:selected-cells', [...props.selectedCells, cell])
    }
    return
  }

  // Shift + клик - выделить диапазон
  if (e.shiftKey && props.selectedCells.length > 0) {
    const lastSelected = props.selectedCells[props.selectedCells.length - 1]
    updateSelectionArea(lastSelected, cell)
    return
  }

  // Обычный клик - новое выделение
  isSelecting.value = true
  selectionStart.value = cell
  emit('update:selected-cells', [cell])
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
  event.stopPropagation()
}

const getCellFromEvent = (e) => {
  const cellElement = e.target.closest('.data-cell')
  if (!cellElement) return null
  
  const rowElement = cellElement.parentElement
  const rowIndex = parseInt(rowElement.dataset.rowIndex)
  const colId = cellElement.dataset.colId
  
  return { rowIndex, colId }
}

const updateSelectionArea = (startCell, endCell) => {
  const startRow = Math.min(startCell.rowIndex, endCell.rowIndex)
  const endRow = Math.max(startCell.rowIndex, endCell.rowIndex)
  const startCol = Math.min(parseInt(startCell.colId), parseInt(endCell.colId))
  const endCol = Math.max(parseInt(startCell.colId), parseInt(endCell.colId))

  const selected = []
  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      selected.push({ rowIndex: row, colId: col.toString() })
    }
  }
  
  emit('update:selected-cells', selected)
}

const isSelected = (rowIndex, colId) => {
  return props.selectedCells.some(
    cell => cell.rowIndex === rowIndex && cell.colId === colId.toString()
  )
}

const isInSelectionArea = (rowIndex, colId) => {
  if (!isSelecting.value || !selectionStart.value) return false
  
  const startRow = Math.min(selectionStart.value.rowIndex, rowIndex)
  const endRow = Math.max(selectionStart.value.rowIndex, rowIndex)
  const startCol = Math.min(parseInt(selectionStart.value.colId), parseInt(colId))
  const endCol = Math.max(parseInt(selectionStart.value.colId), parseInt(colId))
  
  return rowIndex >= startRow && rowIndex <= endRow &&
         parseInt(colId) >= startCol && parseInt(colId) <= endCol
}

const isSearchMatch = (rowIndex, colId) => {
  return props.searchResults.some(
    cell => cell.rowIndex === rowIndex && cell.colId === colId.toString()
  )
}

// Автоскрытие длинного текста
const shouldTruncate = (text) => {
  if (!text) return false
  return text.length > 20
}

// Подсказки для обрезанного текста
const showTooltipIfNeeded = (event, text) => {
  if (!shouldTruncate(text)) return
  
  const cell = event.target.closest('.data-cell')
  if (!cell) return
  
  const cellRect = cell.getBoundingClientRect()
  tooltip.value = {
    visible: true,
    text,
    x: cellRect.left,
    y: cellRect.bottom
  }
}

const hideTooltip = () => {
  tooltip.value.visible = false
}

// Стили ячеек
const getCellStyle = (cellData) => {
  const style = cellData?.style || {}
  return {
    fontWeight: style.fontWeight || 'normal',
    textAlign: style.textAlign || 'left',
    color: style.color || '#000000',
    backgroundColor: style.backgroundColor || '#ffffff',
    fontSize: style.fontSize || '14px'
  }
}

const getInputStyle = (cellData) => {
  return {
    ...getCellStyle(cellData),
    width: '100%',
    height: '100%'
  }
}

// Изменение размера столбцов
const startResize = (column, e) => {
  e.preventDefault()
  const startX = e.clientX
  const startWidth = column.width
  
  const doResize = (moveEvent) => {
    const newWidth = startWidth + (moveEvent.clientX - startX)
    if (Math.abs(newWidth - column.width) > 2) {
      emit('resize-column', { 
        colId: column.id, 
        width: Math.max(30, newWidth) 
      })
    }
  }
  
  const stopResize = () => {
    document.removeEventListener('mousemove', doResize)
    document.removeEventListener('mouseup', stopResize)
    document.body.style.cursor = ''
  }
  
  document.body.style.cursor = 'col-resize'
  document.addEventListener('mousemove', doResize)
  document.addEventListener('mouseup', stopResize, { once: true })
}
</script>

<style scoped>
.excel-container {
  position: relative;
  width: 100%;
  height: 100%;
  overflow: hidden;
}

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
  table-layout: fixed;
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
  height: 100%;
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
  overflow: hidden;
}

.data-cell.selected {
  background-color: #e6f3ff;
  outline: 2px solid #42b983;
  outline-offset: -2px;
}

.data-cell.selection-area {
  background-color: #f0f7ff;
}

.data-cell.search-match {
  background-color: #fffacd;
  outline: 2px solid #ffcc00;
  outline-offset: -2px;
}

.cell-content {
  width: 100%;
  height: 100%;
  overflow: hidden;
  white-space: nowrap;
}

.cell-content.text-truncated {
  text-overflow: ellipsis;
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

.tooltip {
  position: fixed;
  background-color: rgba(0, 0, 0, 0.8);
  color: white;
  padding: 4px 8px;
  border-radius: 4px;
  font-size: 12px;
  z-index: 1000;
  pointer-events: none;
  max-width: 300px;
  word-wrap: break-word;
  white-space: normal;
}
</style>