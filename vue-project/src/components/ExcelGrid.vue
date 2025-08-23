<template>
  <div class="excel-container">
    <div
      ref="scrollEl"
      class="excel-grid"
      @scroll="onScroll"
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
                <div class="resize-handle" @mousedown.stop="startResize(col, $event)"></div>
              </div>
            </th>
          </tr>
        </thead>

        <tbody>
          <!-- верхний спейсер -->
          <tr v-if="topSpacer > 0" class="spacer-row">
            <td :colspan="columns.length" :style="{ height: topSpacer + 'px', padding: 0, border: 0 }"></td>
          </tr>

          <!-- видимые строки -->
          <tr
            v-for="rowAbsoluteIndex in visibleIndexes"
            :key="rowAbsoluteIndex"
            class="data-row"
            :data-row-index="rowAbsoluteIndex"
            :style="{ height: rowHeight + 'px' }"
          >
              <td
                v-for="col in columns"
                :key="col.id"
                class="data-cell"
                :class="{
                  selected: isSelected(rowAbsoluteIndex, col.id),
                  'selection-area': isInSelectionArea(rowAbsoluteIndex, col.id),
                  'search-match': isSearchMatch(rowAbsoluteIndex, col.id)
                }"
                :style="{
                  width: col.width + 'px',
                  maxWidth: col.width + 'px',
                  ...getCellStyle(data[rowAbsoluteIndex]?.[col.id])
                }"
                :data-col-id="col.id"
                @click.stop="handleCellClick(rowAbsoluteIndex, col.id, $event)"
                @dblclick.stop="startEditing(rowAbsoluteIndex, col.id, $event)"
                @mouseenter="showTooltipIfNeeded($event, data[rowAbsoluteIndex]?.[col.id]?.value)"
                @mouseleave="hideTooltip"
              >
                  <div v-if="isFirstSearchMatch(rowAbsoluteIndex, col.id)" class="search-indicator">
                    {{ getSearchMatchIndex(rowAbsoluteIndex, col.id) + 1 }}
                  </div>
              <div class="cell-content" :class="{ 'text-truncated': shouldTruncate(data[rowAbsoluteIndex]?.[col.id]?.value) }">
                <template v-if="!isEditing(rowAbsoluteIndex, col.id)">
                  <!-- если есть гиперссылка -->
                  <a
                    v-if="data[rowAbsoluteIndex]?.[col.id]?.meta?.hyperlink"
                    :href="data[rowAbsoluteIndex][col.id].meta.hyperlink"
                    target="_blank"
                    rel="noopener noreferrer"
                  >
                    {{ data[rowAbsoluteIndex][col.id]?.value || '' }}
                  </a>
                  <span v-else>
                    {{ data[rowAbsoluteIndex]?.[col.id]?.value || '' }}
                  </span>
                </template>

                <input
                  v-else
                  ref="editorInput"
                  type="text"
                  v-model="editingValue"
                  @blur="stopEditing"
                  @keyup.enter="stopEditing"
                  @keyup.esc="cancelEditing"
                  :style="getInputStyle(data[rowAbsoluteIndex]?.[col.id])"
                  class="cell-input"
                >
              </div>
            </td>
          </tr>

          <!-- нижний спейсер -->
          <tr v-if="bottomSpacer > 0" class="spacer-row">
            <td :colspan="columns.length" :style="{ height: bottomSpacer + 'px', padding: 0, border: 0 }"></td>
          </tr>
        </tbody>
      </table>
    </div>

    <div v-if="tooltip.visible" class="tooltip" :style="{ left: `${tooltip.x}px`, top: `${tooltip.y}px` }">
      {{ tooltip.text }}
    </div>
  </div>
</template>

<script setup>
import { ref, computed, nextTick, onMounted } from 'vue'

const props = defineProps({
  data: { type: Array, required: true },
  columns: { type: Array, required: true },
  selectedCells: { type: Array, default: () => [] },
  searchResults: { type: Array, default: () => [] },
  rowHeight: { type: Number, default: 30 } // для виртуализации
})

const emit = defineEmits(['update:selected-cells', 'update-cell', 'resize-column'])

/** ---- Редактор ---- */
const editorInput = ref(null)
const editingCell = ref(null)
const editingValue = ref('')

const startEditing = async (rowIndex, colId) => {
  editingCell.value = { rowIndex, colId }
  editingValue.value = props.data[rowIndex]?.[colId]?.value || ''
  await nextTick()
  if (editorInput.value && editorInput.value[0]) editorInput.value[0].focus()
}
const stopEditing = () => {
  if (editingCell.value) {
    emit('update-cell', { rowIndex: editingCell.value.rowIndex, colId: editingCell.value.colId, value: editingValue.value })
    editingCell.value = null
  }
}
const cancelEditing = () => { editingCell.value = null }
const isEditing = (rowIndex, colId) => editingCell.value?.rowIndex === rowIndex && editingCell.value?.colId === colId

/** ---- Выделение ---- */
const isSelecting = ref(false)
const selectionStart = ref(null)

const startSelection = (e) => {
  const cell = getCellFromEvent(e); if (!cell) return
  if (e.ctrlKey || e.metaKey) {
    const i = props.selectedCells.findIndex(c => c.rowIndex === cell.rowIndex && c.colId === cell.colId)
    if (i >= 0) {
      const arr = [...props.selectedCells]; arr.splice(i, 1); emit('update:selected-cells', arr)
    } else emit('update:selected-cells', [...props.selectedCells, cell])
    return
  }
  if (e.shiftKey && props.selectedCells.length > 0) {
    const last = props.selectedCells[props.selectedCells.length - 1]
    updateSelectionArea(last, cell); return
  }
  isSelecting.value = true
  selectionStart.value = cell
  emit('update:selected-cells', [cell])
}

// Новые вычисляемые свойства и методы для поиска
const isFirstSearchMatch = (rowIndex, colId) => {
  const firstMatch = props.searchResults[0]
  return firstMatch && firstMatch.rowIndex === rowIndex && firstMatch.colId === colId
}

const getSearchMatchIndex = (rowIndex, colId) => {
  return props.searchResults.findIndex(
    result => result.rowIndex === rowIndex && result.colId === colId
  )
}

const handleSelection = (e) => {
  if (!isSelecting.value || !selectionStart.value) return
  const cell = getCellFromEvent(e); if (cell) updateSelectionArea(selectionStart.value, cell)
}
const endSelection = () => { isSelecting.value = false }
const handleCellClick = (_rowIndex, _colId, event) => { event.stopPropagation() }

const getCellFromEvent = (e) => {
  const cellEl = e.target.closest('.data-cell'); if (!cellEl) return null
  const rowEl = cellEl.parentElement
  const rowIndex = parseInt(rowEl.dataset.rowIndex)
  const colId = cellEl.dataset.colId
  return { rowIndex, colId }
}

const updateSelectionArea = (startCell, endCell) => {
  const startRow = Math.min(startCell.rowIndex, endCell.rowIndex)
  const endRow = Math.max(startCell.rowIndex, endCell.rowIndex)
  const startCol = Math.min(parseInt(startCell.colId), parseInt(endCell.colId))
  const endCol = Math.max(parseInt(startCell.colId), parseInt(endCell.colId))
  const selected = []
  for (let r = startRow; r <= endRow; r++) {
    for (let c = startCol; c <= endCol; c++) selected.push({ rowIndex: r, colId: c.toString() })
  }
  emit('update:selected-cells', selected)
}

const isSelected = (rowIndex, colId) => props.selectedCells.some(c => c.rowIndex === rowIndex && c.colId === colId.toString())
const isInSelectionArea = (rowIndex, colId) => {
  if (!isSelecting.value || !selectionStart.value) return false
  const startRow = Math.min(selectionStart.value.rowIndex, rowIndex)
  const endRow = Math.max(selectionStart.value.rowIndex, rowIndex)
  const startCol = Math.min(parseInt(selectionStart.value.colId), parseInt(colId))
  const endCol = Math.max(parseInt(selectionStart.value.colId), parseInt(colId))
  return rowIndex >= startRow && rowIndex <= endRow && parseInt(colId) >= startCol && parseInt(colId) <= endCol
}
const isSearchMatch = (rowIndex, colId) => props.searchResults.some(c => c.rowIndex === rowIndex && c.colId === colId.toString())

/** ---- Подсказки + обрезка ---- */
const tooltip = ref({ visible: false, text: '', x: 0, y: 0 })
const shouldTruncate = (t) => !!t && String(t).length > 20
const showTooltipIfNeeded = (event, text) => {
  if (!shouldTruncate(text)) return
  const cell = event.target.closest('.data-cell'); if (!cell) return
  const r = cell.getBoundingClientRect()
  tooltip.value = { visible: true, text, x: r.left, y: r.bottom }
}
const hideTooltip = () => { tooltip.value.visible = false }

/** ---- Стили ячеек ---- */
const getCellStyle = (cellData) => {
  const st = cellData?.style || {}
  return {
    fontWeight: st.fontWeight || 'normal',
    textAlign: st.textAlign || 'left',
    color: st.color || '#000000',
    backgroundColor: st.backgroundColor || 'transparent',
    fontSize: st.fontSize || '14px',
    fontStyle: st.fontStyle || 'normal',
    textDecoration: st.textDecoration || 'none'
  }
}
const getInputStyle = (cellData) => ({ ...getCellStyle(cellData), width: '100%', height: '100%' })

/** ---- Ресайз колонок ---- */
const startResize = (column, e) => {
  e.preventDefault()
  const startX = e.clientX
  const startW = column.width
  const move = (me) => {
    const w = startW + (me.clientX - startX)
    if (Math.abs(w - column.width) > 2) emit('resize-column', { colId: column.id, width: Math.max(30, w) })
  }
  const up = () => {
    document.removeEventListener('mousemove', move)
    document.removeEventListener('mouseup', up)
    document.body.style.cursor = ''
  }
  document.body.style.cursor = 'col-resize'
  document.addEventListener('mousemove', move)
  document.addEventListener('mouseup', up, { once: true })
}

/** -------------------------
 *     Виртуализация строк
 *  ------------------------*/
const scrollEl = ref(null)
const viewportHeight = ref(0)
const scrollTop = ref(0)
const rowHeight = computed(() => props.rowHeight)
const buffer = 10 // дополнительный запас строк сверху/снизу

const totalRows = computed(() => props.data.length)
const rowsPerViewport = computed(() => Math.max(1, Math.ceil(viewportHeight.value / rowHeight.value)))
const startIndex = computed(() => Math.max(0, Math.floor(scrollTop.value / rowHeight.value) - buffer))
const endIndex = computed(() => Math.min(totalRows.value - 1, startIndex.value + rowsPerViewport.value + buffer * 2))
const visibleIndexes = computed(() => {
  const arr = []
  for (let i = startIndex.value; i <= endIndex.value; i++) arr.push(i)
  return arr
})

const topSpacer = computed(() => startIndex.value * rowHeight.value)
const bottomSpacer = computed(() => (totalRows.value - endIndex.value - 1) * rowHeight.value)

const onScroll = () => { scrollTop.value = scrollEl.value?.scrollTop || 0 }
onMounted(() => {
  viewportHeight.value = scrollEl.value?.clientHeight || 0
})

</script>

<style scoped>
.excel-container { position: relative; width: 100%; height: 100%; overflow: hidden; }
.excel-grid { width: 100%; height: 100%; overflow: auto; font-family: Arial, sans-serif; }
.excel-table { border-collapse: collapse; width: max-content; min-width: 100%; table-layout: fixed; }
.header-cell { position: relative; padding: 8px; border: 1px solid #ddd; background-color: #f0f0f0; font-weight: bold; text-align: center; user-select: none; }
.header-content { display: flex; justify-content: center; align-items: center; position: relative; height: 100%; }
.resize-handle { position: absolute; right: 0; top: 0; width: 5px; height: 100%; cursor: col-resize; background-color: rgba(0,0,0,.1); }
.resize-handle:hover { background-color: #42b983; }
.data-row { height: 30px; }
.data-cell {
  position: relative;
  padding: 4px 8px;
  border: 1px solid #ddd;
  vertical-align: middle;
  cursor: cell;
  overflow: hidden;
  background-color: inherit; /* Добавлено */
}

.data-cell.selected {
  background-color: #e6f3ff !important; /* Добавлено !important */
  outline: 2px solid #42b983;
  outline-offset: -2px;
}

.data-cell.selection-area {
  background-color: #f0f7ff !important; /* Добавлено !important */
}

.data-cell.search-match {
  background-color: #fffacd !important; /* Добавлено !important */
  outline: 2px solid #ffcc00;
  outline-offset: -2px;
}
.cell-content { width: 100%; height: 100%; overflow: hidden; white-space: nowrap; }
.cell-content.text-truncated { text-overflow: ellipsis; }
.cell-input { width: 100%; height: 100%; border: none; outline: none; padding: 0; margin: 0; font: inherit; color: inherit; background: transparent; }
.tooltip { position: fixed; background-color: rgba(0,0,0,.8); color: #fff; padding: 4px 8px; border-radius: 4px; font-size: 12px; z-index: 1000; pointer-events: none; max-width: 300px; word-wrap: break-word; white-space: normal; }
.spacer-row td { border: 0 !important; }
a { color: inherit; text-decoration: underline; }

/* Добавляем стили для индикатора поиска */
.search-indicator {
  position: absolute;
  top: 2px;
  right: 2px;
  background: #ffcc00;
  color: #000;
  font-size: 10px;
  font-weight: bold;
  width: 16px;
  height: 16px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  justify-content: center;
  z-index: 2;
}

/* Улучшаем стиль для ячеек с совпадениями поиска */
.data-cell.search-match {
  background-color: #fffacd;
  position: relative;
  box-shadow: 0 0 0 2px #ffcc00 inset;
}

/* Увеличиваем z-index для выделенных ячеек с поиском */
.data-cell.search-match.selected {
  z-index: 3;
  box-shadow: 0 0 0 2px #42b983, 0 0 0 4px #ffcc00;
}
</style>
