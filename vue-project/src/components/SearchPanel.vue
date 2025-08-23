<template>
  <div class="search-panel">
    <div v-if="searchResults.length > 0" class="search-navigation">
      <span>Найдено: {{ searchResults.length }}</span>
      <button @click="prevResult">◀</button>
      <span class="current-position">
        {{ currentIndex + 1 }} / {{ searchResults.length }}
      </span>
      <button @click="nextResult">▶</button>
    </div>

    <div class="search-inputs">
      <div 
        v-for="(input, index) in searchInputs" 
        :key="index"
        class="search-input-group"
      >
        <input
          type="text"
          v-model="searchInputs[index]"
          :placeholder="`Поиск ${index + 1}`"
          @keyup.enter="executeSearch"
        >
        <button 
          class="remove-search-btn"
          @click="removeSearchInput(index)"
          v-if="index > 0"
          title="Удалить строку поиска"
        >
          ×
        </button>
      </div>
    </div>

    <div class="search-controls">
      <button class="search-btn" @click="executeSearch" title="Начать поиск">
        Поиск
      </button>
      <button class="add-search-btn" @click="addSearchInput" title="Добавить строку поиска">
        Добавить поиск
      </button>
      <button class="clear-btn" @click="clearSearch" v-if="hasSearchText" title="Очистить поиск">
        Очистить
      </button>
    </div>

    <div class="search-options">
      
      <label class="checkbox-label">
        <input 
          type="checkbox" 
          v-model="searchInSelectedColumns"
          class="checkbox"
        >
        Искать в выбранных столбцах
      </label>
      
      <div v-if="searchInSelectedColumns" class="column-select">
        <label 
          v-for="col in columns" 
          :key="col.id"
          class="column-checkbox"
        >
          <input 
            type="checkbox" 
            v-model="selectedColumns" 
            :value="col.id"
          >
          {{ col.name }}
        </label>
      </div>
    </div>
  </div>
</template>

<script setup>
import { ref, computed, watch } from 'vue'
const props = defineProps({
  columns: {
    type: Array,
    required: true
  },
  searchResults: {
    type: Array,
    default: () => []   // ✅ чтобы не было undefined
  }
})

const emit = defineEmits(['search', 'navigate'])

const searchInputs = ref([''])
const matchCase = ref(false)
const searchInSelectedColumns = ref(false)
const selectedColumns = ref([])

const hasSearchText = computed(() => {
  return searchInputs.value.some(input => input.trim() !== '')
})


function addSearchInput() {
  searchInputs.value.push('')
}

function removeSearchInput(index) {
  searchInputs.value.splice(index, 1)
}

function executeSearch() {
  const queries = searchInputs.value
    .map(input => input.trim())
    .filter(input => input !== '')
  
  emit('search', {
    queries,
    matchCase: matchCase.value,
    columns: searchInSelectedColumns.value ? selectedColumns.value : 'all'
  })
}

function clearSearch() {
  searchInputs.value = ['']
  matchCase.value = false
  searchInSelectedColumns.value = false
  emit('search', { queries: [], columns: 'all' })
}

const currentIndex = ref(0)

function nextResult() {
  if (!props.searchResults.length) return
  currentIndex.value = (currentIndex.value + 1) % props.searchResults.length
  emit('navigate', props.searchResults[currentIndex.value])
}

function prevResult() {
  if (!props.searchResults.length) return
  currentIndex.value = (currentIndex.value - 1 + props.searchResults.length) % props.searchResults.length
  emit('navigate', props.searchResults[currentIndex.value])
}

</script>

<style scoped>
.search-panel {
  padding: 10px;
  background: #f5f5f5;
  border-bottom: 1px solid #ddd;
}

.search-inputs {
  margin-bottom: 10px;
}

.search-input-group {
  display: flex;
  margin-bottom: 5px;
  align-items: center;
}

.search-input-group input {
  flex: 1;
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 4px;
  outline: none;
}

.search-input-group input:focus {
  border-color: #42b983;
}

.remove-search-btn {
  margin-left: 5px;
  padding: 0 10px;
  border: none;
  background: #e74c3c;
  color: white;
  border-radius: 4px;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  height: 32px;
  width: 32px;
}

.remove-search-btn:hover {
  background: #c0392b;
}

.search-controls {
  display: flex;
  gap: 10px;
  margin-bottom: 10px;
}

.search-btn, .add-search-btn, .clear-btn {
  padding: 5px 10px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
  display: flex;
  align-items: center;
  gap: 5px;
}

.search-btn {
  background: #42b983;
  color: white;
}

.search-btn:hover {
  background: #369f6b;
}

.add-search-btn {
  background: #3498db;
  color: white;
}

.add-search-btn:hover {
  background: #2980b9;
}

.clear-btn {
  background: #e74c3c;
  color: white;
}

.clear-btn:hover {
  background: #c0392b;
}

.search-options {
  display: flex;
  flex-wrap: wrap;
  gap: 15px;
  align-items: center;
}

.checkbox-label {
  display: flex;
  align-items: center;
  gap: 5px;
  cursor: pointer;
  user-select: none;
}

.checkbox {
  margin: 0;
}

.column-select {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  margin-top: 5px;
}

.column-checkbox {
  display: flex;
  align-items: center;
  gap: 5px;
  padding: 3px 8px;
  background: #eee;
  border-radius: 4px;
  cursor: pointer;
}
/* Добавляем стили для отображения текущей позиции */
.current-position {
  margin: 0 10px;
  font-weight: bold;
  min-width: 50px;
  text-align: center;
}
</style>