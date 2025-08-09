<template>
  <div class="search-panel">
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
        >
          <i class="material-icons">close</i>
        </button>
      </div>
    </div>

    <div class="search-controls">
      <button class="search-btn" @click="executeSearch">
        <i class="material-icons">search</i> Поиск
      </button>
      <button class="add-search-btn" @click="addSearchInput">
        <i class="material-icons">add</i> Добавить строку
      </button>
      <button class="clear-btn" @click="clearSearch" v-if="hasSearchText">
        <i class="material-icons">clear</i> Очистить
      </button>
    </div>

    <div class="search-options">
      <label class="checkbox-label">
        <input 
          type="checkbox" 
          v-model="matchCase"
          class="checkbox"
        >
        Учитывать регистр
      </label>
      
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
  }
})

const emit = defineEmits(['search'])

const searchInputs = ref([''])
const matchCase = ref(false)
const searchInSelectedColumns = ref(false)
const selectedColumns = ref([])

const hasSearchText = computed(() => {
  return searchInputs.value.some(input => input.trim() !== '')
})

watch(() => props.columns, (newColumns) => {
  selectedColumns.value = newColumns.map(col => col.id)
}, { immediate: true })

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

.add-search-btn {
  background: #3498db;
  color: white;
}

.clear-btn {
  background: #e74c3c;
  color: white;
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
}

.checkbox {
  margin: 0;
}

.column-select {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
}

.column-checkbox {
  display: flex;
  align-items: center;
  gap: 5px;
  padding: 3px 8px;
  background: #eee;
  border-radius: 4px;
}

.material-icons {
  font-size: 18px;
}
</style>