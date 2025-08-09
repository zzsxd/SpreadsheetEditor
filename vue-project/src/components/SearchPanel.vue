<template>
  <div class="search-panel">
    <div class="search-input-group">
      <input
        type="text"
        v-model="searchQuery"
        placeholder="Поиск по таблице"
        @keyup.enter="executeSearch"
      >
      <button class="search-btn" @click="executeSearch">
        <i class="material-icons">search</i>
      </button>
      <button class="clear-btn" @click="clearSearch" v-if="searchQuery">
        <i class="material-icons">clear</i>
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
import { ref, watch } from 'vue'

const props = defineProps({
  columns: {
    type: Array,
    required: true
  }
})

const emit = defineEmits(['search'])

const searchQuery = ref('')
const matchCase = ref(false)
const searchInSelectedColumns = ref(false)
const selectedColumns = ref([])

watch(() => props.columns, (newColumns) => {
  selectedColumns.value = newColumns.map(col => col.id)
}, { immediate: true })

function executeSearch() {
  if (!searchQuery.value.trim()) return
  
  emit('search', {
    query: searchQuery.value,
    matchCase: matchCase.value,
    columns: searchInSelectedColumns.value ? selectedColumns.value : 'all'
  })
}

function clearSearch() {
  searchQuery.value = ''
  matchCase.value = false
  searchInSelectedColumns.value = false
  emit('search', { query: '', columns: 'all' })
}
</script>

<style scoped>
.search-panel {
  padding: 10px;
  background: #f5f5f5;
  border-bottom: 1px solid #ddd;
}

.search-input-group {
  display: flex;
  margin-bottom: 10px;
}

.search-input-group input {
  flex: 1;
  padding: 8px 12px;
  border: 1px solid #ddd;
  border-radius: 4px 0 0 4px;
  outline: none;
}

.search-input-group input:focus {
  border-color: #42b983;
}

.search-btn, .clear-btn {
  padding: 0 12px;
  border: none;
  background: #42b983;
  color: white;
  cursor: pointer;
  display: flex;
  align-items: center;
}

.search-btn {
  border-radius: 0 4px 4px 0;
}

.clear-btn {
  background: #e74c3c;
  margin-left: 1px;
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