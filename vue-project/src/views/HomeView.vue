<template>
  <div class="home">
    <SheetTabs
      :sheets="sheets"
      :active-sheet="activeSheet"
      @select-sheet="selectSheet"
      @add-sheet="addSheet"
      @rename-sheet="renameSheet"
      @delete-sheet="deleteSheet"
    />
    <Toolbar 
      :current-styles="getCurrentCellStyles()"
      @apply-styles="applyStylesToSelectedCells"
      @import-data="handleImport"
      @export-excel="exportExcel"
    />
    <SearchPanel 
      :columns="columns"
      @search="handleSearch"
    />
    <ExcelGrid
      :data="currentData"
      :columns="columns"
      :selected-cells="selectedCells"
      :search-results="searchResults"
      @update:selected-cells="selectedCells = $event"
      @update-cell="updateCell"
      @resize-column="resizeColumn"
    />
  </div>
</template>

<script setup>
import { ref, computed } from 'vue'
import * as XLSX from 'xlsx'
import ExcelJS from 'exceljs';
import SheetTabs from '@/components/SheetTabs.vue'
import Toolbar from '@/components/Toolbar.vue'
import ExcelGrid from '@/components/ExcelGrid.vue'
import SearchPanel from '@/components/SearchPanel.vue'

const ROWS_COUNT = 100
const COLS_COUNT = 26
const DEFAULT_COL_WIDTH = 120

const nextSheetId = ref(2)
const sheets = ref([])
const activeSheet = ref(null)
const columns = ref([])
const selectedCells = ref([])
const searchResults = ref([])

columns.value = Array.from({ length: COLS_COUNT }, (_, i) => ({
  id: (i + 1).toString(),
  name: String.fromCharCode(65 + i),
  width: DEFAULT_COL_WIDTH
}))

function createNewSheet(name) {
  const sheetId = nextSheetId.value++
  return {
    id: sheetId,
    name: name || `Лист${sheetId}`,
    data: Array(ROWS_COUNT).fill().map(() => {
      const row = {}
      columns.value.forEach(col => {
        row[col.id] = { value: '', style: {} }
      })
      return row
    })
  }
}

sheets.value = [createNewSheet('Лист1')]
activeSheet.value = sheets.value[0].id

const currentData = computed(() => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet ? sheet.data : []
})

const selectSheet = (id) => {
  // Изменяем эту строку:
  activeSheet.value = id; // Просто сохраняем ID вместо всего объекта
};


const addSheet = () => {
  const newSheet = createNewSheet()
  sheets.value.push(newSheet)
  selectSheet(newSheet.id)
}

const renameSheet = ({ id, newName }) => {
  const sheet = sheets.value.find(s => s.id === id)
  if (sheet && newName.trim()) {
    sheet.name = newName.trim()
  }
}

const deleteSheet = (sheetId) => {
  if (sheets.value.length <= 1) return
  
  if (confirm('Вы уверены, что хотите удалить этот лист?')) {
    const index = sheets.value.findIndex(s => s.id === sheetId)
    if (index !== -1) {
      sheets.value.splice(index, 1)
      if (activeSheet.value === sheetId) {
        selectSheet(sheets.value[0].id)
      }
    }
  }
}

const getCurrentCellStyles = () => {
  if (selectedCells.value.length === 0) return {
    fontWeight: 'normal',
    textAlign: 'left',
    color: '#000000',
    backgroundColor: '#ffffff',
    fontSize: '14px'
  }
  
  const firstCell = selectedCells.value[0]
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  return sheet?.data[firstCell.rowIndex]?.[firstCell.colId]?.style || {
    fontWeight: 'normal',
    textAlign: 'left',
    color: '#000000',
    backgroundColor: '#ffffff',
    fontSize: '14px'
  }
}

const updateCell = ({ rowIndex, colId, value }) => {
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (sheet && sheet.data[rowIndex]) {
    sheet.data[rowIndex][colId].value = value
  }
}

const applyStylesToSelectedCells = (styles) => {
  if (selectedCells.value.length === 0) return
  
  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (!sheet) return
  
  selectedCells.value.forEach(({ rowIndex, colId }) => {
    if (sheet.data[rowIndex]?.[colId]) {
      sheet.data[rowIndex][colId].style = {
        ...sheet.data[rowIndex][colId].style,
        ...styles
      }
    }
  })
}

const resizeColumn = ({ colId, width }) => {
  const column = columns.value.find(c => c.id === colId)
  if (column) {
    column.width = Math.max(50, width)
  }
}

const handleSearch = ({ queries, matchCase, columns: searchColumns }) => {
  if (!queries || queries.length === 0) {
    searchResults.value = []
    return
  }

  const sheet = sheets.value.find(s => s.id === activeSheet.value)
  if (!sheet) return

  const results = []

  sheet.data.forEach((row, rowIndex) => {
    Object.entries(row).forEach(([colId, cell]) => {
      if (searchColumns !== 'all' && !searchColumns.includes(colId)) return
      
      const cellText = matchCase 
        ? cell.value 
        : cell.value?.toLowerCase()
      
      const anyMatch = queries.some(query => {
        const searchWord = matchCase ? query : query.toLowerCase()
        return cellText?.includes(searchWord)
      })
      
      if (anyMatch) {
        results.push({ rowIndex, colId })
      }
    })
  })

  searchResults.value = results
  if (results.length > 0) {
    selectedCells.value = results
  }
}

const handleImport = async (event) => {
  const file = event.target.files[0];
  if (!file) return;

  try {
    const buffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    // Очищаем существующие листы
    sheets.value = [];

    workbook.eachSheet((worksheet, sheetId) => {
      // Создаем массив для данных листа
      const sheetData = [];
      const sheetColumns = [];
      
      // Определяем количество колонок
      const columnCount = worksheet.columnCount || COLS_COUNT;
      
      // Создаем колонки
      for (let i = 1; i <= columnCount; i++) {
        sheetColumns.push({
          id: i.toString(),
          name: String.fromCharCode(64 + i), // A, B, C...
          width: DEFAULT_COL_WIDTH
        });
      }

      // Обрабатываем каждую строку
      worksheet.eachRow({ includeEmpty: true }, (row, rowIndex) => {
        const rowData = {};
        
        // Обрабатываем каждую ячейку в строке
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          const colId = colNumber.toString();
          
          // Получаем значение ячейки
          let cellValue = '';
          if (cell.value !== undefined && cell.value !== null) {
            // Для формул берем результат вычисления
            if (cell.formula) {
              cellValue = cell.result !== undefined ? String(cell.result) : '';
            } else {
              cellValue = String(cell.value);
            }
          }
          
          // Создаем объект ячейки со стилями
          rowData[colId] = {
            value: cellValue,
            style: {
              fontWeight: cell.font?.bold ? 'bold' : 'normal',
              textAlign: cell.alignment?.horizontal || 'left',
              color: cell.font?.color?.argb ? `#${cell.font.color.argb.slice(2)}` : '#000000',
              backgroundColor: cell.fill?.fgColor?.argb ? `#${cell.fill.fgColor.argb.slice(2)}` : '#FFFFFF',
              fontSize: cell.font?.size ? `${cell.font.size}px` : '14px'
            }
          };
        });
        
        // Заполняем пустые ячейки в строке
        for (let i = 1; i <= columnCount; i++) {
          if (!rowData[i.toString()]) {
            rowData[i.toString()] = { value: '', style: {} };
          }
        }
        
        sheetData.push(rowData);
      });

      // Создаем новый лист
      const newSheet = {
        id: sheetId,
        name: worksheet.name || `Лист${sheetId}`,
        data: sheetData,
        columns: sheetColumns
      };

      sheets.value.push(newSheet);
    });

    // Активируем первый лист
    if (sheets.value.length > 0) {
      activeSheet.value = sheets.value[0].id;
      columns.value = [...sheets.value[0].columns];
    }

  } catch (error) {
    console.error('Ошибка импорта:', error);
    alert('Не удалось загрузить файл');
  } finally {
    event.target.value = ''; // Сбрасываем значение input для повторной загрузки
  }
};

const exportExcel = async () => {
  const workbook = new ExcelJS.Workbook();

  sheets.value.forEach(sheet => {
    const ws = workbook.addWorksheet(sheet.name);

    // Устанавливаем ширину колонок
    ws.columns = columns.value.map(col => ({
      width: Math.min(50, Math.max(5, col.width / 7))
    }));

    sheet.data.forEach(row => {
      const excelRow = ws.addRow(columns.value.map(col => row[col.id]?.value || ''));

      columns.value.forEach((col, colIndex) => {
        const cellData = row[col.id];
        if (cellData?.style) {
          const cell = excelRow.getCell(colIndex + 1);

          // Шрифт
          cell.font = {
            bold: cellData.style.fontWeight === 'bold',
            color: cellData.style.color
              ? { argb: cellData.style.color.replace('#', '') }
              : undefined,
            size: cellData.style.fontSize
              ? parseInt(cellData.style.fontSize)
              : undefined
          };

          // Выравнивание
          if (cellData.style.textAlign) {
            cell.alignment = {
              horizontal: cellData.style.textAlign,
              vertical: 'middle'
            };
          }

          // Заливка
          if (cellData.style.backgroundColor) {
            cell.fill = {
              type: 'pattern',
              pattern: 'solid',
              fgColor: { argb: cellData.style.backgroundColor.replace('#', '') }
            };
          }

          // Границы
          cell.border = {
            top: { style: 'thin', color: { argb: '000000' } },
            bottom: { style: 'thin', color: { argb: '000000' } },
            left: { style: 'thin', color: { argb: '000000' } },
            right: { style: 'thin', color: { argb: '000000' } }
          };
        }
      });
    });
  });

  // Генерация и скачивание файла
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  });
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = `${sheets.value[0]?.name || 'таблица'}.xlsx`;
  link.click();
};
</script>

<style scoped>
.home {
  display: flex;
  flex-direction: column;
  height: 100vh;
  padding: 0;
}
</style>