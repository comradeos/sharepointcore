import * as XLSX from 'xlsx';
import { GridApi } from 'ag-grid-community';
import { IExcelExportColumn, IGridRow, IRequestGridRow, IRequestGridRowValues } from './ServiceRequestsGrid.types';

const exportableFields: Array<keyof IRequestGridRowValues> = [
  'id',
  'title',
  'description',
  'category',
  'subcategory',
  'status',
  'priority',
  'requester',
  'assignee',
  'dueDate',
  'estimatedHours'
];
const dateFormatter = new Intl.DateTimeFormat('uk-UA', {
  day: '2-digit',
  month: '2-digit',
  year: 'numeric',
  hour: '2-digit',
  minute: '2-digit'
});

// повертає значення комірки у форматі для excel
function getExcelCellValue(
  row: IRequestGridRow,
  field: keyof IRequestGridRowValues
): string | number {
  switch (field) {
    case 'dueDate':
      return row.dueDate ? dateFormatter.format(row.dueDate) : '';
    case 'estimatedHours':
      return row.estimatedHours ?? '';
    default:
      const value = row[field];
      return typeof value === 'number' || typeof value === 'string' ? value : '';
  }
}

// перевіряє чи колонка містить значення заявки для експорту
function isExportableField(field: string | undefined): field is keyof IRequestGridRowValues {
  return field !== undefined && exportableFields.indexOf(field as keyof IRequestGridRowValues) !== -1;
}

// створює назву файла з поточною датою
function getExcelFileName(): string {
  const today = new Date();
  const formatDatePart = (value: number): string => value < 10 ? `0${value}` : String(value);
  const formattedDate = [
    today.getFullYear(),
    formatDatePart(today.getMonth() + 1),
    formatDatePart(today.getDate())
  ].join('-');

  return `Сервісні_заявки_${formattedDate}.xlsx`;
}

// експортує відфільтровані та відсортовані заявки у файл excel
export function exportGridToExcel(api: GridApi<IGridRow>): void {
  const exportColumns: IExcelExportColumn[] = api.getAllDisplayedColumns()
    .map(column => {
      const columnDefinition = column.getColDef();
      const field = columnDefinition.field;

      return isExportableField(field)
        ? {
          field,
          headerName: columnDefinition.headerName ?? field,
          width: column.getActualWidth()
        }
        : undefined;
    })
    .filter((column): column is IExcelExportColumn => column !== undefined);

  if (exportColumns.length === 0) {
    return;
  }

  const excelRows: Array<Array<string | number>> = [
    exportColumns.map(column => column.headerName)
  ];

  api.forEachNodeAfterFilterAndSort(node => {
    const row = node.data;

    if (row && !row.isGroup) {
      excelRows.push(exportColumns.map(column => getExcelCellValue(row, column.field)));
    }
  });

  const worksheet = XLSX.utils.aoa_to_sheet(excelRows);
  worksheet['!cols'] = exportColumns.map(column => ({
    wch: Math.max(12, Math.min(60, Math.round(column.width / 7)))
  }));

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Заявки');
  XLSX.writeFile(workbook, getExcelFileName(), { compression: true });
}
