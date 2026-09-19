import * as React from 'react';
import {
  AllCommunityModule,
  ColDef,
  themeQuartz,
  ValueFormatterParams
} from 'ag-grid-community';
import { AG_GRID_LOCALE_UA } from '@ag-grid-community/locale';
import { AgGridProvider, AgGridReact } from 'ag-grid-react';
import { IServiceRequest } from '../models/ServiceDeskModels';
import styles from './ServiceRequestsGrid.module.scss';

// значення одного рядка таблиці після перетворення даних sharepoint
interface IRequestGridRow {
  id: number;
  title: string;
  category: string;
  subcategory: string;
  status: string;
  priority: string;
  requester: string;
  assignee: string;
  dueDate?: number;
  estimatedHours?: number;
}

// вхідні дані таблиці заявок
interface IServiceRequestsGridProps {
  requests: IServiceRequest[];
}

const modules = [AllCommunityModule];
const localeText = {
  ...AG_GRID_LOCALE_UA,
  noRowsToShow: 'Заявок поки немає'
};
const dateFormatter = new Intl.DateTimeFormat('uk-UA', {
  day: '2-digit',
  month: '2-digit',
  year: 'numeric',
  hour: '2-digit',
  minute: '2-digit'
});
const hoursFormatter = new Intl.NumberFormat('uk-UA', {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1
});

// перетворює заявку sharepoint на рядок із назвами полів із довідників
function toGridRow(request: IServiceRequest): IRequestGridRow {
  const parsedDueDate = Date.parse(request.DueDate);

  return {
    id: request.Id,
    title: request.Title,
    category: request.Category?.Title ?? '',
    subcategory: request.Subcategory?.Title ?? '',
    status: request.Status,
    priority: request.Priority,
    requester: request.Requester?.Title ?? '',
    assignee: request.Assignee?.Title ?? '',
    dueDate: Number.isNaN(parsedDueDate) ? undefined : parsedDueDate,
    estimatedHours: request.EstimatedHours ?? undefined
  };
}

// показує дату українською мовою без зміни числового значення для сортування
function formatDueDate(params: ValueFormatterParams<IRequestGridRow, number>): string {
  return typeof params.value === 'number' ? dateFormatter.format(params.value) : '';
}

// показує кількість годин з українським десятковим роздільником
function formatEstimatedHours(params: ValueFormatterParams<IRequestGridRow, number>): string {
  return typeof params.value === 'number' ? hoursFormatter.format(params.value) : '';
}

const columnDefs: ColDef<IRequestGridRow>[] = [
  { field: 'id', headerName: 'ID', width: 90 },
  { field: 'title', headerName: 'Назва заявки', minWidth: 220, flex: 2 },
  { field: 'category', headerName: 'Категорія', minWidth: 160 },
  { field: 'subcategory', headerName: 'Підкатегорія', minWidth: 170 },
  { field: 'status', headerName: 'Статус', minWidth: 130 },
  { field: 'priority', headerName: 'Пріоритет', minWidth: 130 },
  { field: 'requester', headerName: 'Заявник', minWidth: 170 },
  { field: 'assignee', headerName: 'Виконавець', minWidth: 170 },
  {
    field: 'dueDate',
    headerName: 'Кінцевий термін',
    minWidth: 180,
    valueFormatter: formatDueDate
  },
  {
    field: 'estimatedHours',
    headerName: 'Оцінка годин',
    minWidth: 140,
    valueFormatter: formatEstimatedHours
  }
];

const defaultColDef: ColDef<IRequestGridRow> = {
  sortable: true,
  resizable: true,
  filter: false
};

// відображає заявки в ag grid із сортуванням та українською локалізацією
export default function ServiceRequestsGrid(props: IServiceRequestsGridProps): React.ReactElement {
  const rows = props.requests.map(toGridRow);

  return (
    <AgGridProvider modules={modules}>
      <div className={styles.grid}>
        <AgGridReact<IRequestGridRow>
          rowData={rows}
          columnDefs={columnDefs}
          defaultColDef={defaultColDef}
          localeText={localeText}
          theme={themeQuartz}
          accentedSort={true}
        />
      </div>
    </AgGridProvider>
  );
}
