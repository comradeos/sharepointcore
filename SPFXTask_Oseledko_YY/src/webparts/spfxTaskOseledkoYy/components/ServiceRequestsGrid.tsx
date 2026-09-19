import * as React from 'react';
import {
  AllCommunityModule,
  ColDef,
  themeQuartz,
  ValueFormatterParams
} from 'ag-grid-community';
import { AG_GRID_LOCALE_UA } from '@ag-grid-community/locale';
import {
  AgGridProvider,
  AgGridReact,
  CustomCellRendererProps
} from 'ag-grid-react';
import { IconButton, TooltipHost } from '@fluentui/react';
import { IServiceRequest } from '../models/ServiceDeskModels';
import styles from './ServiceRequestsGrid.module.scss';

// значення одного рядка таблиці після перетворення даних sharepoint
interface IRequestGridRow {
  request: IServiceRequest;
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
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
}

// дії таблиці доступні клітинці через контекст ag grid
interface IRequestGridContext {
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
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
    request,
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

// показує кнопки дій для одного рядка таблиці
function RequestActionsRenderer(
  props: CustomCellRendererProps<IRequestGridRow, undefined, IRequestGridContext>
): React.ReactElement | undefined {
  if (!props.data) {
    return undefined;
  }

  // передає вибрану заявку обробнику сторінки
  const handleView = (): void => {
    props.context.onView(props.data?.request as IServiceRequest);
  };

  // передає вибрану заявку до форми редагування
  const handleEdit = (): void => {
    props.context.onEdit(props.data?.request as IServiceRequest);
  };

  return (
    <div className={styles.actions}>
      <TooltipHost
        content="Переглянути"
        styles={{ root: { display: 'flex', alignItems: 'center', height: '100%' } }}
      >
        <IconButton
          className={styles.actionButton}
          iconProps={{ iconName: 'RedEye' }}
          ariaLabel="Переглянути"
          onClick={handleView}
        />
      </TooltipHost>
      <TooltipHost
        content="Редагувати"
        styles={{ root: { display: 'flex', alignItems: 'center', height: '100%' } }}
      >
        <IconButton
          className={styles.actionButton}
          iconProps={{ iconName: 'Edit' }}
          ariaLabel="Редагувати"
          onClick={handleEdit}
        />
      </TooltipHost>
    </div>
  );
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
  {
    field: 'id',
    headerName: 'ID',
    width: 50,
    pinned: 'left',
    lockPinned: true,
    suppressMovable: true
  },
  {
    field: 'title',
    headerName: 'Назва заявки',
    width: 150,
    pinned: 'left',
    lockPinned: true,
    suppressMovable: true
  },
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
  },
  {
    headerName: 'Дії',
    width: 104,
    minWidth: 104,
    maxWidth: 104,
    pinned: 'right',
    sortable: false,
    resizable: false,
    suppressMovable: true,
    cellRenderer: RequestActionsRenderer
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
  const gridContext: IRequestGridContext = { onView: props.onView, onEdit: props.onEdit };

  return (
    <AgGridProvider modules={modules}>
      <div className={styles.grid}>
        <AgGridReact<IRequestGridRow>
          rowData={rows}
          columnDefs={columnDefs}
          defaultColDef={defaultColDef}
          localeText={localeText}
          context={gridContext}
          theme={themeQuartz}
          accentedSort={true}
        />
      </div>
    </AgGridProvider>
  );
}
