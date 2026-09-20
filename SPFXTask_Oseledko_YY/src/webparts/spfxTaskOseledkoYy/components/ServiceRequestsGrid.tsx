import * as React from 'react';
import {
  AllCommunityModule,
  ColDef,
  GridApi,
  GridReadyEvent,
  ModelUpdatedEvent,
  themeQuartz,
  ValueFormatterParams
} from 'ag-grid-community';
import { AG_GRID_LOCALE_UA } from '@ag-grid-community/locale';
import {
  AgGridProvider,
  AgGridReact,
  CustomCellRendererProps
} from 'ag-grid-react';
import {
  DefaultButton,
  IconButton,
  SearchBox,
  Text,
  TooltipHost
} from '@fluentui/react';
import { IServiceRequest } from '../models/ServiceDeskModels';
import ServiceRequestChoiceFilter from './ServiceRequestChoiceFilter';
import styles from './ServiceRequestsGrid.module.scss';

// значення одного рядка таблиці після перетворення даних sharepoint
interface IRequestGridRow {
  request: IServiceRequest;
  id: number;
  title: string;
  description: string;
  category: string;
  subcategory: string;
  status: string;
  priority: string;
  requester: string;
  assignee: string;
  dueDate?: Date;
  estimatedHours?: number;
}

// вхідні дані таблиці заявок
interface IServiceRequestsGridProps {
  requests: IServiceRequest[];
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
  onDelete: (request: IServiceRequest) => void;
}

// дії таблиці доступні клітинці через контекст ag grid
interface IRequestGridContext {
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
  onDelete: (request: IServiceRequest) => void;
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
    description: request.Description,
    category: request.Category?.Title ?? '',
    subcategory: request.Subcategory?.Title ?? '',
    status: request.Status,
    priority: request.Priority,
    requester: request.Requester?.Title ?? '',
    assignee: request.Assignee?.Title ?? '',
    dueDate: Number.isNaN(parsedDueDate) ? undefined : new Date(parsedDueDate),
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

  // передає вибрану заявку до підтвердження видалення
  const handleDelete = (): void => {
    props.context.onDelete(props.data?.request as IServiceRequest);
  };

  return (
    <div className={styles.actions}>
      <TooltipHost
        content="Переглянути"
        styles={{ root: { display: 'flex', alignItems: 'center', height: '100%' } }}
      >
        <IconButton
          className={`${styles.actionButton} ${styles.viewButton}`}
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
          className={`${styles.actionButton} ${styles.editButton}`}
          iconProps={{ iconName: 'Edit' }}
          ariaLabel="Редагувати"
          onClick={handleEdit}
        />
      </TooltipHost>
      <TooltipHost
        content="Видалити"
        styles={{ root: { display: 'flex', alignItems: 'center', height: '100%' } }}
      >
        <IconButton
          className={`${styles.actionButton} ${styles.deleteButton}`}
          iconProps={{ iconName: 'Delete' }}
          ariaLabel="Видалити"
          onClick={handleDelete}
        />
      </TooltipHost>
    </div>
  );
}

// показує дату українською мовою без зміни числового значення для сортування
function formatDueDate(params: ValueFormatterParams<IRequestGridRow, Date>): string {
  return params.value instanceof Date ? dateFormatter.format(params.value) : '';
}

// порівнює кінцевий термін з датою фільтра без урахування часу
function compareDueDate(filterDate: Date, cellValue: Date): number {
  if (!(cellValue instanceof Date)) {
    return -1;
  }

  const cellDate = new Date(
    cellValue.getFullYear(),
    cellValue.getMonth(),
    cellValue.getDate()
  );
  return cellDate.getTime() - filterDate.getTime();
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
    suppressMovable: true,
    filter: 'agTextColumnFilter'
  },
  {
    field: 'description',
    headerName: 'Опис',
    minWidth: 240,
    filter: 'agTextColumnFilter'
  },
  {
    field: 'category',
    headerName: 'Категорія',
    minWidth: 160,
    filter: ServiceRequestChoiceFilter,
    filterParams: { allowMultiple: false }
  },
  {
    field: 'subcategory',
    headerName: 'Підкатегорія',
    minWidth: 170,
    filter: ServiceRequestChoiceFilter,
    filterParams: { allowMultiple: false }
  },
  {
    field: 'status',
    headerName: 'Статус',
    minWidth: 130,
    filter: ServiceRequestChoiceFilter,
    filterParams: { allowMultiple: true }
  },
  {
    field: 'priority',
    headerName: 'Пріоритет',
    minWidth: 130,
    filter: ServiceRequestChoiceFilter,
    filterParams: { allowMultiple: false }
  },
  {
    field: 'requester',
    headerName: 'Заявник',
    minWidth: 170,
    filter: 'agTextColumnFilter'
  },
  {
    field: 'assignee',
    headerName: 'Виконавець',
    minWidth: 170,
    filter: 'agTextColumnFilter'
  },
  {
    field: 'dueDate',
    headerName: 'Кінцевий термін',
    minWidth: 180,
    valueFormatter: formatDueDate,
    filter: 'agDateColumnFilter',
    filterParams: {
      browserDatePicker: true,
      comparator: compareDueDate
    }
  },
  {
    field: 'estimatedHours',
    headerName: 'Оцінка годин',
    minWidth: 140,
    valueFormatter: formatEstimatedHours,
    filter: 'agNumberColumnFilter'
  },
  {
    headerName: 'Дії',
    width: 136,
    minWidth: 136,
    maxWidth: 136,
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
  const [searchText, setSearchText] = React.useState('');
  const [visibleCount, setVisibleCount] = React.useState(props.requests.length);
  const gridApi = React.useRef<GridApi<IRequestGridRow>>();
  const rows = props.requests.map(toGridRow);
  const gridContext: IRequestGridContext = {
    onView: props.onView,
    onEdit: props.onEdit,
    onDelete: props.onDelete
  };

  // зберігає введений текст загального пошуку
  const handleSearchChange = (_event?: React.ChangeEvent<HTMLInputElement>, value?: string): void => {
    setSearchText(value || '');
  };

  // зберігає api таблиці після її створення
  const handleGridReady = (event: GridReadyEvent<IRequestGridRow>): void => {
    gridApi.current = event.api;
    setVisibleCount(event.api.getDisplayedRowCount());
  };

  // оновлює кількість рядків після пошуку або фільтрації
  const handleModelUpdated = (event: ModelUpdatedEvent<IRequestGridRow>): void => {
    setVisibleCount(event.api.getDisplayedRowCount());
  };

  // очищає загальний пошук та фільтри колонок
  const handleClearFilters = (): void => {
    setSearchText('');
    gridApi.current?.setFilterModel(null);
  };

  return (
    <AgGridProvider modules={modules}>
      <div className={styles.searchRow}>
        <Text className={styles.searchLabel}>Пошук</Text>
        <div className={styles.searchControls}>
          <SearchBox
            className={styles.searchControl}
            placeholder="Введіть текст"
            ariaLabel="Пошук у заявках"
            value={searchText}
            onChange={handleSearchChange}
          />
          <DefaultButton text="Очистити фільтри" onClick={handleClearFilters} />
        </div>
      </div>
      <div className={styles.resultRow}>
        <Text>Знайдено заявок: {visibleCount}</Text>
      </div>
      <div className={styles.grid}>
        <AgGridReact<IRequestGridRow>
          rowData={rows}
          columnDefs={columnDefs}
          defaultColDef={defaultColDef}
          localeText={localeText}
          context={gridContext}
          theme={themeQuartz}
          accentedSort={true}
          pagination={true}
          paginationPageSize={10}
          paginationPageSizeSelector={[10, 25, 50]}
          quickFilterText={searchText}
          cacheQuickFilter={true}
          onGridReady={handleGridReady}
          onModelUpdated={handleModelUpdated}
        />
      </div>
    </AgGridProvider>
  );
}
