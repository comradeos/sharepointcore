import * as React from 'react';
import {
  AllCommunityModule,
  ColDef,
  ColumnMovedEvent,
  ColumnResizedEvent,
  GetRowIdParams,
  GridApi,
  GridReadyEvent,
  themeQuartz,
  ValueFormatterParams
} from 'ag-grid-community';
import { AG_GRID_LOCALE_UA } from '@ag-grid-community/locale';
import { AgGridProvider, AgGridReact, CustomCellRendererProps } from 'ag-grid-react';
import { DefaultButton, Dropdown, Icon, IconButton, IDropdownOption, SearchBox, Text, TooltipHost } from '@fluentui/react';
import { IServiceRequest } from '../models/ServiceDeskModels';
import ServiceRequestChoiceFilter from './ServiceRequestChoiceFilter';
import styles from './ServiceRequestsGrid.module.scss';

// значення одного рядка таблиці після перетворення даних sharepoint
interface IRequestGridRowValues {
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

// звичайний рядок таблиці з даними заявки
interface IRequestGridRow extends IRequestGridRowValues {
  request: IServiceRequest;
  isGroup?: false;
}

// повноширинний рядок заголовка групи
interface IGroupGridRow extends IRequestGridRowValues {
  isGroup: true;
  groupKey: string;
  groupLabel: string;
  groupCount: number;
  isCollapsed: boolean;
}

type IGridRow = IRequestGridRow | IGroupGridRow;
type TGroupBy = 'none' | 'category' | 'subcategory' | 'status' | 'priority' | 'assignee' | 'requester';

// вхідні дані таблиці заявок
interface IServiceRequestsGridProps {
  requests: IServiceRequest[];
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
  onDelete: (request: IServiceRequest) => void;
  onVisibleRequestIdsChange: (requestIds: number[]) => void;
}

// дії таблиці доступні клітинці через контекст ag grid
interface IRequestGridContext {
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
  onDelete: (request: IServiceRequest) => void;
  onToggleGroup: (groupKey: string) => void;
}

// кеш одного перетвореного рядка таблиці
interface ICachedRequestGridRow {
  request: IServiceRequest;
  row: IRequestGridRow;
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
const columnStateStorageKey = 'spfx-task-oseledko-yy.service-requests-grid.columns.v1';
const groupingStorageKey = 'spfx-task-oseledko-yy.service-requests-grid.grouping.v1';
const groupByOptions: IDropdownOption[] = [
  { key: 'none', text: 'Без групування' },
  { key: 'category', text: 'Категорія' },
  { key: 'subcategory', text: 'Підкатегорія' },
  { key: 'status', text: 'Статус' },
  { key: 'priority', text: 'Пріоритет' },
  { key: 'assignee', text: 'Виконавець' },
  { key: 'requester', text: 'Заявник' }
];

// збережений порядок і ширина колонок таблиці
interface IStoredColumnState {
  colId: string;
  width?: number;
}

// читає збережений порядок і ширину колонок із localstorage
function getStoredColumnState(): IStoredColumnState[] | undefined {
  try {
    const serializedState = window.localStorage.getItem(columnStateStorageKey);

    if (!serializedState) {
      return undefined;
    }

    const parsedState: unknown = JSON.parse(serializedState);
    if (!Array.isArray(parsedState)) {
      return undefined;
    }

    const columnState = parsedState.filter((value): value is IStoredColumnState => {
      if (!value || typeof value !== 'object') {
        return false;
      }

      const state = value as IStoredColumnState;
      return typeof state.colId === 'string'
        && (state.width === undefined || (Number.isFinite(state.width) && state.width > 0));
    });

    return columnState.length > 0 ? columnState : undefined;
  } catch {
    return undefined;
  }
}

// застосовує збережений порядок і ширину колонок
function applyStoredColumnState(api: GridApi<IGridRow>): void {
  const columnState = getStoredColumnState();

  if (columnState) {
    api.applyColumnState({ state: columnState, applyOrder: true });
  }
}

// зберігає поточний порядок і ширину колонок у localstorage
function saveColumnState(api: GridApi<IGridRow>): void {
  const columnState: IStoredColumnState[] = api.getColumnState().map(({ colId, width }) => ({
    colId,
    width: width ?? undefined
  }));

  try {
    window.localStorage.setItem(columnStateStorageKey, JSON.stringify(columnState));
  } catch {
    // браузер може заборонити доступ до localstorage
  }
}

// читає збережене поле групування з localstorage
function getStoredGroupBy(): TGroupBy {
  try {
    const groupBy = window.localStorage.getItem(groupingStorageKey);
    const isAvailable = groupByOptions.some(option => option.key === groupBy);

    return isAvailable ? groupBy as TGroupBy : 'none';
  } catch {
    return 'none';
  }
}

// зберігає вибране поле групування в localstorage
function saveGroupBy(groupBy: TGroupBy): void {
  try {
    window.localStorage.setItem(groupingStorageKey, groupBy);
  } catch {
    // браузер може заборонити доступ до localstorage
  }
}

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

// повторно використовує незмінені рядки та перетворює тільки нові дані
function getGridRows(
  requests: IServiceRequest[],
  rowCache: Map<number, ICachedRequestGridRow>
): IRequestGridRow[] {
  const activeRequestIds = new Set<number>();
  const rows: IRequestGridRow[] = [];

  for (const request of requests) {
    activeRequestIds.add(request.Id);

    const cachedRow = rowCache.get(request.Id);
    const row = cachedRow?.request === request
      ? cachedRow.row
      : toGridRow(request);

    rowCache.set(request.Id, { request, row });
    rows.push(row);
  }

  for (const requestId of Array.from(rowCache.keys())) {
    if (!activeRequestIds.has(requestId)) {
      rowCache.delete(requestId);
    }
  }

  return rows;
}

// повертає значення заявки для вибраного поля групування
function getGroupValue(row: IRequestGridRow, groupBy: TGroupBy): string {
  switch (groupBy) {
    case 'category':
      return row.category;
    case 'subcategory':
      return row.subcategory;
    case 'status':
      return row.status;
    case 'priority':
      return row.priority;
    case 'assignee':
      return row.assignee;
    case 'requester':
      return row.requester;
    default:
      return '';
  }
}

// додає до списку розгортані заголовки вибраних груп
function getGroupedRows(
  rows: IRequestGridRow[],
  groupBy: TGroupBy,
  collapsedGroupKeys: Set<string>
): IGridRow[] {
  if (groupBy === 'none') {
    return rows;
  }

  const groups = new Map<string, { label: string; rows: IRequestGridRow[] }>();

  for (const row of rows) {
    const label = getGroupValue(row, groupBy) || 'Не вказано';
    const key = `${groupBy}:${label}`;
    const group = groups.get(key);

    if (group) {
      group.rows.push(row);
    } else {
      groups.set(key, { label, rows: [row] });
    }
  }

  const groupedRows: IGridRow[] = [];
  let groupIndex = 0;

  for (const [key, group] of Array.from(groups.entries())) {
    const isCollapsed = collapsedGroupKeys.has(key);
    groupedRows.push({
      id: -(groupIndex + 1),
      isGroup: true,
      groupKey: key,
      groupLabel: group.label,
      groupCount: group.rows.length,
      isCollapsed,
      title: group.label,
      description: group.label,
      category: group.label,
      subcategory: group.label,
      status: group.label,
      priority: group.label,
      requester: group.label,
      assignee: group.label
    });

    if (!isCollapsed) {
      groupedRows.push(...group.rows);
    }

    groupIndex += 1;
  }

  return groupedRows;
}

// повертає стабільний ідентифікатор рядка для локального оновлення ag grid
function getRequestRowId(params: GetRowIdParams<IGridRow>): string {
  return params.data.isGroup
    ? `group:${params.data.groupKey}:${params.data.isCollapsed ? 'collapsed' : 'expanded'}`
    : String(params.data.id);
}

// повертає ім'я власника чинного блокування заявки
function getActiveEditLockOwner(request: IServiceRequest): string | undefined {
  const expiresAt = request.EditLockExpiresAt
    ? new Date(request.EditLockExpiresAt)
    : undefined;
  const isExpiryDateValid = expiresAt !== undefined && !Number.isNaN(expiresAt.getTime());
  const hasActiveLock = Boolean(request.EditLockOwnerId)
    && Boolean(request.EditLockToken?.trim())
    && isExpiryDateValid
    && expiresAt !== undefined
    && expiresAt > new Date();

  return hasActiveLock
    ? request.EditLockOwner?.Title ?? 'інший користувач'
    : undefined;
}

// показує кнопки дій для одного рядка таблиці
function RequestActionsRenderer(
  props: CustomCellRendererProps<IGridRow, undefined, IRequestGridContext>
): React.ReactElement {
  if (!props.data || props.data.isGroup) {
    return <></>;
  }

  const request = props.data.request;

  // передає вибрану заявку обробнику сторінки
  const handleView = (): void => {
    props.context.onView(request);
  };

  // передає вибрану заявку до форми редагування
  const handleEdit = (): void => {
    props.context.onEdit(request);
  };

  // передає вибрану заявку до підтвердження видалення
  const handleDelete = (): void => {
    props.context.onDelete(request);
  };

  const editLockOwner = getActiveEditLockOwner(request);
  const isDeleteDisabled = editLockOwner !== undefined;
  const deleteTooltip = isDeleteDisabled
    ? `Заявку редагує ${editLockOwner}`
    : 'Видалити';

  return (
    <div className={styles.actions}>
      <TooltipHost
        content="Переглянути"
        hostClassName={styles.actionTooltip}
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
        hostClassName={styles.actionTooltip}
      >
        <IconButton
          className={`${styles.actionButton} ${styles.editButton}`}
          iconProps={{ iconName: 'Edit' }}
          ariaLabel="Редагувати"
          onClick={handleEdit}
        />
      </TooltipHost>
      <TooltipHost
        content={deleteTooltip}
        hostClassName={styles.actionTooltip}
      >
        <IconButton
          className={`${styles.actionButton} ${styles.deleteButton}`}
          iconProps={{ iconName: 'Delete' }}
          ariaLabel="Видалити"
          onClick={handleDelete}
          disabled={isDeleteDisabled}
        />
      </TooltipHost>
    </div>
  );
}

// показує ідентифікатор заявки або заголовок групи в першій колонці
function RequestIdRenderer(
  props: CustomCellRendererProps<IGridRow, undefined, IRequestGridContext>
): React.ReactElement {
  if (!props.data) {
    return <></>;
  }

  if (!props.data.isGroup) {
    return <>{props.valueFormatted ?? props.value}</>;
  }

  const group = props.data;
  const iconName = group.isCollapsed ? 'ChevronRight' : 'ChevronDown';
  const actionName = group.isCollapsed ? 'Розгорнути' : 'Згорнути';

  return (
    <button
      type="button"
      className={styles.groupHeader}
      onClick={() => props.context.onToggleGroup(group.groupKey)}
      aria-label={`${actionName} групу ${group.groupLabel}`}
    >
      <Icon iconName={iconName} className={styles.groupHeaderIcon} />
      <span>{group.groupLabel}</span>
      <span className={styles.groupCount}>{group.groupCount}</span>
    </button>
  );
}

// показує дату українською мовою без зміни числового значення для сортування
function formatDueDate(params: ValueFormatterParams<IGridRow, Date>): string {
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
function formatEstimatedHours(params: ValueFormatterParams<IGridRow, number>): string {
  return typeof params.value === 'number' ? hoursFormatter.format(params.value) : '';
}

const columnDefs: ColDef<IGridRow>[] = [
  {
    field: 'id',
    headerName: 'ID',
    width: 90,
    pinned: 'left',
    lockPinned: true,
    suppressMovable: true,
    cellRenderer: RequestIdRenderer
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
    filter: ServiceRequestChoiceFilter
  },
  {
    field: 'subcategory',
    headerName: 'Підкатегорія',
    minWidth: 170,
    filter: ServiceRequestChoiceFilter
  },
  {
    field: 'status',
    headerName: 'Статус',
    minWidth: 130,
    filter: ServiceRequestChoiceFilter
  },
  {
    field: 'priority',
    headerName: 'Пріоритет',
    minWidth: 130,
    filter: ServiceRequestChoiceFilter
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
    colId: 'actions',
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

const defaultColDef: ColDef<IGridRow> = {
  sortable: true,
  resizable: true,
  filter: false,
  cellClassRules: {
    [styles.groupRowCell]: params => Boolean(params.data?.isGroup)
  }
};

// відображає заявки в ag grid із сортуванням та українською локалізацією
export default function ServiceRequestsGrid(props: IServiceRequestsGridProps): React.ReactElement {
  const [searchText, setSearchText] = React.useState('');
  const [groupBy, setGroupBy] = React.useState<TGroupBy>(getStoredGroupBy);
  const [collapsedGroupKeys, setCollapsedGroupKeys] = React.useState<Set<string>>(new Set());
  const gridApi = React.useRef<GridApi<IGridRow>>();
  const rowCache = React.useRef<Map<number, ICachedRequestGridRow>>(new Map());
  const requestRows = getGridRows(props.requests, rowCache.current);
  const rows = getGroupedRows(requestRows, groupBy, collapsedGroupKeys);

  const gridContext: IRequestGridContext = {
    onView: props.onView,
    onEdit: props.onEdit,
    onDelete: props.onDelete,
    onToggleGroup: groupKey => {
      setCollapsedGroupKeys(currentKeys => {
        const nextKeys = new Set(currentKeys);

        if (nextKeys.has(groupKey)) {
          nextKeys.delete(groupKey);
        } else {
          nextKeys.add(groupKey);
        }

        return nextKeys;
      });
    }
  };
  // зберігає введений текст загального пошуку
  const handleSearchChange = (_event?: React.ChangeEvent<HTMLInputElement>, value?: string): void => {
    setSearchText(value || '');
  };

  // передає ідентифікатори рядків поточної сторінки для періодичного оновлення
  const reportVisibleRequestIds = (): void => {
    const api = gridApi.current;

    if (!api) {
      return;
    }

    const pageSize = api.paginationGetPageSize();
    const pageStartIndex = api.paginationGetCurrentPage() * pageSize;
    const pageEndIndex = pageStartIndex + pageSize;
    const visibleRequestIds: number[] = [];

    api.forEachNodeAfterFilterAndSort(node => {
      if (
        node.data
        && !node.data.isGroup
        && node.rowIndex !== null
        && node.rowIndex >= pageStartIndex
        && node.rowIndex < pageEndIndex
      ) {
        visibleRequestIds.push(node.data.request.Id);
      }
    });

    props.onVisibleRequestIdsChange(visibleRequestIds);
  };

  // змінює поле групування та розгортає всі його групи
  const handleGroupByChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (!option || !groupByOptions.some(groupOption => groupOption.key === option.key)) {
      return;
    }

    const nextGroupBy = option.key as TGroupBy;
    setGroupBy(nextGroupBy);
    setCollapsedGroupKeys(new Set());
    saveGroupBy(nextGroupBy);
  };

  // згортає всі групи вибраного поля
  const handleCollapseAllGroups = (): void => {
    setCollapsedGroupKeys(new Set(
      requestRows.map(row => `${groupBy}:${getGroupValue(row, groupBy) || 'Не вказано'}`)
    ));
  };

  // розгортає всі групи вибраного поля
  const handleExpandAllGroups = (): void => {
    setCollapsedGroupKeys(new Set());
  };

  // зберігає api таблиці для очищення фільтрів
  const handleGridReady = (event: GridReadyEvent<IGridRow>): void => {
    gridApi.current = event.api;
    applyStoredColumnState(event.api);
    reportVisibleRequestIds();
  };

  // зберігає ширину колонки після завершення зміни розміру
  const handleColumnResized = (event: ColumnResizedEvent<IGridRow>): void => {
    if (event.finished) {
      saveColumnState(event.api);
    }
  };

  // зберігає порядок колонок після перетягування
  const handleColumnMoved = (event: ColumnMovedEvent<IGridRow>): void => {
    saveColumnState(event.api);
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

          <DefaultButton
            className={styles.secondaryActionButton}
            text="Очистити фільтри"
            onClick={handleClearFilters}
          />
        </div>
      </div>

      <div className={styles.groupingControls}>
        <Dropdown
          className={styles.groupingSelect}
          label="Групувати за"
          selectedKey={groupBy}
          options={groupByOptions}
          onChange={handleGroupByChange}
        />

        {groupBy !== 'none' && (
          <div className={styles.groupingActions}>
            <DefaultButton
              className={styles.secondaryActionButton}
              text="Розгорнути всі"
              onClick={handleExpandAllGroups}
            />
            <DefaultButton
              className={styles.secondaryActionButton}
              text="Згорнути всі"
              onClick={handleCollapseAllGroups}
            />
          </div>
        )}
      </div>

      <div className={styles.grid}>
        <AgGridReact<IGridRow>
          rowData={rows}
          getRowId={getRequestRowId}
          columnDefs={columnDefs}
          defaultColDef={defaultColDef}
          localeText={localeText}
          context={gridContext}
          theme={themeQuartz}
          accentedSort={true}
          suppressCellFocus={groupBy !== 'none'}
          pagination={true}
          paginationPageSize={10}
          paginationPageSizeSelector={[10, 25, 50]}
          quickFilterText={searchText}
          cacheQuickFilter={true}
          getRowHeight={params => params.data?.isGroup ? 44 : undefined}
          onGridReady={handleGridReady}
          onColumnResized={handleColumnResized}
          onColumnMoved={handleColumnMoved}
          onPaginationChanged={reportVisibleRequestIds}
          onFilterChanged={reportVisibleRequestIds}
          onSortChanged={reportVisibleRequestIds}
          onModelUpdated={reportVisibleRequestIds}
        />
      </div>
      
    </AgGridProvider>
  );
}
