import * as React from 'react';
import {
  AllCommunityModule,
  ColumnMovedEvent,
  ColumnResizedEvent,
  GridApi,
  GridReadyEvent,
  themeQuartz
} from 'ag-grid-community';
import { AG_GRID_LOCALE_UA } from '@ag-grid-community/locale';
import { AgGridProvider, AgGridReact } from 'ag-grid-react';
import { DefaultButton, Dropdown, IDropdownOption, SearchBox, Text } from '@fluentui/react';
import { requestGridColumnDefs, requestGridDefaultColumnDef } from './ServiceRequestsGrid.columns';
import { exportGridToExcel } from './ServiceRequestsGrid.export';
import { getGridRows, getGroupedRows, getGroupValue, getRequestRowId } from './ServiceRequestsGrid.grouping';
import { applyStoredColumnState, getStoredGroupBy, saveColumnState, saveGroupBy } from './ServiceRequestsGrid.storage';
import {
  ICachedRequestGridRow,
  IGridRow,
  IRequestGridContext,
  IServiceRequestsGridProps,
  TGroupBy
} from './ServiceRequestsGrid.types';
import styles from './ServiceRequestsGrid.module.scss';

const modules = [AllCommunityModule];
const localeText = {
  ...AG_GRID_LOCALE_UA,
  noRowsToShow: 'Заявок поки немає'
};
const groupByOptions: IDropdownOption[] = [
  { key: 'none', text: 'Без групування' },
  { key: 'category', text: 'Категорія' },
  { key: 'subcategory', text: 'Підкатегорія' },
  { key: 'status', text: 'Статус' },
  { key: 'priority', text: 'Пріоритет' },
  { key: 'assignee', text: 'Виконавець' },
  { key: 'requester', text: 'Заявник' }
];
const availableGroupBy = groupByOptions.map(option => option.key as TGroupBy);

// відображає заявки в ag grid із сортуванням та українською локалізацією
export default function ServiceRequestsGrid(props: IServiceRequestsGridProps): React.ReactElement {
  const [searchText, setSearchText] = React.useState('');
  const [groupBy, setGroupBy] = React.useState<TGroupBy>(() => getStoredGroupBy(availableGroupBy));
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

  // експортує відфільтровані та відсортовані заявки у файл excel
  const handleExportToExcel = (): void => {
    const api = gridApi.current;

    if (api) {
      exportGridToExcel(api);
    }
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
          columnDefs={requestGridColumnDefs}
          defaultColDef={requestGridDefaultColumnDef}
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

      <div className={styles.exportControls}>
        <DefaultButton
          className={styles.exportButton}
          text="Експортувати в Excel"
          iconProps={{ iconName: 'ExcelLogo' }}
          onClick={handleExportToExcel}
        />
      </div>
    </AgGridProvider>
  );
}
