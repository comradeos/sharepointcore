import { GridApi } from 'ag-grid-community';
import { IGridRow, TGroupBy } from './ServiceRequestsGrid.types';

const columnStateStorageKey = 'spfx-task-oseledko-yy.service-requests-grid.columns.v1';
const groupingStorageKey = 'spfx-task-oseledko-yy.service-requests-grid.grouping.v1';

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
export function applyStoredColumnState(api: GridApi<IGridRow>): void {
  const columnState = getStoredColumnState();

  if (columnState) {
    api.applyColumnState({ state: columnState, applyOrder: true });
  }
}

// зберігає поточний порядок і ширину колонок у localstorage
export function saveColumnState(api: GridApi<IGridRow>): void {
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
export function getStoredGroupBy(availableGroupBy: readonly TGroupBy[]): TGroupBy {
  try {
    const groupBy = window.localStorage.getItem(groupingStorageKey) as TGroupBy | null;
    return groupBy && availableGroupBy.indexOf(groupBy) !== -1 ? groupBy : 'none';
  } catch {
    return 'none';
  }
}

// зберігає вибране поле групування в localstorage
export function saveGroupBy(groupBy: TGroupBy): void {
  try {
    window.localStorage.setItem(groupingStorageKey, groupBy);
  } catch {
    // браузер може заборонити доступ до localstorage
  }
}
