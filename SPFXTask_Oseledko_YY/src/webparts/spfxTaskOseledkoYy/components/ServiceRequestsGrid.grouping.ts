import { GetRowIdParams } from 'ag-grid-community';
import { IServiceRequest } from '../models/ServiceDeskModels';
import {
  ICachedRequestGridRow,
  IGridRow,
  IRequestGridRow,
  TGroupBy
} from './ServiceRequestsGrid.types';

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
export function getGridRows(
  requests: IServiceRequest[],
  rowCache: Map<number, ICachedRequestGridRow>
): IRequestGridRow[] {
  const activeRequestIds = new Set<number>();
  const rows: IRequestGridRow[] = [];

  for (const request of requests) {
    activeRequestIds.add(request.Id);
    const cachedRow = rowCache.get(request.Id);
    const row = cachedRow?.request === request ? cachedRow.row : toGridRow(request);

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
export function getGroupValue(row: IRequestGridRow, groupBy: TGroupBy): string {
  switch (groupBy) {
    case 'category': return row.category;
    case 'subcategory': return row.subcategory;
    case 'status': return row.status;
    case 'priority': return row.priority;
    case 'assignee': return row.assignee;
    case 'requester': return row.requester;
    default: return '';
  }
}

// додає до списку розгортані заголовки вибраних груп
export function getGroupedRows(
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
export function getRequestRowId(params: GetRowIdParams<IGridRow>): string {
  return params.data.isGroup
    ? `group:${params.data.groupKey}:${params.data.isCollapsed ? 'collapsed' : 'expanded'}`
    : String(params.data.id);
}
