import * as React from 'react';
import { ColDef, ValueFormatterParams } from 'ag-grid-community';
import { CustomCellRendererProps } from 'ag-grid-react';
import { Icon, IconButton, TooltipHost } from '@fluentui/react';
import { IServiceRequest } from '../models/ServiceDeskModels';
import ServiceRequestChoiceFilter from './ServiceRequestChoiceFilter';
import { IGridRow, IRequestGridContext } from './ServiceRequestsGrid.types';
import styles from './ServiceRequestsGrid.module.scss';

const dateFormatter = new Intl.DateTimeFormat('uk-UA', {
  day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit'
});
const hoursFormatter = new Intl.NumberFormat('uk-UA', {
  minimumFractionDigits: 1, maximumFractionDigits: 1
});

// повертає ім'я власника чинного блокування заявки
function getActiveEditLockOwner(request: IServiceRequest): string | undefined {
  const expiresAt = request.EditLockExpiresAt ? new Date(request.EditLockExpiresAt) : undefined;
  const isExpiryDateValid = expiresAt !== undefined && !Number.isNaN(expiresAt.getTime());
  const hasActiveLock = Boolean(request.EditLockOwnerId)
    && Boolean(request.EditLockToken?.trim())
    && isExpiryDateValid
    && expiresAt !== undefined
    && expiresAt > new Date();

  return hasActiveLock ? request.EditLockOwner?.Title ?? 'інший користувач' : undefined;
}

// показує кнопки дій для одного рядка таблиці
function RequestActionsRenderer(
  props: CustomCellRendererProps<IGridRow, undefined, IRequestGridContext>
): React.ReactElement {
  if (!props.data || props.data.isGroup) {
    return <></>;
  }

  const request = props.data.request;
  const editLockOwner = getActiveEditLockOwner(request);
  const isDeleteDisabled = editLockOwner !== undefined;
  const deleteTooltip = isDeleteDisabled ? `Заявку редагує ${editLockOwner}` : 'Видалити';

  return (
    <div className={styles.actions}>
      <TooltipHost content="Переглянути" hostClassName={styles.actionTooltip}>
        <IconButton className={`${styles.actionButton} ${styles.viewButton}`} iconProps={{ iconName: 'RedEye' }} ariaLabel="Переглянути" onClick={() => props.context.onView(request)} />
      </TooltipHost>
      <TooltipHost content="Редагувати" hostClassName={styles.actionTooltip}>
        <IconButton className={`${styles.actionButton} ${styles.editButton}`} iconProps={{ iconName: 'Edit' }} ariaLabel="Редагувати" onClick={() => props.context.onEdit(request)} />
      </TooltipHost>
      <TooltipHost content={deleteTooltip} hostClassName={styles.actionTooltip}>
        <IconButton className={`${styles.actionButton} ${styles.deleteButton}`} iconProps={{ iconName: 'Delete' }} ariaLabel="Видалити" onClick={() => props.context.onDelete(request)} disabled={isDeleteDisabled} />
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
    <button type="button" className={styles.groupHeader} onClick={() => props.context.onToggleGroup(group.groupKey)} aria-label={`${actionName} групу ${group.groupLabel}`}>
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

  const cellDate = new Date(cellValue.getFullYear(), cellValue.getMonth(), cellValue.getDate());
  return cellDate.getTime() - filterDate.getTime();
}

// показує кількість годин з українським десятковим роздільником
function formatEstimatedHours(params: ValueFormatterParams<IGridRow, number>): string {
  return typeof params.value === 'number' ? hoursFormatter.format(params.value) : '';
}

export const requestGridColumnDefs: ColDef<IGridRow>[] = [
  { field: 'id', headerName: 'ID', width: 90, pinned: 'left', lockPinned: true, suppressMovable: true, cellRenderer: RequestIdRenderer },
  { field: 'title', headerName: 'Назва заявки', width: 150, filter: 'agTextColumnFilter' },
  { field: 'description', headerName: 'Опис', minWidth: 240, filter: 'agTextColumnFilter' },
  { field: 'category', headerName: 'Категорія', minWidth: 160, filter: ServiceRequestChoiceFilter },
  { field: 'subcategory', headerName: 'Підкатегорія', minWidth: 170, filter: ServiceRequestChoiceFilter },
  { field: 'status', headerName: 'Статус', minWidth: 130, filter: ServiceRequestChoiceFilter },
  { field: 'priority', headerName: 'Пріоритет', minWidth: 130, filter: ServiceRequestChoiceFilter },
  { field: 'requester', headerName: 'Заявник', minWidth: 170, filter: 'agTextColumnFilter' },
  { field: 'assignee', headerName: 'Виконавець', minWidth: 170, filter: 'agTextColumnFilter' },
  { field: 'dueDate', headerName: 'Кінцевий термін', minWidth: 180, valueFormatter: formatDueDate, filter: 'agDateColumnFilter', filterParams: { browserDatePicker: true, comparator: compareDueDate } },
  { field: 'estimatedHours', headerName: 'Оцінка годин', minWidth: 140, valueFormatter: formatEstimatedHours, filter: 'agNumberColumnFilter' },
  { colId: 'actions', headerName: 'Дії', width: 136, minWidth: 136, maxWidth: 136, pinned: 'right', sortable: false, resizable: false, suppressMovable: true, cellRenderer: RequestActionsRenderer }
];

export const requestGridDefaultColumnDef: ColDef<IGridRow> = {
  sortable: true,
  resizable: true,
  filter: false,
  cellClassRules: {
    [styles.groupRowCell]: params => Boolean(params.data?.isGroup)
  }
};
