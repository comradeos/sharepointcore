import { IServiceRequest } from '../models/ServiceDeskModels';

// значення одного рядка таблиці після перетворення даних sharepoint
export interface IRequestGridRowValues {
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
export interface IRequestGridRow extends IRequestGridRowValues {
  request: IServiceRequest;
  isGroup?: false;
}

// рядок заголовка групи
export interface IGroupGridRow extends IRequestGridRowValues {
  isGroup: true;
  groupKey: string;
  groupLabel: string;
  groupCount: number;
  isCollapsed: boolean;
}

export type IGridRow = IRequestGridRow | IGroupGridRow;
export type TGroupBy = 'none' | 'category' | 'subcategory' | 'status' | 'priority' | 'assignee' | 'requester';

// вхідні дані таблиці заявок
export interface IServiceRequestsGridProps {
  requests: IServiceRequest[];
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
  onDelete: (request: IServiceRequest) => void;
  onVisibleRequestIdsChange: (requestIds: number[]) => void;
}

// дії таблиці доступні клітинці через контекст ag grid
export interface IRequestGridContext {
  onView: (request: IServiceRequest) => void;
  onEdit: (request: IServiceRequest) => void;
  onDelete: (request: IServiceRequest) => void;
  onToggleGroup: (groupKey: string) => void;
}

// кеш одного перетвореного рядка таблиці
export interface ICachedRequestGridRow {
  request: IServiceRequest;
  row: IRequestGridRow;
}

// опис колонки для експорту в excel
export interface IExcelExportColumn {
  field: keyof IRequestGridRowValues;
  headerName: string;
  width: number;
}
