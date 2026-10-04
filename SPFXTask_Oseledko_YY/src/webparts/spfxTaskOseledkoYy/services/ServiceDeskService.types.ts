import { SharePointNullable } from '../models/ServiceDeskModels';

// користувач якого повертає метод ensureuser
export interface IEnsuredUser {
  Id: number;
}

// поля потрібні для підрахунку заявок за датою створення
export interface IRequestDateMetadata {
  Id: number;
  Created: string;
}

// дані заявки у форматі внутрішніх полів списку sharepoint
export interface IServiceRequestPayload {
  Title: string;
  Description: string;
  CategoryId: number;
  SubcategoryId: number;
  Status: string;
  Priority: string;
  RequesterId: number;
  AssigneeId: SharePointNullable<number>;
  PlannedStart: SharePointNullable<string>;
  DueDate: string;
  EstimatedHours: SharePointNullable<number>;
  ContactEmail: SharePointNullable<string>;
  RequiresOnsiteVisit: boolean;
}

// поля потрібні для тимчасового блокування редагування заявки
export interface IRequestEditLockPayload {
  EditLockOwnerId?: SharePointNullable<number>;
  EditLockExpiresAt?: SharePointNullable<string>;
  EditLockToken?: SharePointNullable<string>;
}

// дані для часткового оновлення заявки та керування її блокуванням
export type IServiceRequestUpdatePayload = Partial<IServiceRequestPayload> & IRequestEditLockPayload;
