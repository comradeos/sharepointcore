// порожнє значення sharepoint rest api повертає як null
export type SharePointNullable<T> = T | null;

// значення поля lookup яке повертає sharepoint rest api
export interface ILookupValue {
  Id: number;
  Title: string;
}

// користувач із поля person якого повертає sharepoint rest api
export interface IPersonValue extends ILookupValue {
  EMail: string;
}

// елемент довідника категорій
export interface IRequestCategory {
  Id: number;
  Title: string;
  IsActive: boolean;
}

// елемент довідника підкатегорій та його належність до категорії
export interface IRequestSubcategory {
  Id: number;
  Title: string;
  IsActive: boolean;
  CategoryId: number;
  Category: ILookupValue;
}

// заявка з розгорнутими значеннями lookup і person
export interface IServiceRequest {
  Id: number;
  Title: string;
  Description: string;
  CategoryId: number;
  Category: ILookupValue;
  SubcategoryId: number;
  Subcategory: ILookupValue;
  Status: string;
  Priority: string;
  RequesterId: number;
  Requester: IPersonValue;
  AssigneeId: SharePointNullable<number>;
  Assignee: SharePointNullable<IPersonValue>;
  PlannedStart: SharePointNullable<string>;
  DueDate: string;
  EstimatedHours: SharePointNullable<number>;
  ContactEmail: SharePointNullable<string>;
  RequiresOnsiteVisit: boolean;
}

// перевірені дані форми для створення заявки
export interface IServiceRequestDraft {
  title: string;
  description: string;
  categoryId: number;
  subcategoryId: number;
  status: string;
  priority: string;
  requesterIdentity: string;
  assigneeIdentity?: string;
  plannedStart?: string;
  dueDate: string;
  estimatedHours?: number;
  contactEmail?: string;
  requiresOnsiteVisit: boolean;
}

// дані потрібні для першого відображення вебчастини
export interface IServiceDeskData {
  categories: IRequestCategory[];
  subcategories: IRequestSubcategory[];
  requests: IServiceRequest[];
}
