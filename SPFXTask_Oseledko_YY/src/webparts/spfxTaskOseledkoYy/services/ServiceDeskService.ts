import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IRequestCategory,
  IRequestSubcategory,
  IServiceDeskData,
  IServiceRequest,
  IServiceRequestDraft,
  SharePointNullable
} from '../models/ServiceDeskModels';

// сторінка результатів sharepoint rest api без службових метаданих
interface IODataPage<T> {
  value: T[];
  '@odata.nextLink'?: string;
}

// користувач якого повертає метод ensureuser
interface IEnsuredUser {
  Id: number;
}

// дані заявки у форматі внутрішніх полів списку sharepoint
interface IServiceRequestPayload {
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

// читає заявки та довідники із сайту де розміщена вебчастина
export default class ServiceDeskService {
  private readonly userIdCache = new Map<string, Promise<number>>();

  // зберігає клієнт spfx та адресу поточного сайту для всіх запитів
  public constructor(
    private readonly client: SPHttpClient,
    private readonly webUrl: string
  ) {}

  // завантажує три списки паралельно та повертає їх як один набір даних
  public async loadData(): Promise<IServiceDeskData> {
    const [categories, subcategories, requests] = await Promise.all([
      this.getItems<IRequestCategory>(
        'RequestCategories',
        ['Id', 'Title', 'IsActive']
      ),
      this.getItems<IRequestSubcategory>(
        'RequestSubcategories',
        ['Id', 'Title', 'IsActive', 'CategoryId', 'Category/Id', 'Category/Title'],
        ['Category']
      ),
      this.getItems<IServiceRequest>(
        'ServiceRequests',
        [
          'Id', 'Title', 'Description',
          'CategoryId', 'Category/Id', 'Category/Title',
          'SubcategoryId', 'Subcategory/Id', 'Subcategory/Title',
          'Status', 'Priority',
          'RequesterId', 'Requester/Id', 'Requester/Title', 'Requester/EMail',
          'AssigneeId', 'Assignee/Id', 'Assignee/Title', 'Assignee/EMail',
          'PlannedStart', 'DueDate', 'EstimatedHours', 'ContactEmail',
          'RequiresOnsiteVisit'
        ],
        ['Category', 'Subcategory', 'Requester', 'Assignee']
      )
    ]);

    return { categories, subcategories, requests };
  }

  // створює нову заявку у списку servicerequests
  public async createRequest(draft: IServiceRequestDraft): Promise<void> {
    const payload = await this.createRequestPayload(draft);
    const listUrl = `${this.webUrl.replace(/\/$/, '')}/_api/web/lists/getbytitle('ServiceRequests')/items`;
    const response = await this.client.post(
      listUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(payload)
      }
    );

    if (!response.ok) {
      throw new Error(`Не вдалося створити заявку (HTTP ${response.status})`);
    }
  }

  // оновлює наявну заявку у списку servicerequests
  public async updateRequest(itemId: number, draft: IServiceRequestDraft): Promise<void> {
    const payload = await this.createRequestPayload(draft);
    const itemUrl = `${this.webUrl.replace(/\/$/, '')}/_api/web/lists/getbytitle('ServiceRequests')/items(${itemId})`;
    const response = await this.client.post(
      itemUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
      }
    );

    if (!response.ok) {
      throw new Error(`Не вдалося оновити заявку (HTTP ${response.status})`);
    }
  }

  // видаляє заявку зі списку servicerequests
  public async deleteRequest(itemId: number): Promise<void> {
    const itemUrl = `${this.webUrl.replace(/\/$/, '')}/_api/web/lists/getbytitle('ServiceRequests')/items(${itemId})`;
    const response = await this.client.post(
      itemUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'IF-MATCH': '*',
          'X-HTTP-Method': 'DELETE'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Не вдалося видалити заявку (HTTP ${response.status})`);
    }
  }

  // готує поля заявки та визначає ідентифікатори користувачів
  private async createRequestPayload(draft: IServiceRequestDraft): Promise<IServiceRequestPayload> {
    const requesterIdPromise = this.ensureUser(draft.requesterIdentity);
    const assigneeIdPromise = draft.assigneeIdentity
      ? this.ensureUser(draft.assigneeIdentity)
      : Promise.resolve<number | null>(null);
    const [requesterId, assigneeId] = await Promise.all([
      requesterIdPromise,
      assigneeIdPromise
    ]);

    return {
      Title: draft.title,
      Description: draft.description,
      CategoryId: draft.categoryId,
      SubcategoryId: draft.subcategoryId,
      Status: draft.status,
      Priority: draft.priority,
      RequesterId: requesterId,
      AssigneeId: assigneeId,
      PlannedStart: draft.plannedStart ? new Date(draft.plannedStart).toISOString() : null,
      DueDate: new Date(draft.dueDate).toISOString(),
      EstimatedHours: draft.estimatedHours ?? null,
      ContactEmail: draft.contactEmail ?? null,
      RequiresOnsiteVisit: draft.requiresOnsiteVisit
    };
  }

  // додає користувача до сайту та повертає його числовий ідентифікатор
  private async ensureUser(identity: string): Promise<number> {
    const cacheKey = identity.trim().toLowerCase();
    let userIdPromise = this.userIdCache.get(cacheKey);

    if (!userIdPromise) {
      userIdPromise = this.requestUserId(identity);
      this.userIdCache.set(cacheKey, userIdPromise);
    }

    try {
      return await userIdPromise;
    } catch (error) {
      this.userIdCache.delete(cacheKey);
      throw error;
    }
  }

  // виконує запит sharepoint для визначення ідентифікатора користувача
  private async requestUserId(identity: string): Promise<number> {
    const ensureUserUrl = `${this.webUrl.replace(/\/$/, '')}/_api/web/ensureuser`;
    const response = await this.client.post(
      ensureUserUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({ logonName: identity })
      }
    );

    if (!response.ok) {
      throw new Error(`Не вдалося визначити користувача (HTTP ${response.status})`);
    }

    const user = await response.json() as IEnsuredUser;
    if (!Number.isInteger(user.Id)) {
      throw new Error('SharePoint повернув некоректний ідентифікатор користувача');
    }

    return user.Id;
  }

  // читає всі сторінки списку щоб отримати кожен елемент
  private async getItems<T>(
    listTitle: string,
    selectFields: string[],
    expandFields: string[] = []
  ): Promise<T[]> {
    const listName = listTitle.replace(/'/g, "''");
    const query = [`$select=${selectFields.join(',')}`, '$top=5000'];

    if (expandFields.length > 0) {
      query.push(`$expand=${expandFields.join(',')}`);
    }

    let nextUrl: string | undefined =
      `${this.webUrl.replace(/\/$/, '')}/_api/web/lists/getbytitle('${listName}')/items?${query.join('&')}`;
    const items: T[] = [];

    while (nextUrl) {
      const response: SPHttpClientResponse = await this.client.get(
        nextUrl,
        SPHttpClient.configurations.v1,
        { headers: { Accept: 'application/json;odata=nometadata' } }
      );

      if (!response.ok) {
        throw new Error(`Не вдалося завантажити список ${listTitle} (HTTP ${response.status}).`);
      }

      const page = await response.json() as IODataPage<T>;
      if (!Array.isArray(page.value)) {
        throw new Error(`Список ${listTitle} повернув неочікуваний формат даних.`);
      }

      items.push(...page.value);
      nextUrl = page['@odata.nextLink'];
    }

    return items;
  }
}
