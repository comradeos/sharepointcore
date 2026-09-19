import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  IRequestCategory,
  IRequestSubcategory,
  IServiceDeskData,
  IServiceRequest,
  IServiceRequestDraft
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

// читає заявки та довідники із сайту де розміщена вебчастина
export default class ServiceDeskService {
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
    const requesterIdPromise = this.ensureUser(draft.requesterIdentity);
    const assigneeIdPromise = draft.assigneeIdentity
      ? this.ensureUser(draft.assigneeIdentity)
      : Promise.resolve<number | null>(null);
    const [requesterId, assigneeId] = await Promise.all([
      requesterIdPromise,
      assigneeIdPromise
    ]);
    const listUrl = `${this.webUrl.replace(/\/$/, '')}/_api/web/lists/getbytitle('ServiceRequests')/items`;
    const response = await this.client.post(
      listUrl,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata',
          'Content-Type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify({
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
        })
      }
    );

    if (!response.ok) {
      throw new Error(`Не вдалося створити заявку (HTTP ${response.status})`);
    }
  }

  // додає користувача до сайту та повертає його числовий ідентифікатор
  private async ensureUser(identity: string): Promise<number> {
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
