import { SPHttpClient } from '@microsoft/sp-http';
import { IRequestCategory, IRequestSubcategory, IServiceDeskData, IServiceRequest, IServiceRequestDraft, SharePointNullable } from '../models/ServiceDeskModels';
import SharePointListUtils from '../utils/SharePointListUtils';

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

const serviceRequestSelectFields = [
  'Id', 'Title', 'Description',
  'CategoryId', 'Category/Id', 'Category/Title',
  'SubcategoryId', 'Subcategory/Id', 'Subcategory/Title',
  'Status', 'Priority',
  'RequesterId', 'Requester/Id', 'Requester/Title', 'Requester/EMail',
  'AssigneeId', 'Assignee/Id', 'Assignee/Title', 'Assignee/EMail',
  'PlannedStart', 'DueDate', 'EstimatedHours', 'ContactEmail',
  'RequiresOnsiteVisit'
];

const serviceRequestExpandFields = [
  'Category',
  'Subcategory',
  'Requester',
  'Assignee'
];

// читає заявки та довідники із сайту де розміщена вебчастина
export default class ServiceDeskService {
  private readonly userIdCache = new Map<string, Promise<number>>();
  private readonly categoriesList: SharePointListUtils<IRequestCategory>;
  private readonly subcategoriesList: SharePointListUtils<IRequestSubcategory>;
  private readonly requestsList: SharePointListUtils<IServiceRequest, IServiceRequestPayload>;

  // зберігає клієнт spfx та адресу поточного сайту для всіх запитів
  public constructor(
    private readonly client: SPHttpClient,
    private readonly webUrl: string
  ) {
    this.categoriesList = new SharePointListUtils<IRequestCategory>(
      client,
      webUrl,
      {
        listTitle: 'RequestCategories',
        selectFields: ['Id', 'Title', 'IsActive']
      }
    );
    this.subcategoriesList = new SharePointListUtils<IRequestSubcategory>(
      client,
      webUrl,
      {
        listTitle: 'RequestSubcategories',
        selectFields: [
          'Id',
          'Title',
          'IsActive',
          'CategoryId',
          'Category/Id',
          'Category/Title'
        ],
        expandFields: ['Category']
      }
    );
    this.requestsList = new SharePointListUtils<IServiceRequest, IServiceRequestPayload>(
      client,
      webUrl,
      {
        listTitle: 'ServiceRequests',
        selectFields: serviceRequestSelectFields,
        expandFields: serviceRequestExpandFields
      }
    );
  }

  // завантажує три списки паралельно та повертає їх як один набір даних
  public async loadData(): Promise<IServiceDeskData> {
    const [categories, subcategories, requests] = await Promise.all([
      this.categoriesList.getAll(),
      this.subcategoriesList.getAll(),
      this.requestsList.getAll()
    ]);

    return { categories, subcategories, requests };
  }

  // створює нову заявку у списку servicerequests
  public async createRequest(draft: IServiceRequestDraft): Promise<IServiceRequest> {
    const payload = await this.createRequestPayload(draft);

    return this.requestsList.create(payload);
  }

  // оновлює наявну заявку у списку servicerequests
  public async updateRequest(
    itemId: number,
    draft: IServiceRequestDraft
  ): Promise<IServiceRequest> {
    const payload = await this.createRequestPayload(draft);

    return this.requestsList.update(itemId, payload);
  }

  // видаляє заявку зі списку servicerequests
  public async deleteRequest(itemId: number): Promise<void> {
    await this.requestsList.delete(itemId);
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

}
