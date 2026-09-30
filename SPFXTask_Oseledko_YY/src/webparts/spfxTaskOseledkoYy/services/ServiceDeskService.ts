import { SPHttpClient } from '@microsoft/sp-http';
import {
  IRequestCategory,
  IRequestCreatedDateRange,
  IRequestEditSession,
  IRequestPageCursor,
  IRequestSubcategory,
  IServiceDeskDictionaries,
  IServiceRequest,
  IServiceRequestDraft,
  IServiceRequestPage,
  SharePointNullable
} from '../models/ServiceDeskModels';
import SharePointListUtils from '../utils/SharePointListUtils';

// користувач якого повертає метод ensureuser
interface IEnsuredUser {
  Id: number;
}

// поля потрібні для підрахунку заявок за датою створення
interface IRequestDateMetadata {
  Id: number;
  Created: string;
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

// поля потрібні для тимчасового блокування редагування заявки
interface IRequestEditLockPayload {
  EditLockOwnerId?: SharePointNullable<number>;
  EditLockExpiresAt?: SharePointNullable<string>;
  EditLockToken?: SharePointNullable<string>;
}

// дані для часткового оновлення заявки та керування її блокуванням
type IServiceRequestUpdatePayload = Partial<IServiceRequestPayload> & IRequestEditLockPayload;

const serviceRequestSelectFields = [
  'Id', 'Title', 'Description',
  'CategoryId', 'Category/Id', 'Category/Title',
  'SubcategoryId', 'Subcategory/Id', 'Subcategory/Title',
  'Status', 'Priority',
  'RequesterId', 'Requester/Id', 'Requester/Title', 'Requester/EMail',
  'AssigneeId', 'Assignee/Id', 'Assignee/Title', 'Assignee/EMail',
  'PlannedStart', 'DueDate', 'EstimatedHours', 'ContactEmail',
  'RequiresOnsiteVisit', 'Created',
  'EditLockOwnerId', 'EditLockOwner/Id', 'EditLockOwner/Title', 'EditLockOwner/EMail',
  'EditLockExpiresAt', 'EditLockToken'
];

const serviceRequestExpandFields = [
  'Category',
  'Subcategory',
  'Requester',
  'Assignee',
  'EditLockOwner'
];
const initialRequestsPageSize = 50;
const requestIdBlockSize = 4000;
const editLockDurationMinutes = 5;

// читає заявки та довідники із сайту де розміщена вебчастина
export default class ServiceDeskService {
  private readonly userIdCache = new Map<string, Promise<number>>();
  private readonly categoriesList: SharePointListUtils<IRequestCategory>;
  private readonly subcategoriesList: SharePointListUtils<IRequestSubcategory>;
  private readonly requestMetadataList: SharePointListUtils<IRequestDateMetadata>;
  private readonly requestsList: SharePointListUtils<IServiceRequest, IServiceRequestUpdatePayload>;

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
    this.requestMetadataList = new SharePointListUtils<IRequestDateMetadata>(
      client,
      webUrl,
      {
        listTitle: 'ServiceRequests',
        selectFields: ['Id', 'Created'],
        pageSize: requestIdBlockSize
      }
    );
    this.requestsList = new SharePointListUtils<IServiceRequest, IServiceRequestUpdatePayload>(
      client,
      webUrl,
      {
        listTitle: 'ServiceRequests',
        selectFields: serviceRequestSelectFields,
        expandFields: serviceRequestExpandFields,
        pageSize: initialRequestsPageSize
      }
    );
  }

  // завантажує лише довідники потрібні для форми заявки
  public async loadDictionaries(): Promise<IServiceDeskDictionaries> {
    const [categories, subcategories] = await Promise.all([
      this.categoriesList.getAll(),
      this.subcategoriesList.getAll()
    ]);

    return { categories, subcategories };
  }

  // завантажує заявки за вибраним діапазоном дат створення
  public async loadRequestsPage(
    createdDateRange: IRequestCreatedDateRange,
    cursor?: IRequestPageCursor,
    pageSize: number = initialRequestsPageSize
  ): Promise<IServiceRequestPage> {
    if (cursor) {
      return this.loadMatchingRequestsPage(createdDateRange, cursor, pageSize);
    }

    try {
      return await this.loadFirstServerFilteredRequestsPage(createdDateRange);
    } catch (error) {
      const isListViewThresholdExceeded = this.isListViewThresholdError(error);

      if (!isListViewThresholdExceeded) {
        throw error;
      }

      const initialCursor = this.createInitialRequestPageCursor(
        await this.getLatestRequestId()
      );

      return initialCursor
        ? this.loadMatchingRequestsPage(createdDateRange, initialCursor, pageSize)
        : { requests: [] };
    }
  }

  // завантажує першу сторінку з серверним фільтром за датою створення
  private async loadFirstServerFilteredRequestsPage(
    createdDateRange: IRequestCreatedDateRange
  ): Promise<IServiceRequestPage> {
    const filter = this.buildCreatedDateFilter(createdDateRange);
    const page = await this.requestsList.getPage({
      filter,
      orderBy: 'Id desc'
    });

    return {
      requests: page.items,
      cursor: page.nextPageUrl
        ? this.createCursorAfterRequests(page.items)
        : undefined
    };
  }

  // перевіряє чи є помилка перевищення порогу подання списку
  private isListViewThresholdError(error: unknown): boolean {
    if (!(error instanceof Error)) {
      return false;
    }

    const errorMessage = error.message.toLowerCase();

    return errorMessage.includes('граничне значення подання списку')
      || errorMessage.includes('list view threshold')
      || errorMessage.includes('spquerythrottledexception');
  }

  // повертає найбільший ідентифікатор заявки для початку читання списку
  private async getLatestRequestId(): Promise<number | undefined> {
    const latestRequestsPage = await this.requestMetadataList.getPage({
      orderBy: 'Id desc',
      pageSize: 1
    });

    return latestRequestsPage.items[0]?.Id;
  }

  // створює стан першого безпечного діапазону id
  private createInitialRequestPageCursor(
    latestRequestId: number | undefined
  ): IRequestPageCursor | undefined {
    if (!latestRequestId) {
      return undefined;
    }

    const blockLowerId = Math.max(1, latestRequestId - requestIdBlockSize + 1);

    return {
      blockLowerId,
      nextItemUpperId: latestRequestId
    };
  }

  // створює стан після першої серверної сторінки заявок
  private createCursorAfterRequests(
    requests: IServiceRequest[]
  ): IRequestPageCursor | undefined {
    const lastRequestId = requests[requests.length - 1]?.Id;

    return lastRequestId
      ? this.createInitialRequestPageCursor(lastRequestId - 1)
      : undefined;
  }

  // завантажує сторінку заявок з безпечних діапазонів id
  private async loadMatchingRequestsPage(
    createdDateRange: IRequestCreatedDateRange,
    cursor: IRequestPageCursor,
    pageSize: number
  ): Promise<IServiceRequestPage> {
    const requests: IServiceRequest[] = [];
    let currentCursor: IRequestPageCursor | undefined = cursor;

    while (currentCursor && requests.length < pageSize) {
      const remainingRequestsCount = pageSize - requests.length;
      const filter = this.buildRequestBlockFilter(currentCursor);
      const page = await this.requestsList.getPage({
        filter,
        orderBy: 'Id desc',
        pageSize: remainingRequestsCount
      });
      const matchingRequests = page.items.filter(request =>
        this.isRequestCreatedInRange(request, createdDateRange)
      );
      requests.push(...matchingRequests);
      currentCursor = this.createNextRequestPageCursor(currentCursor, page.items);
    }

    return { requests, cursor: currentCursor };
  }

  // створює стан для наступної частини безпечного діапазону id
  private createNextRequestPageCursor(
    cursor: IRequestPageCursor,
    requests: IServiceRequest[]
  ): IRequestPageCursor | undefined {
    const blockLowerId = cursor.blockLowerId;
    const nextItemUpperId = cursor.nextItemUpperId;

    if (blockLowerId === undefined || nextItemUpperId === undefined) {
      return undefined;
    }

    const lastRequestId = requests[requests.length - 1]?.Id;
    const nextRequestUpperId = lastRequestId
      ? lastRequestId - 1
      : blockLowerId - 1;

    if (nextRequestUpperId >= blockLowerId) {
      return {
        blockLowerId,
        nextItemUpperId: nextRequestUpperId
      };
    }

    const nextBlockUpperId = blockLowerId - 1;

    if (nextBlockUpperId <= 0) {
      return undefined;
    }

    return {
      blockLowerId: Math.max(1, nextBlockUpperId - requestIdBlockSize + 1),
      nextItemUpperId: nextBlockUpperId
    };
  }

  // будує умову sharepoint для безпечного діапазону id
  private buildRequestBlockFilter(cursor: IRequestPageCursor): string {
    return `Id ge ${cursor.blockLowerId} and Id le ${cursor.nextItemUpperId}`;
  }

  // підраховує всі заявки вибраного діапазону без серверного фільтра за датою
  public async countRequests(createdDateRange: IRequestCreatedDateRange): Promise<number> {
    const latestRequestId = await this.getLatestRequestId();

    if (!latestRequestId) {
      return 0;
    }

    let requestsCount = 0;

    for (let blockUpperId = latestRequestId; blockUpperId > 0; blockUpperId -= requestIdBlockSize) {
      const blockLowerId = Math.max(1, blockUpperId - requestIdBlockSize + 1);
      const filter = `Id ge ${blockLowerId} and Id le ${blockUpperId}`;
      const page = await this.requestMetadataList.getPage({
        filter,
        orderBy: 'Id desc',
        pageSize: requestIdBlockSize
      });
      const matchingRequestsCount = page.items.filter(request =>
        this.isRequestCreatedInRange(request, createdDateRange)
      ).length;

      requestsCount += matchingRequestsCount;
    }

    return requestsCount;
  }

  // будує умову sharepoint для вибраного діапазону дат створення
  private buildCreatedDateFilter(createdDateRange: IRequestCreatedDateRange): string {
    const { from, to } = createdDateRange;
    const filterConditions: string[] = [];

    if (from) {
      const fromDate = this.getLocalDayStart(from).toISOString();
      filterConditions.push(`Created ge datetime'${fromDate}'`);
    }

    if (to) {
      const toDate = this.getNextLocalDayStart(to).toISOString();
      filterConditions.push(`Created lt datetime'${toDate}'`);
    }

    return filterConditions.join(' and ');
  }

  // перевіряє чи входить дата створення заявки до вибраного діапазону
  private isRequestCreatedInRange(
    request: IRequestDateMetadata,
    createdDateRange: IRequestCreatedDateRange
  ): boolean {
    const createdDate = new Date(request.Created);
    const isCreatedDateInvalid = Number.isNaN(createdDate.getTime());

    if (isCreatedDateInvalid) {
      return false;
    }

    const { from, to } = createdDateRange;
    const fromDate = from ? this.getLocalDayStart(from) : undefined;
    const toDate = to ? this.getNextLocalDayStart(to) : undefined;
    const isBeforeFromDate = fromDate !== undefined && createdDate < fromDate;
    const isOnOrAfterToDate = toDate !== undefined && createdDate >= toDate;

    return !isBeforeFromDate && !isOnOrAfterToDate;
  }

  // повертає початок вибраної дати у локальному часовому поясі
  private getLocalDayStart(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  // повертає початок наступної локальної дати для включення всього дня
  private getNextLocalDayStart(date: Date): Date {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate() + 1);
  }

  // створює нову заявку у списку servicerequests
  public async createRequest(draft: IServiceRequestDraft): Promise<IServiceRequest> {
    const payload = await this.createRequestPayload(draft);

    return this.requestsList.create(payload);
  }

  // оновлює наявну заявку у списку servicerequests
  public async updateRequest(
    itemId: number,
    draft: IServiceRequestDraft,
    editLockToken: string
  ): Promise<IServiceRequest> {
    const lockedRequest = await this.requestsList.getByIdWithETag(itemId);
    this.ensureEditLockOwnership(lockedRequest.item, editLockToken);
    const payload = await this.createRequestPayload(draft);

    try {
      return await this.requestsList.update(itemId, {
        ...payload,
        ...this.createEmptyEditLockPayload()
      }, lockedRequest.eTag);
    } catch (error) {
      const isRequestChangedDuringSave = this.isVersionConflictError(error);
      const requestSaveConflictErrorMessage = 'Заявку було змінено під час збереження Відкрийте її для редагування ще раз';

      if (isRequestChangedDuringSave) {
        throw new Error(requestSaveConflictErrorMessage);
      }

      throw error;
    }
  }

  // захоплює тимчасове блокування заявки для поточної вкладки
  public async acquireEditLock(
    itemId: number,
    currentUserEmail: string
  ): Promise<IRequestEditSession> {
    const normalizedUserEmail = currentUserEmail.trim();
    const isCurrentUserEmailMissing = !normalizedUserEmail;
    const missingCurrentUserEmailErrorMessage = 'Не вдалося визначити поточного користувача для редагування заявки';

    if (isCurrentUserEmailMissing) {
      throw new Error(missingCurrentUserEmailErrorMessage);
    }

    const currentUserId = await this.ensureUser(normalizedUserEmail);

    return this.tryAcquireEditLock(itemId, currentUserId, 0);
  }

  // продовжує блокування відкритої форми редагування
  public async renewEditLock(itemId: number, editLockToken: string): Promise<void> {
    const lockedRequest = await this.requestsList.getByIdWithETag(itemId);
    this.ensureEditLockOwnership(lockedRequest.item, editLockToken);

    await this.updateEditLockExpiry(itemId, lockedRequest.eTag);
  }

  // знімає блокування якщо воно належить поточній вкладці
  public async releaseEditLock(itemId: number, editLockToken: string): Promise<void> {
    const lockedRequest = await this.requestsList.getByIdWithETag(itemId);
    const isLockOwnedByCurrentTab = lockedRequest.item.EditLockToken === editLockToken;

    if (!isLockOwnedByCurrentTab) {
      return;
    }

    try {
      await this.requestsList.update(
        itemId,
        this.createEmptyEditLockPayload(),
        lockedRequest.eTag
      );
    } catch (error) {
      const isLockAlreadyChanged = this.isVersionConflictError(error);

      if (!isLockAlreadyChanged) {
        throw error;
      }
    }
  }

  // видаляє заявку зі списку servicerequests
  public async deleteRequest(itemId: number): Promise<void> {
    await this.requestsList.delete(itemId);
  }

  // намагається захопити блокування та повторює читання після конфлікту версії
  private async tryAcquireEditLock(
    itemId: number,
    currentUserId: number,
    attemptNumber: number
  ): Promise<IRequestEditSession> {
    const requestWithETag = await this.requestsList.getByIdWithETag(itemId);
    const hasActiveEditLock = this.hasActiveEditLock(requestWithETag.item);

    if (hasActiveEditLock) {
      throw new Error(this.getActiveEditLockErrorMessage(requestWithETag.item));
    }

    const editLockToken = this.createEditLockToken();
    const editLockPayload: IRequestEditLockPayload = {
      EditLockOwnerId: currentUserId,
      EditLockExpiresAt: this.getNextEditLockExpiry().toISOString(),
      EditLockToken: editLockToken
    };

    try {
      await this.requestsList.update(itemId, editLockPayload, requestWithETag.eTag);
    } catch (error) {
      const canRetryAfterVersionConflict = attemptNumber === 0 && this.isVersionConflictError(error);

      if (canRetryAfterVersionConflict) {
        return this.tryAcquireEditLock(itemId, currentUserId, attemptNumber + 1);
      }

      throw error;
    }

    const lockedRequest = await this.requestsList.getById(itemId);
    const isLockTokenStored = lockedRequest.EditLockToken === editLockToken;
    const editLockWasNotStoredErrorMessage = 'Не вдалося встановити блокування заявки Спробуйте відкрити її ще раз';

    if (!isLockTokenStored) {
      throw new Error(editLockWasNotStoredErrorMessage);
    }

    return { request: lockedRequest, token: editLockToken };
  }

  // перевіряє що блокування належить поточній вкладці та не завершилося
  private ensureEditLockOwnership(request: IServiceRequest, editLockToken: string): void {
    const isLockOwnedByCurrentTab = request.EditLockToken === editLockToken;
    const hasActiveEditLock = this.hasActiveEditLock(request);
    const isEditLockInvalid = !isLockOwnedByCurrentTab || !hasActiveEditLock;
    const invalidEditLockErrorMessage = 'Блокування заявки завершилося або було змінено іншим користувачем Відкрийте заявку для редагування ще раз';

    if (isEditLockInvalid) {
      throw new Error(invalidEditLockErrorMessage);
    }
  }

  // перевіряє чи має заявка незавершене блокування редагування
  private hasActiveEditLock(request: IServiceRequest): boolean {
    const lockExpiryDate = request.EditLockExpiresAt
      ? new Date(request.EditLockExpiresAt)
      : undefined;
    const isLockExpiryDateValid = lockExpiryDate !== undefined
      && !Number.isNaN(lockExpiryDate.getTime());
    const isLockExpired = !isLockExpiryDateValid || lockExpiryDate <= new Date();
    const hasLockOwner = Boolean(request.EditLockOwner);
    const hasLockToken = Boolean(request.EditLockToken?.trim());

    return hasLockOwner && hasLockToken && !isLockExpired;
  }

  // створює текст повідомлення про активне блокування заявки
  private getActiveEditLockErrorMessage(request: IServiceRequest): string {
    const lockOwnerName = request.EditLockOwner?.Title || 'інший користувач';
    const lockExpiryDate = request.EditLockExpiresAt
      ? new Date(request.EditLockExpiresAt)
      : undefined;
    const lockExpiryText = lockExpiryDate && !Number.isNaN(lockExpiryDate.getTime())
      ? lockExpiryDate.toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' })
      : 'невідомого часу';

    return `Заявку редагує ${lockOwnerName} до ${lockExpiryText}`;
  }

  // повертає дату завершення нового або продовженого блокування
  private getNextEditLockExpiry(): Date {
    const lockExpiryDate = new Date();
    lockExpiryDate.setMinutes(lockExpiryDate.getMinutes() + editLockDurationMinutes);

    return lockExpiryDate;
  }

  // продовжує строк блокування та пояснює конфлікт версії користувачу
  private async updateEditLockExpiry(itemId: number, eTag: string): Promise<void> {
    try {
      await this.requestsList.update(itemId, {
        EditLockExpiresAt: this.getNextEditLockExpiry().toISOString()
      }, eTag);
    } catch (error) {
      const isLockChangedByAnotherRequest = this.isVersionConflictError(error);
      const editLockRenewalErrorMessage = 'Не вдалося продовжити блокування заявки Перевірте стан заявки та відкрийте її повторно';

      if (isLockChangedByAnotherRequest) {
        throw new Error(editLockRenewalErrorMessage);
      }

      throw error;
    }
  }

  // створює значення для очищення всіх полів блокування
  private createEmptyEditLockPayload(): IRequestEditLockPayload {
    return {
      EditLockOwnerId: null,
      EditLockExpiresAt: null,
      EditLockToken: null
    };
  }

  // створює випадковий ключ для однієї вкладки редагування
  private createEditLockToken(): string {
    const randomBytes = new Uint8Array(16);
    crypto.getRandomValues(randomBytes);
    randomBytes[6] = (randomBytes[6] & 0x0f) | 0x40;
    randomBytes[8] = (randomBytes[8] & 0x3f) | 0x80;
    const hexadecimalBytes = Array.from(randomBytes, byte => {
      const hexadecimalByte = byte.toString(16);

      return hexadecimalByte.length === 1
        ? `0${hexadecimalByte}`
        : hexadecimalByte;
    });

    return [
      hexadecimalBytes.slice(0, 4).join(''),
      hexadecimalBytes.slice(4, 6).join(''),
      hexadecimalBytes.slice(6, 8).join(''),
      hexadecimalBytes.slice(8, 10).join(''),
      hexadecimalBytes.slice(10, 16).join('')
    ].join('-');
  }

  // перевіряє чи пов'язана помилка з одночасною зміною елемента
  private isVersionConflictError(error: unknown): boolean {
    return error instanceof Error && error.message.includes('HTTP 412');
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
