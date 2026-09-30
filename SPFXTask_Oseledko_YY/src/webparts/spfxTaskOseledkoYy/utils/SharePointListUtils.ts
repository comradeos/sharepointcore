import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const odataJsonContentType = 'application/json;odata=nometadata';
const defaultPageSize = 5000;
const matchAnyVersion = '*';
const spHttpClientConfiguration = SPHttpClient.configurations.v1;
const readHeaders = {
  Accept: odataJsonContentType
};
const jsonHeaders = {
  Accept: odataJsonContentType,
  'Content-Type': odataJsonContentType
};

// параметри підключення до одного списку sharepoint
export interface ISharePointListOptions {
  listTitle: string;
  selectFields: readonly string[];
  expandFields?: readonly string[];
  pageSize?: number;
}

// додаткові параметри читання колекції елементів
export interface ISharePointListQueryOptions {
  filter?: string;
  orderBy?: string;
  pageSize?: number;
}

// сторінка елементів списку та посилання на наступну сторінку
export interface ISharePointListPage<TItem> {
  items: TItem[];
  nextPageUrl?: string;
}

// елемент списку та його версія для захисту від паралельних змін
export interface ISharePointListItemWithETag<TItem> {
  item: TItem;
  eTag: string;
}

// сторінка результатів sharepoint rest api без службових метаданих
interface IODataPage<TItem> {
  value: TItem[];
  '@odata.nextLink'?: string;
}

// створений елемент який повертає sharepoint rest api
interface ICreatedListItem {
  Id?: number;
  ID?: number;
}

// помилка яку повертає sharepoint rest api
interface ISharePointErrorResponse {
  error?: {
    message?: {
      value?: string;
    } | string;
  };
}

// виконує типові операції з одним списком sharepoint
export default class SharePointListUtils<TItem, TPayload extends object = Record<string, unknown>> {
  private readonly listUrl: string;
  private readonly itemsQuery: string;
  private readonly pageSize: number;

  // зберігає клієнт адресу сайту та структуру списку
  public constructor(
    private readonly client: SPHttpClient,
    webUrl: string,
    private readonly options: ISharePointListOptions
  ) {
    this.validateOptions(webUrl, options);

    const siteUrl = webUrl.trim().replace(/\/+$/, '');
    const listName = options.listTitle.trim().replace(/'/g, "''");

    this.itemsQuery = this.buildItemsQuery(
      options.selectFields,
      options.expandFields ?? []
    );
    this.listUrl = `${siteUrl}/_api/web/lists/getbytitle('${listName}')`;
    this.pageSize = options.pageSize ?? defaultPageSize;
  }

  // читає всі сторінки налаштованого списку
  public async getAll(queryOptions: ISharePointListQueryOptions = {}): Promise<TItem[]> {
    let nextPageUrl: string | undefined;
    const items: TItem[] = [];

    do {
      const page = await this.getPage(queryOptions, nextPageUrl);
      items.push(...page.items);
      nextPageUrl = page.nextPageUrl;
    } while (nextPageUrl);

    return items;
  }

  // читає одну сторінку списку та повертає посилання на наступну сторінку
  public async getPage(
    queryOptions: ISharePointListQueryOptions = {},
    nextPageUrl?: string
  ): Promise<ISharePointListPage<TItem>> {
    const collectionUrl = nextPageUrl ?? this.buildCollectionUrl(queryOptions);
    const response: SPHttpClientResponse = await this.client.get(
      collectionUrl,
      spHttpClientConfiguration,
      { headers: readHeaders }
    );

    await this.ensureSuccessfulResponse(response, 'завантажити список');

    const page = await response.json() as IODataPage<TItem>;

    if (!Array.isArray(page.value)) {
      throw new Error(`Список ${this.options.listTitle} повернув неочікуваний формат даних`);
    }

    return {
      items: page.value,
      nextPageUrl: page['@odata.nextLink']
    };
  }

  // читає один елемент списку за його числовим ідентифікатором
  public async getById(itemId: number): Promise<TItem> {
    const itemWithETag = await this.getByIdWithETag(itemId);

    return itemWithETag.item;
  }

  // читає елемент списку разом з версією для безпечного оновлення
  public async getByIdWithETag(itemId: number): Promise<ISharePointListItemWithETag<TItem>> {
    this.validateItemId(itemId);

    const response = await this.client.get(
      this.getItemUrl(itemId),
      spHttpClientConfiguration,
      { headers: readHeaders }
    );

    await this.ensureSuccessfulResponse(response, 'завантажити елемент списку');

    const item = await response.json() as TItem;

    if (!item) {
      throw new Error(`Елемент списку ${this.options.listTitle} з ідентифікатором ${itemId} не знайдено`);
    }

    const responseETag = response.headers.get('ETag');
    const eTag = responseETag?.trim();
    const isETagMissing = !eTag;
    const missingETagErrorMessage = `SharePoint не повернув версію елемента списку ${this.options.listTitle}`;

    if (isETagMissing) {
      throw new Error(missingETagErrorMessage);
    }

    return { item, eTag };
  }

  // створює елемент та повертає його актуальні дані зі списку
  public async create(payload: TPayload): Promise<TItem> {
    const response = await this.client.post(
      `${this.listUrl}/items`,
      spHttpClientConfiguration,
      {
        headers: jsonHeaders,
        body: JSON.stringify(payload)
      }
    );

    await this.ensureSuccessfulResponse(response, 'створити елемент списку');

    const createdItem = await response.json() as ICreatedListItem;
    const createdItemId = createdItem.Id ?? createdItem.ID;
    const isCreatedItemIdInvalid = typeof createdItemId !== 'number'
      || !Number.isInteger(createdItemId);
    const missingCreatedItemIdErrorMessage = `SharePoint не повернув ідентифікатор створеного елемента списку ${this.options.listTitle}`;

    if (isCreatedItemIdInvalid) {
      throw new Error(missingCreatedItemIdErrorMessage);
    }

    return this.getById(createdItemId);
  }

  // оновлює елемент та повертає його актуальні дані зі списку
  public async update(
    itemId: number,
    payload: TPayload,
    eTag: string = matchAnyVersion
  ): Promise<TItem> {
    this.validateItemId(itemId);
    const isETagEmpty = !eTag.trim();
    const emptyETagErrorMessage = 'Версія елемента не може бути порожньою';

    if (isETagEmpty) {
      throw new Error(emptyETagErrorMessage);
    }

    const response = await this.client.post(
      this.getItemUrl(itemId, false),
      spHttpClientConfiguration,
      {
        headers: {
          ...jsonHeaders,
          'IF-MATCH': eTag,
          'X-HTTP-Method': 'MERGE'
        },
        body: JSON.stringify(payload)
      }
    );

    await this.ensureSuccessfulResponse(response, 'оновити елемент списку');

    return this.getById(itemId);
  }

  // видаляє елемент списку за його числовим ідентифікатором
  public async delete(itemId: number): Promise<void> {
    this.validateItemId(itemId);

    const response = await this.client.post(
      this.getItemUrl(itemId, false),
      spHttpClientConfiguration,
      {
        headers: {
          ...readHeaders,
          'IF-MATCH': matchAnyVersion,
          'X-HTTP-Method': 'DELETE'
        }
      }
    );

    await this.ensureSuccessfulResponse(response, 'видалити елемент списку');
  }

  // перевіряє налаштування списку до виконання першого запиту
  private validateOptions(webUrl: string, options: ISharePointListOptions): void {
    const isWebUrlEmpty = !webUrl.trim();
    const emptyWebUrlErrorMessage = 'Адреса сайту SharePoint не може бути порожньою';

    if (isWebUrlEmpty) {
      throw new Error(emptyWebUrlErrorMessage);
    }

    const isListTitleEmpty = !options.listTitle.trim();
    const emptyListTitleErrorMessage = 'Назва списку SharePoint не може бути порожньою';

    if (isListTitleEmpty) {
      throw new Error(emptyListTitleErrorMessage);
    }

    const hasNoSelectFields = options.selectFields.length === 0;
    const hasEmptySelectField = options.selectFields.some(field => !field.trim());
    const hasInvalidSelectFields = hasNoSelectFields || hasEmptySelectField;
    const invalidSelectFieldsErrorMessage = 'Потрібно вказати щонайменше одне поле для завантаження';

    if (hasInvalidSelectFields) {
      throw new Error(invalidSelectFieldsErrorMessage);
    }

    const hasEmptyExpandField = options.expandFields?.some(field => !field.trim()) ?? false;
    const emptyExpandFieldErrorMessage = 'Поля розгортання не можуть бути порожніми';

    if (hasEmptyExpandField) {
      throw new Error(emptyExpandFieldErrorMessage);
    }

    const hasInvalidPageSize = options.pageSize !== undefined && (
      !Number.isInteger(options.pageSize) || options.pageSize <= 0
    );
    const invalidPageSizeErrorMessage = 'Розмір сторінки має бути додатним цілим числом';

    if (hasInvalidPageSize) {
      throw new Error(invalidPageSizeErrorMessage);
    }
  }

  // перевіряє числовий ідентифікатор елемента списку
  private validateItemId(itemId: number): void {
    const isItemIdInvalid = !Number.isInteger(itemId) || itemId <= 0;
    const invalidItemIdErrorMessage = 'Ідентифікатор елемента має бути додатним цілим числом';

    if (isItemIdInvalid) {
      throw new Error(invalidItemIdErrorMessage);
    }
  }

  // перевіряє успішність відповіді sharepoint
  private async ensureSuccessfulResponse(
    response: SPHttpClientResponse,
    action: string
  ): Promise<void> {
    const hasRequestFailed = !response.ok;

    if (hasRequestFailed) {
      const errorDetails = await this.getResponseErrorDetails(response);
      const requestErrorMessage = errorDetails
        ? `Не вдалося ${action} ${this.options.listTitle} (HTTP ${response.status} ${errorDetails})`
        : `Не вдалося ${action} ${this.options.listTitle} (HTTP ${response.status})`;

      throw new Error(requestErrorMessage);
    }
  }

  // повертає текст помилки зі відповіді sharepoint
  private async getResponseErrorDetails(response: SPHttpClientResponse): Promise<string | undefined> {
    const responseText = await response.text();

    if (!responseText) {
      return undefined;
    }

    try {
      const responseError = JSON.parse(responseText) as ISharePointErrorResponse;
      const errorMessage = responseError.error?.message;

      return typeof errorMessage === 'string'
        ? errorMessage
        : errorMessage?.value;
    } catch {
      return undefined;
    }
  }

  // будує параметри вибору та розгортання полів
  private buildItemsQuery(
    selectFields: readonly string[],
    expandFields: readonly string[]
  ): string {
    const query = [`$select=${selectFields.join(',')}`];

    if (expandFields.length > 0) {
      query.push(`$expand=${expandFields.join(',')}`);
    }

    return query.join('&');
  }

  // будує адресу колекції з фільтром сортуванням та розміром сторінки
  private buildCollectionUrl(queryOptions: ISharePointListQueryOptions): string {
    const query = [this.itemsQuery];
    const filter = queryOptions.filter?.trim();
    const orderBy = queryOptions.orderBy?.trim();

    if (filter) {
      query.push(`$filter=${encodeURIComponent(filter)}`);
    }

    if (orderBy) {
      query.push(`$orderby=${encodeURIComponent(orderBy)}`);
    }

    const pageSize = queryOptions.pageSize ?? this.pageSize;
    query.push(`$top=${pageSize}`);

    return `${this.listUrl}/items?${query.join('&')}`;
  }

  // будує адресу одного елемента зі службовими параметрами за потреби
  private getItemUrl(itemId: number, includeQuery: boolean = true): string {
    const itemUrl = `${this.listUrl}/items(${itemId})`;

    return includeQuery
      ? `${itemUrl}?${this.itemsQuery}`
      : itemUrl;
  }
}

// результат однієї фонової частини завантаження списку
export interface ISharePointBackgroundPage<TItem, TCursor> {
  items: TItem[];
  cursor?: TCursor;
}

// налаштування фонового посторінкового завантаження
export interface ISharePointBackgroundPageLoaderOptions<TItem, TCursor> {
  delayMilliseconds: number;
  pageSize: number;
  loadPage: (cursor: TCursor, pageSize: number) => Promise<ISharePointBackgroundPage<TItem, TCursor>>;
  onPageLoaded: (page: ISharePointBackgroundPage<TItem, TCursor>) => void;
  onError: (error: unknown) => void;
  onCompleted?: () => void;
}

// сесія тимчасового блокування одного елемента списку
export interface ISharePointListEditSession<TItem> {
  item: TItem;
  token: string;
}

// значення полів тимчасового блокування
export interface ISharePointListEditLockValues {
  ownerId?: number;
  expiresAt?: string;
  token?: string;
}

// налаштування тимчасового блокування для конкретного типу елементів списку
export interface ISharePointListEditLockOptions<TItem, TPayload extends object> {
  createPayload: (values: ISharePointListEditLockValues) => TPayload;
  getOwnerId: (item: TItem) => number | undefined;
  getOwnerName: (item: TItem) => string | undefined;
  getExpiresAt: (item: TItem) => string | undefined;
  getToken: (item: TItem) => string | undefined;
  itemName?: string;
  itemNameGenitive?: string;
  durationMilliseconds?: number;
}

const defaultBackgroundPageDelayMilliseconds = 3000;
const defaultEditLockDurationMilliseconds = 5 * 60 * 1000;

// завантажує сторінки у фоні та відкидає відповіді скасованих запусків
export class SharePointBackgroundPageLoader<TItem, TCursor> {
  private timeout?: ReturnType<typeof setTimeout>;
  private loadingVersion = 0;

  public constructor(
    private readonly options: ISharePointBackgroundPageLoaderOptions<TItem, TCursor>
  ) {
    this.validateOptions();
  }

  // запускає завантаження з курсора та повертає версію запуску
  public start(cursor: TCursor): number {
    this.cancel();
    const loadingVersion = this.loadingVersion;

    this.schedule(cursor, loadingVersion);

    return loadingVersion;
  }

  // скасовує очікування та забороняє застосування попередніх відповідей
  public cancel(): void {
    this.loadingVersion += 1;

    if (this.timeout) {
      clearTimeout(this.timeout);
      this.timeout = undefined;
    }
  }

  // повертає актуальну версію запуску для пов'язаних асинхронних операцій
  public getVersion(): number {
    return this.loadingVersion;
  }

  // планує читання наступної сторінки після заданої затримки
  private schedule(cursor: TCursor, loadingVersion: number): void {
    this.timeout = setTimeout(() => {
      this.loadNextPage(cursor, loadingVersion).catch(error => {
        if (loadingVersion === this.loadingVersion) {
          this.options.onError(error);
        }
      });
    }, this.options.delayMilliseconds);
  }

  // читає сторінку та планує наступну якщо курсор ще наявний
  private async loadNextPage(cursor: TCursor, loadingVersion: number): Promise<void> {
    const page = await this.options.loadPage(cursor, this.options.pageSize);
    const isLoadingCancelled = loadingVersion !== this.loadingVersion;

    if (isLoadingCancelled) {
      return;
    }

    this.options.onPageLoaded(page);

    if (page.cursor !== undefined) {
      this.schedule(page.cursor, loadingVersion);
      return;
    }

    this.options.onCompleted?.();
  }

  // перевіряє параметри перед першим фоновим запитом
  private validateOptions(): void {
    const delayMilliseconds = this.options.delayMilliseconds ?? defaultBackgroundPageDelayMilliseconds;
    const isDelayInvalid = !Number.isFinite(delayMilliseconds) || delayMilliseconds < 0;

    if (isDelayInvalid) {
      throw new Error("Затримка фонового завантаження має бути невід'ємним числом");
    }

    const isPageSizeInvalid = !Number.isInteger(this.options.pageSize) || this.options.pageSize <= 0;

    if (isPageSizeInvalid) {
      throw new Error('Розмір фонової сторінки має бути додатним цілим числом');
    }
  }
}

// налаштування автоматичного продовження блокування редагування
export interface ISharePointEditLockRenewalOptions {
  intervalMilliseconds: number;
  renew: () => Promise<void>;
  onRenewed?: () => void;
  onError: (error: unknown) => void;
}

// періодично продовжує блокування та припиняється після помилки
export class SharePointEditLockRenewal {
  private timer?: ReturnType<typeof setInterval>;
  private renewalVersion = 0;
  private isRenewalInProgress = false;

  public constructor(
    private readonly options: ISharePointEditLockRenewalOptions
  ) {
    this.validateOptions();
  }

  // запускає періодичне продовження блокування
  public start(): void {
    this.stop();
    this.timer = setInterval(this.renew, this.options.intervalMilliseconds);
  }

  // зупиняє продовження та відкидає результат поточного запиту
  public stop(): void {
    this.renewalVersion += 1;

    if (this.timer) {
      clearInterval(this.timer);
      this.timer = undefined;
    }
  }

  // повідомляє чи виконується запит продовження у цей момент
  public getIsRenewalInProgress(): boolean {
    return this.isRenewalInProgress;
  }

  // виконує один запит продовження блокування
  private readonly renew = async (): Promise<void> => {
    if (this.isRenewalInProgress) {
      return;
    }

    this.isRenewalInProgress = true;
    const renewalVersion = this.renewalVersion;

    try {
      await this.options.renew();

      if (renewalVersion === this.renewalVersion) {
        this.options.onRenewed?.();
      }
    } catch (error) {
      if (renewalVersion === this.renewalVersion) {
        this.stop();
        this.options.onError(error);
      }
    } finally {
      this.isRenewalInProgress = false;
    }
  };

  // перевіряє параметри перед першим продовженням
  private validateOptions(): void {
    const isIntervalInvalid = !Number.isFinite(this.options.intervalMilliseconds)
      || this.options.intervalMilliseconds <= 0;

    if (isIntervalInvalid) {
      throw new Error('Інтервал продовження блокування має бути додатним числом');
    }
  }
}

// керує тимчасовим блокуванням елемента та перевіркою його версії
export class SharePointListEditLockUtils<TItem, TPayload extends object> {
  private readonly itemName: string;
  private readonly itemNameGenitive: string;
  private readonly durationMilliseconds: number;

  public constructor(
    private readonly list: SharePointListUtils<TItem, TPayload>,
    private readonly options: ISharePointListEditLockOptions<TItem, TPayload>
  ) {
    this.itemName = options.itemName?.trim() || 'елемента списку';
    this.itemNameGenitive = options.itemNameGenitive?.trim() || this.itemName;
    this.durationMilliseconds = options.durationMilliseconds ?? defaultEditLockDurationMilliseconds;
    this.validateOptions();
  }

  // захоплює блокування для користувача та повертає сесію вкладки
  public async acquire(itemId: number, ownerId: number): Promise<ISharePointListEditSession<TItem>> {
    this.validateOwnerId(ownerId);

    return this.tryAcquire(itemId, ownerId, 0);
  }

  // продовжує блокування якщо воно ще належить поточній вкладці
  public async renew(itemId: number, token: string): Promise<void> {
    const itemWithETag = await this.list.getByIdWithETag(itemId);
    this.ensureOwnership(itemWithETag.item, token);

    try {
      await this.list.update(
        itemId,
        this.options.createPayload({
          ownerId: this.options.getOwnerId(itemWithETag.item),
          expiresAt: this.getNextExpiry().toISOString(),
          token
        }),
        itemWithETag.eTag
      );
    } catch (error) {
      if (isSharePointListVersionConflictError(error)) {
        throw new Error(`Не вдалося продовжити блокування ${this.itemNameGenitive} Перевірте стан елемента та відкрийте його повторно`);
      }

      throw error;
    }
  }

  // знімає блокування якщо воно все ще належить поточній вкладці
  public async release(itemId: number, token: string): Promise<void> {
    const itemWithETag = await this.list.getByIdWithETag(itemId);
    const isLockOwnedByCurrentTab = this.options.getToken(itemWithETag.item) === token;

    if (!isLockOwnedByCurrentTab) {
      return;
    }

    try {
      await this.list.update(itemId, this.createEmptyPayload(), itemWithETag.eTag);
    } catch (error) {
      if (!isSharePointListVersionConflictError(error)) {
        throw error;
      }
    }
  }

  // оновлює елемент якщо чинне блокування належить поточній вкладці
  public async updateLocked(itemId: number, token: string, payload: TPayload): Promise<TItem> {
    const itemWithETag = await this.list.getByIdWithETag(itemId);
    this.ensureOwnership(itemWithETag.item, token);

    try {
      return await this.list.update(itemId, {
        ...payload,
        ...this.createEmptyPayload()
      }, itemWithETag.eTag);
    } catch (error) {
      if (isSharePointListVersionConflictError(error)) {
        throw new Error(`${this.capitalizeItemName()} було змінено під час збереження Відкрийте його для редагування ще раз`);
      }

      throw error;
    }
  }

  // намагається створити блокування та повторює читання після конфлікту версії
  private async tryAcquire(
    itemId: number,
    ownerId: number,
    attemptNumber: number
  ): Promise<ISharePointListEditSession<TItem>> {
    const itemWithETag = await this.list.getByIdWithETag(itemId);

    if (this.hasActiveLock(itemWithETag.item)) {
      throw new Error(this.getActiveLockErrorMessage(itemWithETag.item));
    }

    const token = this.createToken();

    try {
      await this.list.update(itemId, this.options.createPayload({
        ownerId,
        expiresAt: this.getNextExpiry().toISOString(),
        token
      }), itemWithETag.eTag);
    } catch (error) {
      const canRetryAfterVersionConflict = attemptNumber === 0
        && isSharePointListVersionConflictError(error);

      if (canRetryAfterVersionConflict) {
        return this.tryAcquire(itemId, ownerId, attemptNumber + 1);
      }

      throw error;
    }

    const lockedItem = await this.list.getById(itemId);
    const isLockTokenStored = this.options.getToken(lockedItem) === token;

    if (!isLockTokenStored) {
      throw new Error(`Не вдалося встановити блокування ${this.itemName} Спробуйте відкрити його ще раз`);
    }

    return { item: lockedItem, token };
  }

  // перевіряє що блокування належить поточній вкладці та не завершилося
  private ensureOwnership(item: TItem, token: string): void {
    const isLockOwnedByCurrentTab = this.options.getToken(item) === token;
    const isLockInvalid = !isLockOwnedByCurrentTab || !this.hasActiveLock(item);

    if (isLockInvalid) {
      throw new Error(`Блокування ${this.itemNameGenitive} завершилося або було змінено іншим користувачем Відкрийте елемент для редагування ще раз`);
    }
  }

  // перевіряє чи має елемент чинне тимчасове блокування
  private hasActiveLock(item: TItem): boolean {
    const expiresAt = this.options.getExpiresAt(item);
    const expiryDate = expiresAt ? new Date(expiresAt) : undefined;
    const isExpiryDateValid = expiryDate !== undefined && !Number.isNaN(expiryDate.getTime());
    const hasLockOwner = Boolean(this.options.getOwnerId(item));
    const hasLockToken = Boolean(this.options.getToken(item)?.trim());

    return hasLockOwner && hasLockToken && isExpiryDateValid && expiryDate > new Date();
  }

  // формує повідомлення про блокування іншим користувачем
  private getActiveLockErrorMessage(item: TItem): string {
    const ownerName = this.options.getOwnerName(item)?.trim() || 'інший користувач';
    const expiresAt = this.options.getExpiresAt(item);
    const expiryDate = expiresAt ? new Date(expiresAt) : undefined;
    const expiryText = expiryDate && !Number.isNaN(expiryDate.getTime())
      ? expiryDate.toLocaleTimeString('uk-UA', { hour: '2-digit', minute: '2-digit' })
      : 'невідомого часу';

    return `${this.capitalizeItemName()} редагує ${ownerName} до ${expiryText}`;
  }

  // починає назву елемента з великої літери для повідомлення користувачу
  private capitalizeItemName(): string {
    return `${this.itemName.charAt(0).toUpperCase()}${this.itemName.slice(1)}`;
  }

  // повертає значення для очищення полів блокування
  private createEmptyPayload(): TPayload {
    return this.options.createPayload({});
  }

  // повертає момент завершення нового або продовженого блокування
  private getNextExpiry(): Date {
    return new Date(Date.now() + this.durationMilliseconds);
  }

  // генерує випадковий ключ для однієї вкладки редагування
  private createToken(): string {
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

  // перевіряє налаштування блокування до першого запиту
  private validateOptions(): void {
    const isDurationInvalid = !Number.isFinite(this.durationMilliseconds)
      || this.durationMilliseconds <= 0;

    if (isDurationInvalid) {
      throw new Error('Тривалість блокування має бути додатним числом');
    }
  }

  // перевіряє ідентифікатор користувача для встановлення блокування
  private validateOwnerId(ownerId: number): void {
    const isOwnerIdInvalid = !Number.isInteger(ownerId) || ownerId <= 0;

    if (isOwnerIdInvalid) {
      throw new Error('Ідентифікатор користувача для блокування має бути додатним цілим числом');
    }
  }
}

// перевіряє чи пов'язана помилка з одночасною зміною елемента
export function isSharePointListVersionConflictError(error: unknown): boolean {
  return error instanceof Error && error.message.includes('HTTP 412');
}

// додає нові елементи на початок локального списку
export function addListItems<TItem>(items: TItem[], newItems: TItem[]): TItem[] {
  return [...newItems, ...items];
}

// замінює один елемент локального списку за вказаним ключем
export function replaceListItem<TItem, TKey extends keyof TItem>(
  items: TItem[],
  updatedItem: TItem,
  key: TKey
): TItem[] {
  return items.map(item =>
    item[key] === updatedItem[key]
      ? updatedItem
      : item
  );
}

// видаляє один елемент локального списку за вказаним ключем
export function removeListItem<TItem, TKey extends keyof TItem>(
  items: TItem[],
  itemKey: TItem[TKey],
  key: TKey
): TItem[] {
  return items.filter(item => item[key] !== itemKey);
}
