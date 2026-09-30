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
