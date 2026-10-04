import SharePointListUtils from './SharePointListClient';

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

const defaultEditLockDurationMilliseconds = 5 * 60 * 1000;

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

  // видаляє елемент лише якщо його не редагує жодна вкладка
  public async deleteUnlocked(itemId: number): Promise<void> {
    const itemWithETag = await this.list.getByIdWithETag(itemId);

    if (this.hasActiveLock(itemWithETag.item)) {
      throw new Error(this.getActiveLockErrorMessage(itemWithETag.item));
    }

    try {
      await this.list.delete(itemId, itemWithETag.eTag);
    } catch (error) {
      if (isSharePointListVersionConflictError(error)) {
        const latestItem = await this.list.getById(itemId);

        if (this.hasActiveLock(latestItem)) {
          throw new Error(this.getActiveLockErrorMessage(latestItem));
        }

        throw new Error(`${this.capitalizeItemName()} було змінено перед видаленням Спробуйте виконати дію ще раз`);
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
