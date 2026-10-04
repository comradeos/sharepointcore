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

const defaultBackgroundPageDelayMilliseconds = 3000;

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
