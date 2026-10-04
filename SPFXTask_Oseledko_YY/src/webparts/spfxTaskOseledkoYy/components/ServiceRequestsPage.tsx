import * as React from 'react';
import packageInfo from '../../../../package.json';
import {
  MessageBar,
  MessageBarType,
  Spinner,
  Text
} from '@fluentui/react';
import styles from './ServiceRequestsPage.module.scss';
import { IServiceRequestsPageProps } from './IServiceRequestsPageProps';
import {
  IRequestCreatedDateRange,
  IRequestPageCursor,
  IServiceDeskDictionaries,
  IServiceRequest,
  IServiceRequestDraft,
  IServiceRequestPage
} from '../models/ServiceDeskModels';
import ServiceDeskService from '../services/ServiceDeskService';
import {
  addListItems,
  ISharePointBackgroundPage,
  removeListItem,
  replaceListItem,
  SharePointBackgroundPageLoader,
  SharePointEditLockRenewal
} from '../utils/SharePointListUtils';
import ServiceRequestsGrid from './ServiceRequestsGrid';
import ServiceRequestForm from './ServiceRequestForm';
import ServiceRequestView from './ServiceRequestView';
import ServiceRequestDeleteDialog from './ServiceRequestDeleteDialog';
import ServiceRequestsPageHeader from './ServiceRequestsPageHeader';
import { IServiceRequestsPageState } from './ServiceRequestsPage.types';

const backgroundRequestsPageSize = 4000;
const backgroundRequestsDelay = 150;
const editLockRenewalInterval = 60 * 1000;
const visibleRequestsRefreshInterval = 30 * 1000;

// показує таблицю заявок із даними трьох списків sharepoint
export default class ServiceRequestsPage extends React.Component<
  IServiceRequestsPageProps,
  IServiceRequestsPageState
> {
  private readonly service: ServiceDeskService;
  private readonly backgroundRequestsLoader: SharePointBackgroundPageLoader<
    IServiceRequest,
    IRequestPageCursor
  >;
  private editLockRenewal?: SharePointEditLockRenewal;
  private visibleRequestIds: number[] = [];
  private visibleRequestsRefreshTimer?: ReturnType<typeof setInterval>;
  private isVisibleRequestsRefreshInProgress = false;
  private isComponentMounted = false;
  private latestKnownRequestId?: number;

  // створює сервіс для сайту вебчастини та початковий стан екрана
  public constructor(props: IServiceRequestsPageProps) {
    super(props);
    this.service = new ServiceDeskService(props.spHttpClient, props.webUrl);
    this.backgroundRequestsLoader = new SharePointBackgroundPageLoader({
      delayMilliseconds: backgroundRequestsDelay,
      pageSize: backgroundRequestsPageSize,
      loadPage: async (cursor, pageSize) => {
        const page = await this.service.loadRequestsPage(
          { from: this.state.createdFrom, to: this.state.createdTo },
          cursor,
          pageSize
        );

        return { items: page.requests, cursor: page.cursor };
      },
      onPageLoaded: this.handleBackgroundRequestsPageLoadSuccess,
      onError: this.handleBackgroundRequestsPageLoadError,
      onCompleted: this.handleBackgroundRequestsLoadComplete
    });
    const currentDate = this.getCurrentDate();

    this.state = {
      createdFrom: currentDate,
      createdTo: currentDate,
      hasLoadedRequests: false,
      isLoading: false,
      isBackgroundRequestsLoading: false,
      isCountingRequests: false,
      isCreateOpen: false,
      isAcquiringEditLock: false,
      isEditLockValid: false,
      isDeleting: false
    };
  }

  // завантажує довідники після появи вебчастини на сторінці
  public componentDidMount(): void {
    this.isComponentMounted = true;
    this.loadDictionaries();
    this.startVisibleRequestsRefresh();
  }

  // зупиняє фонові запити перед видаленням вебчастини зі сторінки
  public componentWillUnmount(): void {
    this.isComponentMounted = false;
    this.backgroundRequestsLoader.cancel();
    this.stopVisibleRequestsRefresh();
    this.stopEditLockRenewal();
    this.releaseActiveEditLock();
  }

  // завантажує довідники без читання великого списку заявок
  private loadDictionaries(): void {
    this.setState({ isLoading: true, error: undefined });
    this.service.loadDictionaries()
      .then(this.handleDictionariesLoadSuccess)
      .catch(this.handleLoadError);
  }

  // зберігає довідники та створює порожній стан списку заявок
  private readonly handleDictionariesLoadSuccess = (
    dictionaries: IServiceDeskDictionaries
  ): void => {
    this.setState({
      data: {
        ...dictionaries,
        requests: []
      },
      isLoading: false
    });
  };

  // завантажує першу сторінку заявок за вибраним діапазоном дат створення
  private loadRequests(createdDateRange: IRequestCreatedDateRange): void {
    this.setState({
      isLoading: true,
      isBackgroundRequestsLoading: false,
      isCountingRequests: false,
      totalRequests: undefined,
      error: undefined
    });
    this.service.loadRequestsPage(createdDateRange)
      .then(this.handleRequestsLoadSuccess)
      .catch(this.handleLoadError);
  }

  // зберігає першу сторінку заявок та запускає фонове завантаження
  private readonly handleRequestsLoadSuccess = (page: IServiceRequestPage): void => {
    this.latestKnownRequestId = undefined;
    this.initializeLatestKnownRequestId();

    this.setState(previousState => ({
      data: previousState.data
        ? {
          ...previousState.data,
          requests: page.requests
        }
        : previousState.data,
      hasLoadedRequests: true,
      isLoading: false,
      isBackgroundRequestsLoading: Boolean(page.cursor),
      isCountingRequests: Boolean(page.cursor)
    }), () => {
      const backgroundLoadingVersion = page.cursor
        ? this.backgroundRequestsLoader.start(page.cursor)
        : this.backgroundRequestsLoader.getVersion();

      this.startRequestsCount(backgroundLoadingVersion);
    });
  };

  // додає наступну частину заявок до вже завантажених рядків
  private readonly handleBackgroundRequestsPageLoadSuccess = (
    page: ISharePointBackgroundPage<IServiceRequest, IRequestPageCursor>
  ): void => {
    this.setState(previousState => ({
      data: previousState.data
        ? {
          ...previousState.data,
          requests: [...previousState.data.requests, ...page.items]
        }
        : previousState.data,
      error: undefined
    }), this.completeBackgroundLoadingIfAllRequestsLoaded);
  };

  // показує помилку фонової частини заявок та зупиняє завантаження
  private readonly handleBackgroundRequestsPageLoadError = (error: unknown): void => {
    this.setState({
      isBackgroundRequestsLoading: false,
      isCountingRequests: false,
      error: error instanceof Error ? error.message : 'Невідома помилка завантаження.'
    });
  };

  // завершує стан фонового завантаження після останньої сторінки
  private readonly handleBackgroundRequestsLoadComplete = (): void => {
    this.setState({ isBackgroundRequestsLoading: false });
  };

  // зберігає рядки поточної сторінки таблиці для періодичного оновлення
  private readonly handleVisibleRequestIdsChange = (requestIds: number[]): void => {
    this.visibleRequestIds = Array.from(new Set(requestIds));
  };

  // запускає періодичне оновлення заявок які користувач бачить у таблиці
  private startVisibleRequestsRefresh(): void {
    this.stopVisibleRequestsRefresh();
    this.visibleRequestsRefreshTimer = setInterval(() => {
      this.refreshVisibleRequests().catch(() => undefined);
    }, visibleRequestsRefreshInterval);
  }

  // зупиняє періодичне оновлення перед видаленням вебчастини
  private stopVisibleRequestsRefresh(): void {
    if (this.visibleRequestsRefreshTimer) {
      clearInterval(this.visibleRequestsRefreshTimer);
      this.visibleRequestsRefreshTimer = undefined;
    }
  }

  // перечитує лише заявки поточної сторінки без перезапуску основного завантаження
  private readonly refreshVisibleRequests = async (): Promise<void> => {
    const { createdFrom, createdTo, data, hasLoadedRequests, isLoading } = this.state;
    const requestIds = this.visibleRequestIds;
    const latestKnownRequestId = this.latestKnownRequestId;

    if (
      this.isVisibleRequestsRefreshInProgress
      || !hasLoadedRequests
      || isLoading
      || !data
      || (requestIds.length === 0 && latestKnownRequestId === undefined)
    ) {
      return;
    }

    this.isVisibleRequestsRefreshInProgress = true;

    try {
      const [refreshedRequests, newRequestsResult] = await Promise.all([
        requestIds.length > 0
          ? this.service.loadRequestsByIds(requestIds)
          : Promise.resolve([]),
        latestKnownRequestId !== undefined
          ? this.service.loadNewRequestsAfterId(
            { from: createdFrom, to: createdTo },
            latestKnownRequestId
          )
          : Promise.resolve({
            requests: [] as IServiceRequest[],
            latestRequestId: undefined as number | undefined
          })
      ]);

      const isDateRangeChanged = this.state.createdFrom.getTime() !== createdFrom.getTime()
        || this.state.createdTo.getTime() !== createdTo.getTime();

      if (!this.isComponentMounted || isDateRangeChanged) {
        return;
      }

      if (newRequestsResult.latestRequestId) {
        this.latestKnownRequestId = newRequestsResult.latestRequestId;
      }

      const refreshedRequestsById = new Map(
        refreshedRequests.map(request => [request.Id, request])
      );
      const refreshedRequestIds = new Set(requestIds);

      this.setState(previousState => {
        if (!previousState.data) {
          return { data: previousState.data };
        }

        const existingRequestIds = new Set(
          previousState.data.requests.map(request => request.Id)
        );
        const newRequests = newRequestsResult.requests.filter(request =>
          !existingRequestIds.has(request.Id)
        );

        return {
          data: {
            ...previousState.data,
            requests: [
              ...newRequests,
              ...previousState.data.requests
                .filter(request =>
                  !refreshedRequestIds.has(request.Id) || refreshedRequestsById.has(request.Id)
                )
                .map(request => refreshedRequestsById.get(request.Id) ?? request)
            ]
          },
          totalRequests: previousState.totalRequests === undefined
            ? undefined
            : previousState.totalRequests + newRequests.length
        };
      });
    } catch {
      // ігнорує тимчасову помилку щоб не переривати роботу користувача
    } finally {
      this.isVisibleRequestsRefreshInProgress = false;
    }
  };

  // визначає початковий ідентифікатор для відстеження нових заявок
  private initializeLatestKnownRequestId(): void {
    this.service.getLatestRequestId()
      .then(latestRequestId => {
        this.latestKnownRequestId = latestRequestId;
      })
      .catch(() => undefined);
  }

  // запускає підрахунок загальної кількості заявок вибраного діапазону
  private startRequestsCount(backgroundLoadingVersion: number): void {
    const { createdFrom, createdTo, isBackgroundRequestsLoading } = this.state;

    if (!isBackgroundRequestsLoading) {
      this.setState(previousState => ({
        totalRequests: previousState.data?.requests.length ?? 0,
        isCountingRequests: false
      }));
      return;
    }

    this.service.countRequests({ from: createdFrom, to: createdTo })
      .then(totalRequests => this.handleRequestsCountSuccess(
        totalRequests,
        backgroundLoadingVersion
      ))
      .catch(() => this.handleRequestsCountError(backgroundLoadingVersion));
  }

  // зберігає обчислену загальну кількість заявок
  private readonly handleRequestsCountSuccess = (
    totalRequests: number,
    backgroundRequestsVersion: number
  ): void => {
    if (backgroundRequestsVersion === this.backgroundRequestsLoader.getVersion()) {
      this.setState(
        { totalRequests, isCountingRequests: false },
        this.completeBackgroundLoadingIfAllRequestsLoaded
      );
    }
  };

  // зупиняє фонове читання коли кількість завантажених заявок збігається з загальною
  private readonly completeBackgroundLoadingIfAllRequestsLoaded = (): void => {
    const { data, totalRequests, isBackgroundRequestsLoading } = this.state;
    const loadedRequestsCount = data?.requests.length ?? 0;
    const areAllRequestsLoaded = totalRequests !== undefined
      && loadedRequestsCount >= totalRequests;

    if (isBackgroundRequestsLoading && areAllRequestsLoaded) {
      this.backgroundRequestsLoader.cancel();
      this.setState({ isBackgroundRequestsLoading: false });
    }
  };

  // завершує підрахунок якщо не вдалося визначити загальну кількість
  private readonly handleRequestsCountError = (backgroundRequestsVersion: number): void => {
    if (backgroundRequestsVersion === this.backgroundRequestsLoader.getVersion()) {
      this.setState({ isCountingRequests: false });
    }
  };

  // скасовує фонове завантаження та залишає в таблиці вже отримані заявки
  private readonly handleCancelRequestsLoading = (): void => {
    this.backgroundRequestsLoader.cancel();

    this.setState(previousState => ({
      isBackgroundRequestsLoading: false,
      isCountingRequests: false,
      totalRequests: previousState.data?.requests.length
    }));
  };

  // показує повідомлення якщо один із запитів завершився помилкою
  private readonly handleLoadError = (error: unknown): void => {
    this.setState({
      isLoading: false,
      error: error instanceof Error ? error.message : 'Невідома помилка завантаження.'
    });
  };

  // перевіряє межі дат та завантажує список заявок
  private readonly handleRefresh = (): void => {
    const { createdFrom, createdTo } = this.state;
    const isCreatedDateRangeInvalid = createdFrom > createdTo;

    if (isCreatedDateRangeInvalid) {
      this.setState({
        error: 'Дата створення від не може бути пізніше за дату створення до'
      });
      return;
    }

    this.backgroundRequestsLoader.cancel();
    this.setState({ success: undefined });
    this.loadRequests({ from: createdFrom, to: createdTo });
  };

  // зберігає нижню межу дат та очищає завантажені заявки
  private readonly handleCreatedFromChange = (date: Date | null | undefined): void => {
    if (date) {
      this.backgroundRequestsLoader.cancel();
      this.latestKnownRequestId = undefined;
      this.setState(previousState => ({
        createdFrom: date,
        data: previousState.data
          ? { ...previousState.data, requests: [] }
          : previousState.data,
        hasLoadedRequests: false,
        isBackgroundRequestsLoading: false,
        isCountingRequests: false,
        totalRequests: undefined,
        error: undefined
      }));
    }
  };

  // зберігає верхню межу дат та очищає завантажені заявки
  private readonly handleCreatedToChange = (date: Date | null | undefined): void => {
    if (date) {
      this.backgroundRequestsLoader.cancel();
      this.latestKnownRequestId = undefined;
      this.setState(previousState => ({
        createdTo: date,
        data: previousState.data
          ? { ...previousState.data, requests: [] }
          : previousState.data,
        hasLoadedRequests: false,
        isBackgroundRequestsLoading: false,
        isCountingRequests: false,
        totalRequests: undefined,
        error: undefined
      }));
    }
  };

  // повертає поточну дату без часу для початкового діапазону
  private getCurrentDate(): Date {
    const currentDate = new Date();

    return new Date(
      currentDate.getFullYear(),
      currentDate.getMonth(),
      currentDate.getDate()
    );
  }

  // форматує дату для відображення у полі fluent ui
  private readonly formatDate = (date?: Date): string => {
    return date ? date.toLocaleDateString('uk-UA') : '';
  };

  // форматує стан фонового завантаження заявок для відображення користувачу
  private getRequestsLoadingText(
    loadedRequestsCount: number,
    totalRequests: number | undefined,
    isCountingRequests: boolean
  ): string {
    if (totalRequests === undefined) {
      return isCountingRequests
        ? `Завантажено заявок: ${loadedRequestsCount} ... визначаємо загальну кількість`
        : `Завантажено заявок: ${loadedRequestsCount}`;
    }

    const loadingProgress = totalRequests === 0
      ? 100
      : Math.min(100, Math.round((loadedRequestsCount / totalRequests) * 100));

    return `Завантажено заявок: ${loadedRequestsCount} із ${totalRequests} (${loadingProgress}%)`;
  }

  // відкриває форму створення нової заявки
  private readonly handleOpenCreate = (): void => {
    this.setState({
      isCreateOpen: true,
      selectedRequest: undefined,
      editingRequest: undefined,
      deletingRequest: undefined,
      success: undefined
    });
  };

  // закриває форму створення заявки
  private readonly handleCloseCreate = (): void => {
    this.setState({ isCreateOpen: false });
  };

  // створює заявку та додає її до локального списку
  private readonly handleCreateRequest = async (draft: IServiceRequestDraft): Promise<void> => {
    const createdRequest = await this.service.createRequest(draft);
    this.latestKnownRequestId = Math.max(this.latestKnownRequestId ?? 0, createdRequest.Id);

    this.setState(previousState => ({
      data: previousState.data
        ? {
          ...previousState.data,
          requests: addListItems(previousState.data.requests, [createdRequest])
        }
        : previousState.data,
      success: 'Заявку успішно створено',
      hasLoadedRequests: true,
      error: undefined
    }));
  };

  // створює тестові заявки групами та додає їх до локального списку
  private readonly handleGenerateRequest = async (drafts: IServiceRequestDraft[]): Promise<void> => {
    try {
      const batchSize = 10;
      const createdRequests: IServiceRequest[] = [];

      for (let startIndex = 0; startIndex < drafts.length; startIndex += batchSize) {
        const createRequestPromises: Array<Promise<IServiceRequest>> = [];
        const batchEnd = Math.min(startIndex + batchSize, drafts.length);

        for (let draftIndex = startIndex; draftIndex < batchEnd; draftIndex += 1) {
          createRequestPromises.push(this.service.createRequest(drafts[draftIndex]));
        }

        const createdBatch = await Promise.all(createRequestPromises);
        createdRequests.push(...createdBatch);
      }

      if (createdRequests.length > 0) {
        const latestCreatedRequestId = Math.max(...createdRequests.map(request => request.Id));
        this.latestKnownRequestId = Math.max(
          this.latestKnownRequestId ?? 0,
          latestCreatedRequestId
        );
      }

      this.setState(previousState => ({
        data: previousState.data
          ? {
            ...previousState.data,
            requests: addListItems(
              previousState.data.requests,
              createdRequests
            )
          }
          : previousState.data,
        success: `Створено заявок: ${drafts.length}`,
        hasLoadedRequests: true,
        error: undefined
      }));
    } catch (error) {
      this.setState({
        success: undefined,
        error: error instanceof Error
          ? error.message
          : 'Не вдалося згенерувати заявку'
      });
    }
  };

  // відкриває вибрану заявку у режимі перегляду
  private readonly handleOpenView = (request: IServiceRequest): void => {
    this.setState({
      selectedRequest: request,
      editingRequest: undefined,
      deletingRequest: undefined,
      success: undefined
    });
  };

  // закриває вікно перегляду заявки
  private readonly handleCloseView = (): void => {
    this.setState({ selectedRequest: undefined });
  };

  // захоплює блокування та відкриває вибрану заявку у режимі редагування
  private readonly handleOpenEdit = async (request: IServiceRequest): Promise<void> => {
    if (this.state.isAcquiringEditLock) {
      return;
    }

    this.setState({
      isCreateOpen: false,
      selectedRequest: undefined,
      isAcquiringEditLock: true,
      isEditLockValid: false,
      deletingRequest: undefined,
      success: undefined,
      error: undefined
    });

    try {
      const editSession = await this.service.acquireEditLock(
        request.Id,
        this.props.currentUserEmail
      );

      this.setState({
        editingRequest: editSession.item,
        editSession,
        isAcquiringEditLock: false,
        isEditLockValid: true
      }, this.startEditLockRenewal);
    } catch (error) {
      this.setState({
        isAcquiringEditLock: false,
        isEditLockValid: false,
        error: error instanceof Error
          ? error.message
          : 'Не вдалося заблокувати заявку для редагування'
      });
    }
  };

  // закриває форму редагування та звільняє блокування заявки
  private readonly handleCloseEdit = (): void => {
    this.releaseActiveEditLock();
    this.stopEditLockRenewal();
    this.setState({
      editingRequest: undefined,
      editSession: undefined,
      isEditLockValid: false
    });
  };

  // оновлює заявку та замінює її у локальному списку
  private readonly handleUpdateRequest = async (draft: IServiceRequestDraft): Promise<void> => {
    const { editingRequest, editSession, isEditLockValid } = this.state;

    const isEditSessionMissing = !editingRequest || !editSession || !isEditLockValid;
    const missingEditSessionErrorMessage = 'Блокування заявки недійсне Відкрийте заявку для редагування ще раз';
    const isEditLockRenewalInProgress = this.editLockRenewal?.getIsRenewalInProgress() ?? false;
    const editLockRenewalInProgressErrorMessage = 'Зачекайте завершення продовження блокування та збережіть заявку ще раз';

    if (isEditSessionMissing) {
      throw new Error(missingEditSessionErrorMessage);
    }

    if (isEditLockRenewalInProgress) {
      throw new Error(editLockRenewalInProgressErrorMessage);
    }

    const updatedRequest = await this.service.updateRequest(
      editingRequest.Id,
      draft,
      editSession.token
    );

    this.stopEditLockRenewal();
    this.setState(previousState => ({
      data: previousState.data
        ? {
          ...previousState.data,
          requests: replaceListItem(
            previousState.data.requests,
            updatedRequest,
            'Id'
          )
        }
        : previousState.data,
      editingRequest: undefined,
      editSession: undefined,
      isEditLockValid: false,
      success: 'Заявку успішно оновлено',
      error: undefined
    }));
  };

  // запускає періодичне продовження блокування відкритої форми
  private startEditLockRenewal = (): void => {
    this.stopEditLockRenewal();
    const { editSession } = this.state;

    if (!editSession) {
      return;
    }

    this.editLockRenewal = new SharePointEditLockRenewal({
      intervalMilliseconds: editLockRenewalInterval,
      renew: () => this.service.renewEditLock(editSession.item.Id, editSession.token),
      onRenewed: () => this.setState({ isEditLockValid: true }),
      onError: this.handleEditLockRenewalError
    });
    this.editLockRenewal.start();
  };

  // зупиняє періодичне продовження блокування
  private stopEditLockRenewal(): void {
    this.editLockRenewal?.stop();
    this.editLockRenewal = undefined;
  }

  // показує помилку якщо блокування не вдалося продовжити
  private readonly handleEditLockRenewalError = (error: unknown): void => {
    this.setState({
      isEditLockValid: false,
      error: error instanceof Error
        ? error.message
        : 'Не вдалося продовжити блокування заявки'
    });
  };

  // звільняє блокування поточної форми без очікування відповіді сервера
  private releaseActiveEditLock(): void {
    const { editSession } = this.state;

    if (editSession) {
      this.service.releaseEditLock(editSession.item.Id, editSession.token)
        .catch(() => undefined);
    }
  }

  // відкриває підтвердження видалення вибраної заявки
  private readonly handleOpenDelete = (request: IServiceRequest): void => {
    this.setState({
      isCreateOpen: false,
      selectedRequest: undefined,
      editingRequest: undefined,
      deletingRequest: request,
      isDeleting: false,
      deleteError: undefined,
      success: undefined
    });
  };

  // закриває підтвердження видалення заявки
  private readonly handleCloseDelete = (): void => {
    if (!this.state.isDeleting) {
      this.setState({ deletingRequest: undefined, deleteError: undefined });
    }
  };

  // видаляє заявку та прибирає її з локального списку
  private readonly handleDeleteRequest = async (): Promise<void> => {
    const { deletingRequest } = this.state;

    if (!deletingRequest) {
      return;
    }

    this.setState({ isDeleting: true, deleteError: undefined });

    try {
      await this.service.deleteRequest(deletingRequest.Id);

      this.setState(previousState => ({
        data: previousState.data
          ? {
            ...previousState.data,
            requests: removeListItem(
              previousState.data.requests,
              deletingRequest.Id,
              'Id'
            )
          }
          : previousState.data,
        deletingRequest: undefined,
        isDeleting: false,
        success: 'Заявку успішно видалено',
        error: undefined
      }));
    } catch (error) {
      this.setState({
        isDeleting: false,
        deleteError: error instanceof Error
          ? error.message
          : 'Не вдалося видалити заявку'
      });
    }
  };

  // відображає стан завантаження кількість заявок і таблицю
  public render(): React.ReactElement<IServiceRequestsPageProps> {
    const { showRequestGenerator } = this.props;
    const {
      data, error, success, isLoading, isCreateOpen, selectedRequest, editingRequest,
      deletingRequest, isDeleting, deleteError, hasLoadedRequests, createdFrom, createdTo,
      isBackgroundRequestsLoading, totalRequests, isCountingRequests
    } = this.state;

    return (
      <section className={styles.page}>
        <ServiceRequestsPageHeader
          version={packageInfo.version}
          showRequestGenerator={showRequestGenerator}
          data={data}
          createdFrom={createdFrom}
          createdTo={createdTo}
          isLoading={isLoading}
          isBackgroundRequestsLoading={isBackgroundRequestsLoading}
          formatDate={this.formatDate}
          onCreatedFromChange={this.handleCreatedFromChange}
          onCreatedToChange={this.handleCreatedToChange}
          onRefresh={this.handleRefresh}
          onOpenCreate={this.handleOpenCreate}
          onCancelLoading={this.handleCancelRequestsLoading}
          onGenerate={this.handleGenerateRequest}
        />

        {isLoading && <Spinner className={styles.loading} label="Завантажуємо дані..." />}

        {this.state.isAcquiringEditLock && (
          <Spinner className={styles.loading} label="Перевіряємо блокування заявки..." />
        )}

        {error && (
          <MessageBar className={styles.statusMessage} messageBarType={MessageBarType.error}>
            {error}
          </MessageBar>
        )}

        {success && (
          <MessageBar className={styles.statusMessage} messageBarType={MessageBarType.success}>
            {success}
          </MessageBar>
        )}

        {data && (
          <div className={styles.content}>
            {hasLoadedRequests ? (
              <>
                <Text variant="medium">
                  {this.getRequestsLoadingText(
                    data.requests.length,
                    totalRequests,
                    isCountingRequests
                  )}
                </Text>

                <ServiceRequestsGrid
                  requests={data.requests}
                  onView={this.handleOpenView}
                  onEdit={this.handleOpenEdit}
                  onDelete={this.handleOpenDelete}
                  onVisibleRequestIdsChange={this.handleVisibleRequestIdsChange}
                />
              </>
            ) : (
              <Text variant="medium">Заявки ще не завантажено Натисніть Оновити</Text>
            )}

            {isCreateOpen && (
              <ServiceRequestForm
                isOpen={isCreateOpen}
                onDismiss={this.handleCloseCreate}
                categories={data.categories}
                subcategories={data.subcategories}
                peoplePickerContext={this.props.peoplePickerContext}
                currentUserEmail={this.props.currentUserEmail}
                onSubmit={this.handleCreateRequest}
              />
            )}

            {selectedRequest && (
              <ServiceRequestView
                request={selectedRequest}
                onDismiss={this.handleCloseView}
              />
            )}

            {editingRequest && (
              <ServiceRequestForm
                isOpen
                request={editingRequest}
                onDismiss={this.handleCloseEdit}
                categories={data.categories}
                subcategories={data.subcategories}
                peoplePickerContext={this.props.peoplePickerContext}
                currentUserEmail={this.props.currentUserEmail}
                isEditLockValid={this.state.isEditLockValid}
                onSubmit={this.handleUpdateRequest}
              />
            )}
            
            {deletingRequest && (
              <ServiceRequestDeleteDialog
                request={deletingRequest}
                isDeleting={isDeleting}
                error={deleteError}
                onDismiss={this.handleCloseDelete}
                onConfirm={this.handleDeleteRequest}
              />
            )}
          </div>
        )}
      </section>
    );
  }
}
