import * as React from 'react';
import {
  DatePicker,
  DefaultButton,
  DayOfWeek,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  Stack,
  Text
} from '@fluentui/react';
import styles from './ServiceRequestsPage.module.scss';
import { IServiceRequestsPageProps } from './IServiceRequestsPageProps';
import {
  IRequestCreatedDateRange,
  IRequestEditSession,
  IRequestPageCursor,
  IServiceDeskDictionaries,
  IServiceDeskData,
  IServiceRequest,
  IServiceRequestDraft,
  IServiceRequestPage
} from '../models/ServiceDeskModels';
import ServiceDeskService from '../services/ServiceDeskService';
import {
  addListItems,
  removeListItem,
  replaceListItem
} from '../utils/SharePointListUtils';
import ServiceRequestsGrid from './ServiceRequestsGrid';
import ServiceRequestForm from './ServiceRequestForm';
import ServiceRequestView from './ServiceRequestView';
import ServiceRequestDeleteDialog from './ServiceRequestDeleteDialog';
import ServiceRequestGenerator from './ServiceRequestGenerator';

const backgroundRequestsPageSize = 1000;
const backgroundRequestsDelay = 3000;
const editLockRenewalInterval = 60 * 1000;

// стан екрана сервісних заявок
interface IServiceRequestsPageState {
  data?: IServiceDeskData;
  createdFrom: Date;
  createdTo: Date;
  nextRequestsPageCursor?: IRequestPageCursor;
  hasLoadedRequests: boolean;
  isLoading: boolean;
  isLoadingMoreRequests: boolean;
  isCountingRequests: boolean;
  totalRequests?: number;
  isCreateOpen: boolean;
  selectedRequest?: IServiceRequest;
  editingRequest?: IServiceRequest;
  editSession?: IRequestEditSession;
  isAcquiringEditLock: boolean;
  isEditLockValid: boolean;
  deletingRequest?: IServiceRequest;
  isDeleting: boolean;
  deleteError?: string;
  error?: string;
  success?: string;
}

// показує таблицю заявок із даними трьох списків sharepoint
export default class ServiceRequestsPage extends React.Component<
  IServiceRequestsPageProps,
  IServiceRequestsPageState
> {
  private readonly service: ServiceDeskService;
  private backgroundRequestsTimeout?: ReturnType<typeof setTimeout>;
  private backgroundRequestsVersion = 0;
  private editLockRenewalTimer?: ReturnType<typeof setInterval>;
  private isEditLockRenewalInProgress = false;

  // створює сервіс для сайту вебчастини та початковий стан екрана
  public constructor(props: IServiceRequestsPageProps) {
    super(props);
    this.service = new ServiceDeskService(props.spHttpClient, props.webUrl);
    const currentDate = this.getCurrentDate();

    this.state = {
      createdFrom: currentDate,
      createdTo: currentDate,
      hasLoadedRequests: false,
      isLoading: false,
      isLoadingMoreRequests: false,
      isCountingRequests: false,
      isCreateOpen: false,
      isAcquiringEditLock: false,
      isEditLockValid: false,
      isDeleting: false
    };
  }

  // завантажує довідники після появи вебчастини на сторінці
  public componentDidMount(): void {
    this.loadDictionaries();
  }

  // зупиняє фонові запити перед видаленням вебчастини зі сторінки
  public componentWillUnmount(): void {
    this.cancelBackgroundRequestsLoading();
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
      isLoadingMoreRequests: false,
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
    this.setState(previousState => ({
      data: previousState.data
        ? {
          ...previousState.data,
          requests: page.requests
        }
        : previousState.data,
      nextRequestsPageCursor: page.cursor,
      hasLoadedRequests: true,
      isLoading: false,
      isCountingRequests: Boolean(page.cursor)
    }), () => {
      this.startRequestsCount();
      this.scheduleBackgroundRequestsLoading();
    });
  };

  // завантажує наступну фонову частину заявок
  private readonly loadNextRequestsPage = (): void => {
    const { createdFrom, createdTo, isLoadingMoreRequests, nextRequestsPageCursor } = this.state;
    const canLoadNextRequestsPage = Boolean(nextRequestsPageCursor) && !isLoadingMoreRequests;

    if (!canLoadNextRequestsPage || !nextRequestsPageCursor) {
      return;
    }

    this.setState({ isLoadingMoreRequests: true, error: undefined });
    const currentBackgroundRequestsVersion = this.backgroundRequestsVersion;

    this.service.loadRequestsPage(
      { from: createdFrom, to: createdTo },
      nextRequestsPageCursor,
      backgroundRequestsPageSize
    )
      .then(page => this.handleNextRequestsPageLoadSuccess(
        page,
        currentBackgroundRequestsVersion
      ))
      .catch(error => this.handleNextRequestsPageLoadError(
        error,
        currentBackgroundRequestsVersion
      ));
  };

  // додає наступну частину заявок до вже завантажених рядків
  private readonly handleNextRequestsPageLoadSuccess = (
    page: IServiceRequestPage,
    backgroundRequestsVersion: number
  ): void => {
    if (backgroundRequestsVersion !== this.backgroundRequestsVersion) {
      return;
    }

    this.setState(previousState => ({
      data: previousState.data
        ? {
          ...previousState.data,
          requests: [...previousState.data.requests, ...page.requests]
        }
        : previousState.data,
      nextRequestsPageCursor: page.cursor,
      isLoadingMoreRequests: false
    }), this.scheduleBackgroundRequestsLoading);
  };

  // показує помилку фонової частини заявок та зупиняє завантаження
  private readonly handleNextRequestsPageLoadError = (
    error: unknown,
    backgroundRequestsVersion: number
  ): void => {
    if (backgroundRequestsVersion !== this.backgroundRequestsVersion) {
      return;
    }

    this.setState({
      isLoadingMoreRequests: false,
      error: error instanceof Error ? error.message : 'Невідома помилка завантаження.'
    });
  };

  // запускає підрахунок загальної кількості заявок вибраного діапазону
  private startRequestsCount(): void {
    const { createdFrom, createdTo, nextRequestsPageCursor } = this.state;

    if (!nextRequestsPageCursor) {
      this.setState(previousState => ({
        totalRequests: previousState.data?.requests.length ?? 0,
        isCountingRequests: false
      }));
      return;
    }

    const currentBackgroundRequestsVersion = this.backgroundRequestsVersion;

    this.service.countRequests({ from: createdFrom, to: createdTo })
      .then(totalRequests => this.handleRequestsCountSuccess(
        totalRequests,
        currentBackgroundRequestsVersion
      ))
      .catch(() => this.handleRequestsCountError(currentBackgroundRequestsVersion));
  }

  // зберігає обчислену загальну кількість заявок
  private readonly handleRequestsCountSuccess = (
    totalRequests: number,
    backgroundRequestsVersion: number
  ): void => {
    if (backgroundRequestsVersion === this.backgroundRequestsVersion) {
      this.setState({ totalRequests, isCountingRequests: false });
    }
  };

  // завершує підрахунок якщо не вдалося визначити загальну кількість
  private readonly handleRequestsCountError = (backgroundRequestsVersion: number): void => {
    if (backgroundRequestsVersion === this.backgroundRequestsVersion) {
      this.setState({ isCountingRequests: false });
    }
  };

  // планує наступне фонове завантаження після затримки
  private scheduleBackgroundRequestsLoading = (): void => {
    if (!this.state.nextRequestsPageCursor) {
      return;
    }

    this.cancelBackgroundRequestsTimeout();
    this.backgroundRequestsTimeout = setTimeout(
      this.loadNextRequestsPage,
      backgroundRequestsDelay
    );
  };

  // зупиняє таймер фонового завантаження
  private cancelBackgroundRequestsTimeout(): void {
    if (this.backgroundRequestsTimeout) {
      clearTimeout(this.backgroundRequestsTimeout);
      this.backgroundRequestsTimeout = undefined;
    }
  }

  // скасовує поточне фонове завантаження та захищає від застарілих відповідей
  private cancelBackgroundRequestsLoading(): void {
    this.backgroundRequestsVersion += 1;
    this.cancelBackgroundRequestsTimeout();
  }

  // скасовує фонове завантаження та залишає в таблиці вже отримані заявки
  private readonly handleCancelRequestsLoading = (): void => {
    this.cancelBackgroundRequestsLoading();

    this.setState(previousState => ({
      nextRequestsPageCursor: undefined,
      isLoadingMoreRequests: false,
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

    this.cancelBackgroundRequestsLoading();
    this.setState({ success: undefined });
    this.loadRequests({ from: createdFrom, to: createdTo });
  };

  // зберігає нижню межу дат та очищає завантажені заявки
  private readonly handleCreatedFromChange = (date: Date | null | undefined): void => {
    if (date) {
      this.cancelBackgroundRequestsLoading();
      this.setState(previousState => ({
        createdFrom: date,
        data: previousState.data
          ? { ...previousState.data, requests: [] }
          : previousState.data,
        nextRequestsPageCursor: undefined,
        hasLoadedRequests: false,
        isLoadingMoreRequests: false,
        isCountingRequests: false,
        totalRequests: undefined,
        error: undefined
      }));
    }
  };

  // зберігає верхню межу дат та очищає завантажені заявки
  private readonly handleCreatedToChange = (date: Date | null | undefined): void => {
    if (date) {
      this.cancelBackgroundRequestsLoading();
      this.setState(previousState => ({
        createdTo: date,
        data: previousState.data
          ? { ...previousState.data, requests: [] }
          : previousState.data,
        nextRequestsPageCursor: undefined,
        hasLoadedRequests: false,
        isLoadingMoreRequests: false,
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
        ? `Завантажено заявок: ${loadedRequestsCount} Визначаємо загальну кількість`
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
        editingRequest: editSession.request,
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
    const isEditLockRenewalInProgress = this.isEditLockRenewalInProgress;
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
    this.editLockRenewalTimer = setInterval(this.renewEditLock, editLockRenewalInterval);
  };

  // зупиняє періодичне продовження блокування
  private stopEditLockRenewal(): void {
    if (this.editLockRenewalTimer) {
      clearInterval(this.editLockRenewalTimer);
      this.editLockRenewalTimer = undefined;
    }
  }

  // продовжує строк блокування поки користувач редагує заявку
  private readonly renewEditLock = async (): Promise<void> => {
    const { editSession } = this.state;
    const cannotRenewEditLock = !editSession || this.isEditLockRenewalInProgress;

    if (cannotRenewEditLock) {
      return;
    }

    this.isEditLockRenewalInProgress = true;

    try {
      await this.service.renewEditLock(editSession.request.Id, editSession.token);
      this.setState({ isEditLockValid: true });
    } catch (error) {
      this.stopEditLockRenewal();
      this.setState({
        isEditLockValid: false,
        error: error instanceof Error
          ? error.message
          : 'Не вдалося продовжити блокування заявки'
      });
    } finally {
      this.isEditLockRenewalInProgress = false;
    }
  };

  // звільняє блокування поточної форми без очікування відповіді сервера
  private releaseActiveEditLock(): void {
    const { editSession } = this.state;

    if (editSession) {
      this.service.releaseEditLock(editSession.request.Id, editSession.token)
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
      nextRequestsPageCursor, isLoadingMoreRequests, totalRequests, isCountingRequests
    } = this.state;
    const isBackgroundRequestsLoading = Boolean(nextRequestsPageCursor);

    return (
      <section className={styles.page}>
        <Stack className={styles.header}>
          <Text className={styles.pageTitle} variant="xLarge">Сервісні заявки</Text>

          <Stack className={styles.headerActions}>
            {isBackgroundRequestsLoading ? (
              <DefaultButton
                text="Скасувати завантаження"
                onClick={this.handleCancelRequestsLoading}
                className={styles.wideActionButton}
              />
            ) : (
              <>
                {showRequestGenerator && data && (
                  <ServiceRequestGenerator
                    categories={data.categories}
                    subcategories={data.subcategories}
                    disabled={isLoading || isLoadingMoreRequests}
                    onGenerate={this.handleGenerateRequest}
                  />
                )}

                <div className={styles.dateFilterActions}>
                  <DatePicker
                    className={styles.createdDateFilter}
                    label="Дата створення від"
                    value={createdFrom}
                    onSelectDate={this.handleCreatedFromChange}
                    formatDate={this.formatDate}
                    firstDayOfWeek={DayOfWeek.Monday}
                    disabled={isLoading || isLoadingMoreRequests}
                  />

                  <DatePicker
                    className={styles.createdDateFilter}
                    label="Дата створення до"
                    value={createdTo}
                    onSelectDate={this.handleCreatedToChange}
                    formatDate={this.formatDate}
                    firstDayOfWeek={DayOfWeek.Monday}
                    disabled={isLoading || isLoadingMoreRequests}
                  />

                  <DefaultButton
                    text="Оновити"
                    onClick={this.handleRefresh}
                    disabled={isLoading || isLoadingMoreRequests}
                    className={styles.refreshButton}
                  />
                </div>

                <PrimaryButton
                  text="Створити"
                  onClick={this.handleOpenCreate}
                  disabled={!data || isLoading || isLoadingMoreRequests}
                  className={styles.wideActionButton}
                />
              </>
            )}
          </Stack>
        </Stack>

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
