import * as React from 'react';
import {
  DefaultButton,
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
  IServiceDeskData,
  IServiceRequest,
  IServiceRequestDraft
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

// стан екрана сервісних заявок
interface IServiceRequestsPageState {
  data?: IServiceDeskData;
  isLoading: boolean;
  isCreateOpen: boolean;
  selectedRequest?: IServiceRequest;
  editingRequest?: IServiceRequest;
  deletingRequest?: IServiceRequest;
  isDeleting: boolean;
  deleteError?: string;
  error?: string;
  success?: string;
}

// тимчасовий функціонал для швидкого створення тестових записів у базі
const showRequestGenerator = false;

// показує таблицю заявок із даними трьох списків sharepoint
export default class ServiceRequestsPage extends React.Component<
  IServiceRequestsPageProps,
  IServiceRequestsPageState
> {
  private readonly service: ServiceDeskService;

  // створює сервіс для сайту вебчастини та початковий стан екрана
  public constructor(props: IServiceRequestsPageProps) {
    super(props);
    this.service = new ServiceDeskService(props.spHttpClient, props.webUrl);
    this.state = {
      isLoading: false,
      isCreateOpen: false,
      isDeleting: false
    };
  }

  // запускає завантаження після появи вебчастини на сторінці
  public componentDidMount(): void {
    this.loadData();
  }

  // запускає запити до списків і передає результат відповідним обробникам
  private loadData(): void {
    this.setState({ isLoading: true, error: undefined });
    this.service.loadData()
      .then(this.handleLoadSuccess)
      .catch(this.handleLoadError);
  }

  // зберігає завантажені дані та прибирає індикатор завантаження
  private readonly handleLoadSuccess = (data: IServiceDeskData): void => {
    this.setState({ data, isLoading: false });
  };

  // показує повідомлення якщо один із запитів завершився помилкою
  private readonly handleLoadError = (error: unknown): void => {
    this.setState({
      isLoading: false,
      error: error instanceof Error ? error.message : 'Невідома помилка завантаження.'
    });
  };

  // повторює запити після натискання кнопки оновлення
  private readonly handleRefresh = (): void => {
    this.setState({ success: undefined });
    this.loadData();
  };

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

  // відкриває вибрану заявку у режимі редагування
  private readonly handleOpenEdit = (request: IServiceRequest): void => {
    this.setState({
      isCreateOpen: false,
      selectedRequest: undefined,
      editingRequest: request,
      deletingRequest: undefined,
      success: undefined
    });
  };

  // закриває форму редагування заявки
  private readonly handleCloseEdit = (): void => {
    this.setState({ editingRequest: undefined });
  };

  // оновлює заявку та замінює її у локальному списку
  private readonly handleUpdateRequest = async (draft: IServiceRequestDraft): Promise<void> => {
    const { editingRequest } = this.state;

    if (!editingRequest) {
      throw new Error('Не вдалося визначити заявку для оновлення');
    }

    const updatedRequest = await this.service.updateRequest(editingRequest.Id, draft);

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
      success: 'Заявку успішно оновлено',
      error: undefined
    }));
  };

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
    const {
      data, error, success, isLoading, isCreateOpen, selectedRequest, editingRequest,
      deletingRequest, isDeleting, deleteError
    } = this.state;

    return (
      <section className={styles.page}>
        <Stack
          className={styles.header}
          horizontal
          horizontalAlign="space-between"
          verticalAlign="center"
          wrap
        >
          <Text className={styles.pageTitle} variant="xLarge">Сервісні заявки</Text>

          <Stack className={styles.headerActions} horizontal wrap tokens={{ childrenGap: 8 }}>
            {showRequestGenerator && data && (
              <ServiceRequestGenerator
                categories={data.categories}
                subcategories={data.subcategories}
                disabled={isLoading}
                onGenerate={this.handleGenerateRequest}
              />
            )}

            <DefaultButton 
              text="Оновити" 
              onClick={this.handleRefresh} 
              disabled={isLoading} 
            />

            <PrimaryButton
              text="Створити"
              onClick={this.handleOpenCreate}
              disabled={!data || isLoading}
            />
          </Stack>
        </Stack>

        {isLoading && <Spinner className={styles.loading} label="Завантажуємо списки..." />}

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
            <Text variant="medium">Усього заявок: {data.requests.length}</Text>

            <ServiceRequestsGrid
              requests={data.requests}
              onView={this.handleOpenView}
              onEdit={this.handleOpenEdit}
              onDelete={this.handleOpenDelete}
            />

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
