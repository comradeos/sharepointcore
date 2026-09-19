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
  IServiceRequestDraft
} from '../models/ServiceDeskModels';
import ServiceDeskService from '../services/ServiceDeskService';
import ServiceRequestsGrid from './ServiceRequestsGrid';
import ServiceRequestForm from './ServiceRequestForm';

// стан екрана сервісних заявок
interface IServiceRequestsPageState {
  data?: IServiceDeskData;
  isLoading: boolean;
  isCreateOpen: boolean;
  error?: string;
  success?: string;
}

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
    this.state = { isLoading: false, isCreateOpen: false };
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
    this.setState({ isCreateOpen: true, success: undefined });
  };

  // закриває форму створення заявки
  private readonly handleCloseCreate = (): void => {
    this.setState({ isCreateOpen: false });
  };

  // створює заявку та оновлює таблицю після успішної відповіді
  private readonly handleCreateRequest = async (draft: IServiceRequestDraft): Promise<void> => {
    await this.service.createRequest(draft);
    this.setState({
      success: 'Заявку успішно створено',
      error: undefined
    });
    this.loadData();
  };

  // відображає стан завантаження кількість заявок і таблицю
  public render(): React.ReactElement<IServiceRequestsPageProps> {
    const { data, error, success, isLoading, isCreateOpen } = this.state;

    return (
      <section className={styles.page}>
        <Stack horizontal horizontalAlign="space-between" verticalAlign="center" wrap>
          <Text variant="xLarge">Сервісні заявки</Text>
          <Stack horizontal tokens={{ childrenGap: 8 }}>
            <PrimaryButton
              text="Створити заявку"
              onClick={this.handleOpenCreate}
              disabled={!data || isLoading}
            />
            <DefaultButton text="Оновити" onClick={this.handleRefresh} disabled={isLoading} />
          </Stack>
        </Stack>

        {isLoading && <Spinner label="Завантажуємо списки..." />}
        {error && <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>}
        {success && (
          <MessageBar className={styles.successMessage} messageBarType={MessageBarType.success}>
            {success}
          </MessageBar>
        )}

        {data && (
          <div className={styles.content}>
            <Text variant="medium">Усього заявок: {data.requests.length}</Text>
            <ServiceRequestsGrid requests={data.requests} />
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
          </div>
        )}
      </section>
    );
  }
}
