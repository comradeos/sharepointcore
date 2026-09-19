import * as React from 'react';
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IPersonaProps,
  IconButton,
  MessageBar,
  MessageBarType,
  Modal,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  TextField,
  Toggle
} from '@fluentui/react';
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType
} from '@pnp/spfx-controls-react/lib/PeoplePicker';
import {
  IRequestCategory,
  IRequestSubcategory,
  IServiceRequestDraft
} from '../models/ServiceDeskModels';
import styles from './ServiceRequestForm.module.scss';

// вхідні дані форми створення заявки
export interface IServiceRequestFormProps {
  isOpen: boolean;
  onDismiss: () => void;
  categories: IRequestCategory[];
  subcategories: IRequestSubcategory[];
  peoplePickerContext: IPeoplePickerContext;
  currentUserEmail: string;
  onSubmit: (draft: IServiceRequestDraft) => Promise<void>;
}

// значення полів нової заявки до збереження
interface IServiceRequestFormState {
  title: string;
  description: string;
  categoryId?: number;
  subcategoryId?: number;
  status: string;
  priority: string;
  requester: IPersonaProps[];
  assignee: IPersonaProps[];
  plannedStart: string;
  dueDate: string;
  estimatedHours: string;
  contactEmail: string;
  requiresOnsiteVisit: boolean;
  validationErrors: IServiceRequestFormErrors;
  validationMessage?: string;
  isSubmitting: boolean;
  submitError?: string;
}

// значення користувача яке повертає peoplepicker
interface IPeoplePickerPersona extends IPersonaProps {
  loginName?: string;
}

// помилки перевірки полів нової заявки
interface IServiceRequestFormErrors {
  title?: string;
  description?: string;
  category?: string;
  subcategory?: string;
  requester?: string;
  assignee?: string;
  plannedStart?: string;
  dueDate?: string;
  estimatedHours?: string;
  contactEmail?: string;
}

// доступні значення статусу зі списку sharepoint
const statusOptions: IDropdownOption[] = [
  { key: 'Нова', text: 'Нова' },
  { key: 'В роботі', text: 'В роботі' },
  { key: 'Вирішена', text: 'Вирішена' },
  { key: 'Закрита', text: 'Закрита' }
];

// доступні значення пріоритету зі списку sharepoint
const priorityOptions: IDropdownOption[] = [
  { key: 'Низький', text: 'Низький' },
  { key: 'Середній', text: 'Середній' },
  { key: 'Високий', text: 'Високий' }
];

// показує поля нової заявки у модальному вікні
export default class ServiceRequestForm extends React.Component<
  IServiceRequestFormProps,
  IServiceRequestFormState
> {
  // задає початкові значення полів відповідно до схеми списку
  public constructor(props: IServiceRequestFormProps) {
    super(props);
    this.state = {
      title: '',
      description: '',
      status: 'Нова',
      priority: 'Середній',
      requester: props.currentUserEmail
        ? [{ text: props.currentUserEmail, secondaryText: props.currentUserEmail }]
        : [],
      assignee: [],
      plannedStart: '',
      dueDate: '',
      estimatedHours: '',
      contactEmail: '',
      requiresOnsiteVisit: false,
      validationErrors: {},
      isSubmitting: false
    };
  }

  // прибирає помилку зміненого поля та повідомлення перевірки
  private clearValidation(...fields: Array<keyof IServiceRequestFormErrors>): void {
    const validationErrors = { ...this.state.validationErrors };
    for (const field of fields) {
      delete validationErrors[field];
    }
    this.setState({ validationErrors, validationMessage: undefined });
  }

  // зберігає назву заявки у стані форми
  private readonly handleTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.clearValidation('title');
    this.setState({ title: value || '' });
  };

  // зберігає опис заявки у стані форми
  private readonly handleDescriptionChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.clearValidation('description');
    this.setState({ description: value || '' });
  };

  // змінює категорію та очищає підкатегорію попередньої категорії
  private readonly handleCategoryChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.clearValidation('category', 'subcategory');
    this.setState({ categoryId: option ? Number(option.key) : undefined, subcategoryId: undefined });
  };

  // зберігає вибрану підкатегорію у стані форми
  private readonly handleSubcategoryChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.clearValidation('subcategory');
    this.setState({ subcategoryId: option ? Number(option.key) : undefined });
  };

  // зберігає вибраний статус у стані форми
  private readonly handleStatusChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.clearValidation('assignee');
      this.setState({ status: String(option.key) });
    }
  };

  // зберігає вибраний пріоритет у стані форми
  private readonly handlePriorityChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
      this.setState({ priority: String(option.key) });
    }
  };

  // зберігає вибраного заявника у стані форми
  private readonly handleRequesterChange = (people: IPersonaProps[]): void => {
    this.clearValidation('requester');
    this.setState({ requester: people });
  };

  // зберігає вибраного виконавця у стані форми
  private readonly handleAssigneeChange = (people: IPersonaProps[]): void => {
    this.clearValidation('assignee');
    this.setState({ assignee: people });
  };

  // зберігає плановий початок у стані форми
  private readonly handlePlannedStartChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.clearValidation('plannedStart', 'dueDate');
    this.setState({ plannedStart: value || '' });
  };

  // зберігає кінцевий термін у стані форми
  private readonly handleDueDateChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.clearValidation('plannedStart', 'dueDate');
    this.setState({ dueDate: value || '' });
  };

  // зберігає оцінку часу у стані форми
  private readonly handleEstimatedHoursChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.clearValidation('estimatedHours');
    this.setState({ estimatedHours: value || '' });
  };

  // зберігає контактну адресу у стані форми
  private readonly handleContactEmailChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.clearValidation('contactEmail');
    this.setState({ contactEmail: value || '' });
  };

  // зберігає ознаку потреби у виїзді у стані форми
  private readonly handleOnsiteChange = (_event: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    this.setState({ requiresOnsiteVisit: checked || false });
  };

  // повертає поточну локальну дату та час без секунд
  private getMinimumDateTimeValue(): string {
    const currentDate = new Date();
    currentDate.setSeconds(0, 0);
    const localDate = new Date(currentDate.getTime() - currentDate.getTimezoneOffset() * 60000);
    return localDate.toISOString().slice(0, 16);
  }

  // перевіряє значення форми за правилами списку sharepoint
  private validateForm(): IServiceRequestFormErrors {
    const errors: IServiceRequestFormErrors = {};
    const {
      title, description, categoryId, subcategoryId, status, requester, assignee,
      plannedStart, dueDate, estimatedHours, contactEmail
    } = this.state;
    const trimmedTitle = title.trim();
    const trimmedDescription = description.trim();

    if (!trimmedTitle) {
      errors.title = 'Вкажіть назву заявки';
    } else if (trimmedTitle.length < 3) {
      errors.title = 'Назва повинна містити щонайменше 3 символи';
    } else if (trimmedTitle.length > 255) {
      errors.title = 'Назва не може містити більше 255 символів';
    }

    if (!trimmedDescription) {
      errors.description = 'Вкажіть опис заявки';
    }

    if (categoryId === undefined) {
      errors.category = 'Оберіть категорію';
    }

    if (subcategoryId === undefined) {
      errors.subcategory = 'Оберіть підкатегорію';
    } else {
      let isRelatedSubcategory = false;
      for (const subcategory of this.props.subcategories) {
        if (subcategory.Id === subcategoryId && subcategory.CategoryId === categoryId) {
          isRelatedSubcategory = true;
          break;
        }
      }
      if (!isRelatedSubcategory) {
        errors.subcategory = 'Підкатегорія не належить обраній категорії';
      }
    }

    if (requester.length === 0) {
      errors.requester = 'Оберіть заявника';
    }

    if ((status === 'Вирішена' || status === 'Закрита') && assignee.length === 0) {
      errors.assignee = 'Оберіть виконавця для завершеної заявки';
    }

    const plannedStartTime = plannedStart ? new Date(plannedStart).getTime() : undefined;
    const dueDateTime = dueDate ? new Date(dueDate).getTime() : undefined;
    const minimumDateTime = new Date(this.getMinimumDateTimeValue()).getTime();

    if (!dueDate) {
      errors.dueDate = 'Вкажіть кінцевий термін';
    } else if (dueDateTime === undefined || Number.isNaN(dueDateTime)) {
      errors.dueDate = 'Вкажіть коректний кінцевий термін';
    } else if (dueDateTime < minimumDateTime) {
      errors.dueDate = 'Кінцевий термін не може бути в минулому';
    }

    if (plannedStart && (plannedStartTime === undefined || Number.isNaN(plannedStartTime))) {
      errors.plannedStart = 'Вкажіть коректний плановий початок';
    } else if (plannedStartTime !== undefined && plannedStartTime < minimumDateTime) {
      errors.plannedStart = 'Плановий початок не може бути в минулому';
    }

    if (plannedStartTime !== undefined && dueDateTime !== undefined
      && !Number.isNaN(plannedStartTime) && !Number.isNaN(dueDateTime)) {
      if (plannedStartTime > dueDateTime) {
        errors.plannedStart = 'Плановий початок не може бути пізніше кінцевого терміну';
        errors.dueDate = 'Кінцевий термін не може бути раніше планового початку';
      }
    }

    if (estimatedHours) {
      const estimatedHoursPattern = /^\d+(\.\d)?$/;
      const estimatedHoursNumber = Number(estimatedHours);
      if (!estimatedHoursPattern.test(estimatedHours) || estimatedHoursNumber <= 0) {
        errors.estimatedHours = 'Вкажіть додатне число з одним десятковим знаком';
      }
    }

    if (contactEmail) {
      const emailPattern = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailPattern.test(contactEmail.trim())) {
        errors.contactEmail = 'Вкажіть коректну адресу електронної пошти';
      }
    }

    return errors;
  }

  // повертає адресу або логін вибраного користувача
  private getPersonIdentity(people: IPersonaProps[], fallback = ''): string {
    const person = people[0] as IPeoplePickerPersona | undefined;
    if (!person) {
      return fallback;
    }

    return person.loginName || person.secondaryText || String(person.id || fallback);
  }

  // готує незалежну від інтерфейсу модель для сервісу sharepoint
  private createDraft(): IServiceRequestDraft {
    const {
      title, description, categoryId, subcategoryId, status, priority,
      requester, assignee, plannedStart, dueDate, estimatedHours,
      contactEmail, requiresOnsiteVisit
    } = this.state;
    const assigneeIdentity = this.getPersonIdentity(assignee);

    return {
      title: title.trim(),
      description: description.trim(),
      categoryId: categoryId as number,
      subcategoryId: subcategoryId as number,
      status,
      priority,
      requesterIdentity: this.getPersonIdentity(requester, this.props.currentUserEmail),
      assigneeIdentity: assigneeIdentity || undefined,
      plannedStart: plannedStart || undefined,
      dueDate,
      estimatedHours: estimatedHours ? Number(estimatedHours) : undefined,
      contactEmail: contactEmail.trim() || undefined,
      requiresOnsiteVisit
    };
  }

  // перевіряє форму та передає заявку для створення
  private readonly handleSubmit = async (): Promise<void> => {
    const validationErrors = this.validateForm();
    if (Object.keys(validationErrors).length > 0) {
      this.setState({
        validationErrors,
        validationMessage: 'Перевірте виділені поля'
      });
      return;
    }

    this.setState({
      validationErrors: {},
      validationMessage: undefined,
      isSubmitting: true,
      submitError: undefined
    });

    try {
      await this.props.onSubmit(this.createDraft());
      this.props.onDismiss();
    } catch (error) {
      this.setState({
        isSubmitting: false,
        submitError: error instanceof Error
          ? error.message
          : 'Не вдалося створити заявку'
      });
    }
  };

  // показує поля форми та обмежує підкатегорії вибраною категорією
  public render(): React.ReactElement<IServiceRequestFormProps> {
    const { categories, subcategories, peoplePickerContext, currentUserEmail } = this.props;
    const {
      title, description, categoryId, subcategoryId, status, priority,
      plannedStart, dueDate, estimatedHours, contactEmail, requiresOnsiteVisit,
      validationErrors, validationMessage, isSubmitting, submitError
    } = this.state;
    const categoryOptions: IDropdownOption[] = [];
    const subcategoryOptions: IDropdownOption[] = [];
    const minimumDateTime = this.getMinimumDateTimeValue();

    for (const category of categories) {
      if (category.IsActive) {
        categoryOptions.push({ key: category.Id, text: category.Title });
      }
    }

    for (const subcategory of subcategories) {
      if (subcategory.IsActive && subcategory.CategoryId === categoryId) {
        subcategoryOptions.push({ key: subcategory.Id, text: subcategory.Title });
      }
    }

    return (
      <Modal
        isOpen={this.props.isOpen}
        onDismiss={isSubmitting ? undefined : this.props.onDismiss}
        isBlocking
        containerClassName={styles.modal}
        scrollableContentClassName={styles.content}
        styles={{ main: { overflowY: 'hidden' }, scrollableContent: { overflowY: 'hidden' } }}
        titleAriaId="service-request-form-title"
      >
        <div className={styles.header}>
          <h2 id="service-request-form-title" className={styles.title}>Нова сервісна заявка</h2>
          <IconButton
            iconProps={{ iconName: 'Cancel' }}
            ariaLabel="Закрити"
            onClick={this.props.onDismiss}
            disabled={isSubmitting}
          />
        </div>
        <Stack className={styles.body} tokens={{ childrenGap: 14 }}>
          {validationMessage && (
            <MessageBar messageBarType={Object.keys(validationErrors).length > 0 ? MessageBarType.error : MessageBarType.success}>
              {validationMessage}
            </MessageBar>
          )}
          {submitError && (
            <MessageBar messageBarType={MessageBarType.error}>{submitError}</MessageBar>
          )}
          <TextField label="Назва заявки" required value={title} errorMessage={validationErrors.title} onChange={this.handleTitleChange} />
          <TextField label="Опис" required multiline rows={4} value={description} errorMessage={validationErrors.description} onChange={this.handleDescriptionChange} />
          <Dropdown label="Категорія" required placeholder="Оберіть категорію" options={categoryOptions} selectedKey={categoryId} errorMessage={validationErrors.category} onChange={this.handleCategoryChange} />
          <Dropdown label="Підкатегорія" required placeholder="Оберіть підкатегорію" options={subcategoryOptions} selectedKey={subcategoryId} errorMessage={validationErrors.subcategory} onChange={this.handleSubcategoryChange} disabled={categoryId === undefined} />
          <Dropdown label="Статус" required options={statusOptions} selectedKey={status} onChange={this.handleStatusChange} />
          <Dropdown label="Пріоритет" required options={priorityOptions} selectedKey={priority} onChange={this.handlePriorityChange} />
          <PeoplePicker
            context={peoplePickerContext}
            titleText="Заявник"
            placeholder="Оберіть заявника"
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            defaultSelectedUsers={currentUserEmail ? [currentUserEmail] : []}
            ensureUser
            required
            errorMessage={validationErrors.requester}
            onChange={this.handleRequesterChange}
          />
          <PeoplePicker
            context={peoplePickerContext}
            titleText="Виконавець"
            placeholder="Оберіть виконавця"
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            ensureUser
            errorMessage={validationErrors.assignee}
            onChange={this.handleAssigneeChange}
          />
          <TextField label="Плановий початок" type="datetime-local" min={minimumDateTime} value={plannedStart} errorMessage={validationErrors.plannedStart} onChange={this.handlePlannedStartChange} />
          <TextField label="Кінцевий термін" required type="datetime-local" min={minimumDateTime} value={dueDate} errorMessage={validationErrors.dueDate} onChange={this.handleDueDateChange} />
          <TextField label="Оцінка часу в годинах" type="number" min={0.1} step={0.1} value={estimatedHours} errorMessage={validationErrors.estimatedHours} onChange={this.handleEstimatedHoursChange} />
          <TextField label="Контактна електронна пошта" type="email" value={contactEmail} errorMessage={validationErrors.contactEmail} onChange={this.handleContactEmailChange} />
          <Toggle label="Потрібен виїзд" checked={requiresOnsiteVisit} onText="Так" offText="Ні" onChange={this.handleOnsiteChange} />
        </Stack>
        <div className={styles.footer}>
          <PrimaryButton text="Створити" onClick={this.handleSubmit} disabled={isSubmitting} />
          <DefaultButton text="Закрити" onClick={this.props.onDismiss} disabled={isSubmitting} />
          {isSubmitting && (
            <Spinner size={SpinnerSize.small} label="Зберігаємо заявку" />
          )}
        </div>
      </Modal>
    );
  }
}
