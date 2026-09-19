import * as React from 'react';
import {
  DefaultButton,
  Dropdown,
  IDropdownOption,
  IPersonaProps,
  IconButton,
  Modal,
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
  IRequestSubcategory
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
      requester: [],
      assignee: [],
      plannedStart: '',
      dueDate: '',
      estimatedHours: '',
      contactEmail: '',
      requiresOnsiteVisit: false
    };
  }

  // зберігає назву заявки у стані форми
  private readonly handleTitleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ title: value || '' });
  };

  // зберігає опис заявки у стані форми
  private readonly handleDescriptionChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ description: value || '' });
  };

  // змінює категорію та очищає підкатегорію попередньої категорії
  private readonly handleCategoryChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ categoryId: option ? Number(option.key) : undefined, subcategoryId: undefined });
  };

  // зберігає вибрану підкатегорію у стані форми
  private readonly handleSubcategoryChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ subcategoryId: option ? Number(option.key) : undefined });
  };

  // зберігає вибраний статус у стані форми
  private readonly handleStatusChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    if (option) {
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
    this.setState({ requester: people });
  };

  // зберігає вибраного виконавця у стані форми
  private readonly handleAssigneeChange = (people: IPersonaProps[]): void => {
    this.setState({ assignee: people });
  };

  // зберігає плановий початок у стані форми
  private readonly handlePlannedStartChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ plannedStart: value || '' });
  };

  // зберігає кінцевий термін у стані форми
  private readonly handleDueDateChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ dueDate: value || '' });
  };

  // зберігає оцінку часу у стані форми
  private readonly handleEstimatedHoursChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ estimatedHours: value || '' });
  };

  // зберігає контактну адресу у стані форми
  private readonly handleContactEmailChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ contactEmail: value || '' });
  };

  // зберігає ознаку потреби у виїзді у стані форми
  private readonly handleOnsiteChange = (_event: React.MouseEvent<HTMLElement>, checked?: boolean): void => {
    this.setState({ requiresOnsiteVisit: checked || false });
  };

  // показує поля форми та обмежує підкатегорії вибраною категорією
  public render(): React.ReactElement<IServiceRequestFormProps> {
    const { categories, subcategories, peoplePickerContext, currentUserEmail } = this.props;
    const {
      title, description, categoryId, subcategoryId, status, priority,
      plannedStart, dueDate, estimatedHours, contactEmail, requiresOnsiteVisit
    } = this.state;
    const categoryOptions: IDropdownOption[] = [];
    const subcategoryOptions: IDropdownOption[] = [];

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
        onDismiss={this.props.onDismiss}
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
          />
        </div>
        <Stack className={styles.body} tokens={{ childrenGap: 14 }}>
          <TextField label="Назва заявки" required value={title} onChange={this.handleTitleChange} />
          <TextField label="Опис" required multiline rows={4} value={description} onChange={this.handleDescriptionChange} />
          <Dropdown label="Категорія" required placeholder="Оберіть категорію" options={categoryOptions} selectedKey={categoryId} onChange={this.handleCategoryChange} />
          <Dropdown label="Підкатегорія" required placeholder="Оберіть підкатегорію" options={subcategoryOptions} selectedKey={subcategoryId} onChange={this.handleSubcategoryChange} disabled={categoryId === undefined} />
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
            onChange={this.handleRequesterChange}
          />
          <PeoplePicker
            context={peoplePickerContext}
            titleText="Виконавець"
            placeholder="Оберіть виконавця"
            personSelectionLimit={1}
            principalTypes={[PrincipalType.User]}
            ensureUser
            onChange={this.handleAssigneeChange}
          />
          <TextField label="Плановий початок" type="datetime-local" value={plannedStart} onChange={this.handlePlannedStartChange} />
          <TextField label="Кінцевий термін" required type="datetime-local" value={dueDate} onChange={this.handleDueDateChange} />
          <TextField label="Оцінка часу в годинах" type="number" min={0.1} step={0.1} value={estimatedHours} onChange={this.handleEstimatedHoursChange} />
          <TextField label="Контактна електронна пошта" type="email" value={contactEmail} onChange={this.handleContactEmailChange} />
          <Toggle label="Потрібен виїзд" checked={requiresOnsiteVisit} onText="Так" offText="Ні" onChange={this.handleOnsiteChange} />
        </Stack>
        <div className={styles.footer}>
          <DefaultButton text="Закрити" onClick={this.props.onDismiss} />
        </div>
      </Modal>
    );
  }
}
