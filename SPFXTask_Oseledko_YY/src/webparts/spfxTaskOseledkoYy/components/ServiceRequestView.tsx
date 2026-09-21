import * as React from 'react';
import { DefaultButton, IconButton, Label, Modal, Text } from '@fluentui/react';
import { IServiceRequest, SharePointNullable } from '../models/ServiceDeskModels';
import styles from './ServiceRequestView.module.scss';

// вхідні дані вікна перегляду заявки
export interface IServiceRequestViewProps {
  request: IServiceRequest;
  onDismiss: () => void;
}

// вхідні дані одного поля у режимі перегляду
interface IReadOnlyFieldProps {
  label: string;
  value: string;
  wide?: boolean;
}

const dateFormatter = new Intl.DateTimeFormat('uk-UA', {
  day: '2-digit',
  month: '2-digit',
  year: 'numeric',
  hour: '2-digit',
  minute: '2-digit'
});

const hoursFormatter = new Intl.NumberFormat('uk-UA', {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1
});

// перетворює дату sharepoint на український формат
function formatDate(value: SharePointNullable<string>): string {
  if (!value) {
    return 'Не вказано';
  }

  const timestamp = Date.parse(value);

  return Number.isNaN(timestamp) ? 'Не вказано' : dateFormatter.format(timestamp);
}

// перетворює оцінку часу на український числовий формат
function formatHours(value: SharePointNullable<number>): string {
  return typeof value === 'number' ? hoursFormatter.format(value) : 'Не вказано';
}

// показує підпис та значення одного поля заявки
function ReadOnlyField(props: IReadOnlyFieldProps): React.ReactElement {
  const className = props.wide ? `${styles.field} ${styles.wideField}` : styles.field;

  return (
    <div className={className}>
      <Label>{props.label}</Label>
      <Text className={styles.value}>{props.value}</Text>
    </div>
  );
}

// показує всі значення вибраної заявки без редагування
export default function ServiceRequestView(
  props: IServiceRequestViewProps
): React.ReactElement {
  const { request } = props;

  return (
    <Modal
      isOpen
      onDismiss={props.onDismiss}
      isBlocking
      containerClassName={styles.modal}
      scrollableContentClassName={styles.content}
      titleAriaId="service-request-view-title"
    >
      <div className={styles.header}>
        <h2 id="service-request-view-title" className={styles.title}>Перегляд сервісної заявки</h2>
        
        <IconButton
          iconProps={{ iconName: 'Cancel' }}
          ariaLabel="Закрити"
          onClick={props.onDismiss}
        />
      </div>

      <div className={styles.body}>
        <ReadOnlyField label="ID" value={String(request.Id)} />
        <ReadOnlyField label="Назва заявки" value={request.Title} />
        <ReadOnlyField label="Опис" value={request.Description} wide />
        <ReadOnlyField label="Категорія" value={request.Category?.Title || 'Не вказано'} />
        <ReadOnlyField label="Підкатегорія" value={request.Subcategory?.Title || 'Не вказано'} />
        <ReadOnlyField label="Статус" value={request.Status} />
        <ReadOnlyField label="Пріоритет" value={request.Priority} />
        <ReadOnlyField label="Заявник" value={request.Requester?.Title || 'Не вказано'} />
        <ReadOnlyField label="Виконавець" value={request.Assignee?.Title || 'Не призначено'} />
        <ReadOnlyField label="Плановий початок" value={formatDate(request.PlannedStart)} />
        <ReadOnlyField label="Кінцевий термін" value={formatDate(request.DueDate)} />
        <ReadOnlyField label="Оцінка часу в годинах" value={formatHours(request.EstimatedHours)} />
        <ReadOnlyField label="Контактна електронна пошта" value={request.ContactEmail || 'Не вказано'} />
        <ReadOnlyField label="Потрібен виїзд" value={request.RequiresOnsiteVisit ? 'Так' : 'Ні'} />
      </div>

      <div className={styles.footer}>
        <DefaultButton text="Закрити" onClick={props.onDismiss} />
      </div>
    </Modal>
  );
}
