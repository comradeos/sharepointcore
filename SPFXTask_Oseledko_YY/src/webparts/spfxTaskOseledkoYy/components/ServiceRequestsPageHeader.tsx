import * as React from 'react';
import {
  DatePicker,
  DefaultButton,
  DayOfWeek,
  PrimaryButton,
  Stack,
  Text
} from '@fluentui/react';
import { IServiceDeskData, IServiceRequestDraft } from '../models/ServiceDeskModels';
import ServiceRequestGenerator from './ServiceRequestGenerator';
import styles from './ServiceRequestsPage.module.scss';

// властивості заголовка сторінки заявок
interface IServiceRequestsPageHeaderProps {
  version: string;
  showRequestGenerator: boolean;
  data?: IServiceDeskData;
  createdFrom: Date;
  createdTo: Date;
  isLoading: boolean;
  isBackgroundRequestsLoading: boolean;
  formatDate: (date?: Date) => string;
  onCreatedFromChange: NonNullable<React.ComponentProps<typeof DatePicker>['onSelectDate']>;
  onCreatedToChange: NonNullable<React.ComponentProps<typeof DatePicker>['onSelectDate']>;
  onRefresh: () => void;
  onOpenCreate: () => void;
  onCancelLoading: () => void;
  onGenerate: (drafts: IServiceRequestDraft[]) => Promise<void>;
}

// відображає назву сторінки та дії з датами
export default function ServiceRequestsPageHeader(
  props: IServiceRequestsPageHeaderProps
): React.ReactElement {
  const isDisabled = props.isLoading || props.isBackgroundRequestsLoading;

  return (
    <Stack className={styles.header}>
      <Text className={styles.pageTitle} variant="xLarge">
        Сервісні заявки · v{props.version}
      </Text>

      <Stack className={styles.headerActions}>
        {props.isBackgroundRequestsLoading ? (
          <DefaultButton
            text="Скасувати завантаження"
            onClick={props.onCancelLoading}
            className={styles.wideActionButton}
          />
        ) : (
          <>
            {props.showRequestGenerator && props.data && (
              <ServiceRequestGenerator
                categories={props.data.categories}
                subcategories={props.data.subcategories}
                disabled={isDisabled}
                onGenerate={props.onGenerate}
              />
            )}

            <div className={styles.dateFilterActions}>
              <DatePicker
                className={styles.createdDateFilter}
                label="Дата створення від"
                value={props.createdFrom}
                onSelectDate={props.onCreatedFromChange}
                formatDate={props.formatDate}
                firstDayOfWeek={DayOfWeek.Monday}
                disabled={isDisabled}
              />

              <DatePicker
                className={styles.createdDateFilter}
                label="Дата створення до"
                value={props.createdTo}
                onSelectDate={props.onCreatedToChange}
                formatDate={props.formatDate}
                firstDayOfWeek={DayOfWeek.Monday}
                disabled={isDisabled}
              />

              <DefaultButton
                text="Оновити"
                onClick={props.onRefresh}
                disabled={isDisabled}
                className={styles.refreshButton}
              />

              <PrimaryButton
                text="Створити"
                onClick={props.onOpenCreate}
                disabled={!props.data || isDisabled}
                className={styles.refreshButton}
              />
            </div>
          </>
        )}
      </Stack>
    </Stack>
  );
}
