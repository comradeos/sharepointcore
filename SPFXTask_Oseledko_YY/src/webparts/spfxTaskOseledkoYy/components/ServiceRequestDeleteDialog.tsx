import * as React from 'react';
import { DefaultButton, Dialog, DialogFooter, DialogType, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, Stack } from '@fluentui/react';
import { IServiceRequest } from '../models/ServiceDeskModels';

// вхідні дані вікна підтвердження видалення
interface IServiceRequestDeleteDialogProps {
  request: IServiceRequest;
  isDeleting: boolean;
  error?: string;
  onDismiss: () => void;
  onConfirm: () => void;
}

// показує підтвердження перед видаленням заявки
export default function ServiceRequestDeleteDialog(
  props: IServiceRequestDeleteDialogProps
): React.ReactElement<IServiceRequestDeleteDialogProps> {
  return (
    <Dialog
      hidden={false}
      onDismiss={props.isDeleting ? undefined : props.onDismiss}
      dialogContentProps={{
        type: DialogType.normal,
        title: 'Видалення сервісної заявки',
        closeButtonAriaLabel: 'Закрити',
        subText: `Ви впевнені що хочете видалити заявку "${props.request.Title}"`
      }}
      modalProps={{ isBlocking: true }}
    >
      <Stack tokens={{ childrenGap: 12 }}>
        {props.error && (
          <MessageBar messageBarType={MessageBarType.error}>{props.error}</MessageBar>
        )}
        {props.isDeleting && (
          <Spinner size={SpinnerSize.small} label="Видаляємо заявку" />
        )}
      </Stack>

      <DialogFooter>
        <PrimaryButton
          text="Видалити"
          onClick={props.onConfirm}
          disabled={props.isDeleting}
        />

        <DefaultButton
          text="Скасувати"
          onClick={props.onDismiss}
          disabled={props.isDeleting}
        />
      </DialogFooter>
    </Dialog>
  );
}
