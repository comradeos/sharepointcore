import { IDropdownOption, IPersonaProps } from '@fluentui/react';
import { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { requestPriorities, requestStatuses } from '../models/ServiceDeskConstants';
import {
  IRequestCategory,
  IRequestSubcategory,
  IServiceRequest,
  IServiceRequestDraft
} from '../models/ServiceDeskModels';

// вхідні дані форми створення заявки
export interface IServiceRequestFormProps {
  isOpen: boolean;
  request?: IServiceRequest;
  onDismiss: () => void;
  categories: IRequestCategory[];
  subcategories: IRequestSubcategory[];
  peoplePickerContext: IPeoplePickerContext;
  currentUserEmail: string;
  isEditLockValid?: boolean;
  onSubmit: (draft: IServiceRequestDraft) => Promise<void>;
}

// значення полів заявки до збереження
export interface IServiceRequestFormState {
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
export interface IPeoplePickerPersona extends IPersonaProps {
  loginName?: string;
}

// помилки перевірки полів заявки
export interface IServiceRequestFormErrors {
  title?: string;
  description?: string;
  category?: string;
  subcategory?: string;
  status?: string;
  requester?: string;
  assignee?: string;
  plannedStart?: string;
  dueDate?: string;
  estimatedHours?: string;
  contactEmail?: string;
}

// доступні значення статусу зі списку sharepoint
export const statusOptions: IDropdownOption[] = [
  { key: requestStatuses.new, text: requestStatuses.new },
  { key: requestStatuses.inProgress, text: requestStatuses.inProgress },
  { key: requestStatuses.resolved, text: requestStatuses.resolved },
  { key: requestStatuses.closed, text: requestStatuses.closed }
];

// доступні значення пріоритету зі списку sharepoint
export const priorityOptions: IDropdownOption[] = [
  { key: requestPriorities.low, text: requestPriorities.low },
  { key: requestPriorities.medium, text: requestPriorities.medium },
  { key: requestPriorities.high, text: requestPriorities.high }
];
