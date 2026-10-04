import {
  IServiceDeskData,
  IServiceRequest,
  IRequestPageCursor
} from '../models/ServiceDeskModels';
import { ISharePointListEditSession } from '../utils/SharePointListUtils';

// стан екрана сервісних заявок
export interface IServiceRequestsPageState {
  data?: IServiceDeskData;
  createdFrom: Date;
  createdTo: Date;
  hasLoadedRequests: boolean;
  isLoading: boolean;
  isBackgroundRequestsLoading: boolean;
  isCountingRequests: boolean;
  totalRequests?: number;
  isCreateOpen: boolean;
  selectedRequest?: IServiceRequest;
  editingRequest?: IServiceRequest;
  editSession?: ISharePointListEditSession<IServiceRequest>;
  isAcquiringEditLock: boolean;
  isEditLockValid: boolean;
  deletingRequest?: IServiceRequest;
  isDeleting: boolean;
  deleteError?: string;
  error?: string;
  success?: string;
}

// дані одного запиту фонової сторінки заявок
export interface IBackgroundRequestsLoadResult {
  requests: IServiceRequest[];
  cursor?: IRequestPageCursor;
}
