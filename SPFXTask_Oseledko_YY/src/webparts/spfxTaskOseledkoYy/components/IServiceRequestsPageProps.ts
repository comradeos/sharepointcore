import { SPHttpClient } from '@microsoft/sp-http';
import { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';

export interface IServiceRequestsPageProps {
  spHttpClient: SPHttpClient;
  webUrl: string;
  peoplePickerContext: IPeoplePickerContext;
  currentUserEmail: string;
}
