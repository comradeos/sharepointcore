import type { IProductSource } from '../../services/products/IProductSource';
import type { ISiteInfoSource } from '../../services/site-info/ISiteInfoSource';
import type { ISiteListSource } from '../../services/site-lists/ISiteListSource';
import type { IUserLookupSource } from '../../services/user-lookup/IUserLookupSource';
import type { IUserCreateSource } from '../../services/user-create/IUserCreateSource';

export interface IOseledkoYyTest1Props {
  description: string;
  source: IProductSource;
  siteInfoSource: ISiteInfoSource;
  siteListSource: ISiteListSource;
  userLookupSource: IUserLookupSource;
  userCreateSource: IUserCreateSource;
}
