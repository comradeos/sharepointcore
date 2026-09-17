import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFx, spfi, type SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import type { ISiteInfo, ISiteInfoSource } from './ISiteInfoSource';

interface IWebProperties {
  Id: string;
  Title: string;
  Description?: string;
  Url: string;
  Created?: string;
  LastItemModifiedDate?: string;
}

/** Отримує відомості про поточний сайт через SharePoint REST API та PnPjs. */
export class PnpSiteInfoSource implements ISiteInfoSource {
  private readonly _sp: SPFI;

  /** Створює PnPjs-клієнт із поточним контекстом та авторизацією SPFx. */
  public constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  /** Повертає нормалізовані властивості поточного SharePoint-сайту. */
  public async getSiteInfo(): Promise<ISiteInfo> {
    const properties = await this._sp.web
      .select('Id', 'Title', 'Description', 'Url', 'Created', 'LastItemModifiedDate')<IWebProperties>();

    return {
      id: properties.Id,
      name: properties.Title,
      displayName: properties.Title,
      webUrl: properties.Url,
      description: properties.Description,
      createdDateTime: properties.Created,
      lastModifiedDateTime: properties.LastItemModifiedDate,
      hostname: new URL(properties.Url).hostname
    };
  }
}
