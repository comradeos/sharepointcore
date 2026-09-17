import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFx, spfi, type SPFI } from '@pnp/sp';
import '@pnp/sp/lists';
import type { ISiteList, ISiteListSource } from './ISiteListSource';

interface IListProperties {
  Id: string;
  Title: string;
  ItemCount: number;
  BaseTemplate: number;
  DefaultViewUrl?: string;
  LastItemModifiedDate?: string;
}

/** Отримує видимі списки поточного сайту через SharePoint REST API та PnPjs. */
export class PnpSiteListSource implements ISiteListSource {
  private readonly _sp: SPFI;

  /** Створює PnPjs-клієнт із поточним контекстом та авторизацією SPFx. */
  public constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  /** Повертає нормалізований список видимих списків поточного SharePoint-сайту. */
  public async getSiteLists(): Promise<ISiteList[]> {
    const lists = await this._sp.web.lists
      .filter('Hidden eq false')
      .select('Id', 'Title', 'ItemCount', 'BaseTemplate', 'DefaultViewUrl', 'LastItemModifiedDate')
      .orderBy('Title')<IListProperties[]>();

    return lists.map(list => ({
      id: list.Id,
      title: list.Title,
      itemCount: list.ItemCount,
      template: list.BaseTemplate,
      webUrl: list.DefaultViewUrl,
      lastModifiedDateTime: list.LastItemModifiedDate
    }));
  }
}
