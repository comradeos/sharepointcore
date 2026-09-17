import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFx, spfi, type SPFI } from '@pnp/sp';
import '@pnp/sp/items';
import '@pnp/sp/lists';
import type { IUserCreateSource } from './IUserCreateSource';

const USERS_LIST_TITLE = 'Users';

/** Створює записи у списку Users через SharePoint REST API та PnPjs. */
export class PnpUserCreateSource implements IUserCreateSource {
  private readonly _sp: SPFI;

  /** Створює PnPjs-клієнт із поточним контекстом та авторизацією SPFx. */
  public constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  /** Додає новий запис до списку Users та повертає ID створеного елемента. */
  public async createUser(title: string): Promise<number> {
    const createdItem = await this._sp.web.lists
      .getByTitle(USERS_LIST_TITLE)
      .items
      .add({ Title: title });

    return createdItem.Id;
  }
}
