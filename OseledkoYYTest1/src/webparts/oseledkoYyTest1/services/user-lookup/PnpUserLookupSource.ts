import type { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFx, spfi, type SPFI } from '@pnp/sp';
import '@pnp/sp/items';
import '@pnp/sp/lists';
import type { IUserListItem, IUserLookupSource } from './IUserLookupSource';

const USERS_LIST_TITLE = 'Users';

/** Перевіряє, чи є помилка відповіддю SharePoint зі статусом 404. */
function isNotFoundError(error: unknown): boolean {
  return typeof error === 'object'
    && error !== null
    && 'status' in error
    && (error as { status?: unknown }).status === 404;
}

/** Отримує записи за ID зі списку Users через SharePoint REST API та PnPjs. */
export class PnpUserLookupSource implements IUserLookupSource {
  private readonly _sp: SPFI;

  /** Створює PnPjs-клієнт із поточним контекстом та авторизацією SPFx. */
  public constructor(context: WebPartContext) {
    this._sp = spfi().using(SPFx(context));
  }

  /** Повертає запис Users за його числовим ID або undefined для відсутнього запису. */
  public async getUserById(id: number): Promise<IUserListItem | undefined> {
    try {
      const fields = await this._sp.web.lists
        .getByTitle(USERS_LIST_TITLE)
        .items
        .getById(id)<Record<string, unknown>>();

      return {
        id,
        fields
      };
    } catch (error) {
      if (isNotFoundError(error)) {
        return undefined;
      }

      throw error;
    }
  }
}
