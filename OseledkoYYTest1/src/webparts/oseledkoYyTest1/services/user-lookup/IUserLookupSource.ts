/** Описує запис зі списку Users, який можна показати у віджеті. */
export interface IUserListItem {
  id: number;
  fields: Record<string, unknown>;
}

/** Визначає спільний контракт пошуку запису за ідентифікатором. */
export interface IUserLookupSource {
  /** Повертає запис зі списку Users або undefined, якщо запису з таким ID не існує. */
  getUserById(id: number): Promise<IUserListItem | undefined>;
}
