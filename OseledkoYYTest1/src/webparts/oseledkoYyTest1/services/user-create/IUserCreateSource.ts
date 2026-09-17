/** Визначає спільний контракт створення запису у списку Users. */
export interface IUserCreateSource {
  /** Створює запис зі значенням Title та повертає ID нового елемента списку. */
  createUser(title: string): Promise<number>;
}
