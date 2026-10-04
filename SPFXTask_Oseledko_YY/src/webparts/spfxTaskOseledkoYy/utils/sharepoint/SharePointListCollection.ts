export function addListItems<TItem>(items: TItem[], newItems: TItem[]): TItem[] {
  return [...newItems, ...items];
}

// замінює один елемент локального списку за вказаним ключем
export function replaceListItem<TItem, TKey extends keyof TItem>(
  items: TItem[],
  updatedItem: TItem,
  key: TKey
): TItem[] {
  return items.map(item =>
    item[key] === updatedItem[key]
      ? updatedItem
      : item
  );
}

// видаляє один елемент локального списку за вказаним ключем
export function removeListItem<TItem, TKey extends keyof TItem>(
  items: TItem[],
  itemKey: TItem[TKey],
  key: TKey
): TItem[] {
  return items.filter(item => item[key] !== itemKey);
}
