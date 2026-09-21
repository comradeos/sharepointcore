// значення статусів які зберігаються у sharepoint
export const requestStatuses = {
  new: 'Нова',
  inProgress: 'В роботі',
  resolved: 'Вирішена',
  closed: 'Закрита'
} as const;

// значення пріоритетів які зберігаються у sharepoint
export const requestPriorities = {
  low: 'Низький',
  medium: 'Середній',
  high: 'Високий'
} as const;
