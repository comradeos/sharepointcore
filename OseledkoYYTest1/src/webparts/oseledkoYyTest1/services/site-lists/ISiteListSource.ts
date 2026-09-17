/** Описує список SharePoint, потрібний для відображення у віджеті. */
export interface ISiteList {
  id: string;
  title: string;
  itemCount: number;
  template: number;
  webUrl?: string;
  lastModifiedDateTime?: string;
}

/** Визначає спільний контракт джерела списків поточного SharePoint-сайту. */
export interface ISiteListSource {
  /** Повертає видимі списки сайту в алфавітному порядку. */
  getSiteLists(): Promise<ISiteList[]>;
}
