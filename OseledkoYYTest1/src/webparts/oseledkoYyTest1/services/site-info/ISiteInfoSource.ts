/** Описує дані поточного SharePoint-сайту, потрібні для відображення у віджеті. */
export interface ISiteInfo {
  id: string;
  name: string;
  displayName: string;
  webUrl: string;
  description?: string;
  createdDateTime?: string;
  lastModifiedDateTime?: string;
  hostname?: string;
}

/** Визначає спільний контракт джерела відомостей про SharePoint-сайт. */
export interface ISiteInfoSource {
  /** Повертає відомості про сайт, у контексті якого відкрито вебчастину. */
  getSiteInfo(): Promise<ISiteInfo>;
}
