import * as React from 'react';
import styles from './SiteLists.module.scss';
import type { ISiteList, ISiteListSource } from '../../services/site-lists/ISiteListSource';

const NOT_PROVIDED_TEXT = 'Not provided';
const DATE_FORMAT_OPTIONS: Intl.DateTimeFormatOptions = {
  timeZone: 'UTC',
  timeZoneName: 'short'
};
const LIST_TEMPLATE_NAMES: Record<number, string> = {
  100: 'List',
  101: 'Document library',
  109: 'Picture library',
  119: 'Wiki library',
  8500: 'Publishing pages'
};

interface ISiteListsProps {
  source: ISiteListSource;
}

/** Форматує дату зміни списку англійською мовою із зазначенням часового поясу. */
function formatListDate(value: string | undefined): string {
  if (!value) {
    return NOT_PROVIDED_TEXT;
  }

  const date = new Date(value);

  return Number.isNaN(date.getTime())
    ? NOT_PROVIDED_TEXT
    : date.toLocaleString('en-US', DATE_FORMAT_OPTIONS);
}

/** Повертає зрозумілу англомовну назву типу списку SharePoint. */
function getListTemplateName(template: number): string {
  return LIST_TEMPLATE_NAMES[template] || `Template ${template}`;
}

/** Відображає таблицю видимих списків поточного SharePoint-сайту. */
export default function SiteLists({ source }: ISiteListsProps): React.ReactElement {
  const [lists, setLists] = React.useState<ISiteList[]>([]);
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');

  /** Завантажує списки сайту після монтування компонента або зміни джерела. */
  React.useEffect(() => {
    let isCurrent = true;

    /** Оновлює стан лише якщо компонент ще відображається на сторінці. */
    const loadSiteLists = async (): Promise<void> => {
      setLoading(true);
      setError('');

      try {
        const result = await source.getSiteLists();

        if (isCurrent) {
          setLists(result);
        }
      } catch {
        if (isCurrent) {
          setLists([]);
          setError('Unable to load SharePoint lists. Check your permissions and try again.');
        }
      } finally {
        if (isCurrent) {
          setLoading(false);
        }
      }
    };

    loadSiteLists().catch(() => undefined);

    /** Позначає результат запиту неактуальним після демонтування компонента. */
    return () => {
      isCurrent = false;
    };
  }, [source]);

  return (
    <section className={styles.section} aria-label="SharePoint lists" lang="en">
      <h2 className={styles.heading}>SharePoint lists</h2>

      {loading && (
        <p className={styles.feedback} role="status">
          Loading SharePoint lists…
        </p>
      )}

      {error && (
        <p className={styles.error} role="alert">
          {error}
        </p>
      )}

      {!loading && !error && lists.length === 0 && (
        <p className={styles.feedback}>No visible SharePoint lists were found.</p>
      )}

      {lists.length > 0 && (
        <div className={styles.tableWrapper}>
          <table className={styles.table}>
            <thead>
              <tr>
                <th scope="col">Name</th>
                <th scope="col">Type</th>
                <th scope="col">Items</th>
                <th scope="col">Last modified</th>
              </tr>
            </thead>
            <tbody>
              {lists.map(list => (
                <tr key={list.id}>
                  <td>
                    {list.webUrl ? (
                      <a href={list.webUrl} target="_blank" rel="noreferrer">
                        {list.title}
                      </a>
                    ) : list.title}
                  </td>
                  <td>{getListTemplateName(list.template)}</td>
                  <td>{list.itemCount}</td>
                  <td>{formatListDate(list.lastModifiedDateTime)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </section>
  );
}
