import * as React from 'react';
import styles from './UserLookup.module.scss';
import type { IUserListItem, IUserLookupSource } from '../../services/user-lookup/IUserLookupSource';

const JSON_INDENTATION = 2;
const USER_TITLE_FIELD = 'Title';

interface IUserLookupProps {
  source: IUserLookupSource;
}

/** Повертає введене значення лише з цифр для поля ідентифікатора. */
function keepDigits(value: string): string {
  return value.replace(/\D/g, '');
}

/** Повертає назву запису Users, якщо поле Title доступне. */
function getUserTitle(user: IUserListItem): string | undefined {
  const title = user.fields[USER_TITLE_FIELD];

  return typeof title === 'string' && title.trim() ? title : undefined;
}

/** Відображає форму пошуку запису в списку Users за числовим ID. */
export default function UserLookup({ source }: IUserLookupProps): React.ReactElement {
  const [userId, setUserId] = React.useState('');
  const [user, setUser] = React.useState<IUserListItem>();
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');
  const [notFound, setNotFound] = React.useState(false);

  /** Очищує нецифрові символи з введення та скидає попередній результат. */
  const changeUserId = (event: React.ChangeEvent<HTMLInputElement>): void => {
    setUserId(keepDigits(event.target.value));
    setError('');
    setNotFound(false);
  };

  /** Шукає запис Users за введеним ID і відображає результат запиту. */
  const submit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();

    const id = Number(userId);

    if (!Number.isSafeInteger(id) || id < 1) {
      setUser(undefined);
      setNotFound(false);
      setError('Enter a valid numeric user ID.');
      return;
    }

    setLoading(true);
    setError('');
    setNotFound(false);
    setUser(undefined);

    try {
      const result = await source.getUserById(id);

      if (result) {
        setUser(result);
      } else {
        setNotFound(true);
      }
    } catch {
      setError('Unable to load the user record. Check your permissions and try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <section className={styles.section} aria-label="User lookup" lang="en">
      <h2 className={styles.heading}>User lookup</h2>

      <form className={styles.form} onSubmit={submit} noValidate>
        <label className={styles.label}>
          User ID
          <input
            className={styles.input}
            type="text"
            value={userId}
            onChange={changeUserId}
            inputMode="numeric"
            pattern="[0-9]*"
            autoComplete="off"
            disabled={loading}
            aria-required="true"
          />
        </label>

        <button className={styles.button} type="submit" disabled={loading}>
          Get user
        </button>
      </form>

      {loading && (
        <p className={styles.feedback} role="status">
          Loading user record…
        </p>
      )}

      {error && (
        <p className={styles.error} role="alert">
          {error}
        </p>
      )}

      {notFound && (
        <p className={styles.feedback} role="status">
          No user record was found for this ID.
        </p>
      )}

      {user && (
        <div className={styles.result}>
          <p className={styles.feedback} role="status">
            User record loaded.
          </p>
          <dl className={styles.details}>
            <dt>User ID</dt>
            <dd>{user.id}</dd>
            {getUserTitle(user) && (
              <>
                <dt>Title</dt>
                <dd>{getUserTitle(user)}</dd>
              </>
            )}
          </dl>
          <details className={styles.json}>
            <summary>JSON response</summary>
            <pre>{JSON.stringify(user.fields, undefined, JSON_INDENTATION)}</pre>
          </details>
        </div>
      )}
    </section>
  );
}
