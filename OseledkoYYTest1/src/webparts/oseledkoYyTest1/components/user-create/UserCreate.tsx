import * as React from 'react';
import styles from './UserCreate.module.scss';
import type { IUserCreateSource } from '../../services/user-create/IUserCreateSource';

interface IUserCreateProps {
  source: IUserCreateSource;
}

/** Відображає форму створення нового запису в списку Users. */
export default function UserCreate({ source }: IUserCreateProps): React.ReactElement {
  const [title, setTitle] = React.useState('');
  const [createdId, setCreatedId] = React.useState<number>();
  const [loading, setLoading] = React.useState(false);
  const [error, setError] = React.useState('');

  /** Зберігає назву запису та очищує попередні повідомлення форми. */
  const changeTitle = (event: React.ChangeEvent<HTMLInputElement>): void => {
    setTitle(event.target.value);
    setCreatedId(undefined);
    setError('');
  };

  /** Створює запис Users та показує ID, повернений SharePoint. */
  const submit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();

    const trimmedTitle = title.trim();

    if (!trimmedTitle) {
      setCreatedId(undefined);
      setError('Enter a user name.');
      return;
    }

    setLoading(true);
    setCreatedId(undefined);
    setError('');

    try {
      const id = await source.createUser(trimmedTitle);

      setCreatedId(id);
      setTitle('');
    } catch {
      setError('Unable to create the user record. Check your permissions and try again.');
    } finally {
      setLoading(false);
    }
  };

  return (
    <section className={styles.section} aria-label="Create user" lang="en">
      <h2 className={styles.heading}>Create user</h2>

      <form className={styles.form} onSubmit={submit} noValidate>
        <label className={styles.label}>
          User name
          <input
            className={styles.input}
            type="text"
            value={title}
            onChange={changeTitle}
            autoComplete="off"
            disabled={loading}
            aria-required="true"
          />
        </label>

        <button className={styles.button} type="submit" disabled={loading}>
          Create user
        </button>
      </form>

      {loading && (
        <p className={styles.feedback} role="status">
          Creating user record…
        </p>
      )}

      {error && (
        <p className={styles.error} role="alert">
          {error}
        </p>
      )}

      {createdId !== undefined && (
        <p className={styles.feedback} role="status">
          User record created. New ID: {createdId}.
        </p>
      )}
    </section>
  );
}
