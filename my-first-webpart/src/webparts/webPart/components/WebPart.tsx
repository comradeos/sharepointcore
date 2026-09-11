import * as React from 'react';
import styles from './WebPart.module.scss';
import type { IWebPartProps } from './IWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import welcomeDark from '../assets/welcome-dark.png';
import welcomeLight from '../assets/welcome-light.png';

export default class WebPart extends React.Component<IWebPartProps> {
  public render(): React.ReactElement<IWebPartProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.webPart}`}>
        <div className={styles.welcome}>
          <h2>Well done, {escape(userDisplayName)}!</h2>
          <div>{environmentMessage}</div>
          <div>Web part property value: <strong>{escape(description)}</strong></div>
        </div>
      </section>
    );
  }
}
