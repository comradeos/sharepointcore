import * as React from 'react';
import Products from '../products/Products';
import SiteInfo from '../site-info/SiteInfo';
import styles from './OseledkoYyTest1.module.scss';
import type { IOseledkoYyTest1Props } from './IOseledkoYyTest1Props';

/** Розміщує незалежні компоненти товарів і відомостей сайту в одній вебчастині. */
export default function OseledkoYyTest1({ description, source }: IOseledkoYyTest1Props): React.ReactElement {
  return (
    <div className={styles.layout}>
      <SiteInfo />
      <Products description={description} source={source} />
    </div>
  );
}
