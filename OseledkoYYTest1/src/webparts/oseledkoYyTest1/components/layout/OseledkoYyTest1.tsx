import * as React from 'react';
import Products from '../products/Products';
import SiteInfo from '../site-info/SiteInfo';
import SiteLists from '../site-lists/SiteLists';
import UserCreate from '../user-create/UserCreate';
import UserLookup from '../user-lookup/UserLookup';
import styles from './OseledkoYyTest1.module.scss';
import type { IOseledkoYyTest1Props } from './IOseledkoYyTest1Props';

/** Розміщує незалежні компоненти товарів і відомостей сайту в одній вебчастині. */
export default function OseledkoYyTest1({
  description,
  source,
  siteInfoSource,
  siteListSource,
  userLookupSource,
  userCreateSource
}: IOseledkoYyTest1Props): React.ReactElement {
  return (
    <div className={styles.layout}>
      <SiteInfo source={siteInfoSource} />
      <SiteLists source={siteListSource} />
      <UserLookup source={userLookupSource} />
      <UserCreate source={userCreateSource} />
      <Products description={description} source={source} />
    </div>
  );
}
