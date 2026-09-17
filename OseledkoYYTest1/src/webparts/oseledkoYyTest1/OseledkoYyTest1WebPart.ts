import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'OseledkoYyTest1WebPartStrings';
import OseledkoYyTest1 from './components/layout/OseledkoYyTest1';
import { IOseledkoYyTest1Props } from './components/layout/IOseledkoYyTest1Props';
import { DummyJsonProductSource } from './services/products/DummyJsonProductSource';
import { PnpSiteInfoSource } from './services/site-info/PnpSiteInfoSource';
import type { ISiteInfoSource } from './services/site-info/ISiteInfoSource';
import { PnpSiteListSource } from './services/site-lists/PnpSiteListSource';
import type { ISiteListSource } from './services/site-lists/ISiteListSource';
import { PnpUserLookupSource } from './services/user-lookup/PnpUserLookupSource';
import type { IUserLookupSource } from './services/user-lookup/IUserLookupSource';
import { PnpUserCreateSource } from './services/user-create/PnpUserCreateSource';
import type { IUserCreateSource } from './services/user-create/IUserCreateSource';

export interface IOseledkoYyTest1WebPartProps {
  description: string;
}

/** Керує життєвим циклом вебчастини та передає джерело даних React-компоненту. */
export default class OseledkoYyTest1WebPart extends BaseClientSideWebPart<IOseledkoYyTest1WebPartProps> {

  private readonly _productSource = new DummyJsonProductSource();
  private _siteInfoSource!: ISiteInfoSource;
  private _siteListSource!: ISiteListSource;
  private _userLookupSource!: IUserLookupSource;
  private _userCreateSource!: IUserCreateSource;

  /** Створює джерело сайту після отримання SharePoint-контексту вебчастини. */
  protected async onInit(): Promise<void> {
    await super.onInit();
    this._siteInfoSource = new PnpSiteInfoSource(this.context);
    this._siteListSource = new PnpSiteListSource(this.context);
    this._userLookupSource = new PnpUserLookupSource(this.context);
    this._userCreateSource = new PnpUserCreateSource(this.context);
  }

  /** Відображає контейнер із компонентами товарів і відомостей сайту. */
  public render(): void {
    const element: React.ReactElement<IOseledkoYyTest1Props> = React.createElement(
      OseledkoYyTest1,
      {
        description: this.properties.description,
        source: this._productSource,
        siteInfoSource: this._siteInfoSource,
        siteListSource: this._siteListSource,
        userLookupSource: this._userLookupSource,
        userCreateSource: this._userCreateSource
      }
    );

    ReactDom.render(element, this.domElement);
  }

  /** Оновлює спільні кольори обох компонентів відповідно до теми SharePoint. */
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--bodyBackground', semanticColors.bodyBackground || null);
      this.domElement.style.setProperty('--mutedText', semanticColors.bodySubtext || null);
      this.domElement.style.setProperty('--border', semanticColors.bodyDivider || null);
      this.domElement.style.setProperty('--rowBackground', semanticColors.bodyStandoutBackground || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  /** Демонтує React-компонент і запускає очищення його ресурсів. */
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /** Повертає версію формату збережених властивостей вебчастини. */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /** Налаштовує панель редагування заголовка каталогу. */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
