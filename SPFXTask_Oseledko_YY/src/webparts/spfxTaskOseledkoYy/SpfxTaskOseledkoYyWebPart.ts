import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPeoplePickerContext } from '@pnp/spfx-controls-react/lib/PeoplePicker';

import ServiceRequestsPage from './components/ServiceRequestsPage';
import { IServiceRequestsPageProps } from './components/IServiceRequestsPageProps';

// підключає екран заявок до сторінки sharepoint
export default class SpfxTaskOseledkoYyWebPart extends BaseClientSideWebPart<Record<string, never>> {
  // передає екрану клієнт sharepoint та адресу поточного сайту
  public render(): void {
    // узгоджує типи різних версій пакетів sharepoint
    const peoplePickerContext = {
      absoluteUrl: this.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.context.msGraphClientFactory,
      spHttpClient: this.context.spHttpClient
    } as unknown as IPeoplePickerContext;
    const element: React.ReactElement<IServiceRequestsPageProps> = React.createElement(
      ServiceRequestsPage,
      {
        spHttpClient: this.context.spHttpClient,
        webUrl: this.context.pageContext.web.absoluteUrl,
        peoplePickerContext,
        currentUserEmail: this.context.pageContext.user.email
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // оновлює колір тексту після зміни теми сайту
  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    const bodyText = currentTheme?.semanticColors?.bodyText;
    if (bodyText) {
      this.domElement.style.setProperty('--bodyText', bodyText);
    }
  }

  // прибирає екран після видалення вебчастини зі сторінки
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // повертає версію збережених властивостей вебчастини
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  // повертає порожню панель налаштувань бо екран не має параметрів
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [] };
  }
}
