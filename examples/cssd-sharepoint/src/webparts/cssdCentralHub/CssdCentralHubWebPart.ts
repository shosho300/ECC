import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneLabel } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';

import * as strings from 'CssdCentralHubWebPartStrings';
import { HubStore, type HttpLike } from './app/store';
import { startHub } from './app/hubApp';
import logoUrl from './assets/logo.png';
import heroUrl from './assets/hero.jpg';

export default class CssdCentralHubWebPart extends BaseClientSideWebPart<Record<string, never>> {
  private _dispose: (() => void) | undefined;

  public render(): void {
    // The hub manages its own DOM; render once and keep its state across property pane refreshes.
    if (this._dispose) return;
    const client = this.context.spHttpClient;
    const config = SPHttpClient.configurations.v1;
    const http: HttpLike = {
      get: url => client.get(url, config),
      post: (url, body, headers) => client.post(url, config, { body, headers })
    };
    const web = this.context.pageContext.web;
    const store = new HubStore(http, web.absoluteUrl, web.serverRelativeUrl);
    this._dispose = startHub({
      root: this.domElement,
      store,
      userName: this.context.pageContext.user.displayName,
      logoUrl,
      heroUrl
    });
  }

  protected onDispose(): void {
    if (this._dispose) this._dispose();
    this._dispose = undefined;
    super.onDispose();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.SetupGroupName,
              groupFields: [PropertyPaneLabel('setupHelp', { text: strings.SetupHelp })]
            }
          ]
        }
      ]
    };
  }
}
