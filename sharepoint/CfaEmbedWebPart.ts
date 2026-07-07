import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

/**
 * CFA Transfer embed web part.
 *
 * Hosts the CFA app in an iframe AND — because this code runs INSIDE the SharePoint page — it can
 * read the signed-in user (no login) and hand it to the app. It passes the display name two ways:
 *   1. as `?user=` on the iframe URL (the app reads it on load), and
 *   2. via postMessage once the iframe has loaded (belt-and-suspenders).
 *
 * The app side already listens for both (URL param `user` and postMessage `{type:'cfa-user', user}`).
 */

export interface ICfaEmbedWebPartProps {
  appUrl: string;
  height: string;
}

export default class CfaEmbedWebPart extends BaseClientSideWebPart<ICfaEmbedWebPartProps> {

  public render(): void {
    const base = (this.properties.appUrl || 'https://dynaautomation.pythonanywhere.com/').trim();
    const user = this.context.pageContext.user.displayName || '';
    const email = this.context.pageContext.user.email
      || this.context.pageContext.user.loginName || '';

    const sep = base.indexOf('?') >= 0 ? '&' : '?';
    const src = base + sep + 'user=' + encodeURIComponent(user);
    const heightPx = (this.properties.height || '1000').replace(/[^0-9]/g, '') || '1000';

    this.domElement.innerHTML =
      '<iframe id="cfaFrame" src="' + src + '" ' +
      'style="width:100%;height:' + heightPx + 'px;border:0;" ' +
      'title="CFA Transfer" allow="clipboard-read; clipboard-write"></iframe>';

    // Also postMessage the user once the iframe loads (covers the case where the URL param is stripped).
    const frame = this.domElement.querySelector('#cfaFrame') as HTMLIFrameElement | null;
    if (frame) {
      frame.addEventListener('load', () => {
        try {
          const origin = new URL(base).origin;
          frame.contentWindow?.postMessage({ type: 'cfa-user', user: user, email: email }, origin);
        } catch (e) { /* ignore */ }
      });
    }
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'CFA Transfer embed settings' },
        groups: [{
          groupName: 'Settings',
          groupFields: [
            PropertyPaneTextField('appUrl', {
              label: 'App URL',
              description: 'e.g. https://dynaautomation.pythonanywhere.com/'
            }),
            PropertyPaneTextField('height', { label: 'Height in pixels (e.g. 1000)' })
          ]
        }]
      }]
    };
  }
}
