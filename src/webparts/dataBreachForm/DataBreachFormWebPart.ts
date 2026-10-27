import * as React from 'react';
import * as ReactDom from 'react-dom';
// eslint-disable-next-line @typescript-eslint/no-var-requires
require('govuk-frontend/dist/govuk/govuk-frontend.min.css');
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'DataBreachFormWebPartStrings';
import DataBreachForm from './components/DataBreachForm';
import { IDataBreachFormProps } from './components/IDataBreachFormProps';
import { initSharePoint, checkIsAdmin } from './services/SharePointService';
import { DEFAULT_ENDPOINT_URLS } from './config';

export interface IDataBreachFormWebPartProps {
  description: string;
  logicAppUrl: string;
  senderEmail: string;
  sendMailApiKey: string;
  appendCommLogUrl: string;
  appendCommLogApiKey: string;
  updatePanelRecommendationsUrl: string;
  updatePanelRecommendationsApiKey: string;
  uploadAttachmentUrl: string;
  uploadAttachmentApiKey: string;
}

export default class DataBreachFormWebPart extends BaseClientSideWebPart<IDataBreachFormWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IDataBreachFormProps> = React.createElement(
      DataBreachForm,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        spContext: this.context,
        logicAppUrl: this.properties.logicAppUrl || DEFAULT_ENDPOINT_URLS.sendMail,
        senderEmail: this.properties.senderEmail,
        sendMailApiKey: this.properties.sendMailApiKey,
        appendCommLogUrl: this.properties.appendCommLogUrl || DEFAULT_ENDPOINT_URLS.appendCommLog,
        appendCommLogApiKey: this.properties.appendCommLogApiKey,
        updatePanelRecommendationsUrl: this.properties.updatePanelRecommendationsUrl || DEFAULT_ENDPOINT_URLS.updatePanelRecommendations,
        updatePanelRecommendationsApiKey: this.properties.updatePanelRecommendationsApiKey,
        uploadAttachmentUrl: this.properties.uploadAttachmentUrl || DEFAULT_ENDPOINT_URLS.uploadAttachment,
        uploadAttachmentApiKey: this.properties.uploadAttachmentApiKey,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    initSharePoint(this.context);
    const envPromise = this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
    if (this.context.isServedFromLocalhost) {
      return envPromise;
    }
    this._hideTopChrome();
    this._hideSideNav();
    // Fire-and-forget: the admin check only toggles cosmetic chrome tweaks,
    // so it must never gate the webpart's initial render. Awaiting it here
    // previously meant a slow/hung group-membership lookup (more likely for
    // admin accounts, which tend to belong to far more SharePoint/AD groups)
    // could leave the whole page blank indefinitely.
    // applyChromeVisibility is also passed as the late-result callback, so
    // if the check times out and assumes non-admin (hiding the nav/top
    // chrome), the real result arriving afterwards can un-hide it — no
    // refresh needed.
    const applyChromeVisibility = (isAdmin: boolean): void => {
      if (isAdmin) {
        this._showSideNav();
        // Owners (Data Breach Owners / site admins) get the full SharePoint
        // chrome back — suite nav, site header, command bar, authoring
        // controls — so they can use SharePoint's own page-editing tools
        // (add web parts, edit the page, etc.) instead of it being hidden
        // for everyone the way it is for regular submitters/reviewers.
        this._showTopChrome();
      } else {
        this._hideSideNav();
        this._hideTopChrome();
      }
    };
    checkIsAdmin(applyChromeVisibility)
      .then(applyChromeVisibility)
      .catch(() => { /* non-critical: chrome just stays visible */ });
    return envPromise;
  }

  private _hideTopChrome(): void {
    const styleId = 'db-hide-top-chrome';
    if (document.getElementById(styleId)) return;
    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = [
      '#SuiteNavWrapper { display: none !important; }',
      '#spSiteHeader { display: none !important; }',
      '#spCommandBar { display: none !important; }',
      '[class*="commandBarWrapper"] { display: none !important; }',
      '[data-automation-id="pageCommandBar"] { display: none !important; }',
      '[class*="authoringControls"] { display: none !important; }',
      '.ms-HideForPrint { display: none !important; }',
      'body { margin-top: 0 !important; }',
    ].join(' ');
    document.head.appendChild(style);
  }

  private _showTopChrome(): void {
    const style = document.getElementById('db-hide-top-chrome');
    if (style) style.parentNode?.removeChild(style);
  }

  private _hideSideNav(): void {
    const styleId = 'db-hide-side-nav';
    if (document.getElementById(styleId)) return;
    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = [
      '#spLeftNav { display: none !important; }',
      'div[data-automation-id="navigationPanel"] { display: none !important; }',
    ].join(' ');
    document.head.appendChild(style);
  }

  private _showSideNav(): void {
    const style = document.getElementById('db-hide-side-nav');
    if (style) style.parentNode?.removeChild(style);
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

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
            },
            {
              groupName: 'Email Settings',
              groupFields: [
                PropertyPaneTextField('senderEmail', {
                  label: 'Sender email address',
                  description: 'The email address emails will be sent from (your email for testing, shared mailbox for production)'
                }),
                PropertyPaneTextField('logicAppUrl', {
                  label: 'Logic App URL',
                  description: 'The APIM URL for the sendmail operation'
                }),
                PropertyPaneTextField('sendMailApiKey', {
                  label: 'Logic App API key',
                  description: 'The APIM subscription key for the sendmail operation'
                }),
                PropertyPaneTextField('appendCommLogUrl', {
                  label: 'Communication log API URL',
                  description: 'The APIM URL for the append-comm-log operation. Lets submitter/person responsible reply without needing Edit rights on the raw list item.'
                }),
                PropertyPaneTextField('appendCommLogApiKey', {
                  label: 'Communication log API key',
                  description: 'The APIM subscription key for the append-comm-log operation'
                }),
                PropertyPaneTextField('updatePanelRecommendationsUrl', {
                  label: 'Panel recommendations API URL',
                  description: 'The APIM URL for the update-panel-recommendations operation. Lets reviewers edit the Panel recommendations section without needing Edit rights on the raw list item.'
                }),
                PropertyPaneTextField('updatePanelRecommendationsApiKey', {
                  label: 'Panel recommendations API key',
                  description: 'The APIM subscription key for the update-panel-recommendations operation'
                }),
                PropertyPaneTextField('uploadAttachmentUrl', {
                  label: 'Attachment upload API URL',
                  description: 'The APIM URL for the upload-attachment operation. Lets submitters attach files to their own report without needing Edit rights on the raw list item.'
                }),
                PropertyPaneTextField('uploadAttachmentApiKey', {
                  label: 'Attachment upload API key',
                  description: 'The APIM subscription key for the upload-attachment operation'
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
