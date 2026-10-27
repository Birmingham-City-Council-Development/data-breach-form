import * as React from 'react';
import * as ReactDom from 'react-dom';
// eslint-disable-next-line @typescript-eslint/no-var-requires
require('govuk-frontend/dist/govuk/govuk-frontend.min.css');
import { Version } from '@microsoft/sp-core-library';
import { type IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import ReportViewer from './components/ReportViewer';
import { IReportViewerProps, IEmailAttachment } from './components/IReportViewerProps';
import { initSharePoint, appendCommLogEntryElevated, updatePanelRecommendationsElevated } from '../dataBreachForm/services/SharePointService';
import { DEFAULT_ENDPOINT_URLS } from '../dataBreachForm/config';

interface IReportViewerWebPartProps {
  sharedMailbox: string;
  appendCommLogUrl: string;
  appendCommLogApiKey: string;
  updatePanelRecommendationsUrl: string;
  updatePanelRecommendationsApiKey: string;
}

export default class ReportViewerWebPart extends BaseClientSideWebPart<IReportViewerWebPartProps> {
  private _graphClient: MSGraphClientV3 | undefined;

  protected async onInit(): Promise<void> {
    initSharePoint(this.context);
    try {
      this._graphClient = await this.context.msGraphClientFactory.getClient('3');
    } catch (_e) {
      // Mail.Send permission not yet approved — email features will be unavailable
    }
    return super.onInit();
  }

  public render(): void {
    const graphClient = this._graphClient;
    const sendEmailFromMailbox = graphClient
      ? async (to: string, subject: string, body: string, attachment?: IEmailAttachment): Promise<void> => {
          const mailbox = this.properties.sharedMailbox || '';
          if (!mailbox) throw new Error('No shared mailbox configured. Edit the webpart settings.');
          await graphClient.api(`/users/${mailbox}/sendMail`).post({
            message: {
              subject,
              body: { contentType: 'Text', content: body },
              toRecipients: [{ emailAddress: { address: to } }],
              // Graph's sendMail natively supports file attachments — unlike
              // the Logic App path (DataBreachForm.tsx's sendEmailFromMailbox),
              // this needs no external/server-side change to work.
              attachments: attachment
                ? [{
                    '@odata.type': '#microsoft.graph.fileAttachment',
                    name: attachment.filename,
                    contentType: attachment.contentType,
                    contentBytes: attachment.base64,
                  }]
                : undefined,
            },
            saveToSentItems: true,
          });
        }
      : undefined;

    const appendCommLogUrl = this.properties.appendCommLogUrl || DEFAULT_ENDPOINT_URLS.appendCommLog;
    const appendCommLogApiKey = this.properties.appendCommLogApiKey;
    const updatePanelRecommendationsUrl = this.properties.updatePanelRecommendationsUrl || DEFAULT_ENDPOINT_URLS.updatePanelRecommendations;
    const updatePanelRecommendationsApiKey = this.properties.updatePanelRecommendationsApiKey;
    const element: React.ReactElement<IReportViewerProps> = React.createElement(
      ReportViewer,
      {
        spContext: this.context,
        sharedMailbox: this.properties.sharedMailbox || '',
        sendEmailFromMailbox,
        appendCommLogEntryElevated: appendCommLogUrl
          ? (itemId, callerEmail, entry) => appendCommLogEntryElevated(appendCommLogUrl, appendCommLogApiKey || '', itemId, callerEmail, entry)
          : undefined,
        updatePanelRecommendationsElevated: updatePanelRecommendationsUrl
          ? (itemId, callerEmail, fields) => updatePanelRecommendationsElevated(updatePanelRecommendationsUrl, updatePanelRecommendationsApiKey || '', itemId, callerEmail, fields)
          : undefined,
      }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        groups: [{
          groupFields: [
            PropertyPaneTextField('sharedMailbox', {
              label: 'Shared mailbox email address',
              description: 'Emails requesting additional information will be sent from this address (e.g. dataprotection@example.gov.uk). The mailbox must grant this webpart\'s service account Send As permission.',
            }),
            PropertyPaneTextField('appendCommLogUrl', {
              label: 'Communication log API URL',
              description: 'The APIM URL for the append-comm-log operation. Lets submitter/person responsible reply without needing Edit rights on the raw list item.',
            }),
            PropertyPaneTextField('appendCommLogApiKey', {
              label: 'Communication log API key',
              description: 'The subscription key for the append-comm-log operation.',
            }),
            PropertyPaneTextField('updatePanelRecommendationsUrl', {
              label: 'Panel recommendations API URL',
              description: 'The APIM URL for the update-panel-recommendations operation. Lets reviewers edit the Panel recommendations section without needing Edit rights on the raw list item.',
            }),
            PropertyPaneTextField('updatePanelRecommendationsApiKey', {
              label: 'Panel recommendations API key',
              description: 'The subscription key for the update-panel-recommendations operation.',
            }),
          ],
        }],
      }],
    };
  }
}
