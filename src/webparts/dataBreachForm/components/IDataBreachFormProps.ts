import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IDataBreachFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  spContext: WebPartContext;
  logicAppUrl?: string;
  senderEmail?: string;
  sendMailApiKey?: string;
  appendCommLogUrl?: string;
  appendCommLogApiKey?: string;
  updatePanelRecommendationsUrl?: string;
  updatePanelRecommendationsApiKey?: string;
  uploadAttachmentUrl?: string;
  uploadAttachmentApiKey?: string;
}
