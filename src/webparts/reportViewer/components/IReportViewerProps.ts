import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICommLogEntry, IPanelRecommendationsFields } from '../../dataBreachForm/services/SharePointService';

// base64 is the raw file content (no "data:...;base64," prefix) — each
// sendEmailFromMailbox implementation is responsible for putting it wherever
// its own transport (Graph API / Logic App) expects attachment bytes.
export interface IEmailAttachment {
  filename: string;
  contentType: string;
  base64: string;
}

export interface IReportViewerProps {
  spContext: WebPartContext;
  sharedMailbox?: string;
  sendEmailFromMailbox?: (to: string, subject: string, body: string, attachment?: IEmailAttachment) => Promise<void>;
  targetItemId?: number | null;
  onTargetItemOpened?: () => void;
  // Named individuals (submitter, person responsible) only hold Read at the
  // item level, not Edit — see appendCommLogEntryElevated in
  // SharePointService.ts. When provided, their reply flow routes through
  // this instead of the direct item.update() used by admins. Falls back to
  // the direct path if omitted, so a webpart that hasn't wired this up yet
  // degrades to the old behaviour rather than breaking outright.
  appendCommLogEntryElevated?: (itemId: number, callerEmail: string, entry: ICommLogEntry) => Promise<ICommLogEntry[]>;
  // Reviewers only hold Read at the item level too — see
  // updatePanelRecommendationsElevated in SharePointService.ts. When
  // provided, saving the "Panel recommendations" section as a reviewer
  // (not an admin) routes through this instead of the direct item.update().
  updatePanelRecommendationsElevated?: (itemId: number, callerEmail: string, fields: IPanelRecommendationsFields) => Promise<void>;
  // When true, communication notices ("keep everyone in the loop" copies)
  // aren't filtered to exclude the message's direct sender/recipient — lets
  // testing with a single email across all three roles still show the
  // notice mechanism firing, instead of always computing zero recipients.
  testMode?: boolean;
}
