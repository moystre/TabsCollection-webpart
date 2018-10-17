
/**
 * Valid Wizard message event types
 */
export const wizardResultEventTypes:string[] = [
  'docProvisioningWizard',
  'ChangeStageWizard',
  'RelationsWizard',
  'Wizard',
  'templatemanagement'
];

export interface IWorkPointMessageData {
  type: "expressPanel"|"favorites"|"stage";
}

export interface IIncomingExpressPanelMessage extends IWorkPointMessageData {
  action: "open"|"close"|"signalReady"|"toggle"|"openUrl"|"checkIfWorkPointSettingsAreOutOfSync";
  redirectUrl?: string;
  SettingsChangedUTCTimeStamp?: string;
}

export interface IIncomingFavoriteMessage extends IWorkPointMessageData {
  action: "RemovedFavorite"|"ClearFavoriteCache"|"AddedFavorite";
}

export interface IOutGoingExpressPanelMessage {
  method: "ToggleExpress";
  expressVisibleInWorkPoint: boolean;
  tab: "entities"|"documents"|"favorites";
  eventKey: number;
}

export interface IWizardStartMessage {
  url: string;
  type: "workpointwizard";
}

export interface IWizardMessageData {
  action: "closeAndRefresh"|"openClient"|"openOnline"|"closeDialogAndOpenFile";
  type: string;
  redirectUrl?: string;
  templateServerRelativeUrl?: string;
}

export interface IBasicStageFilterMessage extends IWorkPointMessageData {
  action: "addStageFilter"|"removeStageFilter";
  entityListId:string;
  entityItemId:number;
}

export interface IAddStageFilterMessage extends IBasicStageFilterMessage {
  stage: string;
}