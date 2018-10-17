import { AadHttpClient } from "@microsoft/sp-http";
import { IWorkPointApp } from '../../config/webapi-config.json';

export interface IEMMEntry {
  ConversationID: string;
  Subject: string;
  JournalizeDate: string; // Date?
  OrigDate: string;
  Direction: boolean; // What is this?
  //Author: string;
  From: string;
  To: string;
  CC: string;
  HasAttachment: boolean;
  Attachment: boolean;
  JournalInfo: string;
  Category:  string;
}

interface IEMMResponse {
  Data:IEMMEntry[];
}

/**
 * 
 * @param uniqueId String unique id for the entity we want to search for. Found on the business module list for each entity.
 * @param app Desired EMM webservice app to fetch results from.
 * @param solutionUrl The root site collection URL of the WorkPoint 365 solution.
 * @param amount How many results to fetch.
 * @param aadHttpClient AadHttpClient instance from Application customizer.
 */
export const getJournalItemsEMM = async (uniqueId:string, app:IWorkPointApp, solutionUrl:string, amount:number, aadHttpClient:AadHttpClient):Promise<IEMMEntry[]> => {

  let emmJournalData:IEMMEntry[] = null;

  try {

    /**
     * TODO: REMOVE
     * DEBUG
     */
    //uniqueId = "e069c4de-ddc7-4abf-a419-e3971d1cac36";
    //solutionUrl= "https://wp365dev.sharepoint.com/sites/webpartdevelopment";

    let requestHeaders: HeadersInit = new Headers();
    requestHeaders.append("WorkPoint365Url", solutionUrl);
    requestHeaders.append('Content-Type', 'application/json');
  
    let requestBody: any = {
      JournalID: uniqueId,
      PageSize: amount,
      GetMailsAcrossFolders: true,
      IncludeJournalInfo: true,
      SortDirection: 0,
      AdvancedSearch: true,
      SearchCriteria: { SearchSubject: null }
    };
  
    const emmResponse = await aadHttpClient.post(`${app.url}/SearchJournalItems`, AadHttpClient.configurations.v1, { headers: requestHeaders, body: JSON.stringify(requestBody) });
  
    if (emmResponse.ok) {
      let emmResultsResponse:IEMMResponse = await emmResponse.json();
      emmJournalData = emmResultsResponse.Data;
    } else {
      throw "Could not fetch EMM emails for this journal.";
    }

  } catch (exception) {
    emmJournalData = [];
  }

  return emmJournalData;
};