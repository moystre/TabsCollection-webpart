import pnp, { SearchQuery, SearchResults } from 'sp-pnp-js/lib/pnp';

export enum JournalItemType {
  All,
  Documents,
  Emails,
  Images,
  Events,
  Tasks
}

/**
 * Searches a given string, using PnP search API.
 * 
 * @param query A SharePoint search query string.
 * @param rowLimit Amount of search results to return
 * @param selectProperties Optional. Comma-separated string of properties to select in search.
 * @param refiners Optional. Comma-separated string of refiners to return in the search result.
 */
const search = async (query:string, rowLimit:number, selectProperties?:string, refiners?:string):Promise<SearchResults> => {
 
  try {

    const searchResults:SearchResults = await pnp.sp.search(<SearchQuery>{
      Querytext: query,
      RowLimit: rowLimit,
      ...(refiners && { Refiners: refiners }),
      ...(selectProperties && {SelectProperties: selectProperties.split(",")})
    });

    return searchResults;
    
  } catch (exception) {
    // TODO: Show?
    console.warn(exception);
  }
};

/**
 * Fetches journal items. Not to confuse with EMM Journals, this fetches latest modified documents and items using SharePoint Search.
 * 
 * @param wpItemLocation String representation of an entity structure in WorkPoint 365. Can also be the solution wpItemLocation, with no entity information.
 * @param amount Amount of search results to return
 * @param journalItemTypes A list of journal item types to fetch in the request. Defaults to all.
 * @param selectProperties Optional. Comma-separated string of properties to select in search.
 * @param refiners Optional. Comma-separated string of refiners to return in the search result.
 */
export const fetchJournalItems = async (wpItemLocation:string, amount:number, journalItemTypes:JournalItemType[] = [JournalItemType.All], selectProperties?:string, refiners?:string):Promise<SearchResults> => {

  try {

    // get items by using SP Rest API
    let defaultCondition = `wpItemLocationOWSTEXT:${wpItemLocation}*`;

    /**
     * Build conditions for search
     */
    let searchCondition = [];

    // Documents
    if (journalItemTypes.indexOf(JournalItemType.All) !== -1 || journalItemTypes.indexOf(JournalItemType.Documents) !== -1) {
      searchCondition.push('((FileExtension:doc OR FileExtension:docx OR FileExtension:xls OR FileExtension:xlsx OR FileExtension:ppt OR FileExtension:pptx OR FileExtension:pdf) (IsDocument:"True" OR contentclass:"STS_ListItem"))');
    }

    //.msg eml
    if (journalItemTypes.indexOf(JournalItemType.All) !== -1 || journalItemTypes.indexOf(JournalItemType.Emails) !== -1) {
      searchCondition.push('((FileExtension:msg OR FileExtension:eml) (contentclass:"STS_ListItem"))');
    }
    
    // Images
    if (journalItemTypes.indexOf(JournalItemType.All) !== -1 || journalItemTypes.indexOf(JournalItemType.Images) !== -1) {
      searchCondition.push('((FileType:bmp OR FileType:gif OR FileType:jpe OR FileType:jpeg OR FileType:jpg OR FileType:png) (contentclass:"STS_ListItem"))');
    }
    
    // Events
    if (journalItemTypes.indexOf(JournalItemType.All) !== -1 || journalItemTypes.indexOf(JournalItemType.Events) !== -1) {
      searchCondition.push('ContentTypeId:0x010200CE6BB82A30D1F945A39E3CF4C0EC45EF');
    }
    // Tasks
    if (journalItemTypes.indexOf(JournalItemType.All) !== -1 || journalItemTypes.indexOf(JournalItemType.Tasks) !== -1) {
      searchCondition.push('ContentTypeId:0x0108*');
    }
    
    if(searchCondition.length > 0) {
      defaultCondition += ' AND (' + searchCondition.join(' OR ') + ')';
    }

    const results:SearchResults = await search(defaultCondition, amount, selectProperties, refiners);

    return results;

  } catch (exception) {
    console.warn(`Journal exception: ${exception}`);
    return null;
  }
};