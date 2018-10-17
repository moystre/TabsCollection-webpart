import { BasePermissions, PermissionKind } from "sp-pnp-js/lib/pnp";
import { BusinessModuleEntity } from "./BusinessModule";
import { ListObject } from "./List";

/**
 * Given a solution absolute and relative URL, will return an absolute URL created from the tenant base URL and the provided webRelativeUrl.
 * 
 * @param solutionAbsoluteUrl The WP365 solution absolute URL.
 * @param webRelativeUrl Relative URL part to apply to tenant base URL.
 * 
 * @returns string
 */
export function getWebAbsoluteUrl(solutionAbsoluteUrl: string, webRelativeUrl: string): string {
  const tenantAbsoluteUrl: string = getTenantAbsoluteUrl(solutionAbsoluteUrl);
  return `${tenantAbsoluteUrl}/${webRelativeUrl}`;
}

/**
 * Get the tenant absolute URL given the WorkPoint 365 Solution absolute URL.
 * 
 * @param solutionAbsoluteUrl
 * 
 * @returns string
 */
export function getTenantAbsoluteUrl(solutionAbsoluteUrl: string): string {
  const splitSolutionUrl: string[] = solutionAbsoluteUrl.split("/");
  const tenantAbsoluteUrl: string = splitSolutionUrl.splice(0, splitSolutionUrl.length - 2).join("/");

  return tenantAbsoluteUrl;
}

/**
 * Given a solution absolute URL, returns it in a relative format.
 * 
 * @param solutionAbsoluteUrl The absolute URL of the root site collection in the WorkPoint 365 solution.
 */
export function getSolutionRelativeUrl(solutionAbsoluteUrl:string):string {
  
  const splitSolutionUrl:string[] = solutionAbsoluteUrl.split("/");
  const solutionRelativeUrl:string = `/${splitSolutionUrl.splice(splitSolutionUrl.length-2, splitSolutionUrl.length-1).join("/")}`;

  return solutionRelativeUrl;
}

export function getListUrl(listUrl: string): string {
  const listUrlParts = listUrl.split("/_api");
  const webUrl = listUrlParts[0];
  return webUrl;
}

/**
 * Parses an Express Panel slash divided date value to a normal Date object
 * 
 * @param dateCandidate date string in the following format UTC "2018/5/1/8/15/25" -> year, month, day, hour, minute, second
 * 
 * @returns Date in local time if valid or null if invalid
 */
export function parseSlashDividedDateValue(dateCandidate: string): Date {

  if (typeof dateCandidate === "string" && dateCandidate !== "") {
    // We need to create a DateTime object from an UTC DateTime string
    let splitDate: number[] = dateCandidate.split("/").map(d => parseInt(d));
    // Month is 1 indexed in C# but needs to be 0 indexed in JavaScript
    var date = new Date(Date.UTC(splitDate[0], splitDate[1] - 1, splitDate[2], splitDate[3], splitDate[4], splitDate[5]));

    if (isValidDate(date)) {
      return date;
    }
  }

  return null;
}

export function isValidDate(date: Date): boolean {

  if (Object.prototype.toString.call(date) === "[object Date]") {
    // it is a date
    if (isNaN(date.getTime())) {  // d.valueOf() could also work
      return false;
    } else {
      return true;
    }
  } else {
    return false;
  }
}

export function normalizeURLValue(url: string): string {
  if (typeof url === "string") {

    // Remove html entities
    var tmp = document.createElement("DIV");
    tmp.innerHTML = url;
    var normalizedUrlValue = tmp.textContent || tmp.innerText;

    normalizedUrlValue = normalizedUrlValue.replace(/\s/g, " ");
    return normalizedUrlValue;
  } else {
    return "";
  }
}

export function ensureGuidString(nonDashedGuidString: string): string {

  var lengths: number[] = [8, 4, 4, 4, 12];
  let parts: string[] = [];
  var range: number = 0;
  for (let i: number = 0; i < lengths.length; i++) {
    parts.push(nonDashedGuidString.slice(range, range + lengths[i]));
    range += lengths[i];
  }

  return parts.join('-');
}

export function getListByIdOrUrl(lists: ListObject[], listId?: string, listUrl?: string): ListObject {

  const listIdCandidate: string = typeof listId === "string" && listId !== "" ? listId.toLowerCase() : null;
  const listUrlCandidate: string = typeof listUrl === "string" && listUrl !== "" ? listUrl.toLowerCase() : null;

  if (!lists || (lists && lists.length === 0)) {
    throw "No lists provided";
  }

  if (listIdCandidate === null && listUrlCandidate === null) {
    throw "No identifier arguments provided";
  }

  if (listIdCandidate && listUrlCandidate) {
    throw "Too many identifier arguments provided";
  }

  // Filter out the list for the current MyTools button
  const foundList: ListObject[] = lists.filter(list => {

    // List id matching is simple
    if (listIdCandidate) {

      if (listIdCandidate === list.Id) {
        return true;
      }

      return false;
    }

    /**
     * ListUrl matching requires a bit more:
     * 
     * DefaultViewUrl for lists/libraries:
     * DocumentLibrary: "/sites/modernui/Project4/Documents/Forms/AllItems.aspx"
     * List: "/sites/modernui/Project4/Lists/Events/calendar.aspx"
     * 
     * ParentWebUrl - For both List and Document library:
     * "/sites/modernui/Project4"
     * 
     * List url in button setting:
     * DocumentsLibrary button: "Documents"
     * List button: "Lists/Events"
     */

    let basicListUrl: string = list.DefaultViewUrl.replace(list.ParentWebUrl, "").toLowerCase();

    /**
     * Document library: "/Documents/Forms/AllItems.aspx"
     * List: "/Lists/Events/calendar.aspx"
     */

    const urlParts: string[] = basicListUrl.split("/").filter(part => part !== "" && part !== null && part !== undefined);

    // We are dealing with a list
    if (urlParts[0] === "lists") {
      if (`lists/${urlParts[1]}` === listUrlCandidate) {
        return true;
      }
    } else if (urlParts[0] === listUrlCandidate) {
      // Were dealing with a document library
      return true;
    }

    return false;
  });

  if (foundList.length === 0) {
    throw "No lists found";
  }

  return foundList[0];
}

export function appendFilterValueToUrl(url: string, filterField: string, filterValue: string): string {
  return (filterField && filterValue) ? `${url}?FilterField1=${filterField}&FilterValue1=${filterValue}` : url;
}

export function getListIcon(baseTemplateType: number): string {

  let iconClass: string = "";

  /* TODO: Handle more cases?
  100   Generic list
  101   Document library
  102   Survey
  103   Links list
  104   Announcements list
  105   Contacts list
  106   Events list
  107   Tasks list
  108   Discussion board
  109   Picture library
  110   Data sources
  111   Site template gallery
  112   User Information list
  113   Web Part gallery
  114   List template gallery
  115   XML Form library
  116   Master pages gallery
  117   No-Code Workflows
  118   Custom Workflow Process
  119   Wiki Page library
  120   Custom grid for a list
  130   Data Connection library
  140   Workflow History
  150   Gantt Tasks list
  200   Meeting Series list
  201   Meeting Agenda list
  202   Meeting Attendees list
  204   Meeting Decisions list
  207   Meeting Objectives list
  210   Meeting text box
  211   Meeting Things To Bring list
  212   Meeting Workspace Pages list
  301   Blog Posts list
  302   Blog Comments list
  303   Blog Categories list
  1100   Issue tracking
  1200   Administrator tasks list
  */

  switch (baseTemplateType) {
    case 107: // Tasks
    case 171: // Tasks with timeline and hierarchy
      iconClass = "CheckMark";
      break;
    case 101: // Document library
    case 815: // Asset Library //"Billeder til gruppen af websteder"?
      iconClass = "FabricFolderFill";
      break;
    case 119: // WebPageLibrary
      iconClass = "HomeSolid";
      break;
    case 100: // Generic list
    default:
      iconClass = "List";
  }
  return iconClass;
}

/**
 * Get Office open client-side URI-scheme URL.
 * 
 * Valid document URL's are known Word, Excel and PowerPoint files.
 * 
 * @param documentUrl URL to open document.
 */
export function getClientProgramOpenUrl(documentUrl: string): string {

  const documentUrlParts: string[] = documentUrl.split(".");
  const extension: string = documentUrlParts[documentUrlParts.length - 1];

  let uriScheme: string = null;

  switch (extension) {
    case "docx":
    case "docm":
    case "dotx":
    case "dotm":
    case "doc":
    case "dot":
    case "docb":
      uriScheme = "ms-word:ofe|u|";
      break;
    case "xlsx":
    case "xlsm":
    case "xltx":
    case "xltm":
    case "xls":
    case "xlt":
    case "xlm":
    case "xlsb":
    case "xlw":
      uriScheme = "ms-excel:ofe|u|";
      break;
    case "pptx":
    case "pptm":
    case "potx":
    case "potm":
    case "ppam":
    case "ppsx":
    case "ppsm":
    case "sldx":
    case "sldm":
    case "ppt":
    case "pot":
    case "pps":
      uriScheme = "ms-powerpoint:ofe|u|";
  }

  return `${uriScheme}${documentUrl}`;
}

export const knownExtensions = {
  "Excel": ["xlsx", "xlsm", "xltx", "xltm", "xls", "xlt", "xlm", "xlsb", "xlw"],
  "Word": ["docx", "docm", "dotx", "dotm", "doc", "dot", "docb"],
  "PowerPoint": ["pptx", "pptm", "potx", "potm", "ppam", "ppsx", "ppsm", "sldx", "sldm", "ppt", "pot", "pps"],
  "OneNote": ["onenote.notebook"],
  "Image": ["jpeg", "jpe", "jpg", "png", "bmp", "tiff", "gif"],
  "Pdf": ["pdf"],
  "Email": ["email", "eml", "msg", "mht"],
  "Text": ["txt"],
  "Aspx": ["asp", "aspx"],
  "Html": ["htm", "html"],
  "Css": ["css"],
  "Js": ["js"],
  "Ts": ["ts"]
};

export const officeFabricIcons = [
  "accdb", "csv", "docx", "dotx", "mpp", "mpt", "odp", "ods", "odt", "one", "onepkg", "onetoc", "potx", "ppsx", "pptx", "pub", "vsdx", "vssx", "vstx", "xls", "xlsx", "xltx", "xsn"
];

/**
 * Determines if a file extension belongs to an Office document file extension.
 * Has legacy support for old "doc" file types to be included as an Office file extension.
 * 
 * @param extension String file extension to match against know Office file extensions.
 */
export const isOfficeDocument = (extension: string): boolean => {

  if (typeof extension !== "string") {
    return null;
  }

  extension = extension.trim().toLowerCase();

  if (extension === "doc") {
    return true;
  }

  return officeFabricIcons.indexOf(extension) !== -1;
};

/**
 * Given an extension, returns the client program responsible for these extensions.
 * 
 * For regular non-office file types, it just returns the basic denominator for these.
 * 
 * @param extension Commonly known extensions in WorkPoint
 */
export const getClientProgramByExtension = (extension: string): string => {

  if (typeof extension !== "string" || extension === "") {
    return null;
  }

  let clientProgram: string = null;

  extension = extension.trim().toLowerCase();
  if (knownExtensions.Word.indexOf(extension) !== -1) {
    clientProgram = "Word";
  } else if (knownExtensions.Excel.indexOf(extension) !== -1) {
    clientProgram = "Excel";
  } else if (knownExtensions.PowerPoint.indexOf(extension) !== -1) {
    clientProgram = "PowerPoint";
  } else if (knownExtensions.OneNote.indexOf(extension) !== -1) {
    clientProgram = "OneNote";
  } else if (knownExtensions.Image.indexOf(extension) !== -1) {
    clientProgram = "Image";
  } else if (knownExtensions.Pdf.indexOf(extension) !== -1) {
    clientProgram = "Pdf";
  } else if (knownExtensions.Email.indexOf(extension) !== -1) {
    clientProgram = "Email";
  } else if (knownExtensions.Text.indexOf(extension) !== -1) {
    clientProgram = "Text";
  } else if (knownExtensions.Aspx.indexOf(extension) != -1) {
    clientProgram = "Aspx";
  } else if (knownExtensions.Html.indexOf(extension) != -1) {
    clientProgram = "Html";
  } else if (knownExtensions.Css.indexOf(extension) != -1) {
    clientProgram = "Css";
  } else if (knownExtensions.Js.indexOf(extension) != -1) {
    clientProgram = "Js";
  } else if (knownExtensions.Ts.indexOf(extension) != -1) {
    clientProgram = "Ts";
  }

  return clientProgram;
};

/**
 * Gets a matched Office icon for a provided extension.
 * 
 * Has support for legacy "doc" file extension to be shown as a "docx", as the former does not exist as an Office UI Fabric Icon.
 * 
 * @param extension String representation of an Office program file extension.
 */
export const getOfficeDocumentIconFromExtension = (extension: string): string => {

  if (typeof extension !== "string" || extension === "") {
    return null;
  }

  extension = extension.trim().toLowerCase();

  if (extension === "doc") {
    return "docx";
  }

  // Use default icon brands from Office UI fabric for all supported known extensions
  if (officeFabricIcons.indexOf(extension) !== -1) {
    return extension;
  }
};

/**
 * Tries to match a file extension against our known file extensions.
 * Default to a blank document icon if a match cannot be found.
 * 
 * @param extension String representation of a file extension.
 */
export const getFileIconFromExtension = (extension: string):string => {

  if (typeof extension !== "string") {
    return null;
  }

  extension = extension.toLowerCase();

  let iconName: string = "Document";

  // Use legacy icon mapping as a fallback.
  if (knownExtensions.Image.indexOf(extension) !== -1) {
    iconName = "FileImage";
  } else if (knownExtensions.Pdf.indexOf(extension) !== -1) {
    iconName = "PDF";
  } else if (knownExtensions.Email.indexOf(extension) !== -1) {
    iconName = "Mail";
  } else if (knownExtensions.Text.indexOf(extension) !== -1) {
    iconName = "TextDocument";
  } else if (knownExtensions.Aspx.indexOf(extension) != -1) {
    iconName = "FileASPX";
  } else if (knownExtensions.Html.indexOf(extension) != -1) {
    iconName = "FileHTML";
  } else if (knownExtensions.Css.indexOf(extension) != -1) {
    iconName = "FileCSS";
  } else if (knownExtensions.Js.indexOf(extension) != -1) {
    iconName = "JS";
  } else if (knownExtensions.Ts.indexOf(extension) != -1) {
    iconName = "TypeScriptLanguage";
  }

  return iconName;
};

/**
 * Matches and returns an icon for a given STS_Content_Class string.
 * 
 * @param contentClass String content class from SharePoint.
 */
export const getIconFromSTSContentClass = (contentClass: string):string => {

  switch (contentClass) {
    case "STS_ListItem_WebPageLibrary":
      return "Page";
    case "STS_ListItem_Links":
      return "Link";
    case "STS_ListItem_GanttTasks":
    case "STS_ListItem_Tasks":
    case "STS_ListItem_TasksWithTimelineAndHierarchy":
      return "TaskLogo";
    case "STS_ListItem_Events":
      return "EventDate";
    case "STS_ListItem_Announcements":
      return "Megaphone";
    case "STS_ListItem_Contacts":
      return "Contact";
    case "STS_ListItem_Survey":
      return "SurveyQuestions";
    case "STS_ListItem_DiscussionBoard":
      return "CustomList";
    case "STS_ListItem_IssueTracking":
      return "Bug";
    case "STS_ListItem_XMLForm":
      return "FileCode";
    case "STS_ListItem_Posts":
      return "PostUpdate";
    case "STS_ListItem_Comments":
      return "Comment";
    case "STS_Web":
      return "FileASPX";
    case "STS_ListItem":
    default:
      return "Document";
  }
};

/**
 * Converts a string with HTML escaped content to its initial state, using 'document.createElement' method.
 * 
 * @param input A string needing escaping.
 * 
 * @returns Unescaped string
 */
export const stripHtml = (input: string): string => {
  var tmp = document.createElement("DIV");
  tmp.innerHTML = input;
  return tmp.textContent || tmp.innerText || "";
};

/**
 * Does target contain correct permission, given permissions, permissionKinds and a target to check permissions for.
 * 
 * @param permissions The permissions to go through.
 * @param permissionKinds The kind of permission(s) were checking.
 * @param permissionTarget List or Business module entity (ListItem) were checking permissions for.
 */
export const actionAllowed = (permissions: BasePermissions, permissionKinds: PermissionKind[], permissionTarget: ListObject | BusinessModuleEntity): boolean => {

  let actionPermitted: boolean = true;

  for (let i = 0; i < permissionKinds.length; i++) {
    if (!permissionTarget.hasPermissions(permissions, permissionKinds[i])) {
      actionPermitted = false;
      break;
    }
  }

  return actionPermitted;
};

/**
 * Case insensitive string compare
 * 
 * @param string1 First string argument to compare.
 * @param string2 Second string argument to compare.
 */
export const caseInsensitiveStringCompare = (string1: string, string2: string): boolean => {

  try {

    if (typeof string1 === "string" && typeof string2 === "string") {

      return string1.toLowerCase() === string2.toLowerCase();
    }

    throw `Both arguments are not provided or are not valid string. string1: '${string1}', string2: '${string2}'`;
  } catch (exception) {
    console.warn(`'caseSensitiveStringCompare' could not complete compare operation. Defaulting to returning false. Error message: ${exception}`);

    return false;
  }
};

/**
 * Checks whether a string candidate is a string and that it's not empty.
 * 
 * @param stringCandidate The string candidate to check for.
 */
export function stringNotEmpty(stringCandidate: any): boolean {
  if (typeof stringCandidate === "string" && stringCandidate.trim() !== "") {
    return true;
  } else {
    return false;
  }
}