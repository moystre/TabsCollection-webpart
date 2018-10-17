import { AadHttpClient } from "@microsoft/sp-http";
import { ActivityItem, IActivityItemProps } from 'office-ui-fabric-react/lib/components/ActivityItem';
import { Spinner, SpinnerSize } from 'office-ui-fabric-react/lib/Spinner';
import * as React from 'react';
import { SearchResult, SearchResults } from 'sp-pnp-js/lib/pnp';
import * as strings from 'WorkPointStrings';
import * as webApis from '../../config/webapi-config.json';
import { IBusinessModuleEntity } from '../workPointLibrary/BusinessModule';
import { getJournalItemsEMM, IEMMEntry } from '../workPointLibrary/EMM';
import { getFileIconFromExtension, getIconFromSTSContentClass, getOfficeDocumentIconFromExtension, isOfficeDocument, stringNotEmpty } from '../workPointLibrary/Helper';
import { fetchJournalItems } from '../workPointLibrary/Journal';
import { getActivityForWeb } from '../workPointLibrary/service';
import { getDateTimeFromString, getUserFriendlyTime } from '../workPointLibrary/Time';
import styles from './Activity.module.scss';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

const LoadingIndicator:React.SFC = ():JSX.Element => {
  return (
    <div className={styles.loadingIndicator}>
      <Spinner className={styles.spinner} size={SpinnerSize.small} />
    </div>
  );
};

export interface IActivityResponse {
  Activities: IActivity[];
}

export interface IActivityItem {
  FileType: string;
  Url: string;
  Title: string;
  Type: string; // "ModernPage" | "Document" | "LegacyPage" ...TODO: More?
}

export interface IActivity {
  ActivityItem: IActivityItem;
  Type: "YouViewedActivity" | "YouModifiedActivity"; //string; // "YouViewedActivity" | "YouModifiedActivity" ...TODO: More?
  Time: string;
}

export interface ISortableActivityItemProps extends IActivityItemProps {
  SortDate: Date;
}

export interface IActivityControlState {
  activities: ISortableActivityItemProps[];
  amount: number;
  loading: boolean;
}

export interface IActivityControlProps extends IWorkPointBaseProps {
  currentEntity: IBusinessModuleEntity;
}

export default class ActivityControl extends React.Component<IActivityControlProps, IActivityControlState> {

  constructor(props:IActivityControlProps) {
    super(props);

    this.state = {
      activities: [],
      loading: true,
      amount: 15
    };
  }

  public async componentDidMount():Promise<void> {

    let activities:ISortableActivityItemProps[] = [];

    /**
     * Experimental. Used to get user activity for a given web
     * activities = await this.getBetaWebActivityForUser();
     */

    try {

      activities = await this.getItemActivityForWeb();
    } catch (exception) {
      console.warn(exception);
    }

    this.setState({
      activities,
      loading: false
    });
  }

  /**
   * Maps a 'SearchResult' to a 'ISortableActivityItemProps'.
   * 
   * Adds a SotrDate property for easier sorting afterwards, when combining with journal items.
   * 
   * @param searchResult SearchResult instance
   * 
   * Used for the ActivityItem output.
   */
  protected mapJournalActivity = (searchResult:SearchResult) => {
    
    const title:string = searchResult.Title;
    const fileType:string = searchResult.FileType;
    let contentClassIconName:string = null;

    if (!fileType) {
      contentClassIconName = getIconFromSTSContentClass(searchResult.contentclass);
    }

    /**
     * Handle verb and time
     */
    let verb:string = null;
    let actor:string = null;
    let document:string = null;
    let userFriendlyTime:string = null;

    // TODO: Determine action
    verb = strings.Modified.toLowerCase();
    
    actor = searchResult.Author === this.props.context.sharePointContext.pageContext.user.displayName ? strings.You : searchResult.Author;
    
    // Show filetype if present.
    document = `${title}${fileType ? `.${fileType}`: ""}`;

    /**
     * This seems a bit shabby, but "LastModifiedTime" in "SearchResult" is typed as a Date, but when returned from search, is a string.
     * TODO: Can we make sure we get the correct date format to do this?
     */
    let sortableDate:Date = getDateTimeFromString(searchResult.LastModifiedTime.toString());
    userFriendlyTime = getUserFriendlyTime(searchResult.LastModifiedTime.toString());
    
    let iconElement:JSX.Element = null;

    // Handle known file types
    if (stringNotEmpty(fileType)) {

      const isOfficeIcon:boolean = isOfficeDocument(fileType);
  
      let icon:string = isOfficeIcon ? getOfficeDocumentIconFromExtension(fileType) : getFileIconFromExtension(fileType);
  
      if (isOfficeIcon) {
        iconElement = <div className={`ms-BrandIcon--icon16 ms-BrandIcon--${icon}`}></div>;
      } else {
        iconElement = <i className={`ms-Icon ms-Icon--${icon}`} aria-hidden="true"></i>;
      }
    } else if (stringNotEmpty(contentClassIconName)) {
      iconElement = <i className={`ms-Icon ms-Icon--${contentClassIconName}`} aria-hidden="true"></i>;
    }

    const descriptionElements:JSX.Element[] = [];

    descriptionElements.push(<span className={styles.nameText}>{actor} </span>);
    descriptionElements.push(<span>{verb} </span>);
    descriptionElements.push(<span className={styles.documentText}>{document} </span>);

    if (userFriendlyTime) {
      descriptionElements.push(<span className={styles.timeText} >{userFriendlyTime}</span>);
    }


    const activityItemProp:ISortableActivityItemProps = {
      activityIcon: iconElement,
      activityDescription: descriptionElements,
      SortDate: sortableDate,
      isCompact: true
    };

    return activityItemProp;
  }

  /**
   * Maps a 'IEMMEntry' to a 'ISortableActivityItemProps'.
   * 
   * Adds a SotrDate property for easier sorting afterwards, when combining with journal items.
   * 
   * @param entry IEMMEntry instance
   * 
   * Used for the ActivityItem output.
   */
  protected mapEMMEntry = (entry:IEMMEntry) => {

    const subject:string = entry.Subject;

    let sent:string = strings.Sent.toLowerCase();
    let to:string = strings.To.toLowerCase();
    let sender:string = null;
    let reciever:string = null;
    let userFriendlyTime:string = null;
    
    sender = entry.From === this.props.context.sharePointContext.pageContext.user.displayName ? strings.You : entry.From;
    reciever = entry.To === this.props.context.sharePointContext.pageContext.user.displayName ? strings.You : entry.To;

    let sortableDate:Date = getDateTimeFromString(entry.OrigDate);
    userFriendlyTime = getUserFriendlyTime(entry.OrigDate);

    const icon:string = entry.Direction ? "MailForwardMirrored" : "MailForward";

    let iconElement:JSX.Element = <i className={`ms-Icon ms-Icon--${icon}`} aria-hidden="true"></i>;

    const descriptionElements:JSX.Element[] = [];

    descriptionElements.push(<span className={styles.nameText}>{sender} </span>);
    descriptionElements.push(<span>{sent} </span>);
    descriptionElements.push(<span className={styles.documentText}>{subject} </span>);
    descriptionElements.push(<span>{to} </span>);
    descriptionElements.push(<span className={styles.nameText}>{reciever} </span>);

    const activityItemProp:ISortableActivityItemProps = {
      activityIcon: iconElement,
      activityDescription: descriptionElements,
      ...userFriendlyTime ? {timeStamp:userFriendlyTime} : null,
      SortDate: sortableDate
    };

    return activityItemProp;
  }

  /**
   * Fetches latest items via. search and wpItemLocation.
   * 
   * @returns Promise of ISortableActivityItemProps[]
   */
  protected async getItemActivityForWeb():Promise<ISortableActivityItemProps[]> {

    const { context, currentEntity } = this.props;
    let { ItemLocation, UniqueId } = this.props.currentEntity;
    const { amount } = this.state;

    let journalResults:SearchResults = null;
    let emmResults:IEMMEntry[] = [];

    /**
     * TODO: Remove
     * Debug only
     */
    //const isEMMEnabled:boolean = true;
    const isEMMEnabled:boolean = currentEntity.Settings.EnableEMMIntegration;
    //ItemLocation = "aec457ce16b043d3bf2d7c2cc9a93b28;8105d2b7a7f24388bce26437e1914375;6;e0f4b81ae9a54d6b92677cc4383f11ad;6;";

    if (isEMMEnabled) {
      const aadHttpClient: AadHttpClient = new AadHttpClient(context.sharePointContext.serviceScope, webApis[1].id);
      [journalResults, emmResults] = await Promise.all([fetchJournalItems(ItemLocation, amount), getJournalItemsEMM(UniqueId, webApis[1], context.solutionAbsoluteUrl, amount, aadHttpClient)]);
    } else {
      journalResults = await fetchJournalItems(ItemLocation, amount);
    }

    const mappedEMMEntries = emmResults.map(this.mapEMMEntry);
    const mappedActivities: ISortableActivityItemProps[] = journalResults ? journalResults.PrimarySearchResults.map(this.mapJournalActivity) : [];

    const activities = [...mappedEMMEntries, ...mappedActivities]
      .sort((a, b) => b.SortDate.getTime() - a.SortDate.getTime())
      .slice(0, amount);

    return activities;
  }

  /**
   * This is experimental and only works per user.
   * 
   * Not enabled.
   * 
   * It is based on the new SharePoint activity web part, by sniffing out its request.
   * 
   * @see SharePoint ModernUI Activity Webpart
   */
  protected async getBetaWebActivityForUser():Promise<IActivityItemProps[]> {

    const { context } = this.props;

    const tenantUrl:string = context.solutionAbsoluteUrl.replace(context.solutionRelativeUrl, "");
    const activities: IActivity[] = await getActivityForWeb(tenantUrl, context.sharePointContext);

    const mappedActivities: IActivityItemProps[] = activities.map(activity => {

      // Handle icon
      const { FileType } = activity.ActivityItem;

      const isOfficeIcon:boolean = isOfficeDocument(FileType);

      const icon:string = isOfficeIcon ? getOfficeDocumentIconFromExtension(FileType) : getFileIconFromExtension(FileType);

      let iconElement:JSX.Element = null;

      if (isOfficeIcon) {
        iconElement = <div className={`ms-BrandIcon--icon16 ms-BrandIcon--${icon}`}></div>;
      } else {
        iconElement = <i className={`ms-Icon ms-Icon--${icon}`} aria-hidden="true"></i>;
      }

      /**
       * Handle verb and time
       */
      let verb:string = null;
      let actor:string = null;
      let document:string = null;
      const userFriendlyTime:string = getUserFriendlyTime(activity.Time);

      switch (activity.Type) {
        case "YouViewedActivity":
          verb = strings.Viewed.toLowerCase();
          actor = strings.You;
          document = `${activity.ActivityItem.Title}.${activity.ActivityItem.FileType}`;
          break;
        case "YouModifiedActivity":
          verb = strings.Modified.toLowerCase();
          actor = strings.You;
          document = `${activity.ActivityItem.Title}.${activity.ActivityItem.FileType}`;
          break;
      }

      let time:string = "";

      const descriptionElements:JSX.Element[] = [];

      descriptionElements.push(<span className={styles.nameText}>{actor} </span>);
      descriptionElements.push(<span>{verb} </span>);
      if (document) {
        descriptionElements.push(<span className={styles.documentText}>{document} </span>);
      }

      const activityItemProp:IActivityItemProps = {
        activityIcon: iconElement,
        activityDescription: descriptionElements,
        ...userFriendlyTime ? {timeStamp:userFriendlyTime} : null
      };

      return activityItemProp;
    });

    return mappedActivities;
  }

  public render():JSX.Element {

    const { activities, loading } = this.state;

    return (
      <div className={styles.activityMenu}>
        {loading && <LoadingIndicator />}
        {!loading && activities.length > 0 && activities.map(activity => <ActivityItem className={styles.activityItem} {...activity} />)}
        {!loading && activities.length === 0 && <div className={styles.noActivityContainer}><span className={styles.noActivity}>{strings.NoActivity}</span></div>}
      </div>
    );
  }
}