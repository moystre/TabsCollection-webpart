import { WorkPointContentType } from "./List";


export enum StageType {
  Start = 0,
  StartAndStandard,
  Standard,
  End
}

export enum StageStatus {
  Unknown = 0,
  History,
  Active,
  Future
}

export interface IStage {
  Type: StageType;
  Sorting: number;
  Url: string;
  ContentTypeId: string;
}

export interface ITransition {
  FromStageId: string;
  ToStageId: string;
  Workflow: string;
}

export enum ConstraintType {
  Warning = 0,
  Error
}

export enum ConstraintSourceType {
  List = 0,
  Metadata
}

export enum ConstraintDefinitionType {
  Required = 0,
  Prohibited
}

export class StageHistory {
  public fromStage: string;
  public toStage: string;
  public time: Date; // String
  public deadline: Date; // String
  public constraints: any; // Object with not-known property names

  constructor(fromStage:string, toStage:string, time:string, deadline:string, constraints:any) {
    this.fromStage = fromStage;
    this.toStage = toStage;
    this.time = new Date(time);
    this.deadline = new Date(deadline);
    this.constraints = constraints;
  }
}

export interface IStageHistory {
  FromStage: string;
  ToStage: string;
  Time: string;
  Deadline: string;
  Constraints: any; // Object with not-known property names
}

export class StageObject {
  public id: string;
  public type: StageType;
  public name: string;
  public status: StageStatus;
  public transitionable: boolean;
  public link:string;
  public filterable: boolean;
  public history:StageHistory;

  /**
   * Constructor sets simple values.
   * 
   * History needs to be set explicitly.
   */
  constructor(id:string, type:StageType, name:string, status:StageStatus = StageStatus.Unknown, canTransitionTo:boolean = false, url:string = null) {
    this.id = id;
    this.type = type;
    this.name = name;
    this.status = status;
    this.transitionable = canTransitionTo;
    this.link = url;
  }
}

export interface IConstraint {
  Title: string;
  StageId: string;
  Active: boolean;
  Type: ConstraintType;
  ConstraintSource: ConstraintSourceType;
  ListTitle: string;
  ListUrl: string;
  DefinitionType: ConstraintDefinitionType;
  Definition: string;
  MetadataField: string;
  MetadataValue: string;
  MetadataOperator: FieldOperator;
  ErrorMessage: string;
  ErrorMessageResource: object;
}

export interface IHistoryConstraintStatus extends IConstraint {
  upholded:boolean;
}

export interface IStageSettings {
  Enabled: boolean;
  Stages: IStage[];
  Transitions: ITransition[];
  Constraints: IConstraint[];
}

export enum FieldOperator
{
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LowerThan,
    LowerThanOrEqual,
    IsNull,
    IsNotNull,
    BeginsWith,
    Contains,
    DateRangesOverlap
}

/**
 * Detect if stage can be navigated to from the currently active stage.
 * 
 * @param targetStageId The stage we need to know if we can transition to.
 * @param sourceStageId Stage that were on right now.
 * @param transitions A map of all configured stage transitions.
 * 
 * @returns Boolean.
 */
const transitionable = (targetStageId:string, sourceStageId:string, transitions:ITransition[]):boolean => {

  for (let transition of transitions) {
    if (targetStageId === transition.ToStageId && sourceStageId === transition.FromStageId) {
      return true;
    }
  }
  
  return false;
};

/**
 * Fetches stage history objects for a given stage in a specific entity.
 * 
 * @param stageId Id of stage to find stage history for.
 * @param fullEntityStageHistory Full stage history for the business module entity.
 * 
 * @returns The latest stage history object for the stage.
 */
export const getStageHistory = (stageId:string, fullEntityStageHistory:IStageHistory[]):IStageHistory => {

  // Backwards iterate to find latest version of history.
  for (let i = fullEntityStageHistory.length - 1; i >= 0; --i) {
    let history = fullEntityStageHistory[i];

    if (stageId === history.ToStage) {
      return history;
    }
  }
  
  return null;
};

/**
 * Fetches all constraints for a given stage.
 * 
 * @param stageId Id of stage to find stage constraints for.
 * @param constraints All configured constraints for this business module.
 * 
 * @returns All matching constraints for the stage id argument.
 */
export const getStageConstraints = (stageId:string, constraints:IConstraint[]):IConstraint[] => {
  const constraintsForThisStage:IConstraint[] = constraints.filter(constraint => constraint.Active && constraint.StageId === stageId);
  return constraintsForThisStage;
};

/**
 * Get stages as objects with id, type, name and history, sorted by their sorting key.
 * 
 * @param businessModuleContentTypes Content types in a business module list.
 * @param stages Stages collected as multiple IStage[] from the main StageSettings object.
 * @param transitions Stage transitions for this business module from the main StageSettings object.
 * @param constraints Stage constraints for this business module from the main StageSettings object.
 * @param currentEntityContentTypeId The id string of the current contenttype.
 * @param fullEntityStageHistory Stage history for the entity.
 * 
 * @returns A list of StageObject[]
 */
export const getOrderedStageMap = (businessModuleContentTypes:WorkPointContentType[], stages:IStage[], transitions:ITransition[], constraints:IConstraint[], currentEntityContentTypeId:string, fullEntityStageHistory:IStageHistory[]):StageObject[] => {

  // Sort the stages by the "Sorting" key.
  stages = stages.sort((a, b) => {
    if (a && a.hasOwnProperty("Sorting") && b && b.hasOwnProperty("Sorting")) {
      return a.Sorting - b.Sorting;
    } else {
        return 0;
    }
  });

  // Does this entity have any stage history? If not, then a boolean value is easier to check later on.
  const stageHistoryAvailable:boolean = (fullEntityStageHistory && Array.isArray(fullEntityStageHistory) && fullEntityStageHistory.length > 0);

  /**
   * These are not needed yet
   */
  // Same goes for constraints, evaluate if they should be checked later on.
  //const constraintsAvailable:boolean = (constraints && Array.isArray(constraints) && constraints.length > 0);

  // Find out if current contentTypeId is a stage
  let currentContentTypeIsAStage:boolean = false;
  for (let i = stages.length-1; i >= 0; i--) {
    const stage:IStage = stages[i];
    if (stage.ContentTypeId === currentEntityContentTypeId) {
      currentContentTypeIsAStage = true;
      break;
    }
  }

  let currentStageFound:boolean = false;

  // Iterate and map all known stages from StageSettings
  const mappedStages:StageObject[] = stages.map((stage:IStage) => {

    const matchingContentTypes:WorkPointContentType[] = businessModuleContentTypes.filter(contentType => contentType.StringId === stage.ContentTypeId);

    // The stage setting does not match a list content type, so it is unusable.
    if (matchingContentTypes.length < 1) {
      return null;
    }

    const stageContentType:WorkPointContentType = matchingContentTypes[0];

    // Define the stage status
    let stageStatus:StageStatus = StageStatus.Unknown;
    
    // This is the currently active stage.
    if (currentEntityContentTypeId === stage.ContentTypeId) {
      stageStatus = StageStatus.Active;
      currentStageFound = true;
    } else if (!currentStageFound) {
      // We have not reached the current stage yet, so this must be prior to it
      stageStatus = StageStatus.History;
    } else if (currentStageFound) {
      // This is after the currently active stage, ergo: future
      stageStatus = StageStatus.Future;
    }

    // Can this stage be navigated to from the currently active stage/content type?
    let isTransitionable:boolean = false;

    if (currentContentTypeIsAStage) {
      isTransitionable = transitionable(stage.ContentTypeId, currentEntityContentTypeId, transitions);
    } else if (stage.Type === StageType.Start || stage.Type === StageType.StartAndStandard) {
      isTransitionable = true;
    }

    // The stage content type has the localized name values
    const stageObject = new StageObject(
      stageContentType.StringId,
      stage.Type,
      stageContentType.Name,
      stageStatus,
      isTransitionable,
      stage.Url
    );

    // Stage history is available, so check if this stage has any
    if (stageHistoryAvailable) {

      let filterable:boolean = false;

      // Try to see if this stage has been active at some point
      for (let history of fullEntityStageHistory) {
        if (history.FromStage === stageObject.id) {
          filterable = true;
          break;
        }
      }

      stageObject.filterable = filterable;
      
      try {
        
        const stageCandidate:IStageHistory = getStageHistory(stageObject.id, fullEntityStageHistory);

        const stageHistoryObject:StageHistory = new StageHistory(
          stageCandidate.FromStage,
          stageCandidate.ToStage,
          stageCandidate.Time,
          stageCandidate.Deadline,
          stageCandidate.Constraints
        );
        
        stageObject.history = stageHistoryObject;
      } catch (exception) {}
    }

    return stageObject;
  });

  return mappedStages.filter(stage => stage !== null);
};