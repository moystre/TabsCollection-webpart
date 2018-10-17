import { PersonaSize } from 'office-ui-fabric-react/lib/components/Persona';
import { Facepile, IFacepilePersona, OverflowButtonType } from 'office-ui-fabric-react/lib/Facepile';
import * as React from 'react';
import * as strings from 'WorkPointStrings';
import { IFieldValue } from '../workPointLibrary/EntityDetails';
import { IFieldMappingType } from '../workPointLibrary/FieldMappings';
import { stripHtml } from '../workPointLibrary/Helper';
import styles from './EntityDetails.module.scss';
import { Icon, IconSize } from './Icon';
import { IWorkPointBaseProps } from './WorkPointNavBarInterfaces';

export interface IEntityDetailRowProps {
  label: string;
  valueElement: JSX.Element;
}

export const EntityDetailRow:React.SFC<IEntityDetailRowProps> = (props):JSX.Element => {
  
  return (
  <p className={styles.row}>
    <span className={styles.label} title={props.label}>{props.label}</span>:
    {props.valueElement}
  </p>
  );
};

export class EntityDetailRowControl extends React.Component<IEntityDetailsRowControlProps> {

  private openUserPage = (userId:string) => {
    window.location.href = `${this.props.context.sharePointContext.pageContext.site.serverRelativeUrl}/_layouts/15/userdisp.aspx?ID=${userId}`;
  }

  /**
   * @deprecation-warning
   * This is volatile as this structure could easily be deprecated.
   */
  private getUserFromUserField = (userValue: string): IWorkPointFacepilePersona => {

    try {

      let presence:string = null;
      let email:string = null;
    
      const presenceAndEmailRegex:RegExp = /alt=\'(\b.+?)\'\s*sip='(.+?\b)'/;
      const presenceAndEmailMatch:string[] = presenceAndEmailRegex.exec(userValue);
  
      // Some users do not have a presence and email ('SharePoint' app forinstance)
      if (presenceAndEmailMatch !== null) {
        presence = presenceAndEmailMatch[1];
        email = presenceAndEmailMatch[2];
      }
    
      const userIdAndFullNameRegex:RegExp = /userdisp\.aspx\?ID=([0-9]+?)">(\b.+?)</;
      const userIdAndFullNameRegexMatch:string[] = userIdAndFullNameRegex.exec(userValue);
  
      // We could not parse value to get a meaningful display value, so abort
      if (userIdAndFullNameRegexMatch === null) {
        return null;
      }
    
      const userId:string = userIdAndFullNameRegexMatch[1];
      const userFullName:string = userIdAndFullNameRegexMatch[2];
    
      const persona:IFacepilePersona = {
        personaName: userFullName,
        data: {
          presence, 
          id: userId,
          email
        },
        onClick: (event) => this.openUserPage(userId)
      };

      return persona;

    } catch (exception) {
      return null;
    }
    
  }

  public render(): JSX.Element {

    const { fieldValue } = this.props;

    if (typeof fieldValue.value === "string" && fieldValue.value === "") {
      return <EntityDetailRow label={this.props.fieldValue.displayName} valueElement={<span className={styles.noValue}>{`(${strings.None})`}</span>} />;
    }

    // Initialized with default value
    let displayElement: JSX.Element = <span className={styles.value} title={fieldValue.value}>{stripHtml(fieldValue.value)}</span>;

    // Check for field mappings    
    if (fieldValue.fieldMappingType !== null) {
      switch (fieldValue.fieldMappingType) {
        case IFieldMappingType.Url: {
          const relativeUrlRegExp:RegExp = new RegExp('^(?:[a-z]+:)?//', 'i');
          const isUrlAbsolute:boolean = relativeUrlRegExp.test(fieldValue.value);
          const urlValue:string = isUrlAbsolute ? fieldValue.value : `//${fieldValue.value}`;
          displayElement = <a className={styles.valueLink} href={urlValue} title={`${strings.OpenLinkTo} ${fieldValue.value}`}><span>{fieldValue.value}</span> <Icon size={IconSize.microscopic} iconClass="Link" /></a>;
          break;
        }
        case IFieldMappingType.Email:
          displayElement = <a className={styles.valueLink} href={`mailto:${fieldValue.value}`} title={`${strings.OpenANewMailComposerToTheAddress} '${fieldValue.value}'`}><span>{fieldValue.value}</span> <Icon size={IconSize.microscopic} iconClass="Mail" /></a>;
          break;
        case IFieldMappingType.Phone:
          displayElement = <a className={styles.valueLink} href={`tel:${fieldValue.value}`} title={`${strings.Call}: ${fieldValue.value}`}><span>{fieldValue.value}</span> <Icon size={IconSize.microscopic} iconClass="Phone" /></a>;
          break;
      }
    } else {

      switch (fieldValue.type) {

        /**
         * Single user field
         */
        case "User": {
          const user:IWorkPointFacepilePersona = this.getUserFromUserField(fieldValue.value);
          if (user) {
            displayElement = <span><Facepile personaSize={PersonaSize.size16} personas={[user]} className={styles.facepileStyle} /><span className={styles.fullName}>{user.personaName}</span></span>;
          }
          break;
        }

        /**
         * Multi user field
         */
        case "UserMulti": {

          // @deprecation-warning: This is volatile as this structure could easily be deprecated.
          const userParts: string[] = fieldValue.value.split("<div class='ms-vb'>");
          let users: IWorkPointFacepilePersona[] = [];
    
          userParts.forEach(userString => {
            if (typeof userString === "string" && userString !== "") {
              const foundUser:IWorkPointFacepilePersona = this.getUserFromUserField(userString);
              if (foundUser) {
                users.push(foundUser);
              }
            }
          });

          if (users.length === 1) {
            displayElement = <span><Facepile className={styles.facepileStyle} personaSize={PersonaSize.size16} personas={users} /><span className={styles.fullName}>{users[0].personaName}</span></span>;
          } else if (users.length > 0) {
            displayElement = <Facepile className={styles.facepileStyle} personaSize={PersonaSize.size16} personas={users} maxDisplayablePersonas={3} overflowButtonType={OverflowButtonType.descriptive} overflowButtonProps={{ariaLabel: strings.MoreUsers}} />;
          }
          break;
        }

        /**
         * Notes
         */
        case "Note": {
          displayElement = <span className={styles.linkContainer} dangerouslySetInnerHTML={{__html:this.props.fieldValue.value}}></span>;
          break;
        }

        /**
         * Lookup field
         */
        case "Lookup":
        case "LookupMulti": {
          displayElement = <span className={styles.linkContainer} dangerouslySetInnerHTML={{__html:this.props.fieldValue.value}}></span>;
          break;
        }

        /**
         * URL field
         */
        case "URL": {
          displayElement = <span className={styles.linkContainer} dangerouslySetInnerHTML={{__html:this.props.fieldValue.value}}></span>;
        }
      }
    }

    return <EntityDetailRow label={this.props.fieldValue.displayName} valueElement={displayElement} />;
  }
}

export interface IEntityDetailsProps extends IWorkPointBaseProps {
  fieldValues: IFieldValue[];
}

export class EntityDetails extends React.Component<IEntityDetailsProps> {

  public render(): JSX.Element {
    return (

      <article className={styles.entityDetails}>
        {this.props.fieldValues && this.props.fieldValues.map(fieldValue => <EntityDetailRowControl fieldValue={fieldValue} context={this.props.context} />)}
      </article>
    );
  }
}

export interface IEntityDetailsRowControlProps extends IWorkPointBaseProps {
  fieldValue: IFieldValue;
}

export interface IPersonaExtraData {
  email: string;
  id: string;
  presence: string;
}

export interface IWorkPointFacepilePersona extends IFacepilePersona {
  data?: IPersonaExtraData;
}