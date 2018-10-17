import * as DataService from '../workPointLibrary/service';
import { EntityWebProperty } from '../workPointLibrary/WebProperties';
import { BusinessModuleFields } from './Fields';

/**
 * Tries to fetch an entity's parents wpSite value from the current entity web's property bag.
 * Defaults to this entitys URL if it fails.
 */
export async function fetchEntityParentWebRelativeURL (solutionAbsoluteUrl:string, webAbsoluteUrl:string): Promise<string> {

    let parentWebUrl:string = webAbsoluteUrl;

    try {

        const propertyBagSelectProperties: string[] = [`AllProperties/${EntityWebProperty.itemLocation}`];
        const entityWebProperties = await DataService.getEntityPropertyBagResource(webAbsoluteUrl, propertyBagSelectProperties);
    
        // Item location contains whole hierarchy of parents for this entity.
        const entityItemLocation = entityWebProperties.WP_ITEM_LOCATION;
    
        let wpItemLocationStringParts:string[] = entityItemLocation.split(";").filter(part => part !== "");

        // Remove currentEntityListItemId
        wpItemLocationStringParts.pop();
        // Remove currentEntityListId
        wpItemLocationStringParts.pop();

        const parentEntityListItemId:number = parseInt(wpItemLocationStringParts.pop());
        const parentEntityListId:string = wpItemLocationStringParts.pop();

        const parentListItem = await DataService.loadEntityInformation(solutionAbsoluteUrl, parentEntityListItemId, parentEntityListId, [BusinessModuleFields.Site]);

        parentWebUrl = parentListItem.wpSite;
    } catch (exception) {
        console.warn("Could not load the entity parent web, due to invalid 'wpItemLocation' value. Defaulting to current entity value.");
    }

    return parentWebUrl;
}