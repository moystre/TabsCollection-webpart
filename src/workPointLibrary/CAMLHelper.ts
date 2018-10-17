import * as CamlBuilder from 'camljs';

const domXMLParser = new DOMParser();
const xmlSerializer = new XMLSerializer();

/**
* Helper method that adds a an expression to a collection of CamlBuilder.IExpression objects.
*
* @param fieldType String representation of a SharePoint field type, eg: SP.FieldChoice, SP.FieldText or SP.FieldUser.
* @param fieldInternalName InternalFieldName of a field to filter.
* @param fieldValue The value to use for filtering.
* @param expressions The collection to append expressions to.
*/
export function addExpression (fieldType: string, fieldInternalName: string, fieldValue: any, expressions: CamlBuilder.IExpression[]): CamlBuilder.IExpression[] {

  if (!expressions || !Array.isArray(expressions)) {
    expressions = [];
  }

  switch (fieldType) {

    /**
    * User or lookup field
    */
    case "SP.FieldUser":
      if (isNaN(fieldValue)) {
        expressions.push(CamlBuilder.Expression().UserField(fieldInternalName).ValueAsText().EqualTo(fieldValue));
      } else {
        expressions.push(CamlBuilder.Expression().UserField(fieldInternalName).Id().EqualTo(fieldValue));
      }
      break;
    case "SP.FieldLookup":
      if (isNaN(fieldValue)) {
        expressions.push(CamlBuilder.Expression().LookupField(fieldInternalName).ValueAsText().EqualTo(fieldValue));
      } else {
        expressions.push(CamlBuilder.Expression().LookupField(fieldInternalName).Id().EqualTo(fieldValue));
      }
      break;
    case "SP.FieldChoice":
      expressions.push(CamlBuilder.Expression().ChoiceField(fieldInternalName).EqualTo(fieldValue));
      break;
    case "SP.FieldNumber":
      expressions.push(CamlBuilder.Expression().NumberField(fieldInternalName).EqualTo(fieldValue));
      break;
    case "SP.FieldBoolean":
      expressions.push(CamlBuilder.Expression().BooleanField(fieldInternalName).EqualTo(fieldValue));
      break;
    case "SP.FieldText":
    default:
      expressions.push(CamlBuilder.Expression().TextField(fieldInternalName).EqualTo(fieldValue));
  }

  return expressions;
}

/**
 * Adds a null value expression to a collection of CamlBuilder.IExpression objects.
 * 
 * @param fieldType String representation of a SharePoint field type, eg: SP.FieldChoice, SP.FieldText or SP.FieldUser.
 * @param fieldInternalName InternalFieldName of a field to filter.
 * @param expressions The collection to append expressions to.
 */
export function addIsNullExpression (fieldType: string, fieldInternalName: string, expressions: CamlBuilder.IExpression[]): CamlBuilder.IExpression[] {

  if (!expressions || !Array.isArray(expressions)) {
    expressions = [];
  }

  switch (fieldType) {

    /**
    * User or lookup field
    */
    case "SP.FieldUser":
    case "SP.FieldLookup":
      expressions.push(CamlBuilder.Expression().LookupField(fieldInternalName).Id().IsNull());
      break;
    case "SP.FieldChoice":
      expressions.push(CamlBuilder.Expression().ChoiceField(fieldInternalName).IsNull());
      break;
    case "SP.FieldNumber":
      expressions.push(CamlBuilder.Expression().NumberField(fieldInternalName).IsNull());
      break;
    case "SP.FieldBoolean":
      expressions.push(CamlBuilder.Expression().BooleanField(fieldInternalName).IsNull());
      break;
    case "SP.FieldText":
    default:
      expressions.push(CamlBuilder.Expression().TextField(fieldInternalName).IsNull());
  }

  return expressions;
}

/**
 * Adds <FieldRef> elements to view XML Caml queries.
 * 
 * @param fields Fields to add.
 * @param viewXmlString The viewXML string to manipulate.
 */
export function addFields (fields: string[], viewXmlString:string):string {

  let viewXml:string = viewXmlString;

  try {
  
    const viewXMLDocument:XMLDocument = domXMLParser.parseFromString(viewXml, "text/xml");
  
    const viewFieldsElement = viewXMLDocument.getElementsByTagName("ViewFields")[0];
  
    fields.map(internalFieldName => {
      const element:HTMLElement = domXMLParser.parseFromString(`<FieldRef Name="${internalFieldName}"/>`, "text/xml").documentElement;
      viewFieldsElement.appendChild(element);
    });
  
    viewXml = xmlSerializer.serializeToString(viewXMLDocument);

  } catch (exception) {
    console.warn(`Could not add the following fields to the XML query: '${fields.join(", ")}'`);
  }

  return viewXml;
}

/**
 * Given a full SharePoint list View XML string, replaces the Query element with the given value.
 * 
 * @param viewXmlString A full SharePoint list <View> XML string.
 * @param queryElementXmlString A SharePoint list <Query> element XML string to override the prior viewXmlString argument.
 */
export function replaceViewQueryElement (viewXmlString:string, queryElementXmlString:string):string {

  let outputXmlString:string = viewXmlString;

  try {

    let view = domXMLParser.parseFromString(viewXmlString, "text/xml");
    let oldQuery = view.getElementsByTagName("Query")[0];
    let newQuery = domXMLParser.parseFromString(queryElementXmlString, "text/xml").documentElement;
    view.documentElement.replaceChild(newQuery, oldQuery);
    
    outputXmlString = xmlSerializer.serializeToString(view);

  } catch (exception) {
    console.warn(`Could not replace the Query element of the given viewXML`);
  }

  return outputXmlString;
}

/**
 * Attempt to build a <Query> element from a given SharePoint list view XML string.
 * 
 * @param viewXmlString View XML string to build upon.
 * @param expressions CamlBuilder expressions to fit into the querys <Where> element.
 * @param sortingField Optional. If present, is used to add sorting filters to the query.
 * @param sortAscending Optional. Controls the sorting direction.
 */
export function buildQueryElement (viewXmlString:string, expressions:CamlBuilder.IExpression[], sortingField?:string, sortAscending?:boolean, groupByField?:string):string {
  
  let outputViewXmlString:string = viewXmlString;
    
  try {

    const viewXMLDocument:XMLDocument = domXMLParser.parseFromString(viewXmlString, "text/xml");

    const queryElement = viewXMLDocument.getElementsByTagName("Query")[0];

    // Apply the sorting condition
    if (sortingField && sortingField !== "") {
      
      const newOrderByElement = domXMLParser.parseFromString(`<OrderBy><FieldRef Name="${sortingField}" Ascending="${sortAscending ? 'True':'False'}" /></OrderBy>`, "text/xml").documentElement;
      
      // Does OrderBy element exist?
      if (queryElement.getElementsByTagName("OrderBy").length > 0) {
        const oldOrderByElement = queryElement.getElementsByTagName("OrderBy")[0];
        queryElement.replaceChild(newOrderByElement, oldOrderByElement);
      } else {
        queryElement.appendChild(newOrderByElement);
      }

    }

    if (groupByField && groupByField !== "") {
      
      const newGroupByElement = domXMLParser.parseFromString(`<GroupBy Collapse="FALSE"><FieldRef Name="${groupByField}" /></GroupBy>`, "text/xml").documentElement;
      
      // Does OrderBy element exist?
      if (queryElement.getElementsByTagName("GroupBy").length > 0) {
        const oldGroupByElement = queryElement.getElementsByTagName("GroupBy")[0];
        queryElement.replaceChild(newGroupByElement, oldGroupByElement);
      } else {
        queryElement.appendChild(newGroupByElement);
      }
    }
    
    let replaceWhereElement:boolean = false;

    if (expressions.length > 0) {

      // If no where element exists, we add it
      if (queryElement.getElementsByTagName("Where").length === 0) {
        
        const whereElement = domXMLParser.parseFromString(`<Where></Where>`, "text/xml").documentElement;
        queryElement.appendChild(whereElement);
        replaceWhereElement = true;
      }
    }
    
    outputViewXmlString = xmlSerializer.serializeToString(queryElement);

    // Handle expressions
    if (expressions.length > 0) {

      let builderExpression:CamlBuilder.IExpression = null;

      const camlBuilderQuery = CamlBuilder.FromXml(outputViewXmlString);
      
      // Should we replace the where element?
      if (replaceWhereElement) {
        builderExpression = camlBuilderQuery.ReplaceWhere().All(expressions);
      } else {
        builderExpression = camlBuilderQuery.ModifyWhere().AppendAnd().All(expressions);
      }

      outputViewXmlString = builderExpression.ToString();
    }

  } catch (exception) {
    console.warn(`Creation of <Query> element went wrong with the following exception: ${exception}`);
  }

  return outputViewXmlString;
}