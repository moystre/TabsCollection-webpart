/// <reference path="@ms/odsp.d.ts" />
declare module "*/config/webapi-config.json"
{
  /**
   * TODO: Add proper named exports for different apps, for easier include operations.
   */
  export interface IWorkPointApp {
    name: string;
    url: string;
    id: string;
  }

  const webApis: IWorkPointApp[];
  
  export default webApis;

}


