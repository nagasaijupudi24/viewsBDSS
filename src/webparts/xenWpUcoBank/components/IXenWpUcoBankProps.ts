import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPFI } from "@pnp/sp";
import { DisplayMode } from "@microsoft/sp-core-library";
import { IPropertyPaneDropdownOption } from "@microsoft/sp-property-pane";

export interface IXenWpUcoBankProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  listName:any;
  libraryName:any;
  formType:string;
  listViews?:string;
  context:WebPartContext;
  sp:SPFI;
  viewType:any;
  noteType:any;
  title: string;
  updateProperty: (value: string) => void;
  displayMode: DisplayMode;
  newPageUrl:string;
  viewPageUrl:string;
  editPage:string;
  CBnewPageUrl:string;
  CBviewPageUrl:string;
  CBeditPage:string;

  fields:string[];
  columnName: string;
  chartType: string;
  chartTitle: string;
  customOption:IPropertyPaneDropdownOption[];
  listId?:string;
  pageType:string;
  committeeType:any;
  httpUrl:string;
  departmentGroupName:string;
  superAdminGroupName:string;
  cmViewType:string //committee view type
  noteListName:any

}
