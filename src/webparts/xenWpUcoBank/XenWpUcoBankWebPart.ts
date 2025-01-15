import * as React from 'react';
import * as ReactDom from 'react-dom';
import {  Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'XenWpUcoBankWebPartStrings';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import SPService from './components/Service/SPservice';
import { PropertyFieldMultiSelect } from '@pnp/spfx-property-controls';
import ChartView from '../../Common/ChartView';
import ListViews from './components/Views/listView';
import XenWpUcoBank from './components/XenWpUcoBank';
import HomePage from './components/eNote/EnoteHomePage';
import PasscodePage from './components/Passcode/PasscodePage';
import ATRViews from './components/ATRViews/ATRViews';
import SearchPage from './components/Search/SearchPage';
import CommitteeMeetingListViews from './components/CommitteeMeeting/CommitteeMeettingViews';

export interface IXenWpUcoBankWebPartProps {
  description: string;
  listName:any;
  libraryName:any;
  formType:string;
  listViews:string;
  sp:SPFI;
  noteType:any;
  viewType:any
  title:string;
  newPageUrl:string;
  viewPageUrl:string;
  editPage:string;
  fields:string[];
  columnName: string;
  chartType: string;
  chartTitle: string;
  customOption:IPropertyPaneDropdownOption[];
  listId?:string;
  pageType:string;
  updateProperty: (value: string) => void;
  committeeType:any;
  httpUrl:string;
  departmentGroupName:string;
  superAdminGroupName:string;
  cmViewType:string;
  CBnewPageUrl:string;
  CBviewPageUrl:string;
  CBeditPage:string;
  noteListName:any
}
export interface IlistDetails {
  id?: string,
  title?: string,
  webUrl?: string
}

export default class XenWpUcoBankWebPart extends BaseClientSideWebPart<IXenWpUcoBankWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _services: SPService;
  private charts: IPropertyPaneDropdownOption[] = [
    { key: "Normal", text: "Boxes" },
    { key: "Pie", text: "Pie" },
    { key: "Bar", text: "Bar" },
    { key: "Doughnut", text: "Doughnut" },
    { key: "Line", text: "Line" },
    { key: "PolarArea", text: "Polar Area" },
   /*  //'line' | 'bar' | 'horizontalBar' | 'radar' | 'doughnut' | 'polarArea' | 'bubble' | 'pie' | 'scatter'; */
  ];

  public render(): void {
    let element:any;
    
   if(this.properties.pageType ==="views"){
      element= React.createElement(
        ListViews,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context:this.context,
          listName:this.properties.listName,
          libraryName:this.properties.libraryName,
          formType:this.properties.formType,
          listViews:this.properties.listViews,
          sp:this.properties.sp,
          noteType:this.properties.noteType,
          viewType:this.properties.viewType,
          title:this.properties.title,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          newPageUrl:this.properties.newPageUrl,
          viewPageUrl:this.properties.viewPageUrl,
          editPage:this.properties.editPage,
          fields:this.properties.fields,
          columnName:this.properties.columnName,
          chartType:this.properties.chartType,
          chartTitle:this.properties.chartTitle,
          customOption:this.properties.customOption,
          listId:this.properties.listId,
          pageType:this.properties.pageType,
          committeeType:this.properties.committeeType,
          httpUrl:this.properties.httpUrl,
          departmentGroupName:this.properties.departmentGroupName,
          superAdminGroupName:this.properties.superAdminGroupName,
          cmViewType:this.properties.cmViewType,
          CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName //noteListName
          
        }
      );
    }else if(this.properties.pageType ==="charts"){
      element= React.createElement(
        ChartView,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context:this.context,
          listName:this.properties.listName,
          libraryName:this.properties.libraryName,
          formType:this.properties.formType,
          listViews:this.properties.listViews,
          sp:this.properties.sp,
          noteType:this.properties.noteType,
          viewType:this.properties.viewType,
          title:this.properties.title,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          newPageUrl:this.properties.newPageUrl,
          viewPageUrl:this.properties.viewPageUrl,
          editPage:this.properties.editPage,
          fields:this.properties.fields,
          columnName:this.properties.columnName,
          chartType:this.properties.chartType,
          chartTitle:this.properties.chartTitle,
          customOption:this.properties.customOption,
          listId:this.properties.listId,
          pageType:this.properties.pageType,
          committeeType:this.properties.committeeType,
          httpUrl:this.properties.httpUrl,
          departmentGroupName:this.properties.departmentGroupName,
          superAdminGroupName:this.properties.superAdminGroupName,
          cmViewType:this.properties.cmViewType,
          CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName //noteListName
          
        }
      );
    }else if(this.properties.pageType ==="home"){
      element= React.createElement(
        HomePage,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context:this.context,
          listName:this.properties.listName,
          libraryName:this.properties.libraryName,
          formType:this.properties.formType,
          listViews:this.properties.listViews,
          sp:this.properties.sp,
          noteType:this.properties.noteType,
          viewType:this.properties.viewType,
          title:this.properties.title,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          newPageUrl:this.properties.newPageUrl,
          viewPageUrl:this.properties.viewPageUrl,
          editPage:this.properties.editPage,
          fields:this.properties.fields,
          columnName:this.properties.columnName,
          chartType:this.properties.chartType,
          chartTitle:this.properties.chartTitle,
          customOption:this.properties.customOption,
          listId:this.properties.listId,
          pageType:this.properties.pageType,
          committeeType:this.properties.committeeType,
          httpUrl:this.properties.httpUrl,
          departmentGroupName:this.properties.departmentGroupName,
          superAdminGroupName:this.properties.superAdminGroupName,
          cmViewType:this.properties.cmViewType,
          CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName 
          
        }
      );
    }
    else if(this.properties.pageType ==="passcode"){
      element= React.createElement(
        PasscodePage,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context:this.context,
          listName:this.properties.listName,
          libraryName:this.properties.libraryName,
          formType:this.properties.formType,
          listViews:this.properties.listViews,
          sp:this.properties.sp,
          noteType:this.properties.noteType,
          viewType:this.properties.viewType,
          title:this.properties.title,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          newPageUrl:this.properties.newPageUrl,
          viewPageUrl:this.properties.viewPageUrl,
          editPage:this.properties.editPage,
          fields:this.properties.fields,
          columnName:this.properties.columnName,
          chartType:this.properties.chartType,
          chartTitle:this.properties.chartTitle,
          customOption:this.properties.customOption,
          listId:this.properties.listId,
          pageType:this.properties.pageType,
          committeeType:this.properties.committeeType,
          httpUrl:this.properties.httpUrl,
          departmentGroupName:this.properties.departmentGroupName,
          superAdminGroupName:this.properties.superAdminGroupName,
          cmViewType:this.properties.cmViewType,
          CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName 
          
        }
      );
    }
    else if(this.properties.pageType ==="atrView"){ 
      element= React.createElement(
        ATRViews,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context:this.context,
          listName:this.properties.listName,
          libraryName:this.properties.libraryName,
          formType:this.properties.formType,
          listViews:this.properties.listViews,
          sp:this.properties.sp,
          noteType:this.properties.noteType,
          viewType:this.properties.viewType,
          title:this.properties.title,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          newPageUrl:this.properties.newPageUrl,
          viewPageUrl:this.properties.viewPageUrl,
          editPage:this.properties.editPage,
          fields:this.properties.fields,
          columnName:this.properties.columnName,
          chartType:this.properties.chartType,
          chartTitle:this.properties.chartTitle,
          customOption:this.properties.customOption,
          listId:this.properties.listId,
          pageType:this.properties.pageType,
          committeeType:this.properties.committeeType,
          httpUrl:this.properties.httpUrl,
          departmentGroupName:this.properties.departmentGroupName,
          superAdminGroupName:this.properties.superAdminGroupName,
          cmViewType:this.properties.cmViewType,
          CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName 
          
        }
      );
    }
    else if(this.properties.pageType ==="search"){ 
      element= React.createElement(
        SearchPage,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context:this.context,
          listName:this.properties.listName,
          libraryName:this.properties.libraryName,
          formType:this.properties.formType,
          listViews:this.properties.listViews,
          sp:this.properties.sp,
          noteType:this.properties.noteType,
          viewType:this.properties.viewType,
          title:this.properties.title,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          newPageUrl:this.properties.newPageUrl,
          viewPageUrl:this.properties.viewPageUrl,
          editPage:this.properties.editPage,
          fields:this.properties.fields,
          columnName:this.properties.columnName,
          chartType:this.properties.chartType,
          chartTitle:this.properties.chartTitle,
          customOption:this.properties.customOption,
          listId:this.properties.listId,
          pageType:this.properties.pageType,
          committeeType:this.properties.committeeType,
          httpUrl:this.properties.httpUrl,
          departmentGroupName:this.properties.departmentGroupName,
          superAdminGroupName:this.properties.superAdminGroupName,
          cmViewType:this.properties.cmViewType,
          CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName 
          
        }
      );
    }
    else if(this.properties.pageType ==="CommitteeViews"){ 
      element= React.createElement(
        CommitteeMeetingListViews,
        {
          description: this.properties.description,
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          context:this.context,
          listName:this.properties.listName,
          libraryName:this.properties.libraryName,
          formType:this.properties.formType,
          listViews:this.properties.listViews,
          sp:this.properties.sp,
          noteType:this.properties.noteType,
          viewType:this.properties.viewType,
          title:this.properties.title,
          displayMode: this.displayMode,
          updateProperty: (value: string) => {
            this.properties.title = value;
          },
          newPageUrl:this.properties.newPageUrl,
          viewPageUrl:this.properties.viewPageUrl,
          editPage:this.properties.editPage,
          fields:this.properties.fields,
          columnName:this.properties.columnName,
          chartType:this.properties.chartType,
          chartTitle:this.properties.chartTitle,
          customOption:this.properties.customOption,
          listId:this.properties.listId,
          pageType:this.properties.pageType,
          committeeType:this.properties.committeeType,
          httpUrl:this.properties.httpUrl,
          departmentGroupName:this.properties.departmentGroupName,
          superAdminGroupName:this.properties.superAdminGroupName,
          cmViewType:this.properties.cmViewType,
          CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName 
          
        }
      );
    }
    else{
        element= React.createElement(
      XenWpUcoBank,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context:this.context,
        listName:this.properties.listName,
        libraryName:this.properties.libraryName,
        formType:this.properties.formType,
        listViews:this.properties.listViews,
        sp:this.properties.sp,
        noteType:this.properties.noteType,
        viewType:this.properties.viewType,
        title:this.properties.title,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        newPageUrl:this.properties.newPageUrl,
        viewPageUrl:this.properties.viewPageUrl,
        editPage:this.properties.editPage,
        fields:this.properties.fields,
        columnName:this.properties.columnName,
        chartType:this.properties.chartType,
        chartTitle:this.properties.chartTitle,
        customOption:this.properties.customOption,
        listId:this.properties.listId,
        pageType:this.properties.pageType,
        committeeType:this.properties.committeeType,
        httpUrl:this.properties.httpUrl,
        departmentGroupName:this.properties.departmentGroupName,
        superAdminGroupName:this.properties.superAdminGroupName,
        cmViewType:this.properties.cmViewType,
        CBnewPageUrl:this.properties.CBnewPageUrl,
          CBviewPageUrl:this.properties.CBviewPageUrl,
          CBeditPage:this.properties.CBeditPage,
          noteListName:this.properties.noteListName 
        
      }
    );

    }
  
    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    this.properties.sp=spfi().using(SPFx(this.context));

    this._services = new SPService(this.context);
    if (this.properties.customOption === undefined) {
      this.properties.customOption = [];
    }
    if (this.properties.listName !== undefined) {
      await this.getListFields();
    }

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
 
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('listName', {
                  label: 'Select a list',
                  selectedList: this.properties.listName,
                  includeHidden: false,
                  includeListTitleAndUrl:true,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.listConfigurationChanged,
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  baseTemplate:100,
                  key: 'listPickerFieldId'
                }),
             
                PropertyPaneDropdown("noteType",{//formType
                  label:"Select Note Type",
                  selectedKey:this.properties.noteType,
                  options:[{
                    key:"enote",text:"eNote",
                  },
                {
                  key:"eCommittee",text:"eCommittee",
                },
                {
                  key:"eCommitteeMeeting",text:"eCommittee Meeting",
                }
              ]
                }),   
                (this.properties.noteType && this.properties.noteType==="eCommittee") &&
                PropertyPaneDropdown("committeeType",{//formType
                  label:"Select Committee Type",
                  selectedKey:this.properties.committeeType,
                  options:[{
                    key:"Committee",text:"Committee",
                  },
                {
                  key:"Board",text:"Board",
                },
               
              ]
                }),
                (this.properties.pageType && this.properties.pageType ==="CommitteeViews") &&
                PropertyPaneDropdown("cmViewType",{
                  label:"Select Committee Meeting View Type",
                  selectedKey:this.properties.cmViewType,
                  options:[{
                    key:"CommitteeUnmappedRecords",text:"Committee Unmapped Records",
                  },
                {
                  key:"MyPendingCommitteeRecords",text:"My Pending Committee Records",
                },
                {
                  key:"AllInprogressCommitteeMeetingRecords",text:"All In-progress Committee Meeting Records",
                },
                {
                  key:"MyApprovedCommitteeRecords",text:"MyApprovedCommitteeRecords",
                },
                {
                  key:"AllApprovedCommitteeMeetings",text:"All Approved Committee Meetings",
                },
               
              ]
                }),
                (this.properties.noteType && this.properties.noteType !=="eCommitteeMeeting") &&
                PropertyPaneDropdown("pageType",{
                  label:"Select Page Type",
                  selectedKey:this.properties.pageType,
                  options:[{
                    key:"home",text:"Home",
                  },
                {
                  key:"views",text:"Views",
                },
                {
                  key:"charts",text:"Charts",
                },
                {
                  key:"passcode",text:"Passcode",
                },
                {
                  key:"atrView",text:"ATR Views",
                },
                {
                  key:"search",text:"Search",
                },
                {
                 key:"CommitteeViews",text:"Committee Views",
                }
              ]
                }),
                (this.properties.pageType && this.properties.pageType ==="atrView")&&
                PropertyFieldListPicker('noteListName', {
                  label: 'Select a Note list',
                  selectedList: this.properties.noteListName,
                  includeHidden: false,
                  includeListTitleAndUrl:true,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.listConfigurationChanged,
                  properties: this.properties,
                  context: this.context,
                  deferredValidationTime: 0,
                  baseTemplate:100,
                  key: 'listPickerFieldIdnoteListName'
                }),
                
                (this.properties.pageType && this.properties.pageType ==="atrView")&&
                PropertyPaneDropdown("viewType",{
                  label:"Select View Type",
                  selectedKey:this.properties.viewType,
                  options:[{
                    key:"pendingATR",text:"Pending ATR",
                  },
                {
                  key:"pendingATRSect",text:"Pending ATR Sect",
                },
                {
                  key:"completedATR",text:"Completed ATR",
                },
                {
                  key:"allATR",text:"All ATR",
                },
                
                
              ]
                }),
                (this.properties.pageType && this.properties.pageType ==="passcode")&&
                PropertyPaneTextField('httpUrl', {
                  label: "Enter PowerAutomate HTTP Url"
                 }),
             
                (this.properties.pageType &&this.properties.pageType ==="views")&&
                PropertyPaneDropdown("viewType",{
                  label:"Select View Type",
                  selectedKey:this.properties.viewType,
                  options:[{
                    key:"All Requests",text:"All Requests",
                  },
                {
                  key:"Draft Requests",text:"Draft Requests",
                },
                {
                  key:"In Progress",text:"In Progress",
                },
                {
                  key:"All Approved",text:"All Approved",
                },
                {
                  key:"All Rejected",text:"All Rejected",
                },
                
                  {
                    key:"Noted Notes",text:"Noted Notes",
                  },
                  {
                    key:"MyNotes",text:"MyNotes",
                  },
                  {
                    key:"MyPendingNotes",text:"My Pending Notes"
                  },
                  {
                    key:"MyReferredNotes",text:"My Recommended / Referred Notes "
                  },
                  {
                    key:"MyReturnedNotes",text:"My Returned Notes"
                  },
                  {
                    key:"MyApprovedNotes",text:"My Approved Notes"
                  },
                  {
                    key:"EDMDNotes",text:"ED/MD Notes"
                  },
                  {
                    key:"PendingWith",text:"Pending With"
                  },
                  {
                    key:"DBNoteReports",text:"Note Reports"
                  },
                  {
                    key:"DBATRReports",text:"ATR Status Reports"
                  }
                
              ]
                }),
                (this.properties.pageType &&this.properties.pageType ==="charts")&&
                PropertyPaneDropdown("chartType", {
                  label: "Chart Type:",
                  options: this.charts,
                  selectedKey: this.properties.chartType,
                }),
                (this.properties.pageType &&this.properties.pageType ==="charts")&&
                PropertyPaneDropdown("columnName", {
                  label: "Column Name:",
                  options: this.properties.customOption,
                  selectedKey: this.properties.columnName,
                }),
            
                (this.properties.pageType &&(this.properties.pageType ==="charts" || this.properties.pageType ==="views"))&&
                PropertyFieldMultiSelect("fields", {
                  key: "multiSelect",
                  label: "Select Columns for Exporting to Excel:",
                  options: this.properties.customOption,
                  selectedKeys: this.properties.fields,
                }),
                (this.properties.pageType && this.properties.pageType ==="Views")&&
                PropertyFieldMultiSelect("fields", {
                  key: "multiSelect",
                  label: "Select Columns for Exporting to Excel:",
                  options: this.properties.customOption,
                  selectedKeys: this.properties.fields,
                }),
                PropertyPaneTextField('newPageUrl', {
                  label: "New Page Url"
                 }),
             PropertyPaneTextField('editPage', {
                  label: "Edit Page Url"
                 }),
                PropertyPaneTextField('viewPageUrl', {
                  label: "View Page Url"
                 }),

                
                 (this.properties.noteType && this.properties.noteType ==="eCommittee")&&
                 PropertyPaneTextField('CBnewPageUrl', {
                  label: "Board New Page Url"
                 }),
                 (this.properties.noteType && this.properties.noteType ==="eCommittee")&&
                 PropertyPaneTextField('CBviewPageUrl', {
                  label: "Board View Page Url"
                 }),
                 (this.properties.noteType && this.properties.noteType ==="eCommittee")&&
                 PropertyPaneTextField('CBeditPage', {
                  label: "Board Edit Page Url"
                 }),
                 PropertyPaneTextField('superAdminGroupName', {
                  label: "Super Admin Group Name"
                 }),
                 PropertyPaneTextField('departmentGroupName', {
                  label: "Depart Admin Group Name"
                 })
                
              ]
            }
          ]
        }
      ]
    };
  }
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }
  public  getListFields=async():Promise<void>=> {
    if (this.properties.listName) {
      const allFields = await this._services.getcolumnInfo(this.properties.listName);
      (this.properties.customOption as []).length = 0;
      this.properties.customOption.push(
        ...allFields.map((field) => ({
          key: field.key,
          text: field.text,
        }))
      );
    }
  }

  private listConfigurationChanged=async (
    propertyPath: string,
    oldValue: IlistDetails,
    newValue: IlistDetails
  ):Promise<void>=> {
    if (propertyPath === "listName" && newValue) {
      this.properties.fields = [];
      this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
       await this.getListFields();
      this.context.propertyPane.refresh();
    } else {
      super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    }
  }
}

