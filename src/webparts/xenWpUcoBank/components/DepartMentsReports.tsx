import * as React from "react";
import styles from "./styles/superAdmin.module.scss";
import type { IXenWpUcoBankProps } from "./IXenWpUcoBankProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "@pnp/sp/site-users/web";

import {
  ChartControl, ChartType,

} from "@pnp/spfx-controls-react/lib/ChartControl";

import {
  IIconProps,
  CommandBar,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Link,
  PrimaryButton,
  SearchBox,
  SelectionMode,
  Selection, Spinner, SpinnerSize
} from "@fluentui/react";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import { Pagination } from "../../../Common/PageNation";

export interface IChartViewData {
  [x: string]: string
}
export interface IChartStates {
  chartViewData: IChartViewData;
  DropDownValue: string;
}
const excelIcon: IIconProps = { iconName: "ExcelDocument" };


export interface ISuperAdminState {
  listItems: any[];
  columns: IColumn[];
  page: number;
  rowsPerPage?: any;
  pageOfItems: any[];
  allItems: any[];
  searchText: string;
  departmentName: string;
  departmentAlias: string;
  selectionDetails: any;
  selectedcount: number;
  currentUserDetails:any;

}
export default class Das extends React.Component<
  IXenWpUcoBankProps,
  ISuperAdminState,
  {}
> {
  private _listName: any;
  private _selection: Selection;
  private _hideCommandOption:boolean=false;
  private chartOptions:any= {
    legend: {
      display: true,
      position: "left",
    },
    title: {
      display: false,
      text: this.props.chartTitle,
    },
  };
  private _columns: IColumn[] = [
    {
      key: "column1",
      name: "S.No",
      fieldName: "Id",
      minWidth: 50,
      maxWidth: 50,
      isRowHeader: true,
      isResizable: true,
      data: "string",
      onRender: (item, index, column) => {
        return item.Id;
      },
    },
    {
      key: "column2",
      name: "Department Name",
      fieldName: "Department",
      minWidth: 150,
      maxWidth: 350,
      isRowHeader: true,
      isResizable: true,
      data: "string",
      isMultiline: true,
    },
    {
      key: "column22",
      name: "DepartmentAlias Name",
      fieldName: "DepartmentAlias",
      minWidth: 150,
      maxWidth: 350,
      isRowHeader: true,
      isResizable: true,
      data: "string",
      isMultiline: true,
    },
    {
      key: "column3",
      name: "Admin",
      fieldName: "Admin",

      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
      data: "string",
      isMultiline: true,
    },
    {
      key: "column3",
      name: "Created By",
      fieldName: "Author",

      minWidth: 150,
      maxWidth: 300,
      isResizable: true,
      data: "string",
      isMultiline: true,
    },
    {
      key: "column4",
      name: "Created Date",
      fieldName: "Created",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      data: "string",
      onRender: (item, index, column) => {
        return (
          <span>
            {new Date(item.Created).toDateString() +
              " " +
              new Date(item.Created).toLocaleTimeString()}
          </span>
        );
      },
    },
    {
      key: "column5",
      name: "Modified By",
      fieldName: "Editor",
      minWidth: 150,
      maxWidth: 300,
      isMultiline: true,
      isResizable: true,
      data: "string",
    },
    {
      key: "column6",
      name: "Modified Date",
      fieldName: "Modified",
      minWidth: 100,
      maxWidth: 150,
      isResizable: true,
      onRender: (item, index, column) => {
        return (
          <span>
            {new Date(item.Modified).toDateString() +
              " " +
              new Date(item.Modified).toLocaleTimeString()}
          </span>
        );
      },
    },
    {
      key: "column7",
      name: "Action",
      fieldName: "ID",
      minWidth: 70,
      maxWidth: 90,

     
    },
  ];
  constructor(props: IXenWpUcoBankProps) {
    super(props);
    this._selection = new Selection({
      onSelectionChanged: () => {
        if (this._selection.getSelectedCount() === 0) {
          this._hideCommandOption = true;
        } else {
           this._hideCommandOption = false;
        }
        this.setState({
          selectionDetails: this._selection.getSelection()[0],
          selectedcount: this._selection.getSelectedCount(),
        });
      },
    });
    this.state = {
      listItems: [],
      columns: this._bandColumns(),
      page: 1,
      rowsPerPage: 10,
      pageOfItems: [],
      allItems: [],
      searchText: "",
      departmentName: "",
      departmentAlias: "",
      selectionDetails: {},
      selectedcount: 0,
      currentUserDetails:""
    };
    const listObj = this.props.listName;
    this._listName = listObj?.title;
    console.log(this.props.listName);

    this.getAllrequestesData();
  }



  private _getChartData = (uniqueValues: string[], allItems: string[] | any, columnName: string): any => {
    const llbArr: string[] = [];
    const dataArr: number[] = [];
    uniqueValues.forEach((uniqueValue: string) => {
      const arr = allItems.filter(
        (item: { [x: string]: string | any; }) => item[columnName] && item[columnName] === uniqueValue //data type channged 
      );
      llbArr.push(uniqueValue);
      dataArr.push(arr.length);
    });
    const chartViewData:any = {
      labels: llbArr,
      datasets: [
        {
          label: "Dataset",
          data: dataArr,
        },
      ],
    };
    return chartViewData;
  }
  private redirectToViewPage=(item:any)=>{
    window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
  }
/*   // Column rendering based on selected  view typ */
  private _bandColumns = () => {
    const { viewType } = this.props;
    switch (viewType) {
      case "All Requests":
        this._columns = [
          {
            key: "column1",
            name: "Note Number",
            fieldName: "Title",
            minWidth: 130,
            maxWidth: 150,
            isRowHeader: true,
            isResizable: true,
            isMultiline: true,
            data: "string",
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return <Link onClick={()=>this.redirectToViewPage(item)}>{item.Title}</Link>;
            },
          },
          {
            key: "column2",
            name: "Requester",
            fieldName: "Author",
            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column3",
            name: "Department",
            fieldName: "Department",
            minWidth: 150,
            maxWidth: 350,
            isRowHeader: true,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column4",
            name: "Subject",
            fieldName: "Subject",

            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column5",
            name: "Current Approver",
            fieldName: "CurrentApprover",

            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column6",
            name: "Status",
            fieldName: "Status",

            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column7",
            name: "Modified Date",
            fieldName: "Modified",
            minWidth: 100,
            maxWidth: 150,
            isMultiline: true,
            isResizable: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return (
                <span>
                  {new Date(item.Modified).toDateString() +
                    " " +
                    new Date(item.Modified).toLocaleTimeString()}
                </span>
              );
            },
          },

          {
            key: "column8",
            name: "Created Date",
            fieldName: "Created",
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            isMultiline: true,
            data: "string",
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return (
                <span>
                  {new Date(item.Created).toDateString() +
                    " " +
                    new Date(item.Created).toLocaleTimeString()}
                </span>
              );
            },
          },
         
        ];

        break;

      case "Noted Notes":
        this._columns = [
          {
            key: "column1",
            name: "Note Number",
            fieldName: "Title",
            minWidth: 130,
            maxWidth: 150,
            isRowHeader: true,
            isResizable: true,
            isMultiline: true,
            data: "string",
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return <Link onClick={()=>this.redirectToViewPage(item)}>{item.Title}</Link>;
            },
          },
          {
            key: "column2",
            name: "Requester",
            fieldName: "Author",
            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column3",
            name: "Department",
            fieldName: "Department",
            minWidth: 150,
            maxWidth: 350,
            isRowHeader: true,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          // DepartmentAlias
          {
            key: "column4",
            name: "Subject",
            fieldName: "Subject",

            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
        
          {
            key: "column7",
            name: "Modified Date",
            fieldName: "Modified",
            minWidth: 100,
            maxWidth: 150,
            isMultiline: true,
            isResizable: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return (
                <span>
                  {new Date(item.Modified).toDateString() +
                    " " +
                    new Date(item.Modified).toLocaleTimeString()}
                </span>
              );
            },
          },

          {
            key: "column8",
            name: "Created Date",
            fieldName: "Created",
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            isMultiline: true,
            data: "string",
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return (
                <span>
                  {new Date(item.Created).toDateString() +
                    " " +
                    new Date(item.Created).toLocaleTimeString()}
                </span>
              );
            },
          },
        
        ];
        break;
        case "Draft Requests":
          this._columns = [
            {
              key: "column1",
              name: "Note Number",
              fieldName: "Title",
              minWidth: 130,
              maxWidth: 150,
              isRowHeader: true,
              isResizable: true,
              isMultiline: true,
              data: "string",
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
              onRender: (item, index, column) => {
                return <Link onClick={()=>this.redirectToViewPage(item)}>{item.Title}</Link>;
              },
            },
            {
              key: "column2",
              name: "Requester",
              fieldName: "Author",
              minWidth: 150,
              maxWidth: 300,
              isResizable: true,
              data: "string",
              isMultiline: true,
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
            },
            {
              key: "column3",
              name: "Department",
              fieldName: "Department",
              minWidth: 150,
              maxWidth: 350,
              isRowHeader: true,
              isResizable: true,
              data: "string",
              isMultiline: true,
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
            },
            // DepartmentAlias
            {
              key: "column4",
              name: "Subject",
              fieldName: "Subject",
  
              minWidth: 150,
              maxWidth: 300,
              isResizable: true,
              data: "string",
              isMultiline: true,
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
            },
            {
              key: "column5",
              name: "Current Approver",
              fieldName: "CurrentApprover",
  
              minWidth: 150,
              maxWidth: 300,
              isResizable: true,
              data: "string",
              isMultiline: true,
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
            },
            {
              key: "column6",
              name: "Final Approver",
              fieldName: "FinalApprover",
  
              minWidth: 150,
              maxWidth: 300,
              isResizable: true,
              data: "string",
              isMultiline: true,
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
            },
            {
              key: "column7",
              name: "Modified Date",
              fieldName: "Modified",
              minWidth: 100,
              maxWidth: 150,
              isMultiline: true,
              isResizable: true,
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
              onRender: (item, index, column) => {
                return (
                  <span>
                    {new Date(item.Modified).toDateString() +
                      " " +
                      new Date(item.Modified).toLocaleTimeString()}
                  </span>
                );
              },
            },
  
            {
              key: "column8",
              name: "Created Date",
              fieldName: "Created",
              minWidth: 100,
              maxWidth: 150,
              isResizable: true,
              isMultiline: true,
              data: "string",
              onColumnClick: this._onColumnClick,
              isSorted: false,
              isSortedDescending: false,
              sortAscendingAriaLabel: "Sorted A to Z",
              sortDescendingAriaLabel: "Sorted Z to A",
              onRender: (item, index, column) => {
                return (
                  <span>
                    {new Date(item.Created).toDateString() +
                      " " +
                      new Date(item.Created).toLocaleTimeString()}
                  </span>
                );
              },
            },
            {
              key: "column9",
              name: "Action",
              fieldName: "ID",
              minWidth: 70,
              maxWidth: 90,
        

            }, 
           
          ];
          break;
      default:
        this._columns = [
          {
            key: "column1",
            name: "Note Number",
            fieldName: "Title",
            minWidth: 100,
            maxWidth: 100,
            isRowHeader: true,
            isResizable: true,
            isMultiline: true,
            data: "string",
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return <Link onClick={()=>this.redirectToViewPage(item)}>{item.Title}</Link>;
            },
          },
          {
            key: "column2",
            name: "Requester",
            fieldName: "Author",
            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column3",
            name: "Department",
            fieldName: "Department",
            minWidth: 150,
            maxWidth: 350,
            isRowHeader: true,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          // DepartmentAlias
          {
            key: "column4",
            name: "Subject",
            fieldName: "Subject",

            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column5",
            name: "Current Approver",
            fieldName: "CurrentApprover",

            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column6",
            name: "Final Approver",
            fieldName: "FinalApprover",

            minWidth: 150,
            maxWidth: 300,
            isResizable: true,
            data: "string",
            isMultiline: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
          },
          {
            key: "column7",
            name: "Modified Date",
            fieldName: "Modified",
            minWidth: 100,
            maxWidth: 150,
            isMultiline: true,
            isResizable: true,
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return (
                <span>
                  {new Date(item.Modified).toDateString() +
                    " " +
                    new Date(item.Modified).toLocaleTimeString()}
                </span>
              );
            },
          },

          {
            key: "column8",
            name: "Created Date",
            fieldName: "Created",
            minWidth: 100,
            maxWidth: 150,
            isResizable: true,
            isMultiline: true,
            data: "string",
            onColumnClick: this._onColumnClick,
            isSorted: false,
            isSortedDescending: false,
            sortAscendingAriaLabel: "Sorted A to Z",
            sortDescendingAriaLabel: "Sorted Z to A",
            onRender: (item, index, column) => {
              return (
                <span>
                  {new Date(item.Created).toDateString() +
                    " " +
                    new Date(item.Created).toLocaleTimeString()}
                </span>
              );
            },
          },
         
        ];
        break;
    }
    return this._columns;
  };

  private _getBylistWithFilterQuery = async () => {
    /* 
    Draft -  100
Call back - 200
Cancel - 300
Submit - 1000
Pending Reviewer - 2000
Pending Approver - 3000
Refer - 4000
Return - 5000
Reject - 8000
Approved - 9000 */
let user = await this.props.sp?.web.currentUser();
console.log(user,"user")
    let filterQury = "";
    switch (this.props.viewType) {
      case "All Requests":
        filterQury = "StatusNumber ne '100' ";
        break;
      case "Draft Requests":
        filterQury = `StatusNumber eq '100' and AuthorId eq ${user?.Id} `;
        break;
      case "In Progress":
        filterQury = `(CurrentApproverId eq ${user?.Id}) and (StatusNumber eq '2000' or StatusNumber eq '3000' or StatusNumber eq '4000')`;
        break;
      case "All Rejected":
        filterQury = "StatusNumber eq '8000' ";
        break;
      case "All Approved":
        filterQury = "StatusNumber eq '9000' ";
        break;
      case "Noted Notes":
        filterQury = "NatureOfNote eq 'Information' ";
        break;

      default:
        break;
    }
    return filterQury;
  };

  private _getItemsCount = async (columnTitle: string): Promise<any> => {
   
/* 
    // API Call. */
    const items = await this.props.sp?.web.lists.getByTitle(this._listName).items();
    const uniqueValues = this._getUniqueValue(items, columnTitle);
  
    return this._getChartData(uniqueValues, items, columnTitle);
  }
  private _getUniqueValue = (items: string[] | any, columnName: string): string[] => {//datatype chenged
    const values: string[] = [];
    const uniqueValues: string[] = [];
    items.forEach((item: { [x: string]: any; }) => {

      values.push(item["" + columnName + ""]);
    });
    values.map((statusValue) => {
      if (uniqueValues.indexOf(statusValue) === -1) {
        uniqueValues.push(statusValue);
      }
    });
    return uniqueValues;
  }
  private getAllrequestesData = async () => {
    const filterQury = await this._getBylistWithFilterQuery();

    const allItems: any = [];
    const items: any = await this.props.sp?.web.lists
      .getByTitle(this._listName)
      .items.filter(filterQury)
      .select(
        `*,Created,Modified,Created,Author/Title,Editor/Title,CurrentApprover/Title,CurrentApprover/EMail,PreviousApprover/Title,PreviousApprover/EMail,FinalApprover/Title,FinalApprover/EMail`
      )
      .expand(`Author,Editor,PreviousApprover,CurrentApprover,FinalApprover`).orderBy("Modified",false)();
    items.map((obj: any) => {
      allItems.push({
        Id: obj.Id,
        Department: obj.Department,
        NoteNumber: obj.NoteNumber,
        Subject: obj.Subject,
        Status: obj.Status,
        CurrentApprover:
          obj.CurrentApprover === null && obj.CurrentApproverId === null
            ? ""
            : obj.CurrentApprover?.Title,
        PreviousApprover:
          obj.PreviousApprover === null && obj.PreviousApproverId === null
            ? ""
            : obj.PreviousApprover?.Title,
        FinalApprover:
          obj.FinalApprover === null && obj.FinalApproverId === null
            ? ""
            : obj.FinalApprover?.Title,
        Title: obj.Title,

        Editor:
          obj.Editor === null && obj.EditorId === null ? "" : obj.Editor?.Title,
        Admin:
          obj.Admin === null && obj.AdminId === null ? "" : obj.Admin?.Title,
        Author:
          obj.Author === null && obj.AuthorId === null ? "" : obj.Author?.Title,
        Created: obj.Created,
        Modified: obj.Modified,
      });
    });

    this.setState({
      listItems: allItems,
      allItems: allItems,
    });
  };

  private paginateFn = (filterItem: any[]) => {
    let { rowsPerPage, page } = this.state;
    console.log(filterItem,"filterItem")
    return rowsPerPage > 0
      ? filterItem.slice(
          (page - 1) * rowsPerPage,
          (page - 1) * rowsPerPage + rowsPerPage
        )
      : filterItem;
      
  };

  private handlePaginationChange(pageNo: number, rowsPerPage: number) {
    this.setState({ page: pageNo, rowsPerPage: rowsPerPage });
  }



 

  // private _deleteSelectedUser = async () => {
  //   this.setState({
  //     hideDeleteDialog: true,
  //   });
  //   try {
  //     await this.props.sp?.web.lists
  //       .getByTitle(this._listName)
  //       .items.getById(this.state.selectedId)
  //       .delete();
  //     this.setState({
  //       hideSuccussDialog: false,
  //       succussMsg: "Request has been deleted successfuly",
  //     });
  //   } catch (err) {
  //     this.setState({
  //       hideWarningDialog: false,
  //       warningMsg: "Failed to delete this request. Please try again",
  //     });
  //   }
  // };

 

  private _onChangeFilterText = (
    event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void => {
    console.log(newValue, "test");
    this.setState({
      listItems: newValue
        ? this.state.allItems.filter((item: any) =>
            Object.values(item).some(
              (value: any) =>
                (value || "")
                  .toString()
                  .toLowerCase()
                  .indexOf(newValue.toLowerCase()) > -1
            )
          )
        : this.state.allItems,
    });
    this.paginateFn(this.state.listItems);
    // }
  };

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { columns, listItems } = this.state;
    const newColumns: IColumn[] = columns.slice();
    const currColumn: any = newColumns.filter(
      (currCol) => column.key === currCol.key
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    const newItems = this._copyAndSort(
      listItems,
      currColumn.fieldName,
      currColumn.isSortedDescending
    );
    this.setState({
      columns: newColumns,
      listItems: newItems,
    });
  };

  private _copyAndSort<T>(
    items: T[],
    columnKey: string,
    isSortedDescending?: boolean
  ): T[] {
    const key = columnKey as keyof T;
    return items.slice(0).sort((a: T, b: T) => {
      if (a[key] > b[key]) {
        return isSortedDescending ? -1 : 1;
      } else if (a[key] < b[key]) {
        return isSortedDescending ? 1 : -1;
      }
      return 0;
    });
  }
  public render(): React.ReactElement<IXenWpUcoBankProps> {
    const { hasTeamsContext } = this.props;

    const _items =this.props.viewType === "Draft Requests"?[
      {
        key: "newItem",
        name: "Create New Request",
        iconProps: {
          iconName: "Add",
        },
        split: true,
        onClick: () => {
          window.location.href =
            this.props.context.pageContext.web.absoluteUrl +
            `/SitePages/${this.props.newPageUrl}.aspx`;
        },
      },
      {
        key: 'EditItem',
        name: 'Edit Request',
        iconProps: {
          iconName: 'Edit'
        },
        disabled: this._hideCommandOption,
        onClick: () => {
          const item = this.state.selectionDetails;
          if (this.state.selectedcount === 0) {
            return
          } else {
            window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`

          }
        }
      },
      {
        key: "ViewItem",
        name: "View Request",
        iconProps: {
          iconName: "View",
        },
        className: "viewBtnDiv",
        disabled: this._hideCommandOption,
        onClick: () => {
          const item = this.state.selectionDetails;
          console.log()
          console.log(item,"item")
          if (this.state.selectedcount === 0) {
            return
          } else {
            window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
          }
        },
      },
    ] :[
      {
        key: "newItem",
        name: "Create New Request",
        iconProps: {
          iconName: "Add",
        },
        split: true,
        onClick: () => {
          window.location.href =
            this.props.context.pageContext.web.absoluteUrl +
            `/SitePages/${this.props.newPageUrl}.aspx`;
        },
      },
     
      {
        key: "ViewItem",
        name: "View Request",
        iconProps: {
          iconName: "View",
        },
        className: "viewBtnDiv",
        disabled: this._hideCommandOption,
        onClick: () => {
          const item = this.state.selectionDetails;
          console.log()
          console.log(item,"item")
          if (this.state.selectedcount === 0) {
            return
          } else {
            window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
          }
        },
      },
    ];
    return (
      <section
        className={`${styles.xenWpUcoBank} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty}
        />
        <div className={styles.commandbarContainer}>
          <div>
            <CommandBar items={_items} />
          </div>
          <div>
            <SearchBox
              placeholder="Search"
              title="Search"
              onSearch={(newValue) => console.log(newValue)}
              onChange={this._onChangeFilterText}
            />
          </div>
        </div>

        {/* <fieldset className={styles._dataTable}> */}
        <div
          id="generateTable"
          className={styles._listviewContainerDataTable}
          data-is-scrollable="true"
        >
          <DetailsList
            data-is-scrollable="true"
            items={this.state.listItems}
            columns={this.state.columns}
            selectionMode={SelectionMode.single}
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selection}
            isHeaderVisible={true}
          />
         
        </div>
        <div>
            <Pagination
              currentPage={this.state.page}
              totalItems={this.state.listItems.length}
              onChange={this.handlePaginationChange.bind(this)}
            />
          </div>
          <div className={styles.ChartViewcontainer}>
        <div className={styles.ContorlSection}>

          <div className={styles.button}>
            <PrimaryButton
              text="Export"
              iconProps={excelIcon}
              allowDisabledFocus
            />
          </div>
        </div>
        <ChartControl
          type={ChartType[this.props.chartType as keyof typeof ChartType] || ChartType.Pie}

          datapromise={this._getItemsCount(
           
            this.props.columnName
          )}
          loadingtemplate={() => (
            <Spinner size={SpinnerSize.large} label="Loading..." />
          )}
          options={this.chartOptions}
        />
      </div>
      </section>

   

    );
  }
}
