import * as React from "react";
import styles from "../styles/superAdmin.module.scss";
import type { IXenWpUcoBankProps } from "../IXenWpUcoBankProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "@pnp/sp/site-users/web";
import "../CustomStyles/Custom.css";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Link,
  SearchBox,
  SelectionMode,
  Selection,
  CommandBar,
  Dropdown,
  Spinner,
  SpinnerSize,
} from "@fluentui/react";
import { Pagination } from "../../../../Common/PageNation";
import * as XLSX from "xlsx";
import * as FileSaver from "file-saver";
import {
  ChartControl,
  ChartType,
} from "@pnp/spfx-controls-react/lib/ChartControl";


const getIdFromUrl = (): any => {
  const params = new URLSearchParams(window.location.search);
  const Id = params.get("noteType");
  return Id;
};
export interface IListViewsState {
  listItems: any[];
  columns: IColumn[];
  page: number;
  rowsPerPage?: any;
  pageOfItems: any[];
  allItems: any[];
  selectionDetails: any;
  selectedcount: number;
  activeBtn: string;
  isSecarory: boolean;
  committeeType: string;
  dashboardCount: number;
  committeeMeetingData: any;
}
export default class HomePage extends React.Component<
  IXenWpUcoBankProps,
  IListViewsState,
  {}
> {
  private _listName: any;
  private _selection: Selection;
  private chartOptions: any = {
    legend: {
      display: true,
      position: "left",
    },
    title: {
      display: true,
      text: "Status",
    },
  };
  private chartOptionsNatureOfNote: any = {
    legend: {
      display: true,
      position: "left",
    },
    title: {
      display: true,
      text: "Nature Of Note",
    },
  };

  private chartOptionsCommitteeNames: any = {
    legend: {
      display: true,
      position: "left",
    },
    title: {
      display: true,
      text: "Committee Name",
    },
  };

  private _columns: IColumn[] = [
    {
      key: "column1",
      name: "S.No",
      fieldName: "Id",
      minWidth: 50,
      maxWidth: 50,
      isRowHeader: false,
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
      isRowHeader: false,
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
      isRowHeader: false,
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
        this.setState({
          selectionDetails: this._selection.getSelection()[0],
          selectedcount: this._selection.getSelectedCount(),
        });
      },
    });
    this.state = {
      listItems: [],
      columns: this._bindColumns(),
      page: 1,
      rowsPerPage: 10,
      pageOfItems: [],
      allItems: [],
      selectionDetails: {},
      selectedcount: 0,
      activeBtn: getIdFromUrl() || "MyNotes",
      isSecarory: false,
      committeeType: "CommitteeNote",
      dashboardCount: 0,
      committeeMeetingData: [],
    };
    const listObj = this.props.listName;
    this._listName = listObj?.title;
    getIdFromUrl();

    this.getSectoryinfo();
    if(this.props.noteType ==="eCommittee"){
      this.committeemeetingData();
    }
    this.getAllrequestesData(getIdFromUrl() || "MyNotes", "CommitteeNote");
  }

  private _formatDate = (date: Date) => {
 
    if (!(date instanceof Date)) {
      throw new Error("Invalid date");  
    }
    const day = date.getDate().toString().padStart(2, "0"); // Get individual date components
    const month = (date.getMonth() + 1).toString().padStart(2, "0"); // Month is zero-indexed
    const year = date.getFullYear().toString();
    const hours = date.getHours().toString().padStart(2, "0");
    const minutes = date.getMinutes().toString().padStart(2, "0");
    const seconds = date.getSeconds().toString().padStart(2, "0");
    return `${day}${month}${year}${hours}${minutes}${seconds}`; // Concatenate them in the desired format
  };

  private _getExcel = async (): Promise<void> => {
    const todayDate = new Date();
    const formatDate = this._formatDate(todayDate);
    const fieldNames = this.state.columns.map((obj) => obj.fieldName);
    const excelData = this.state.allItems?.map((item) => {
      const res: string[] = [];
      fieldNames.forEach((element:any) => {
        if (typeof item[element] === "object" && item[element] !== null) {
          res.push(item[element].join(", "));
        } else {
          res.push(item[element]);
        }
      });
      return res;
    });
    const fileType =
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8";
    const fileExtension = ".xlsx";
    const Heading = [this.state.columns?.map((obj) => obj.name)];
    if (excelData.length > 0) {
      const ws = XLSX.utils.book_new();
      XLSX.utils.sheet_add_aoa(ws, Heading);
      XLSX.utils.sheet_add_json(ws, excelData, {
        origin: "A2",
        skipHeader: true,
      });
      const wb = { Sheets: { data: ws }, SheetNames: ["data"] };
      const excelBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const data = new Blob([excelBuffer], { type: fileType });
      FileSaver(
        data,
        this.props.noteType + this.state.activeBtn + formatDate + fileExtension
      );
    }
  };

  private _commiteeRedirect=(item:any,user:any)=>{
    if (
      (item.StatusNumber === "100" ||
        item.StatusNumber === "200" ||
        item.StatusNumber === "5000") &&
      item.AuthorId === user.Id
    ) {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;
    } else {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
    }
  }

  private _BoardRedirect=(item:any,user:any)=>{
    if (
      (item.StatusNumber === "100" ||
        item.StatusNumber === "200" ||
        item.StatusNumber === "5000") &&
      item.AuthorId === user.Id
    ) {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBeditPage}.aspx?itemId=${item.Id}`;
    } else {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBviewPageUrl}.aspx?itemId=${item.Id}`;
    }
  }

  private _EnoteRedirect=(item:any,user:any)=>{
    if (
      (item.StatusNumber === "100" ||
        item.StatusNumber === "200" ||
        item.StatusNumber === "5000") &&
      item.AuthorId === user.Id
    ) {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;
    } else {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
    }
  }

  private redirectToViewPage = async (item: any) => {
    const user = await this.props.sp?.web.currentUser();
    if (this.props.noteType === "eCommittee") {
      if (item.CommitteeType === "Committee") {
       this._commiteeRedirect(item,user);
      }
      if (item.CommitteeType === "Board") {
     this._BoardRedirect(item,user);
      }
    } else {
     this. _EnoteRedirect(item,user)
      
    }
  };

  private _EnoteDetailslistColumns=()=>{
    const { viewType } = this.props;
    if (viewType === "All Requests") {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
          isResizable: true,
          isMultiline: true,
          data: "string",
          onColumnClick: this._onColumnClick,
          isSorted: false,
          isSortedDescending: true,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
          onRender: (item, index, column) => {
            return (
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
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
    } else if (viewType === "Noted Notes") {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
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
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
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
    }
    
     else if (
      viewType === "In Progress" ||
      viewType === "All Approved" ||
      viewType === "All Rejected" ||  viewType === "Draft Requests"
    ) {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
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
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
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
    } else if (
      viewType === "mypendingnotes" ||
      viewType === "MyReferredNotes" ||
      viewType === "MyReturnedNotes" ||
      viewType === "MyApprovedNotes" ||
      viewType === "EDMDNotes"
    ) {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
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
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
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
          key: "column66",
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
          key: "column5",
          name: "PreviousApprover",
          fieldName: "Previous Approver",

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
    } else if (viewType === "PendingWith") {
      this._PendingWithDetailslistColumns();
    
    } else {
    this._defalutDetailslistColumns();
    }

    return this._columns;
  }

  private _EcommiteeeDetailslistColumns=()=>{
    const { viewType } = this.props;
    if (viewType === "All Requests") {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
          isResizable: true,
          isMultiline: true,
          data: "string",
          onColumnClick: this._onColumnClick,
          isSorted: false,
          isSortedDescending: true,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
          onRender: (item, index, column) => {
            return (
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
          isResizable: true,
          data: "string",
          isMultiline: true,
          onColumnClick: this._onColumnClick,
          isSorted: false,
          isSortedDescending: false,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
        },
        this.props.noteType &&
          this.props.noteType === "eCommittee" && {
            key: "column33",
            name: "Board/Committee Name",
            fieldName: "committeeName",
            minWidth: 150,
            maxWidth: 350,
            isRowHeader: false,
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
    } else if (viewType === "Noted Notes") {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
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
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
          isResizable: true,
          data: "string",
          isMultiline: true,
          onColumnClick: this._onColumnClick,
          isSorted: false,
          isSortedDescending: false,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
        },

        this.props.noteType &&
          this.props.noteType === "eCommittee" && {
            key: "column33",
            name: "Board/Committee Name",
            fieldName: "committeeName",
            minWidth: 150,
            maxWidth: 350,
            isRowHeader: false,
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
    } 
    
     else if (
      viewType === "In Progress" ||
      viewType === "All Approved" ||
      viewType === "All Rejected" || viewType === "Draft Requests"
    ) {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
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
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
          isResizable: true,
          data: "string",
          isMultiline: true,
          onColumnClick: this._onColumnClick,
          isSorted: false,
          isSortedDescending: false,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
        },
        this.props.noteType &&
          this.props.noteType === "eCommittee" && {
            key: "column33",
            name: "Board/Committee Name",
            fieldName: "committeeName",
            minWidth: 150,
            maxWidth: 350,
            isRowHeader: false,
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
    } else if (
      viewType === "mypendingnotes" ||
      viewType === "MyReferredNotes" ||
      viewType === "MyReturnedNotes" ||
      viewType === "MyApprovedNotes" ||
      viewType === "EDMDNotes"
    ) {
      this._columns = [
        {
          key: "column1",
          name: "Note Number",
          fieldName: "Title",
          minWidth: 130,
          maxWidth: 150,
          isRowHeader: false,
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
              <Link onClick={() => this.redirectToViewPage(item)}>
                {item.Title}
              </Link>
            );
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
          isRowHeader: false,
          isResizable: true,
          data: "string",
          isMultiline: true,
          onColumnClick: this._onColumnClick,
          isSorted: false,
          isSortedDescending: false,
          sortAscendingAriaLabel: "Sorted A to Z",
          sortDescendingAriaLabel: "Sorted Z to A",
        },
        this.props.noteType &&
          this.props.noteType === "eCommittee" && {
            key: "column33",
            name: "Board/Committee Name",
            fieldName: "committeeName",
            minWidth: 150,
            maxWidth: 350,
            isRowHeader: false,
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
          key: "column66",
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
          key: "column5",
          name: "PreviousApprover",
          fieldName: "Previous Approver",

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
    } else if (viewType === "PendingWith") {
  this._PendingWithDetailslistColumns();
    } else {
     this. _defalutDetailslistColumns();
    }
    return this._columns;
  }

  private _PendingWithDetailslistColumns=()=>{
    this._columns = [
      {
        key: "column1",
        name: "S.No",
        fieldName: "ID",
        minWidth: 50,
        maxWidth: 100,
        isRowHeader: false,
        isResizable: true,
        isMultiline: true,
        data: "string",
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onRender: (item, index, column) => {
          return index ? index + 1 : null;
        },
      },
      {
        key: "column2",
        name: "Pending With",
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
        key: "column3",
        name: "Designation",
        fieldName: "crntApproverObject",
        minWidth: 150,
        maxWidth: 350,
        isRowHeader: false,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: false,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onRender: (item, index, column) => {
          if (item.crntApproverObject) {
            return item.crntApproverObject?.JobTitle;
          } else {
            return null;
          }
        },
      },
      {
        key: "column4",
        name: "SrNo",
        fieldName: "crntApproverObject",

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
        onRender: (item, index, column) => {
          if (item.crntApproverObject) {
            const SrNo = item.crntApproverObject?.EMail.split("@")[0];
            return SrNo;
          } else {
            return null;
          }
        },
      },
      {
        key: "column5",
        name: "Count",
        fieldName: "count",

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
    ];
  }

  private _defalutDetailslistColumns=()=>{
    this._columns = [
      {
        key: "column1",
        name: "Note Number",
        fieldName: "Title",
        minWidth: 130,
        maxWidth: 150,
        isRowHeader: false,
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
            <Link onClick={() => this.redirectToViewPage(item)}>
              {item.Title}
            </Link>
          );
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
        isRowHeader: false,
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
  }
 
  private _bindColumns = () => {  
    const {  noteType } = this.props;
    if (noteType === "enote") {
     this. _EnoteDetailslistColumns()
    
    } else if (noteType === "eCommittee") {
      this._EcommiteeeDetailslistColumns()
    }

  return this._columns;

  };

  private getSectoryinfo = async () => {
    const user = await this.props.sp?.web.currentUser();
    const items = await this.props.sp.web.lists
      .getByTitle("ApproverMatrix")
      .items.filter(
        `SecretaryId eq ${user.Id}  and ApproverType eq 'Approver' `
      )
      .select("*")();
    if (items && items?.length > 0) {
      this.setState({
        isSecarory: true,
      });
    }
  };

  private _getBylistWithFilterQuery = async (actionbtn: string) => {

    const user = await this.props.sp?.web.currentUser();
    const items = await this.props.sp.web.lists
      .getByTitle("ApproverMatrix")
      .items.filter(`SecretaryId eq ${user.Id} and ApproverType eq 'Approver' `)
      .select("*")();
    let filterQury = "";
    switch (actionbtn) {
      case "MyNotes":
        filterQury = `AuthorId eq ${user?.Id} `;
        break;
      case "mypendingnotes":
        filterQury = `CurrentApproverId eq ${user?.Id} `;
        break;
      case "MyReferredNotes":
        filterQury = `CurrentApproverId eq ${user?.Id} and StatusNumber eq '4000' `;
        break;
      case "MyReturnedNotes":
        filterQury = `AuthorId eq ${user?.Id} and StatusNumber eq '5000' `;
        break;
      case "MyApprovedNotes":
        filterQury = `AuthorId eq ${user?.Id} and StatusNumber eq '9000' `;
        break; 
      case "EDMDNotes":
        filterQury = `CurrentApproverId eq ${items[0].ApproverId} and StatusNumber eq '3000'  `;
        break;
      case "PendingWith":
        filterQury = `StatusNumber eq '2000' or StatusNumber eq '3000' or StatusNumber eq '4000' `;
        break;

      default:
        break;
    }
    return filterQury;
  };


  private _CommitteeNameCondition1 = (obj:any):any=>{
    return obj.BoardName === null
    ? ""
    : obj.BoardName
  }

  private getAllrequestesData = async (
    actionBtn: string,
    committeType?: string
  ) => {
    this.setState({ activeBtn: actionBtn });
    const filterQury = await this._getBylistWithFilterQuery(actionBtn);

    const allItems: any = [];
      const items: any = await this.props.sp?.web.lists
        .getByTitle(this._listName)
        .items.filter(filterQury)
        .select(
          `*,Created,Modified,
        Created,
        Author/Title,
        Editor/Title,
        CurrentApprover/Title,
        CurrentApprover/EMail,
        CurrentApprover/JobTitle,
        PreviousApprover/Title,
        PreviousApprover/EMail,
        FinalApprover/Title,
        FinalApprover/EMail`
        )
        .expand(`Author,Editor,PreviousApprover,CurrentApprover,FinalApprover`)
        .orderBy("Created", false)();
      items.map((obj: any) => {
        allItems.push({
          Id: obj.Id,
          Department: obj.Department,
          NoteNumber: obj.NoteNumber,
          Subject: obj.Subject,
          Status: obj.Status,
          StatusNumber: obj.StatusNumber,
          NatureOfNote: obj.NatureOfNote,
          CurrentApproverObj:
            obj.CurrentApprover === null && obj.CurrentApproverId === null
              ? null
              : obj.CurrentApprover,
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
            obj.Editor === null && obj.EditorId === null
              ? ""
              : obj.Editor?.Title,
          Admin:
            obj.Admin === null && obj.AdminId === null ? "" : obj.Admin?.Title,
          Author:
            obj.Author === null && obj.AuthorId === null
              ? ""
              : obj.Author?.Title,
          Created: obj.Created,
          Modified: obj.Modified,
          committeeName:
            obj.CommitteeName === null
              ?this._CommitteeNameCondition1(obj)
              : obj.CommitteeName,
          CommitteeType:
            obj.CommitteeType === null &&
            typeof obj.CommitteeType === "undefined"
              ? ""
              : obj.CommitteeType,
        });
      });
      if (
        this.props.noteType === "eCommittee" &&
        committeType === "BoardNote"
      ) {
        const boardAllItems = allItems.filter(
          (obj: any) => obj.CommitteeType === "Board"
        );
        this.paginateFn(boardAllItems, 1);
        this.setState({
          allItems: boardAllItems,
        });
      } else if (
        this.props.noteType === "eCommittee" &&
        committeType === "CommitteeNote"
      ) {
        const committeeAllItems = allItems.filter(
          (obj: any) => obj.CommitteeType !== "Board"
        );
        this.paginateFn(committeeAllItems, 1);
        this.setState({
          allItems: committeeAllItems,
        });
      } else {
        this.paginateFn(allItems, 1);
        this.setState({
          allItems: allItems,
        });
      }
    // }
    this.setState({ dashboardCount: allItems.length });
  };

  private paginateFn = (filterItem: any[], pageNo?: any) => {
    const { rowsPerPage } = this.state;
    const items =
      rowsPerPage > 0
        ? filterItem.slice(
            (pageNo - 1) * rowsPerPage,
            (pageNo - 1) * rowsPerPage + rowsPerPage
          )
        : filterItem;
    this.setState({ listItems: items, page: pageNo });
  };

  private handlePaginationChange = (page: any) => {
    this.setState({ page: page });
    this.paginateFn(this.state.allItems, page);
  };



  private _handleCommiteeType = (
    event: React.FormEvent<HTMLDivElement>,
    option?: any,
    index?: number
  ) => {
    if (option.key !== "CommitteeMeetings") {

    this.getAllrequestesData(this.state.activeBtn, option.key);
    }
    else{

      this.setState(
        (prevState) => ({
          activeBtn: prevState.activeBtn, 
        }),)
    }
    this.setState({
      committeeType: option.key,
    });
  };

  private _onChangeFilterText = (
    event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void => {
    this.setState((prevState) => {
      const filteredItems = newValue
        ? prevState.allItems.filter((item: any) =>
            Object.values(item).some(
              (value: any) =>
                (value || "")
                  .toString()
                  .toLowerCase()
                  .indexOf(newValue.toLowerCase()) > -1
            )
          )
        : prevState.allItems;
  
      // Perform pagination based on filtered items and current page
      this.paginateFn(filteredItems, prevState.page);
  
      return {
        listItems: filteredItems,
      };
    });
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

  private committeeChart=()=>{
    return(
      <div className="row">
                <div className="col-md-6 col-sm-12 _homePageCharts ">
                  <ChartControl
                    type={ChartType["Pie"]}
                    datapromise={this._getItemsCountCommittee(
                      "CommitteeMeetingRequests",
                      "MeetingStatus"
                    )}
                    loadingtemplate={() => (
                      <Spinner size={SpinnerSize.large} label="Loading..." />
                    )}
                    options={this.chartOptions}
                  />
                </div>
                <div className="col-md-6 col-sm-12">
                  <ChartControl
                    type={ChartType["Doughnut"]}
                    datapromise={this._getItemsCountCommittee(
                      "CommitteeMeetingRequests",
                      "CommitteeName"
                    )}
                    loadingtemplate={() => (
                      <Spinner size={SpinnerSize.large} label="Loading..." />
                    )}
                    options={this.chartOptionsCommitteeNames}
                  />
                </div>
              </div> 
    )
  }
  private _renderButtons = () => (
    <div className={styles.landingPgTopBtnRow}>
      <button
        className={
          this.state.activeBtn === "MyNotes"
            ? styles.landingPgTopBtn
            : styles.inActivelandingPgTopBtn
        }
        onClick={() => {
          this.getAllrequestesData("MyNotes");
          this.setState({ allItems: [], listItems: [] });
        }}
      >
        <span>
          <span>My Dashboard</span>
          {this.state.activeBtn === "MyNotes" && (
            <span className={styles.landingPgTopBtncontent}>
              {this.state.dashboardCount}
            </span>
          )}
        </span>
      </button>
      <button
        className={
          this.state.activeBtn === "mypendingnotes"
            ? styles.landingPgTopBtn
            : styles.inActivelandingPgTopBtn
        }
        onClick={() => {
          this.getAllrequestesData("mypendingnotes");
          this.setState({ allItems: [], listItems: [] });
        }}
      >
        <span>
          <span> My Pending Notes</span>
          {this.state.activeBtn === "mypendingnotes" && (
            <span className={styles.landingPgTopBtncontent}>
              {this.state.allItems.length}
            </span>
          )}
        </span>
      </button>
      <button
        className={
          this.state.activeBtn === "MyReferredNotes"
            ? styles.landingPgTopBtn
            : styles.inActivelandingPgTopBtn
        }
        onClick={() => {
          this.getAllrequestesData("MyReferredNotes");
          this.setState({ allItems: [], listItems: [] });
        }}
      >
        <span>
          <span>My Recommended / Referred Notes</span>
          {this.state.activeBtn === "MyReferredNotes" && (
            <span className={styles.landingPgTopBtncontent}>
              {this.state.allItems.length}
            </span>
          )}
        </span>
      </button>
      <button
        className={
          this.state.activeBtn === "MyReturnedNotes"
            ? styles.landingPgTopBtn
            : styles.inActivelandingPgTopBtn
        }
        onClick={() => {
          this.getAllrequestesData("MyReturnedNotes");
          this.setState({ allItems: [], listItems: [] });
        }}
      >
        <span>
          <span>My Returned Notes</span>
          {this.state.activeBtn === "MyReturnedNotes" && (
            <span className={styles.landingPgTopBtncontent}>
              {this.state.allItems.length}
            </span>
          )}
        </span>
      </button>
      <button
        className={
          this.state.activeBtn === "MyApprovedNotes"
            ? styles.landingPgTopBtn
            : styles.inActivelandingPgTopBtn
        }
        onClick={() => {
          this.getAllrequestesData("MyApprovedNotes");
          this.setState({ allItems: [], listItems: [] });
        }}
      >
        <span>
          <span> My Approved Notes</span>
          {this.state.activeBtn === "MyApprovedNotes" && (
            <span className={styles.landingPgTopBtncontent}>
              {this.state.allItems.length}
            </span>
          )}
        </span>

      </button>
      {this.state.isSecarory === true && (
        <button
          className={
            this.state.activeBtn === "EDMDNotes"
              ? styles.landingPgTopBtn
              : styles.inActivelandingPgTopBtn
          }
          onClick={() => {
            this.getAllrequestesData("EDMDNotes");
            this.setState({ allItems: [], listItems: [] });
          }}
        >
          <span>
            <span> ED/MD Notes</span>
            {this.state.activeBtn === "EDMDNotes" && (
              <span className={styles.landingPgTopBtncontent}>
                {this.state.allItems.length}
              </span>
            )}
          </span>

        </button>
      )}
    </div>
  );

  private committeemeetingData = async () => {
    const user = await this.props.sp?.web.currentUser();
    const allitems = await this.props.sp?.web.lists
      .getByTitle("CommitteeMeetingRequests")
      .items.filter(`AuthorId eq ${user?.Id} `)();
    this.setState({ committeeMeetingData: allitems || [] });
  };
 
  private _getItemsCount = async (
    listTitle: string,
    columnTitle: string
  ): Promise<any> => {
    const uniqueValues = this._getUniqueValue(
      this.state.allItems,
      columnTitle
    );
    return this._getChartData(
      uniqueValues,
       this.state.allItems,
      columnTitle
    );
  };
 
  private _getUniqueValue = (  
    items: string[] | any,
    columnName: string
  ): string[] => {
    const values: string[] = [];
    const uniqueValues: string[] = [];
    items.forEach((item: { [x: string]: any }) => {
      values.push(item["" + columnName + ""]);
    });
    values.map((statusValue) => {
      if (uniqueValues.indexOf(statusValue) === -1) {
        uniqueValues.push(statusValue);
      }
    });
    return uniqueValues;
  };
 
  private _getChartData = ( 
    uniqueValues: string[],
    allItems: string[] | any,
    columnName: string
  ): any => {
    const llbArr: string[] = [];
    const dataArr: number[] = [];
    uniqueValues.forEach((uniqueValue: string) => {
      const arr = allItems.filter(
        (item: { [x: string]: string | any }) =>
          item[columnName] && item[columnName] === uniqueValue //data type channged
      );
      llbArr.push(uniqueValue);
      dataArr.push(arr.length);
    });
    const chartViewData: any = {
      labels: llbArr,
      datasets: [
        {
          label: "Dataset",
          data: dataArr,
        },
      ],
    };
    return chartViewData;
  };

   private _getItemsCountCommittee = async ( 
    listTitle: string,
    columnTitle: string
  ): Promise<any> => {
    const uniqueValues = this._getUniqueValueCommittee(
      this.state.committeeMeetingData,
      columnTitle
    );
    return this._getChartDataCommittee(
      uniqueValues,
       this.state.committeeMeetingData,
      columnTitle
    );
  };
 
  private _getUniqueValueCommittee = (
    items: string[] | any,
    columnName: string
  ): string[] => {
    const values: string[] = [];
    const uniqueValues: string[] = [];
    items.forEach((item: { [x: string]: any }) => {
      values.push(item["" + columnName + ""]);
    });
    values.map((statusValue) => {
      if (uniqueValues.indexOf(statusValue) === -1) {
        uniqueValues.push(statusValue);
      }
    });
    return uniqueValues;
  };

  private _getChartDataCommittee = (
    uniqueValues: string[],
    allItems: string[] | any,
    columnName: string
  ): any => {
    const llbArr: string[] = [];
    const dataArr: number[] = [];
    uniqueValues.forEach((uniqueValue: string) => {
      const arr = allItems.filter(
        (item: { [x: string]: string | any }) =>
          item[columnName] && item[columnName] === uniqueValue //data type channged
      );
      llbArr.push(uniqueValue);
      dataArr.push(arr.length);
    });
    const chartViewData: any = {
      labels: llbArr,
      datasets: [
        {
          label: "Dataset",
          data: dataArr,
        },
      ],
    };
    return chartViewData;
  };

  private _handleSelectNoteType = (
    event: React.FormEvent<HTMLDivElement>,
    option?: any,
    index?: number
  ) => {
    this.getAllrequestesData(option.key, this.state.committeeType);
  };

  public render(): React.ReactElement<IXenWpUcoBankProps> {
    const { hasTeamsContext } = this.props;
    const _items = [
      {
        key: "Excel",
        name: "Export to Excel",
        iconProps: {
          iconName: "ExcelLogo",
        },
        disabled: this.state.allItems.length === 0,
        onClick: () => {
          this._getExcel()
            .then((res) => res)
            .catch((err) => err);
        },
      },
    ];
    const drpdwnOption =
      this.state.isSecarory === true
        ? [
            { key: "MyNotes", text: "My Dashboard" },
            { key: "mypendingnotes", text: "My Pending Notes" },
            { key: "MyReferredNotes", text: "My Referred Notes" },
            { key: "MyReturnedNotes", text: "My Returned Notes" },
            { key: "MyApprovedNotes", text: "My Approved Notes" },
            { key: "EDMDNotes", text: "ED/MD Notes" },
          ]
        : [
            { key: "MyNotes", text: "My Dashboard" },
            { key: "mypendingnotes", text: "My Pending Notes" },
            { key: "MyReferredNotes", text: "My Referred Notes" },
            { key: "MyReturnedNotes", text: "My Returned Notes" },
            { key: "MyApprovedNotes", text: "My Approved Notes" },
          ];

    return (
      <section
        className={`${styles.xenWpUcoBank} ${
          hasTeamsContext ? styles.teams : ""
        }`}
      >
        <div className="non_mobile_View">
          <div>{this._renderButtons()}</div>
          {this.props.noteType === "eCommittee" &&
            this.state.activeBtn === "MyNotes" && (
              <div style={{ display: "flex", justifyContent: "end" }}>
                <Dropdown
                  label="Select Committee Type"
                  options={[
                    {
                      key: "CommitteeNote",
                      text: "Committee Notes",
                    },
                    {
                      key: "BoardNote",
                      text: "Board Notes",
                    },
                    {
                      key: "CommitteeMeetings",
                      text: "Committee Meetings",
                    },
                  ]}
                  onChange={this._handleCommiteeType}
                  selectedKey={this.state.committeeType}
                />
              </div>
            )}
        </div>

        <div className="mobile_homepage_btn">
          <Dropdown
            options={drpdwnOption}
            onChange={this._handleSelectNoteType}
            selectedKey={this.state.activeBtn}
          />
          {this.props.noteType === "eCommittee" &&
            this.state.activeBtn === "MyNotes" && (
              <Dropdown
                label="Select Committee Type"
                options={[
                  {
                    key: "CommitteeNote",
                    text: "Committee Note",
                  },
                  {
                    key: "BoardNote",
                    text: "Board Note",
                  },
                  {
                    key: "CommitteeMeetings",
                    text: "Committee Meetings",
                  },
                ]}
                onChange={this._handleCommiteeType}
                selectedKey={this.state.committeeType}
              />
            )}
          <br />
          <br />
        </div>

        {this.state.activeBtn === "MyNotes" ? (
          <>
            <br />
            <br />
            {this.state.committeeType === "CommitteeMeetings" ? (
              <>{this.committeeChart()}</>
            
            ) : (
              <div className="row">
                <div className="col-md-6 col-sm-12 _homePageCharts ">
                  <ChartControl
                    type={ChartType["Pie"]}
                    datapromise={this._getItemsCount(this._listName, "Status")}
                    loadingtemplate={() => (
                      <Spinner size={SpinnerSize.large} label="Loading..." />
                    )}
                    options={this.chartOptions}
                  />
                </div>
                <div className="col-md-6 col-sm-12">
                  <ChartControl
                    type={ChartType["Doughnut"]}
                    datapromise={this._getItemsCount(
                      this._listName,
                      "NatureOfNote"
                    )}
                    loadingtemplate={() => (
                      <Spinner size={SpinnerSize.large} label="Loading..." />
                    )}
                    options={this.chartOptionsNatureOfNote}
                  />
                </div>
              </div>
            )}

          </>
        ) : (
          <>
            <div className={styles.commandbarContainer}>
              <div style={{ width: "70%" }}>
                <CommandBar items={_items} />
              </div>
              <div style={{ width: "30%" }}>
                <SearchBox
                  placeholder="Search"
                  title="Search"
                  onSearch={(newValue) => console.log(newValue)}
                  onChange={this._onChangeFilterText}
                />
              </div>
            </div>

            <div
              id="generateTable"
              className={styles._listviewContainerDataTable}
              data-is-scrollable="true"
            >
              <DetailsList
                data-is-scrollable="true"
                items={this.state.listItems}
                columns={this.state.columns}
                selectionMode={SelectionMode.none}
                layoutMode={DetailsListLayoutMode.justified}
                selection={this._selection}
                isHeaderVisible={true}
              />
            </div>

            <div>
              <span className={styles._totalDataCount}>
                {this.state.allItems.length > 0
                  ? `1 - ${this.state.allItems.length}`
                  : null}
              </span>
              <Pagination
                currentPage={this.state.page}
                totalItems={this.state.allItems.length}
                onChange={(page) => this.handlePaginationChange(page)}
              />
            </div>
          </>
        )}
       
      </section>
    );
  }
}
