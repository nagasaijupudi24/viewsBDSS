import * as React from "react";
import styles from "../styles/superAdmin.module.scss";
import type { IXenWpUcoBankProps } from "../IXenWpUcoBankProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "@pnp/sp/site-users/web";
import {
  CommandBar,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Link,
  SearchBox,
  SelectionMode,
  Selection,
} from "@fluentui/react";
import { Pagination } from "../../../../Common/PageNation";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import * as XLSX from "xlsx";
import * as FileSaver from "file-saver";

export interface IListViewsState {
  listItems: any[];
  columns: IColumn[];
  page: number;
  rowsPerPage?: any;
  pageOfItems: any[];
  allItems: any[];
  selectionDetails: any; // isSuperAdmin:boolean,
  selectedcount: number; // isDepartmentAdmin:boolean;

}
export default class CommitteeMeetingListViews extends React.Component<
  IXenWpUcoBankProps,
  IListViewsState,
  {}
> {
  private _listName: any;
  private _selection: Selection;
  private _hideCommandOption: boolean = true;
  private _columns: IColumn[] = [
    {
      key: "column1",
      name: "Note#",
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
      columns: this._bindColumns(),
      page: 1,
      rowsPerPage: 10,
      pageOfItems: [],
      allItems: [],
      selectionDetails: {},
      selectedcount: 0,
    };
    const listObj = this.props.listName;
    this._listName = listObj?.title;

    if(this.props.noteType ==='eCommitteeMeeting' && this.props.cmViewType ==='CommitteeUnmappedRecords'){
      this.getAllCommitteeerequestesData();
    }else{
      
    this.getAllrequestesData();
    }
  }

  private redirectToViewPage = async (item: any) => {
      let user = await this.props.sp?.web.currentUser();
      if (
        (item.StatusNumber === "1000" ||
        item.StatusNumber === "2000" ||
        item.StatusNumber === "3000" ||
        item.StatusNumber === "4000" ||
        item.StatusNumber === "7000"  ) && item.AuthorId === user.Id
    ) {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;
    } else {
      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
    }
  };
 
  private _bindColumns = () => { // Column rendering based on selected  view typ
    const { cmViewType, noteType } = this.props;
 
  if (noteType === "eCommitteeMeeting"){
        if(cmViewType ==="CommitteeUnmappedRecords"){
            this._columns = [
                {
                  key: "column1",
                  name: "Note#",
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
                
                  {  key: "column3",
                    name: "Committee",
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
                },]

        }
        else{
            this._columns = [
                {
                  key: "column1",
                  name: "Title",
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
                  name: "Meeting Subject",
                  fieldName: "MeetingSubject",
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
                  name: "Meeting Mode",
                  fieldName: "MeetingMode",
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
                    name: "Meeting Date",
                    fieldName: "MeetingDate",
        
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
                
                  {  key: "column5",
                    name: "Committee",
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
                  key: "column8",
                  name: "Approved By",
                  fieldName: "PreviousApprover",
      
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
                },]

        }
    }
    return this._columns;
  };

 
  private _formatDate=(date:Date) =>{
    if (!(date instanceof Date)) {
        throw new Error('Invalid date');
    }
    let day = date.getDate().toString().padStart(2, '0');
    let month = (date.getMonth() + 1).toString().padStart(2, '0'); // Month is zero-indexed
    let year = date.getFullYear().toString();
    let hours = date.getHours().toString().padStart(2, '0');
    let minutes = date.getMinutes().toString().padStart(2, '0');
    let seconds = date.getSeconds().toString().padStart(2, '0');

    return `${day}${month}${year}${hours}${minutes}${seconds}`;// Concatenate them in the desired format
    
}

private _getExcel = async (): Promise<void> => {
  const todayDate=new Date();
  const formatDate= this._formatDate(todayDate)

    const fieldNames=this.state.columns.map(obj=>obj.fieldName)
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
  const Heading = [this.state.columns?.map(obj=>obj.name)];
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
      this.props.noteType +this.props.viewType +formatDate+ fileExtension
    );
  }
};


  private _getBylistWithFilterQuery = async () => {
/* 
* Committtee meeting status 
Created - 1000
Published -2000
Meeting Over -3000
MOM Published -4000
Pending Approval -5000
Pending Chairman Approval -6000
Approved -9000
Returned -8000 */
    let user = await this.props.sp?.web.currentUser();
    const userEmail = (user.Email).toLocaleLowerCase()||""
    let filterQury = "";
      switch (this.props.cmViewType) {
        case "CommitteeUnmappedRecords":
          filterQury = `AuthorId eq ${user.Id} and CommitteeType eq 'Committee' and isMapped eq 'false' and StatusNumber eq '9000' `;
          break;
        case "MyPendingCommitteeRecords":
          filterQury = `((StatusNumber  eq '1000' or StatusNumber eq  '2000' or StatusNumber eq '3000' or StatusNumber eq '4000' or StatusNumber eq '8000') and AuthorId eq ${user?.Id}) or (StatusNumber eq '5000' and CommitteeMeetingMembers/EMail eq '${userEmail}') or (StatusNumber eq '6000' and Chairman/EMail eq '${userEmail}')`;
          break;
        case "AllInprogressCommitteeMeetingRecords":
          filterQury = `StatusNumber ne '9000' AuthorId eq ${user.Id}`;
          break;
        case "MyApprovedCommitteeRecords":
          filterQury = `StatusNumber eq '9000' and CommitteeMeetingMembers/EMail eq '${userEmail}' `;
          break;
        case "AllApprovedCommitteeMeetings":
          filterQury = `StatusNumber eq '9000' and AuthorId eq ${user.Id} `;
          break;
        default:
          break;
      }
 
    return filterQury;
  };

  private getAllCommitteeerequestesData = async () => {
    const filterQury = await this._getBylistWithFilterQuery();
    const items: any = await this.props.sp?.web.lists
      .getByTitle("EcommiteeRequests")
      .items.filter(filterQury)
      .select(
        `*`)
      .orderBy("Created", false)();
   
    this.paginateFn(items, 1);
      this.setState({
        allItems: items,
      });
  };
  private getAllrequestesData = async () => {
    const filterQury = await this._getBylistWithFilterQuery();
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
        FinalApprover/EMail
        ,CommitteeMeetingMembers/EMail,
        Chairman/EMail`
      )
      .expand(`Author,Editor,PreviousApprover,CurrentApprover,FinalApprover,Chairman,CommitteeMeetingMembers`)
      .orderBy("Created", false)();
    items.map((obj: any) => {
      allItems.push({
        Id: obj.Id,
        Department: obj.Department,
        MeetingNumber: obj.MeetingNumber,
        MeetingDate:new Date(obj.MeetingDate).toDateString(),
        MeetingLink:obj.MeetingLink,
        MeetingMode:obj.MeetingMode,

        MeetingSubject: obj.MeetingSubject,
        Status: obj.MeetingStatus,
        StatusNumber: obj.StatusNumber,
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
          obj.Editor === null && obj.EditorId === null ? "" : obj.Editor?.Title,
        Admin:
          obj.Admin === null && obj.AdminId === null ? "" : obj.Admin?.Title,
        Author:
          obj.Author === null && obj.AuthorId === null ? "" : obj.Author?.Title,
        Created: obj.Created,
        Modified: obj.Modified,
        committeeName:
          obj.CommitteeName ,
        AuthorId :obj.AuthorId 
      });
    });


    this.paginateFn(allItems, 1);
      this.setState({
        allItems: allItems,
      });
  };

  private paginateFn = (filterItem: any[], pageNo?: any) => {
    let { rowsPerPage } = this.state;
    const items =
      rowsPerPage > 0
        ? filterItem.slice(
            (pageNo - 1) * rowsPerPage,
            (pageNo - 1) * rowsPerPage + rowsPerPage
          )
        : filterItem;
    this.setState({ listItems: items });
  };

  private handlePaginationChange = (page: any) => {
    this.setState({ page: page });
    this.paginateFn(this.state.allItems, page);
  };



  private _onChangeFilterText = (
    event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void => {
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
    this.paginateFn(this.state.allItems, this.state.page);
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
    const { hasTeamsContext} = this.props;
    const _items =
      this.props.viewType === "Draft Requests"
        ? [
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
              key: "EditItem",
              name: "Edit Request",
              iconProps: {
                iconName: "Edit",
              },
              disabled: this._hideCommandOption,
              // split: true,
              onClick: () => {
                const item = this.state.selectionDetails;
                if (this.state.selectedcount === 0) {
                  this._hideCommandOption = true;
                } else {
                  window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;

                }
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
                if (this.state.selectedcount === 0) {
                  return;
                } else {
                  window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                }
              },
            },
          ]
        : this.props.viewType === "PendingWith"
        ? [ {
          key: "Excell",
          name: "Export CSV",
          iconProps: {
            iconName: "ExcelLogo",
          },
          disabled: this.state.allItems.length ===0,
          onClick:() =>{
            this._getExcel().then(res=>res).catch(err=>err)
          },
        },]
        : [
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
                if (this.state.selectedcount === 0) {
                  return;
                } else {

                  window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                }
              },
            },
            {
              key: "Excell",
              name: "Export CSV",
              iconProps: {
                iconName: "ExcelLogo",
              },
              disabled: this.state.allItems.length ===0,
              onClick:() =>{
                this._getExcel().then(res=>res).catch(err=>err)
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
          <div style={{width:"70%"}}>
            <CommandBar items={_items} />
          </div>
          <div style={{width:"30%"}}>
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
            selectionMode={SelectionMode.single}
            layoutMode={DetailsListLayoutMode.justified}
            selection={this._selection}
            isHeaderVisible={true}
          />
        </div>
        <div>
        <span className={styles._totalDataCount}>{this.state.allItems.length>0?`1 - ${this.state.allItems.length}`:null}</span>
          <Pagination
          
            currentPage={this.state.page}
            totalItems={this.state.allItems.length}
            onChange={(page) => this.handlePaginationChange(page)}
          />
        </div>
      </section>
    );
  }
}
