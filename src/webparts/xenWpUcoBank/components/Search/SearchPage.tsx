import * as React from "react";
import styles from "../styles/Search.module.scss";
import { IXenWpUcoBankProps } from "../IXenWpUcoBankProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "@pnp/sp/site-users/web";
import * as XLSX from "xlsx";
import * as FileSaver from "file-saver";
import {
  IPeoplePickerContext,
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import {
  CommandBar,
  DatePicker,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  Dropdown,
  IColumn,
  IconButton,
  Link,
  PrimaryButton,
  SearchBox,
  SelectionMode,
  TextField,
} from "@fluentui/react";
import { Pagination } from "../../../../Common/PageNation";
export interface ISearchState {
  showSuccessPopup: boolean;
  Note: any;
  Requester: any;
  Department: string;
  SearchText: string;
  FromDate: any;
  Todate: any;
  Status: any;
  Subject: string;
  Financial: any;
  FY: any;
  Approvers: any;
  NoteType: any;
  listItems: any;
  allItems: any[];
  searchText: string;
  columns: IColumn[];
  page: number;
  rowsPerPage?: any;
  pageOfItems: any[];
  financialYear: any[];
  noteTypeOption: any[];
  financialOptions: any[];
  departmentOptions: any[];
}

const dragOptions = {
  moveMenuItemText: "Move",
  closeMenuItemText: "Close",
};
const modalPropsStyles = { main: { maxWidth: 600 } };
const dialogContentProps={
  type: DialogType.normal,
  title: (
    <div style={{ display: 'flex', alignItems: 'center', gap: '8px' }}>
      <IconButton
        iconProps={{ iconName: 'Info' }}
      />
      <span style={{ fontSize: '16px', fontWeight: 'bold',marginBottom:'5px' }}>
        Alert
      </span>
    </div>
  ),
}

export default class SearchPage extends React.Component<
  IXenWpUcoBankProps,
  ISearchState,
  {}
> {
  protected ppl: any;
  protected pplRequester: any;
  private _listName: any;
  private _peoplePickerContext: IPeoplePickerContext;
  private _columns: IColumn[] = [];
    private _StatusOption=[
      {
          key:"100",text:"Draft"
      },
      {
        key:"200",text:"Call Back"
    },
    {
      key:"300",text:"Cancel"

  },
  {
    key:"Submitted",text:"Submit"

},
{
  key:"2000",text:"Pending With Reviewer"

},
{
  key:"3000",text:"Pending With Approver"

},
{
  key:"4000",text:"Refer"

},
{
  key:"5000",text:"Return"

},
{
  key:"4900",text:"Refer Back"

},
{
  key:"8000",text:"Reject"

},
{
  key:"9000",text:"Approved"

},
  ]
  constructor(props: IXenWpUcoBankProps) {
    super(props);

    this.state = {
      showSuccessPopup: true,
      Note: "",
      Requester: 0,
      Department: "",
      SearchText: "",
      FromDate: null,
      Todate: null,
      Status: "",
      Subject: "",
      Financial: "",
      FY: "",
      Approvers:0,
      NoteType: "",
      listItems: [],
      allItems: [],
      searchText: "",
      columns: this._bindColumns(),
      page: 1,
      rowsPerPage: 10,
      pageOfItems: [],
      financialYear: this._getRecentFinancialYears(),
      noteTypeOption: [],
      financialOptions: [],
      departmentOptions: [],
    };
    const listObj = this.props.listName;
    this._listName = listObj?.title;
    this._peoplePickerContext = {
      absoluteUrl: this.props.context.pageContext.web.absoluteUrl,
      msGraphClientFactory: this.props.context.msGraphClientFactory,
      spHttpClient: this.props.context.spHttpClient,
    };
    this._getRecentFinancialYears();
    this.getcolumnInfo();
    this._getDepartments();
  }

  private _getDepartments = async () => {
    const departmentOption: { key: any; text: any }[] = [];
    const items = await this.props.sp.web.lists
      .getByTitle("Departments")
      .items.select("Department")();
    if (items && items.length > 0) {
      items?.map((obj) => {
        departmentOption.push({
          key: obj.Department,
          text: obj.Department,
        });
      });
    }
    this.setState({
      departmentOptions: departmentOption,
    });
  };
  public getcolumnInfo = async (): Promise<any> => {
    let _FToptions: { key: string; text: string }[] = [];
    let _NToptions: { key: string; text: string }[] = [];
    const fieldsInfo = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .fields.filter(" Hidden eq false and ReadOnlyField eq false")();
    fieldsInfo.filter((_x) => {
      if (
        _x.TypeDisplayName === "Choice" &&
        _x.InternalName === "FinancialType"
      ) {
        if (_x.Choices) {
          _x.Choices?.map((obj) => {
            _FToptions.push({
              key: obj,
              text: obj,
            });
          });
        }
      }
      if (_x.TypeDisplayName === "Choice" && _x.InternalName === "NoteType") {
        if (_x.Choices) {
          _x.Choices?.map((obj) => {
            _NToptions.push({
              key: obj,
              text: obj,
            });
          });
        }

        console.log(_NToptions,"_NToptions")
        this.setState({
          noteTypeOption: _NToptions,
          financialOptions:_FToptions
        });
      }
    });
  };

  /* To get the financial year viz in 2024-2025 format */
  private _getRecentFinancialYears = () => {
    const currentYear = new Date().getFullYear();
    const years = [];
    for (let i = 0; i < 5; i++) {
      const startYear = currentYear - i;
      const endYear = startYear + 1;
      years.push({ key: startYear, text: `${startYear}-${endYear}` });
    }
    return years;
  };

  private _bindColumns = () => {
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
        key: "column6",
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
        key: "column7",
        name: "Last Approver",
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
        key: "column8",
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
        key: "column9",
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
    return this._columns;
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

  private _onChangeFilterText = (
    event?: React.ChangeEvent<HTMLInputElement>,
    newValue?: string
  ): void => {

    const filteredItems = newValue
    ? this.state.allItems.filter((item: any) =>
        Object.values(item).some(
          (value: any) =>
            (value || "")
              .toString()
              .toLowerCase()
              .indexOf(newValue.toLowerCase()) > -1
        )
      )
    : this.state.allItems;
    this.setState({
      listItems:filteredItems
    })
    this.paginateFn(filteredItems, this.state.page);
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



  private getAllrequestesData = async () => {

    // If all fields are empty, reset popup and exit
    if (this.isAllFieldsEmpty()) {
      this.setState({ showSuccessPopup: false });
      return;
    }
  
    // Build filter conditions
    const filterConditions = this.buildFilterConditions();
  
    // Generate the filter query
    const filterQuery = filterConditions.length > 0 ? filterConditions.join(' and ') : "";
  
    // Fetch data from the list
    const items = await this.fetchFilteredItems(filterQuery);
  
    // Process fetched data
    this._FilterFetchedData(items);
  };
  
  // Helper to check if all fields are empty
  private isAllFieldsEmpty = (): boolean => {
    const { Note, Requester, Department, SearchText, FromDate, Todate, Status, Subject, Financial, FY, Approvers, NoteType } = this.state;
    return (
      Note === "" &&
      Requester === 0 &&
      Department === "" &&
      SearchText === "" &&
      FromDate === null &&
      Todate === null &&
      Status === "" &&
      Subject === "" &&
      Financial === "" &&
      Approvers === 0 &&
      NoteType === "" &&
      FY === ""
    );
  };
  
  // Helper to build filter conditions
  private buildFilterConditions = (): string[] => {
    const { Note, Requester, Department, SearchText, FromDate, Todate, Status, Subject, Financial, FY, Approvers, NoteType } = this.state;
    const conditions: string[] = [];
  
    if (Requester !== 0) conditions.push(`AuthorId eq ${Requester}`);
    if (Status) conditions.push(`StatusNumber eq '${Status}'`);
    if (NoteType) conditions.push(`NoteType eq '${NoteType}'`);
    if (Financial) conditions.push(`Financial eq '${Financial}'`);
    if (Subject) conditions.push(`substringof('${Subject}', Subject)`);
    if (Note) conditions.push(`substringof('${Note}', Title)`);
    if (SearchText) conditions.push(`substringof('${SearchText}', SearchKeyword)`);
    if (FromDate) conditions.push(`Created ge '${this.formatDate(FromDate, true)}'`);
    if (Todate) conditions.push(`Created le '${this.formatDate(Todate, false)}'`);
    if (FY) conditions.push(...this.buildFYCondition(FY));
    if (Department) conditions.push(`Department eq '${Department}'`);
    if (Approvers !== 0) conditions.push(`(ApproversId eq ${Approvers} or ReviewersId eq ${Approvers})`);
  
    return conditions;
  };
  
  // Helper to format date to ISO string
  private formatDate = (date: Date, startOfDay: boolean): string => {
    const newDate = new Date(date);
    if (startOfDay){
      newDate.setHours(0, 0, 0, 1)
    }else{
      newDate.setHours(23, 59, 59, 999)
    }
    return newDate.toISOString();
  };
  
  // Helper to build FY condition
  private buildFYCondition = (FY: string): string[] => {
    const conditions: string[] = [];
    const startYear = new Date();
    startYear.setFullYear(Number(FY), 3, 1); // Start of FY (April 1st)
    const endYear = new Date();
    endYear.setFullYear(Number(FY) + 1, 2, 31); // End of FY (March 31st)
  
    const isCurrentYear = new Date().getFullYear() === Number(FY);
    const finalEndDate = isCurrentYear ? new Date() : endYear;
  
    conditions.push(`Created ge '${this.formatDate(startYear, true)}'`);
    conditions.push(`Created le '${this.formatDate(finalEndDate, false)}'`);
    
    return conditions;
  };
  
  // Helper to fetch filtered items
  private fetchFilteredItems = async (filterQuery: string) => {
    return await this.props.sp?.web.lists
      .getByTitle(this._listName)
      .items.filter(filterQuery)
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
      .orderBy("Created", false)
      .top(500)();
  };

    private _getFieldValue = (field: any, fieldId: any, defaultValue: string = ""): string => {
    return field === null && fieldId === null ? defaultValue : field?.Title || defaultValue;
  };

  private _FilterFetchedData=(filteredData:any)=>{
    const allItems:any[]=[]
    filteredData.map((obj: any) => {

      allItems.push({
        Id: obj.Id,
        Department: obj.Department,
        NoteNumber: obj.NoteNumber,
        Subject: obj.Subject,
        Status: obj.Status,
        StatusNumber: obj.StatusNumber,

        CurrentApproverObj:
          obj.CurrentApprover === null && obj.CurrentApproverId === null
            ? null
            : obj.CurrentApprover,
        CurrentApprover:this._getFieldValue(obj.CurrentApprover, obj.CurrentApproverId),
        PreviousApprover:this._getFieldValue(obj.PreviousApprover, obj.PreviousApproverId),
         
        FinalApprover:this._getFieldValue(obj.FinalApprover, obj.FinalApproverId),
       
        Title: obj.Title,
        AuthorId: obj.AuthorId === null?0:obj.AuthorId,
        Editor:this._getFieldValue(obj.Editor, obj.EditorId),
        Admin:this._getFieldValue(obj.Admin, obj.AdminId),
        
        Author:this._getFieldValue(obj.Author, obj.AuthorId),
       
        Created: obj.Created,
        Modified: obj.Modified,
        committeeName: obj.CommitteeName || obj.BoardName || "",
CommitteeType: obj.CommitteeType ?? "",
       
      });
    
  });

  this.paginateFn(allItems, 1);
  this.setState({
    allItems: allItems,
  });


  }

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
  private _formatDate = (date: Date) => {
    if (!(date instanceof Date)) {
      throw new Error("Invalid date");
    }
   /*  // Get individual date components */
    let day = date.getDate().toString().padStart(2, "0");
    let month = (date.getMonth() + 1).toString().padStart(2, "0"); // Month is zero-indexed
    let year = date.getFullYear().toString();
    let hours = date.getHours().toString().padStart(2, "0");
    let minutes = date.getMinutes().toString().padStart(2, "0");
    let seconds = date.getSeconds().toString().padStart(2, "0");

  /*   // Concatenate them in the desired format */
    return `${day}${month}${year}${hours}${minutes}${seconds}`;
  };

  private _getExcel = async (): Promise<void> => {

    const todayDate = new Date();
    const formatDate = this._formatDate(todayDate);

    const fieldNames = this.state.columns.map((obj) => obj.fieldName);
    const excelData = this.state.allItems?.map((item: any) => {
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
        this.props.noteType + "search"+ formatDate + fileExtension
      );
    }
  };
  private _getPeoplePickerItems = (items: any[]) => {
    if (items) {
      this.setState({
        Requester: items[0].id,
      });
    } else {
      this.setState({
        Requester: 0,
      });
    }
  };

  private _ApprovergetPeoplePickerItems = (items: any[]) => {

    if (items) {
      
      this.setState({
        Approvers: items[0].id,
      });
    } else {
      this.setState({
        Approvers: 0,
      });
    }
  };
  private _onChangeNoteNumber = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: any
  ) => {
    this.setState({ Note: newValue });
  };
  private _onChangeSearchText = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: any
  ) => {
    this.setState({ SearchText: newValue });
  };
  private _onChangeSubject = (
    event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: any
  ) => {
    this.setState({ Subject: newValue });
  };
  private _onChangeNoteType = (
    event: React.FormEvent<HTMLDivElement>,
    option?: any,
    index?: number
  ) => {
    this.setState({
      NoteType: option.key,
    });
  };

  private _onChangefy = (
    event: React.FormEvent<HTMLDivElement>,
    option?: any,
    index?: number
  ) => {
    this.setState({
      FY: option.key,
    });
  };
  private _onChangefinancial = (
    event: React.FormEvent<HTMLDivElement>,
    option?: any,
    index?: number
  ) => {
    this.setState({
      Financial: option.key,
    });
  };
  private _onChangeStatus  = (
    event: React.FormEvent<HTMLDivElement>,
    option?: any,
    index?: number
  ) => {
    this.setState({
      Status: option.key,
    });
  };

  private _onChangeDepartment  = (
    event: React.FormEvent<HTMLDivElement>,
    option?: any,
    index?: number
  ) => {
    this.setState({
      Department: option.key,
    });
  };
  //
  private _onSelectedFromDate = (date: Date | null | undefined) => {
    this.setState({
      FromDate: date,
    });
  };
  private _onSelectedToDate = (date: Date | null | undefined) => {
    this.setState({
      Todate: date,
    });
  };

  private _onClear=()=>{
    this.ppl.state.selectedPersons=[];
    this.ppl.onChange([]);
    this.pplRequester.state.selectedPersons=[];
    this.pplRequester.onChange([]);
    this.setState({
      Requester: 0,
      Department: "",
      SearchText: "",
      FromDate: null,
      Todate: null,
      Status: "",
      Subject: "",
      Financial: "",
      FY: "",
      Approvers: 0,
      NoteType: "",
      listItems: [],
      allItems: [],
      searchText: "",
      Note:""
    });
   

  }

  public render(): React.ReactElement<IXenWpUcoBankProps> {
 

    const modalProps: any = {
      isBlocking: true,
      styles: modalPropsStyles,
      dragOptions: dragOptions,
    };
    const _items = [
      {
        key: "Excell",
        name: "Export CSV",
        iconProps: {
          iconName: "ExcelLogo",
        },
        disabled: this.state.allItems.length ===0,
        onClick: () => {
          this._getExcel()
            .then((res) => res)
            .catch((err) => err);
        },
      },
    ];

    return (
      <section className={styles.SearchWeb}>
        <div className={styles.customHeader}>Search Parameters </div>
        <fieldset className={styles.customFieldSet}>
          <div className="row">
            <div className="col-lg-4 col-md-6 col-sm-12">
              <TextField
                label="Note#:"
                value={this.state.Note}
                onChange={this._onChangeNoteNumber}
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <PeoplePicker
                context={this._peoplePickerContext}
                titleText="Requester"
                personSelectionLimit={1}
                placeholder="Select Requester"
                groupName={""} // Leave this blank in case you want to filter from all users
                showtooltip={true}
                disabled={false}
                searchTextLimit={5}
                onChange={this._getPeoplePickerItems}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                ensureUser={true}
                resolveDelay={1000}
                ref={(c) => (this.pplRequester = c)}
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <Dropdown
                label="Department:"
                options={this.state.departmentOptions}
                selectedKey={this.state.Department}
                onChange={this._onChangeDepartment}

              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <TextField
                label="Search Text:"
                value={this.state.SearchText}
                onChange={this._onChangeSearchText}
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <DatePicker
                label="From Date:"
                onSelectDate={this._onSelectedFromDate}
                maxDate={new Date()}
                value={this.state.FromDate}
                placeholder="Select From date"
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <DatePicker
                label="To Date:"
                onSelectDate={this._onSelectedToDate}
                maxDate={new Date()}
                minDate={new Date(this.state.FromDate)}
                value={this.state.Todate}
                placeholder="Select To date"
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <Dropdown
                label="Status:"
                options={this._StatusOption}
                selectedKey={this.state.Status}
                onChange={this._onChangeStatus}
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <TextField
                label="Subject:"
                value={this.state.Subject}
                onChange={this._onChangeSubject}
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <Dropdown
                label="Financial:"
                options={this.state.financialOptions}
                selectedKey={this.state.Financial}
                onChange={this._onChangefinancial}
              />
            </div>
          
            <div className="col-lg-4 col-md-6 col-sm-12">
              <Dropdown
                label="FY:"
                options={this.state.financialYear}
                selectedKey={this.state.FY}
                onChange={this._onChangefy}
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <PeoplePicker
                context={this._peoplePickerContext}
                titleText="Approver:"
                personSelectionLimit={1}
                groupName={""} 
                showtooltip={true}
                placeholder="Select Approver"
                disabled={false}
                searchTextLimit={5}
                onChange={this._ApprovergetPeoplePickerItems}
                showHiddenInUI={false}
                ensureUser={true}
                principalTypes={[PrincipalType.User]}
                resolveDelay={1000}
                ref={(c) => (this.ppl = c)}
              />
            </div>
            <div className="col-lg-4 col-md-6 col-sm-12">
              <Dropdown
                label="Note Type:"
                options={this.state.noteTypeOption}
                selectedKey={this.state.NoteType}
                onChange={this._onChangeNoteType}
              />
            </div>
          </div>
        </fieldset>
        <div className={styles.btn_Container}>
          <span>
            <PrimaryButton
              iconProps={{ iconName: "Search" }}
              text="Search"
              onClick={this.getAllrequestesData}
            />
          </span>
          <span>
            <DefaultButton
              iconProps={{ iconName: "Reset" }}
              text="Clear"
              onClick={this._onClear}
            />
          </span>
        </div>
        <hr />
        <section>
          <div className={styles.customHeader}>Search Results </div>

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
        </section>

        <Dialog
          hidden={this.state.showSuccessPopup}
          onDismiss={() =>
            this.setState({ showSuccessPopup: !this.state.showSuccessPopup })
          }
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
          maxWidth={600}
        >
          <p className="dialogcontent_">Please fill at least any one of the fields to search.</p>
          <DialogFooter>
            <PrimaryButton onClick={()=>this.setState({showSuccessPopup:!this.state.showSuccessPopup})}>
            Ok
            </PrimaryButton>
          </DialogFooter>
        </Dialog>
      </section>
    );
  }
}
