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
import "../CustomStyles/Custom.css";

export interface IListViewsState {
  listItems: any[];
  columns: IColumn[];
  page: number;
  rowsPerPage?: any;
  pageOfItems: any[];
  allItems: any[];
  selectionDetails: any;
  selectedcount: number;
}
export default class ATRViews extends React.Component<
  IXenWpUcoBankProps,
  IListViewsState,
  {}
> {
  private _listName: any;
  private _selection: Selection;
  private _hideCommandOption: boolean = true;
  private _columns: IColumn[] = [];
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

    this.getAllrequestesData();
    this._columns = [
      {
        key: "column1",
        name: "NOte#",
        fieldName: "ATRNoteID",
        minWidth: 50,
        maxWidth: 50,
        isRowHeader: false,
        isResizable: true,
        data: "string",
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onRender: (item, index, column) => {
          return (
            <Link onClick={() => this.redirectToViewPage(item.Id)}>
              {item.ATRNoteID}
            </Link>
          );
        },
      },
      {
        key: "column2",
        name: "Assignee",
        fieldName: "Assignee",
        minWidth: 150,
        maxWidth: 350,
        isRowHeader: false,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
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
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column4",
        name: "Assigned By",
        fieldName: "AssignedBy",
        minWidth: 150,
        maxWidth: 350,
        isRowHeader: false,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      // DepartmentAlias
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
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column6",
        name: "Subject",
        fieldName: "Subject",

        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column7",
        name: "Remarks",
        fieldName: "Remarks",

        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column7",
        name: "Created Date",
        fieldName: "Created",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        data: "string",
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
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
    this._columns = [
      {
        key: "column1",
        name: "Note#",
        fieldName: "ATRNoteID",
        minWidth: 100,
        maxWidth: 150,
        isRowHeader: false,
        isResizable: true,
        data: "string",
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
        isMultiline: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
        onRender: (item, index, column) => {
          return (
            <Link onClick={() => this.redirectToViewPage(item.Id)}>
              {item.ATRNoteID}
            </Link>
          );
        },
      },
      {
        key: "column2",
        name: "Assignee",
        fieldName: "Assignee",
        minWidth: 150,
        maxWidth: 350,
        isRowHeader: false,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
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
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column4",
        name: "Assigned By",
        fieldName: "AssignedBy",
        minWidth: 150,
        maxWidth: 350,
        isRowHeader: false,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
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
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column6",
        name: "Subject",
        fieldName: "Subject",

        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column7",
        name: "Remarks",
        fieldName: "Remarks",

        minWidth: 150,
        maxWidth: 300,
        isResizable: true,
        data: "string",
        isMultiline: true,
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
        sortAscendingAriaLabel: "Sorted A to Z",
        sortDescendingAriaLabel: "Sorted Z to A",
      },
      {
        key: "column7",
        name: "Created Date",
        fieldName: "Created",
        minWidth: 100,
        maxWidth: 150,
        isResizable: true,
        data: "string",
        onColumnClick: this._onColumnClick,
        isSorted: false,
        isSortedDescending: true,
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

  private redirectToViewPage = (item: any) => {
    window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?atrid=${item}`;
  };

  private _formatDate = (date: Date) => {
    if (!(date instanceof Date)) {
      throw new Error("Invalid date");
    }

    let day = date.getDate().toString().padStart(2, "0"); // Get individual date components
    let month = (date.getMonth() + 1).toString().padStart(2, "0"); // Month is zero-indexed
    let year = date.getFullYear().toString();
    let hours = date.getHours().toString().padStart(2, "0");
    let minutes = date.getMinutes().toString().padStart(2, "0");
    let seconds = date.getSeconds().toString().padStart(2, "0");

    return `${day}${month}${year}${hours}${minutes}${seconds}`; // Concatenate them in the desired format
  };

  private _getExcel = async (): Promise<void> => {
    const todayDate = new Date();
    const formatDate = this._formatDate(todayDate);

    const fieldNames = this.state.columns.map((obj) => obj.fieldName);

    const excelData = this.state.allItems?.map((item) => {
      const res: string[] = [];
      fieldNames.forEach((element: any) => {
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
        this.props.noteType + this.props.viewType + formatDate + fileExtension
      );
    }
  };

  private _getBylistWithFilterQuery = async () => {
    /* Pending -- 1000
Competed -2000
Returned- 3000
Accepted - 4000 */
    let user = await this.props.sp?.web.currentUser();
    let filterQury:any;
    switch (this.props.viewType) {
      case "allATR":
        filterQury = "";
        break;
      case "pendingATR":
        filterQury = `(CurrentApproverId eq ${user?.Id}) and (StatusNumber eq '2000' or StatusNumber eq '1000' or StatusNumber eq '4000')`;
        break;
      case "pendingATRSect":
        filterQury = `StatusNumber eq '3000' and  CurrentApproverId eq ${user?.Id} `;
        break;
      case "completedATR":
        filterQury = `StatusNumber eq '5000'`;
        break;

      default:
        break;
    }
    return filterQury;
  };

  private _Secretory = async (id: any, currentApprover: string) => {
    let isSecarory = false;
    const noteListName = this.props.noteListName;
    const noteListTitle = noteListName?.title;
    const item = await this.props.sp.web.lists
      .getByTitle(noteListTitle)
      .items.filter(`Id eq ${id}`)();
    if (item && item !== undefined) {
      const NoteSecretaryDTO = JSON.parse(item[0].NoteSecretaryDTO || "[]");
      if (
        NoteSecretaryDTO?.some(
          (obj: any) =>
            obj.secretaryEmail.toLowerCase() ===
              this.props.context.pageContext.user.email.toLowerCase() &&
            obj.approverEmail === currentApprover
        )
      )
        isSecarory = true;
    }
    return isSecarory;
  };

  private getAllrequestesData = async () => {
    const filterQury = await this._getBylistWithFilterQuery();
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
        AssignedBy/Title,
       AssignedBy/EMail,
       Assignee/Title,
       Assignee/EMail `
      )
      .expand(`Author,Editor,Assignee,CurrentApprover,AssignedBy`)
      .orderBy("Created", false)();
    let filterData: any;

    switch (this.props.noteType) {
      case "enote":
        filterData = items?.filter(
          (obj: { NoteType: string }) => obj.NoteType === "eNote"
        );
        break;
      case "eCommittee":
        filterData = items?.filter(
          (obj: { NoteType: string }) => obj.NoteType === "committeenote"
        );
        break;
      default:
        filterData = items;
    }
    // const filterData:any= this.pr/*  */ops.noteType ==="enote"?items?.filter((obj: { NoteType: string; })=>obj.NoteType ==="eNote"):this.props.noteType ==="eCommittee"?items?.filter((obj: { NoteType: string; })=>obj.NoteType ==="committeenote"):items;
    this._filterItems(filterData);
  };


  private _filterItems = async (response: any) => {
    const groups = await this.props.sp.web.currentUser.groups();
    const allItems: any[] = [];
  
    const processItem = async (obj: any) => {
      const item = {
        Id: obj.Id,
        Department: obj.Department,
        ATRNoteID: obj.ATRNoteID,
        Subject: obj.Subject,
        Status: obj.Status,
        StatusNumber: obj.StatusNumber,
        Remarks: obj.Remarks,
        CurrentApprover: this._getFieldValue(obj.CurrentApprover, obj.CurrentApproverId),
        Assignee: this._getFieldValue(obj.Assignee, obj.AssigneeId),
        AssignedBy: this._getFieldValue(obj.AssignedBy, obj.AssignedById),
        Title: obj.Title,
        Editor: this._getFieldValue(obj.Editor, obj.EditorId),
        Admin: this._getFieldValue(obj.Admin, obj.AdminId),
        Author: this._getFieldValue(obj.Author, obj.AuthorId),
        Created: obj.Created,
        Modified: obj.Modified,
        AuthorId: obj.AuthorId,
      };
  
      if (this.props.viewType === "allATR") {
        const isSecretaryExist = await this._Secretory(
          obj.NoteID,
          this._getFieldValue(obj.AssignedBy, obj.AssignedById, "")
        );
  
        const isSuperAdmin = groups?.some(
          (group: { Title: string }) => group.Title === this.props.superAdminGroupName
        );
  
        if (isSecretaryExist || isSuperAdmin) {
          allItems.push(item);
        }
      } else {
        allItems.push(item);
      }
    };
  
    await Promise.all(response.map((obj: any) => processItem(obj)));
    
    this.paginateFn(allItems, 1);
    this.setState({ allItems });
  };
  
  private _getFieldValue = (field: any, fieldId: any, defaultValue: string = ""): string => {
    return field === null && fieldId === null ? defaultValue : field?.Title || defaultValue;
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
      listItems: filteredItems,
    });

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

  public render(): React.ReactElement<IXenWpUcoBankProps> {
    const { hasTeamsContext } = this.props;
    const _items = [
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
        disabled: this.state.allItems.length === 0,
        onClick: () => {
          this._getExcel()
            .then((res) => res)
            .catch((err) => err);
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
            selectionMode={SelectionMode.single}
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
      </section>
    );
  }
}
