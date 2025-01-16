import * as React from "react";
import styles from "../styles/superAdmin.module.scss";
import type { IXenWpUcoBankProps } from "../IXenWpUcoBankProps";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "bootstrap/dist/css/bootstrap.min.css";
import "bootstrap/dist/js/bootstrap.bundle.min.js";
import "@pnp/sp/site-users/web";
import { Dialog, DialogType, DialogFooter } from "@fluentui/react/lib/Dialog";
import {
  CommandBar,
  DefaultButton,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  Icon,
  Link,
  PrimaryButton,
  SearchBox,
  SelectionMode,
  Selection,
  IconButton,
} from "@fluentui/react";
import { Pagination } from "../../../../Common/PageNation";
import { WebPartTitle } from "@pnp/spfx-controls-react";
import * as XLSX from "xlsx";
import * as FileSaver from "file-saver";

const dragOptions = {
  moveMenuItemText: "Move",
  closeMenuItemText: "Close",
};
const modalPropsStyles = { main: { maxWidth: 600 } };
const dialogContentProps = {
  type: DialogType.normal,
  title: (
    <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
      <IconButton iconProps={{ iconName: "Info" }} />
      <span
        style={{ fontSize: "16px", fontWeight: "bold", marginBottom: "5px" }}
      >
        Alert
      </span>
    </div>
  ),
};
export interface IListViewsState {
  listItems: any[];
  columns: IColumn[];
  page: number;
  rowsPerPage?: any;
  pageOfItems: any[];
  allItems: any[];
  searchText: string;
  selectedUserId: any;
  hideWarningDialog: boolean;
  hideSuccussDialog: boolean;
  hideDeleteDialog: boolean;
  succussMsg: string;
  warningMsg: string;
  selectedId: any;
  departmentName: string;
  departmentAlias: string;
  selectionDetails: any;
  selectedcount: number;
  currentUserDetails: any;
  isSuperAdmin: boolean;
  isDepartmentAdmin: boolean;
}
export default class ListViews extends React.Component<
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

      onRender: (item, index, column) => {
        return (
          <div>
            <Icon
              iconName="Delete"
              title="Delete"
              onClick={() => this._onclickDelete(item.Id)}
            />
          </div>
        );
      },
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
      searchText: "",
      selectedUserId: undefined,
      hideWarningDialog: true,
      hideSuccussDialog: true,
      succussMsg: "",
      warningMsg: "",
      selectedId: undefined,
      hideDeleteDialog: true,
      departmentName: "",
      departmentAlias: "",
      selectionDetails: {},
      selectedcount: 0,
      currentUserDetails: "",
      isDepartmentAdmin: false,
      isSuperAdmin: false,
    };
    const listObj = this.props.listName;
    this._listName = listObj?.title;

    this.getAllrequestesData();
  }

  private _commiteeRedirect = (item: any, user: any) => {
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
  };

  private _BoardRedirect = (item: any, user: any) => {
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
  };

  private _EnoteRedirect = (item: any, user: any) => {
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
  };

  private redirectToViewPage = async (item: any) => {
    const user = await this.props.sp?.web.currentUser();
    if (this.props.noteType === "eCommittee") {
      if (item.CommitteeType === "Committee") {
        this._commiteeRedirect(item, user);
      }
      if (item.CommitteeType === "Board") {
        this._BoardRedirect(item, user);
      }
    } else {
      this._EnoteRedirect(item, user);
    }
  };

  private _EnoteDetailslistColumns = () => {
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
    } else if (
      viewType === "In Progress" ||
      viewType === "All Approved" ||
      viewType === "All Rejected" ||
      viewType === "Draft Requests"
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
  };

  private _EcommiteeeDetailslistColumns = () => {
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
    } else if (
      viewType === "In Progress" ||
      viewType === "All Approved" ||
      viewType === "All Rejected" ||
      viewType === "Draft Requests"
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
      this._defalutDetailslistColumns();
    }
    return this._columns;
  };

  private _PendingWithDetailslistColumns = () => {
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
  };

  private _defalutDetailslistColumns = () => {
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
  };

  private _bindColumns = () => {
    const { noteType } = this.props;
    if (noteType === "enote") {
      this._EnoteDetailslistColumns();
    } else if (noteType === "eCommittee") {
      this._EcommiteeeDetailslistColumns();
    }

    return this._columns;
  };

  private _formatDate = (date: Date) => {
    if (!(date instanceof Date)) {
      throw new Error("Invalid date");
    }

    let day = date.getDate().toString().padStart(2, "0");
    let month = (date.getMonth() + 1).toString().padStart(2, "0"); // Month is zero-indexed
    let year = date.getFullYear().toString();
    let hours = date.getHours().toString().padStart(2, "0");
    let minutes = date.getMinutes().toString().padStart(2, "0");
    let seconds = date.getSeconds().toString().padStart(2, "0");

    return `${day}${month}${year}${hours}${minutes}${seconds}`;
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
    const user = await this.props.sp?.web.currentUser();
    let filterQury = "";
    const groups = await this.props.sp.web.currentUser.groups();
    if (
      groups &&
      groups?.some(
        (obj: { Title: string }) => obj.Title === this.props.superAdminGroupName
      )
    ) {
      switch (this.props.viewType) {
        case "All Requests":
          filterQury = "StatusNumber ne '100' ";
          break;
        case "Draft Requests":
          filterQury = `StatusNumber eq '100' and AuthorId eq ${user?.Id} `;
          break;
        case "In Progress":
          filterQury = `StatusNumber eq '2000' or StatusNumber eq '3000' or StatusNumber eq '4000' `;
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
        case "MyPendingNotes":
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
        case "PendingWith":
          filterQury = `StatusNumber eq '2000' or StatusNumber eq '3000' or StatusNumber eq '4000' `;
          break;

        default:
          break;
      }
    } else {
      switch (this.props.viewType) {
        case "All Requests":
          filterQury = `StatusNumber ne '100' and AuthorId eq ${user?.Id} `;
          break;
        case "Draft Requests":
          filterQury = `StatusNumber eq '100' and AuthorId eq ${user?.Id} `;
          break;
        case "In Progress":
          filterQury = `StatusNumber eq '2000' or StatusNumber eq '3000' or StatusNumber eq '4000' `;
          break;
        case "All Rejected":
          filterQury = `StatusNumber eq '8000' and AuthorId eq ${user?.Id} `;
          break;
        case "All Approved":
          filterQury = `StatusNumber eq '9000' AuthorId eq ${user?.Id} `;
          break;
        case "Noted Notes":
          filterQury = `NatureOfNote eq 'Information' AuthorId eq ${user?.Id}`;
          break;
        case "MyPendingNotes":
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
        case "PendingWith":
          filterQury = `StatusNumber eq '2000' or StatusNumber eq '3000' or StatusNumber eq '4000' `;
          break;

        default:
          break;
      }
    }
    return filterQury;
  };

  private getAllrequestesData = async () => {
    const filterQury = await this._getBylistWithFilterQuery();
    const groupedData: any = [];

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
        AuthorId: obj.AuthorId,
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
        CommitteeType:
          obj.CommitteeType === null && typeof obj.CommitteeType === "undefined"
            ? ""
            : obj.CommitteeType,
        committeeName:
          obj.CommitteeName === null
            ? obj.BoardName === null
              ? ""
              : obj.BoardName
            : obj.CommitteeName,
      });
    });

    if (this.props.viewType === "PendingWith") {
      allItems?.forEach(
        (_x: { CurrentApprover: any; CurrentApproverObj: any }) => {
          if (groupedData && Array.isArray(groupedData) && _x.CurrentApprover) {
            let found = false;
            groupedData.forEach(
              (obj: {
                count: number;
                CurrentApprover: any;
                crntApproverObject: any;
              }) => {
                if (obj.CurrentApprover === _x.CurrentApprover) {
                  obj.count++; // Increment count if CurrentApprover matches
                  found = true;
                }
              }
            );

            if (!found) {
              groupedData.push({
                CurrentApprover: _x.CurrentApprover,
                count: 1,
                crntApproverObject: _x.CurrentApproverObj,
              });
            }
          }
        }
      );
      this.setState({
        listItems: groupedData,
        allItems: groupedData,
      });
      this.paginateFn(groupedData, 1);
    } else {
      this.paginateFn(allItems, 1);
      this.setState({
        allItems: allItems,
      });
    }
    console.log(groupedData, "groupedData");
  };

  private paginateFn = (filterItem: any[], pageNo?: any) => {
    const { rowsPerPage } = this.state;
    console.log(filterItem, "filterItem");
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
    console.log(page, "page");
    this.setState({ page: page });
    this.paginateFn(this.state.allItems, page);
  };

  private _toggleWarningDialog = () => {
    this.setState({
      hideWarningDialog: !this.state.hideWarningDialog,
    });
  };
  private _toggleSuccussDialog = () => {
    this.setState({
      hideSuccussDialog: !this.state.hideSuccussDialog,
    });
  };

  public _onclickDelete = (id: any) => {
    this.setState({
      selectedId: id,
      hideDeleteDialog: false,
    });
  };

  private _deleteSelectedUser = async () => {
    this.setState({
      hideDeleteDialog: true,
    });
    try {
      await this.props.sp?.web.lists
        .getByTitle(this._listName)
        .items.getById(this.state.selectedId)
        .delete();
      this.setState({
        hideSuccussDialog: false,
        succussMsg: "Request has been deleted successfuly",
      });
    } catch (err) {
      this.setState({
        hideWarningDialog: false,
        warningMsg: "Failed to delete this request. Please try again",
      });
    }
  };

  public handleCloseSuccuss = () => {
    this.getAllrequestesData();
    this.setState({
      hideSuccussDialog: true,
    });
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
    const modalProps: any = {
      isBlocking: true,
      styles: modalPropsStyles,
      dragOptions: dragOptions,
    };
    const _items =
      this.props.noteType === "eCommittee"
        ? this.props.viewType === "Draft Requests"
          ? [
              {
                key: "newItem",
                name: "Create New Request",
                iconProps: {
                  iconName: "Add",
                },
                split: true,
                subMenuProps: {
                  items: [
                    {
                      key: "Committeenotes",
                      text: "Committee Note",
                      onClick: () => {
                        window.location.href =
                          this.props.context.pageContext.web.absoluteUrl +
                          `/SitePages/${this.props.newPageUrl}.aspx`;
                      },
                    },
                    {
                      key: "Boardnotes",
                      text: "Board Notes",
                      onClick: () => {
                        window.location.href =
                          this.props.context.pageContext.web.absoluteUrl +
                          `/SitePages/${this.props.CBnewPageUrl}.aspx`;
                      },
                    },
                  ],
                },
              },
              {
                key: "EditItem",
                name: "Edit Request",
                iconProps: {
                  iconName: "Edit",
                },
                disabled: this._hideCommandOption,
                onClick: () => {
                  const item = this.state.selectionDetails;
                  if (this.state.selectedcount === 0) {
                    this._hideCommandOption = true;
                  } else {
                    if (this.props.noteType === "eCommittee") {
                      if (
                        item.CommitteeType &&
                        item.CommitteeType === "Board"
                      ) {
                        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBeditPage}.aspx?itemId=${item.Id}`;
                      } else {
                        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;
                      }
                    } else {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;
                    }
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
                    this._hideCommandOption = true;
                    return null;
                  } else {
                    if (this.props.noteType === "eCommittee") {
                      if (
                        item.CommitteeType &&
                        item.CommitteeType === "Board"
                      ) {
                        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBviewPageUrl}.aspx?itemId=${item.Id}`;
                      } else {
                        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                      }
                    } else {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                    }
                  }
                },
              },
            ]
          : this.props.viewType === "PendingWith"
          ? [
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
            ]
          : [
              {
                key: "newItem",
                name: "Create New Request",
                iconProps: {
                  iconName: "Add",
                },
                split: true,
                subMenuProps: {
                  items: [
                    {
                      key: "Committeenotes",
                      text: "Committee Note",
                      onClick: () => {
                        window.location.href =
                          this.props.context.pageContext.web.absoluteUrl +
                          `/SitePages/${this.props.newPageUrl}.aspx`;
                      },
                    },
                    {
                      key: "Boardnotes",
                      text: "Board Notes",
                      onClick: () => {
                        window.location.href =
                          this.props.context.pageContext.web.absoluteUrl +
                          `/SitePages/${this.props.CBnewPageUrl}.aspx`;
                      },
                    },
                  ],
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
                    this._hideCommandOption = true;
                    return null;
                  } else {
                    if (this.props.noteType === "eCommittee") {
                      if (
                        item.CommitteeType &&
                        item.CommitteeType === "Board"
                      ) {
                        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBviewPageUrl}.aspx?itemId=${item.Id}`;
                      } else {
                        window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                      }
                    } else {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                    }
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
            ]
        : this.props.viewType === "Draft Requests"
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
              onClick: () => {
                const item = this.state.selectionDetails;
                if (this.state.selectedcount === 0) {
                  this._hideCommandOption = true;
                } else {
                  if (this.props.noteType === "eCommittee") {
                    if (item.CommitteeType && item.CommitteeType === "Board") {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBeditPage}.aspx?itemId=${item.Id}`;
                    } else {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;
                    }
                  } else {
                    window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.editPage}.aspx?itemId=${item.Id}`;
                  }
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
                  this._hideCommandOption = true;
                  return null;
                } else {
                  if (this.props.noteType === "eCommittee") {
                    if (item.CommitteeType && item.CommitteeType === "Board") {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBviewPageUrl}.aspx?itemId=${item.Id}`;
                    } else {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                    }
                  } else {
                    window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                  }
                }
              },
            },
          ]
        : this.props.viewType === "PendingWith"
        ? [
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
          ]
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
                  this._hideCommandOption = true;
                  return null;
                } else {
                  if (this.props.noteType === "eCommittee") {
                    if (item.CommitteeType && item.CommitteeType === "Board") {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.CBviewPageUrl}.aspx?itemId=${item.Id}`;
                    } else {
                      window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                    }
                  } else {
                    window.location.href = `${this.props.context.pageContext.web.absoluteUrl}/SitePages/${this.props.viewPageUrl}.aspx?itemId=${item.Id}`;
                  }
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

        <Dialog
          hidden={this.state.hideSuccussDialog}
          onDismiss={this._toggleSuccussDialog}
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
        >
          <p>{this.state.succussMsg}</p>
          <DialogFooter>
            <PrimaryButton onClick={this.handleCloseSuccuss} text="Ok" />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={this.state.hideWarningDialog}
          onDismiss={this._toggleWarningDialog}
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
        >
          <p>{this.state.warningMsg}</p>
          <DialogFooter>
            <PrimaryButton onClick={this._toggleWarningDialog} text="OK" />
          </DialogFooter>
        </Dialog>

        <Dialog
          hidden={this.state.hideDeleteDialog}
          onDismiss={() =>
            this.setState({ hideDeleteDialog: !this.state.hideDeleteDialog })
          }
          dialogContentProps={dialogContentProps}
          modalProps={modalProps}
        >
          <p>Do you want delete this request?</p>
          <DialogFooter>
            <PrimaryButton onClick={this._deleteSelectedUser} text="Yes" />
            <DefaultButton
              onClick={() => this.setState({ hideDeleteDialog: false })}
              text="No"
            />
          </DialogFooter>
        </Dialog>
      </section>
    );
  }
}
