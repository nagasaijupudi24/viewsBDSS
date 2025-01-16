import * as React from "react";
import styles from "./ChartView.module.scss";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import {
  ChartControl, ChartType,

} from "@pnp/spfx-controls-react/lib/ChartControl";
import * as XLSX from "xlsx";
import * as FileSaver from 'file-saver';
import {
  PrimaryButton,IIconProps,Spinner, SpinnerSize 
} from "@fluentui/react";

import { IXenWpUcoBankProps } from "../webparts/xenWpUcoBank/components/IXenWpUcoBankProps";
import { WebPartTitle } from "@pnp/spfx-controls-react";
const excelIcon: IIconProps = { iconName: "ExcelDocument" };

export interface IChartViewData {
  [x: string]: string
}
export interface IChartStates {
  chartViewData: IChartViewData;
  DropDownValue: string;
}
export default class ChartView extends React.Component<
IXenWpUcoBankProps,
  IChartStates
> {
  private _listName: any;

  private chartOptions: any = {
    legend: {
      display: true,
      position: "left",
    },
    title: {
      display:true,
      text: this.props.chartTitle ||"Status",
    },
  };
  constructor(props: IXenWpUcoBankProps) {
    super(props);
    this.state = {
      chartViewData: {},
      DropDownValue: "Status",

    };
    const listObj = this.props.listName;
    this._listName = listObj?.title;

  }

  public render(): React.ReactElement<IXenWpUcoBankProps> {
   
    return (
      <div className={styles.ChartViewcontainer}>
           <WebPartTitle
          displayMode={this.props.displayMode}
          title={this.props.title}
          updateProperty={this.props.updateProperty}
        />
        <div className={styles.ContorlSection}>

          <div className={styles.button}>
            <PrimaryButton
              text="Export"
              iconProps={excelIcon}
              onClick={this._getExcel}
              allowDisabledFocus
            />
          </div>
        </div>
        <ChartControl
           type={ChartType[this.props.chartType as keyof typeof ChartType] || ChartType.Pie}
          datapromise={this._getItemsCount(
        this._listName,
            this.props.columnName
          )}
          loadingtemplate={() => (
            <Spinner size={SpinnerSize.large} label="Loading..." />
          )}
          options={this.chartOptions}
        />
      </div>
    );
  }

  private filterquery=async ()=>{
    let user = await this.props.sp?.web.currentUser();
    console.log(user, "user");
    const {viewType}=this.props;
    let query="";
    if(viewType ==="MyNotes"){
      query=`AuthorId eq ${user?.Id}`

    }
    else if(viewType ==="DBNoteReports"){
    

      query=`StatusNumber eq '2000' or StatusNumber eq '3000' or StatusNumber eq '4000' or StatusNumber eq '5000' or StatusNumber eq '8000' or StatusNumber eq '9000' `
    }
    
    return query;
  }
  private _getItemsCount = async (listTitle: string, columnTitle: string): Promise<any> => {
const filterQury = await this.filterquery();
    const items = await this.props.sp?.web.lists.getByTitle(listTitle).items.filter(filterQury)();
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
  private _getChartData = (uniqueValues: string[], allItems: string[] | any, columnName: string):any => {
    const llbArr: string[] = [];
    const dataArr: number[] = [];
    uniqueValues.forEach((uniqueValue: string) => {
      const arr = allItems.filter(
        (item: { [x: string]: string | any; }) => item[columnName] && item[columnName] === uniqueValue //data type channged 
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
  }
  private _getExcel = async (): Promise<void> => {
    const filterQury = await this.filterquery();
    const items = await this.props.sp.web.lists
      .getByTitle(this._listName)
      .items.filter(filterQury)();
    const excelData = items?.map((item) => {
      const res: string[] = [];
      this.props.fields.forEach((element) => {

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
    const Heading = [this.props.fields];
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
      FileSaver(data,this.props.noteType + fileExtension);
    }
  }
}