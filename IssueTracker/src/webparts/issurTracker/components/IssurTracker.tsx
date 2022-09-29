import * as React from 'react';
import styles from './IssurTracker.module.scss';
import { IIssurTrackerProps } from './IIssurTrackerProps';
import { escape } from '@microsoft/sp-lodash-subset';
// import * as sp from '@pnp/sp'
// import { Web } from "@pnp/sp/webs"; 
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { DefaultButton, PrimaryButton, CommandBarButton,
 } from '@fluentui/react/lib/Button';
import { Stack, IStackProps, IStackStyles } from '@fluentui/react/lib/Stack';
import { Items, ItemVersion, Web } from "@pnp/sp/presets/all";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {IColumn,DetailsList, SelectionMode,Selection, DetailsListLayoutMode, mergeStyles, Link,Image,ImageFit} from '@fluentui/react';
import * as moment from 'moment';
import { PnPClientStorage, PnPClientStorageWrapper } from '@pnp/common';
import { MarqueeSelection } from '@fluentui/react';
import {Form} from './Form'
export interface ISPList {
  Description: string;
  Priority: string;
  Status: string;
  Assignedto:{
    Title:string
  }; 
DateReported:string;
IssueSource:{

};
Images:any;
Issueloggedby:{
  Title:string
};
Amount:number;
ID:number
}


export interface IDetailsListState {
  Items: ISPList[];
  columns: any;
  isColumnReorderEnabled: boolean;
  disabled:boolean;
  selectionDetails: string;
 isFormEnabled:boolean;
//  selectionMode:boolean,
}


export default class IssurTracker extends React.Component<IIssurTrackerProps ,IDetailsListState> {
  private _selection: Selection;
  constructor(props:any) {
    super(props);
    const columns: IColumn[] = [
      {
        key: "Description",
        name: "Description",
        fieldName: "Description",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "Amount",
        name: "Amount",
        fieldName: "Amount",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
      },
      {
        key: "Priority",
        name: "Priority",
        fieldName: "Priority",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
        isPadded: true,
      },
      {
        key: "Status",
        name: "Status",
        fieldName: "Status",
        minWidth: 70,
        maxWidth: 90,
        isRowHeader: true,
        isResizable: true,
        data: "string",
        isPadded: true
      },
      {
        key: "AssignedTo",
        name: "AssignedTo",
        fieldName: "AssignedTo",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
        isPadded: true
      },

      {
        key: "DateReported",
        name: "DateReported",
        fieldName: "DateReported",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
      },
      {
        key: "Images",
        name: "Images",
        fieldName: "Images",
        minWidth: 210,
        maxWidth: 350,
        isResizable: true,
        data: "string",
      },
      {
        key: "Issueloggedby",
        name: "Issueloggedby",
        fieldName: "Issueloggedby",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
      },
      {
        key: "IssueSource",
        name: "IssueSource",
        fieldName: "IssueSource",
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: "string",
      }
    ];
    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });
    this.state = {
      Items: [],
      columns: columns,
      isColumnReorderEnabled: true,
      disabled:false,
      selectionDetails: this._getSelectionDetails(),
 isFormEnabled:false,
//  selectionMode:false
    };
    // sp.setup({
    //   spfxContext: this.props.spcontext
    // });
  }
  private _getSelectionDetails(): any {
    const selectionCount = this._selection.getSelectedCount();
    const selectedId = this._selection.getItems();
    console.log(selectionCount)
    console.log("selectedId",selectedId)
    
   }
  private web = Web("https://sites.ey.com/sites/testcanda/");
  private async getData() {
    // let web = Web(this.props.webURL);
    const data: ISPList[] = [];
    const items:any[] = await this.web.lists.getByTitle("Issue tracker").items.select("ID","Description","Priority","Status","Assignedto/Title","DateReported","IssueSource","Images","Issueloggedby/Title","Amount").expand("Assignedto","Issueloggedby").getAll();
      console.log(items);
      await items.forEach(async item => {
        await data.push({
          Description: item.Description,
          Priority: item.Priority,
          Status: item.Status,
          Assignedto:item.Assignedto,
          DateReported:moment(item.DateReported).format("DD-MM-YYYY") ,
          IssueSource: item.IssueSource,
          Images: item.Images,
          Issueloggedby:item.Issueloggedby,          
          Amount:item.Amount,
          ID:item.ID
        });
      });
      console.log(data);
      console.log("Items",this.state.Items)
       this.setState({ Items: data });
      console.log("Items",this.state.Items)
    
  }
  public async componentDidMount() {
     await this.getData();
  }
  private _onItemInvoked = (item: ISPList): void => {
    // alert(`Item invoked: ${item.}`);
  };
  public _onRenderItemColumn = (item: ISPList, index: number, column: IColumn): JSX.Element | any => {
    const src = item.Images;
  console.log(column.key);
  console.log("Heelo world");
  console.log("item",item)
     if(column.key=="Amount")
     {
      return item.Amount
     }
     else if(column.key=="Priority")
     {
      return item.Priority
     }
     else if(column.key=="AssignedTo")
     {
      return item.Assignedto?item.Assignedto.Title:""
     }
     else if(column.key=="Status")
     {
      return item.Status?item.Status:""
     }
     else if(column.key=="Description")
     {
      return item.Description?item.Description:""
     }
     else if(column.key=="Images")
     {
      return item.Images?item.Images:""
     }
     else if(column.key=="DateReported")
     {
      return item.DateReported?item.DateReported:""
     }
     else if(column.key=="AssignedTo")
     {
      return item.Issueloggedby?item.Issueloggedby.Title:""
     }
    //  else if(column.key=="IssueSource")
    //  {
    //   return item.IssueSource?item.IssueSource:""
    //  }
     else{
      return ""
     }
    }
  
 
  public render(): React.ReactElement<IIssurTrackerProps> {
  
    return (

 <div><h1>Display SharePoint list data using spfx</h1>
     <PrimaryButton text="New Item" onClick={this.getForm}/>
      <DefaultButton text="Display" onClick={this.getFormToDisplay} />
      <PrimaryButton text="Edit" onClick={this.getFormToEdit}  />
      
  
      <MarqueeSelection selection={this._selection}>
        {console.log("this.state.Items",this.state.Items)}
         <DetailsList
          items={this.state.Items}
          columns={this.state.columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onRenderItemColumn={this._onRenderItemColumn}
          selectionMode={SelectionMode.single}
          selectionZoneProps={{
            selection: this._selection,
            disableAutoSelectOnInputElements: true,
            // selectionMode: selectionMode,
          }}
        /> 
       </MarqueeSelection>
        <h3>Note : Image is clickable.</h3></div>
  
      
   
    );
  }
  getFormToDisplay(){
    return console.log("Hello World")
  }
  getFormToEdit(){
    return console.log("Hello World")
  }

  getForm(){
  this.setState({
    isFormEnabled:true
   });
 
    
  }
 

}

