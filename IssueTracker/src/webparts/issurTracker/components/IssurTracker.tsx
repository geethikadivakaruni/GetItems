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
 import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
 import { IPersonaProps } from '@fluentui/react/lib/Persona';
 import {
  DatePicker,
  DayOfWeek,

  IDropdownOption,
  
  defaultDatePickerStrings,
} from '@fluentui/react';
 import { Dropdown, IDropdownStyles } from '@fluentui/react/lib/Dropdown';
import { Stack, IStackProps, IStackStyles } from '@fluentui/react/lib/Stack';
import { Item, Items, ItemVersion, Web } from "@pnp/sp/presets/all";
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {IColumn,DetailsList, SelectionMode,Selection, DetailsListLayoutMode, mergeStyles, Link,Image,ImageFit} from '@fluentui/react';
import * as moment from 'moment';
import { PnPClientStorage, PnPClientStorageWrapper } from '@pnp/common';
import { MarqueeSelection } from '@fluentui/react';
import { Label } from '@fluentui/react/lib/Label';
import { TextField } from '@fluentui/react/lib/TextField';
// import Form from './Form'
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

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 }, root: { height: 100 } };
export interface IDetailsListState {
  Items: ISPList[];
  columns: any;
  isColumnReorderEnabled: boolean;
  disabled:boolean;
  selectionDetails: string;
 showForm:boolean;
 selectionMode?: SelectionMode;
 IsEditEnabled?:boolean
 DisplaySelectedItem?:boolean;
//  selectionMode:boolean,
}
const StatusOptions = [
  { key: 'A', text: 'Blocked' },
  { key: 'B', text: 'In progress' },
  { key: 'C', text: 'Completed' },
  { key: 'D', text: 'Duplicate' },
  { key: 'E', text: 'By design' },
  { key: 'D', text: "Won't fix" },
  { key: 'E', text: 'New' },
];
const PriorityOptions = [
  { key: 'A', text: 'High' },
  { key: 'B', text: 'Critical' },
  { key: 'C', text: 'Normal' },
  { key: 'D', text: 'Low' },
  { key: 'E', text: 'Option e' },
];
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
 showForm:false,
 selectionMode: SelectionMode.single,
 IsEditEnabled:false,
 DisplaySelectedItem:false
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

   private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
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
   public showForm = () => {
      return (
       
        <div> 
           {alert(this.state.showForm)}
       <form id= "add-app">
   
            <Label>Description</Label>
            {/* <input type="text"> </input> */}
            <TextField></TextField>
   
            <Label>Title </Label>
            <TextField></TextField>
   
            <Label>Priority </Label>
            <Dropdown
             placeholder="Select an option"
             label="Dropdown with error message"
             options={PriorityOptions}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
             styles={dropdownStyles} />

<Label>Status </Label>
            <Dropdown
             placeholder="Select an option"
             label="Dropdown with error message"
             options={StatusOptions}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
             styles={dropdownStyles} />

             <Label>Assigned To</Label>
             {/* <PeoplePicker/> */}
             <Label>Date Reported</Label>
             <DatePicker
        // firstDayOfWeek={firstDayOfWeek}
        placeholder="Select a date..."
        ariaLabel="Select a date"
        // DatePicker uses English strings by default. For localized apps, you must override this prop.
        strings={defaultDatePickerStrings}
      />
      <Label>Issue Source</Label>
      {/* //link */}
      <Label>IssueLoggedBy</Label>
      {/* PeoplePicker */}
<Label>Images</Label>
<Image
      // {...imageProps}
      alt="Example with no image fit value and height or width is specified."
      width={100}
      height={100}
    />
<Label>Amount</Label>
<TextField></TextField>
   
  <DefaultButton onClick={()=>this.setState({showForm:false})}> Cancel</DefaultButton>
  <PrimaryButton onClick={this.createItem}>Create Item</PrimaryButton>
         </form>
         </div>
        );
    }
 
  public render(): React.ReactElement<IIssurTrackerProps> {
  const{selectionMode}=this.state
    return (

 <div><h1>Display SharePoint list data using spfx</h1>
 
   <PrimaryButton text="New Item" onClick={()=>this.setState({showForm:true})}/> 
   <DefaultButton text="Display" onClick={this.getFormToDisplay } />
    <PrimaryButton text="Edit" onClick={this.getFormToEdit}  />
  
  {this.state.showForm?this.showForm():""}
     
      <MarqueeSelection selection={this._selection}>
        {console.log("this.state.Items",this.state.Items)}
         <DetailsList
          items={this.state.Items}
          columns={this.state.columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          isHeaderVisible={true}
          onRenderItemColumn={this._onRenderItemColumn}
          selectionMode={selectionMode}
          selectionZoneProps={{
            selection: this._selection,
            disableAutoSelectOnInputElements: true,
             selectionMode: selectionMode,
          }}
        /> 
       </MarqueeSelection>
        <h3>Note : Image is clickable.</h3>
     
     
        
        
  </div> 
    );
    
  }
  
  getFormToDisplay(){
    return (
       
      <div> 
         {alert(this.state.showForm)}
     <form id= "add-app">
 
          <Label>Description</Label>
          {/* <input type="text"> </input> */}
         {/* {item.Description} */}
 
        <Label>Priority </Label>
          {/* {item.Priority} */}

        <Label>Status </Label>
         {/* {item.Status} */}

           <Label>Assigned To</Label>
          {/* {item.Assignedto.Title} */}
           <Label>Date Reported</Label>
           {/* {item.DateReported} */}
   
    <Label>Issue Source</Label>
{/* {item.IssueSource} */}
    <Label>IssueLoggedBy</Label>
  {/* {item.Issueloggedby} */}
<Label>Images</Label>
{/* {item.Images} */}
<Label>Amount</Label>
{/* {item.Amount} */}
{/* <DefaultButton onClick={()=>this.setState({DisplaySelectedItem:false})}> Cancel</DefaultButton> */}

</form>
</div>
    )
 
  }
  getFormToEdit(){
    return (
      this.state.Items.map(item =>{
      <div> 
         {alert(this.state.showForm)}
     <form id= "add-app">
 
          <Label>Description</Label>
          {/* <input type="text"> </input> */}
          <TextField value={item.Description} ></TextField>
 
         
 
          <Label>Priority </Label>
          <Dropdown
           placeholder="Select an option"
           label="Dropdown with error message"
           options={PriorityOptions}
          //  value={this.state.}
          //  onChange={this.onChange}
          //  errorMessage={showError ? 'This dropdown has an error' : undefined}
           styles={dropdownStyles} />

<Label>Status </Label>
          <Dropdown
           placeholder="Select an option"
           label="Dropdown with error message"
           options={StatusOptions}
          //  errorMessage={showError ? 'This dropdown has an error' : undefined}
           styles={dropdownStyles} />

           <Label>Assigned To</Label>
           <PeoplePicker
 context={this.props.context}
 titleText="People Picker"
 personSelectionLimit={3}
 groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
 showtooltip={true}
//  isRequired={true}
 disabled={false}
//  selectedItems={this._getPeoplePickerItems} 
/>
           <Label>Date Reported</Label>
           <DatePicker
      // firstDayOfWeek={firstDayOfWeek}
      placeholder="Select a date..."
      ariaLabel="Select a date"
      // DatePicker uses English strings by default. For localized apps, you must override this prop.
      strings={defaultDatePickerStrings}
    />
    <Label>Issue Source</Label>
    {/* //link */}
    <Label>IssueLoggedBy</Label>
    {/* PeoplePicker */}
<Label>Images</Label>
<Image
    // {...imageProps}
    alt="Example with no image fit value and height or width is specified."
    width={100}
    height={100}
  />
<Label>Amount</Label>
<TextField></TextField>
 
<DefaultButton onClick={()=>this.setState({IsEditEnabled:false})}> Cancel</DefaultButton>
<PrimaryButton onClick={this.updateItem}>Create Item</PrimaryButton>
</form>
</div>
   } )
    )
  }

  public createItem(){
    return console.log("Added Items")
    // sp.web.lists.getByTitle("Issue Tracker").add({
      
    // })
  }
  public updateItem(){
    return console.log("Added Items")
    // sp.web.lists.getByTitle("Issue Tracker").add({
      
    // })
  }
}

