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
import {
  DefaultButton, PrimaryButton, CommandBarButton,
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
import { IColumn, DetailsList, SelectionMode, Selection, DetailsListLayoutMode, mergeStyles, Link, Image, ImageFit } from '@fluentui/react';
import * as moment from 'moment';
import { PnPClientStorage, PnPClientStorageWrapper } from '@pnp/common';
import { MarqueeSelection } from '@fluentui/react';
import { Label } from '@fluentui/react/lib/Label';
import { TextField } from '@fluentui/react/lib/TextField';
import { ThemeSettingName } from 'office-ui-fabric-react';
import { _Item } from '@pnp/sp/items/types';
// import Form from './Form'
export interface ISPList {
  Description: string;
  Priority: string;
  Status: string;
  Assignedto: {
    Title: string
  };
  DateReported: string;
  IssueSource: {

  };
  Images: any;
  Issueloggedby: {
    Title: string
  };
  Amount: number;
  ID: number
}

const dropdownStyles: Partial<IDropdownStyles> = { dropdown: { width: 300 }, root: { height: 100 } };
export interface IDetailsListState {
  Items: ISPList[];
  SelectedItems: ISPList[];
  columns: any;
  isColumnReorderEnabled: boolean;
  disabled: boolean;
  selectionDetails: string;
  showForm: boolean;
  selectionMode?: SelectionMode;
  IsEditEnabled?: boolean
  DisplaySelectedItem?: boolean;
  Description: string;
  Priority: string;
  Status: string;
  getSelectedData: boolean;
  Assignedto: {
    Title: string
  };
  DateReported: any;
  IssueSource: {

  };
  Images: any;
  Issueloggedby: {
    Title: string
  };
  Amount: string;
  ID: number;
  selectedId: number;
  hideTable:boolean;
  selectedDescription:string;
selectedPriority: string;
      selectedStatus: string;
      selectedAssignedto: {
        Title: string;
      },
      selectedDateReported: string;
     
      selectedIssueSource: string;
      selectedImages: string;
      selectedIssueloggedby: {
        Title: string;
      },
      selectedAmount:string;
  //  value:any
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
]
export default class IssurTracker extends React.Component<IIssurTrackerProps, IDetailsListState, ISPList> {
  private _selection: Selection;
  private web = Web("https://sites.ey.com/sites/testcanda/");

  constructor(props: any) {
    super(props);


    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    this.handleChange = this.handleChange.bind(this)
    this.handleSubmit=this.handleSubmit.bind(this)
    this.handleAMountChange=this.handleAMountChange.bind(this)    
    this.handlePriorityChange=this.handlePriorityChange.bind(this)    
    this.handleStatusChange=this.handleStatusChange.bind(this)    
    this.handleIssueLoggedByChange=this.handleIssueLoggedByChange.bind(this)  
    this.handleImageChange=this.handleImageChange.bind(this)   
     this.handleAssignedToChange=this.handleAssignedToChange.bind(this)  
    this.handleSubmit=this.handleSubmit.bind(this)
    this.updateItem=this.updateItem.bind(this)
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

    this.state = {
      Items: [],
      SelectedItems: [],
      columns: columns,
      isColumnReorderEnabled: true,
      disabled: false,
      Description: "",
      Priority: '',
      Status: '',
      Assignedto: {
        Title: ''
      },
      DateReported: null,
      IssueSource: {},
      Images: '',
      Issueloggedby: {
        Title: ''
      },
      Amount: '0',
      ID: 0,
      selectedDescription: "",
      selectedPriority: '',
      selectedStatus: '',
      selectedAssignedto: {
        Title: ''
      },
      selectedDateReported: null,
     
      selectedIssueSource: "",
      selectedImages: '',
      selectedIssueloggedby: {
        Title: ''
      },
      selectedAmount: '0',
      // value:'',
      selectionDetails: this._getSelectionDetails(),
      showForm: false,
      selectionMode: SelectionMode.single,
      IsEditEnabled: false,
      DisplaySelectedItem: false,
      getSelectedData: false,
      selectedId: 0,
      hideTable:false
      //  selectionMode:false
    };
    // sp.setup({
    //   spfxContext: this.props.spcontext
    // });
  }


 

  public handleChange = (event: any, value: any) => {
    // event.preventDefault();
    // console.log("val",value)
    // console.log("evval",event.target.value)
    this.setState({Description: event.target.value});  
   // this.setState({ Description: value });
    console.log(this.state.Description)

  }
  public handlePriorityChange = (event: any,option:IDropdownOption) => {
let  value=option.text
this.setState({ Priority: value })
   
  }
  public handleStatusChange(event: any,option:IDropdownOption) {
    // event.preventDefault();
    let  value=option.text
   return this.setState({ Status: value })
   
  }
  public handleAssignedToChange(event:any){

    this.setState({Assignedto:event.target.value})
  }
  public handleDateChange = (event: any) => {
    // event.preventDefault();
    // let Date=Da
    // console.log(Date)
    this.setState({ DateReported: event.target.value })
    console.log("Date", this.state.DateReported)
  }

  public handleImageChange = (event: any) => {
    // event.preventDefault();
    this.setState({ Images: event.target.value })
    console.log("Image", this.state.Images)
  }
  public handleAMountChange = (event: any) => {
    // event.preventDefault();
    this.setState({ Amount: event.target.value })
    console.log("Amounty", this.state.Amount)
  }
  public handleIssueSourceChange(event: any) {
    // event.preventDefault();
    this.setState({ IssueSource: event.target.value })
    console.log("Priority", this.state.Priority)
  }
  public handleIssueLoggedByChange(event: any) {
    // event.preventDefault();
    this.setState({ Issueloggedby: event.target.value })
    console.log("Priority", this.state.Priority)
  }

  private _getPeoplePickerItems(items: any[]) {
    console.log('Items:', items);
  }


  public async componentDidMount() {
    await this.getData();
    console.log(this.state)
    //  this.state.getSelectedData?this.getSelectedItemData:""
  }


  public render(): React.ReactElement<IIssurTrackerProps> {

    const { selectionMode, selectionDetails } = this.state
    return (

      <div>
        <h1>Display SharePoint list data using spfx</h1>

        <PrimaryButton text="New Item" onClick={() => this.setState({ showForm: true ,hideTable:true})} />
        <DefaultButton text="Display" onClick={() => this.setState({ DisplaySelectedItem: true ,hideTable:true})} />
        <PrimaryButton text="Edit" onClick={() => this.setState({ IsEditEnabled: true,hideTable:true })} />

        {this.state.showForm ? this.showForm() : ""}
        {this.state.DisplaySelectedItem ? this.getFormToDisplay() : ""}
        {this.state.IsEditEnabled ? this.getFormToEdit() : ""}
        {/* <h1>{this.state.selectionDetails}</h1> */}
        {!(this.state.hideTable)?
        <>
        <MarqueeSelection selection={this._selection}>
          {/* // {console.log("this.state.Items",this.state.Items)} */}
          <DetailsList
            items={this.state.Items}
            columns={this.state.columns}
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            onRenderItemColumn={this._onRenderItemColumn}
            selectionMode={selectionMode}


          />
        </MarqueeSelection>
        
        </>
        :""
  }
        
       



      </div>
    );

  }

  private _getSelectionDetails(): any {
    const selectionCount = this._selection.getSelectedCount();
 
    let SelectedItemId;
    let SelectedItemDescription, SelectedItemAmount,SelectedItemPriority,SelectedItemStatus,SelectedItemDate;

    if (selectionCount == 0) {
      return 'No items selected';

    }
    else {

      return (
        SelectedItemId = (this._selection.getSelection()[0] as IDetailsListState).ID,
        // this.setState({ selectedId: SelectedItemId });
    SelectedItemDescription=(this._selection.getSelection()[0] as IDetailsListState).Description,
    SelectedItemAmount=(this._selection.getSelection()[0] as IDetailsListState).Amount,
SelectedItemPriority=(this._selection.getSelection()[0] as IDetailsListState).Priority,
SelectedItemStatus=(this._selection.getSelection()[0] as IDetailsListState).Status,
SelectedItemDate=(this._selection.getSelection()[0] as IDetailsListState).DateReported,
    this.setState({ selectedId: SelectedItemId ,selectedDescription:SelectedItemDescription,selectedAmount: SelectedItemAmount,selectedPriority:SelectedItemPriority,selectedDateReported:SelectedItemDate,selectedStatus:SelectedItemStatus})




      )
    }

  }


  public async getSelectedItemData() {
    const data: ISPList[] = [];

    const id = this.state.selectedId
    console.log("Id", this.state.selectedId)

    const allListItems: ISPList[] = await this.web.lists.getByTitle("Issue tracker").items.getById(id).select("ID", "Description", "Priority", "Status", "Assignedto/Title", "DateReported", "IssueSource", "Images", "Issueloggedby/Title", "Amount").expand("Assignedto", "Issueloggedby").get();
   console.log(allListItems)
   allListItems.forEach(item=>{
    data.push({
      Description: item.Description,
      Priority: allListItems[0].Priority,
      Status: allListItems[0].Status,
      Assignedto: allListItems[0].Assignedto,
      DateReported: moment(allListItems[0].DateReported).format("DD-MM-YYYY"),
      IssueSource: allListItems[0].IssueSource,
      Images:allListItems[0].Images,
      Issueloggedby: allListItems[0].Issueloggedby,
      Amount: allListItems[0].Amount,
      ID: allListItems[0].ID
    });
   })
    
    console.log(data)
    this.setState({ SelectedItems: data })
    console.log("selectedItems",this.state.SelectedItems)
    
  }




  private _onItemInvoked = (item: IDetailsListState): void => {
    alert(`Item invoked: ${item.ID}`);
    // this.geSelectedItemData()

  };


  getFormToDisplay() 
{
  return(
    <>
<Label>Description</Label>
      {this.state.selectedDescription} 
      
       <Label>Priority </Label>
        {this.state.selectedPriority}
     
       <Label>Status </Label>
     
       {this.state.selectedStatus}
 
         <Label>Assigned To</Label>

       
         <Label>Date Reported</Label>
         {this.state.selectedDateReported}
      {/* {item.DateReported?item.DateReported:""}
         */}
{/*       
       <Label>Issue Source</Label>
      {/* {item.IssueSource?item.IssueSource:''}  */}

      
       {/* <Label>IssueLoggedBy</Label>
       {/* {item.Issueloggedby?item.Issueloggedby.Title:""}   */}
       {/* <p>{this.state.Issueloggedby?this.state.Issueloggedby.Title:''} </p> */} 
       {/* <Label>Images</Label>
       <Image
         // {...imageProps}
         src={this.state.Images}
         // onChange={this.handleImageChange}
         alt="Example with no image fit value and height or width is specified."
         width={100}
         height={100}
       />
        */}
       <Label>Amount</Label>
    
       {this.state.selectedAmount} <br/>
       <DefaultButton onClick={()=>this.setState({DisplaySelectedItem:false,selectedId:0,hideTable:false})}> Cancel</DefaultButton> 
   {console.log(this.state)}
   </>
  )
}
        
     

  getFormToEdit() {
   
    return (

      <div> 
      <form id="add-app" onSubmit={this.updateItem}>

      <Label>Description</Label>
          {/* <input type="text"> </input> */}
          <TextField value={this.state.Description}
          onChange={this.handleChange}
          >
          </TextField>
          <Label>Priority </Label>
          <Dropdown
            placeholder="Select an option"
            //  label="Dropdown with error message"
            defaultValue={this.state.Priority}
            options={PriorityOptions}
             onSelect={()=>this.handlePriorityChange}
             onChange={this.handlePriorityChange}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
            styles={dropdownStyles} />
          <Label>Status </Label>
          <Dropdown
            placeholder="Select an option"
            //  label="Dropdown with error message"
            defaultValue={this.state.Status}
            onSelect={()=>this.handleStatusChange}
            onChange={this.handleStatusChange}
            options={StatusOptions}
           
            styles={dropdownStyles} />
          <Label>Assigned To</Label>
          {/* {console.log("context",this.props.context)} */}
     <PeoplePicker
            context={this.props.context}
             principalTypes={[PrincipalType.User]} 
            // // titleText="People Picker"
            // personSelectionLimit={1}
            // ensureUser={true}    
            // showtooltip={true}
            // groupName={""}
            // //  isRequired={true}
            // disabled={false}
          //  selectedItems={this._getPeoplePickerItems} 
          />
          <Label>Date Reported</Label>
          <DatePicker
            // firstDayOfWeek={firstDayOfWeek}
            value={this.state.DateReported}
            // onChange={this.handleDateChange}
            placeholder="Select a date..."
            ariaLabel="Select a date"
     
            onSelectDate={ date => this.setState({ DateReported: date }) }
            // DatePicker uses English strings by default. For localized apps, you must override this prop.
            strings={defaultDatePickerStrings}
          />
          <Label>Issue Source</Label>
          <TextField ></TextField>
          <Label>IssueLoggedBy</Label>
          <PeoplePicker
            context={this.props.context}
            principalTypes={[PrincipalType.User]}
            // titleText="People Picker"
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            //  isRequired={true}
            disabled={false}
          //  selectedItems={this._getPeoplePickerItems} 
          />
          <Label>Images</Label>
          
          <Label>Amount</Label>
          <TextField type="number" defaultValue={this.state.Amount}
          onChange={this.handleAMountChange}
          ></TextField>
        
          <DefaultButton onClick={() => this.setState({ IsEditEnabled: false ,hideTable:false})}> Cancel</DefaultButton>
          <PrimaryButton type="submit">Update Item</PrimaryButton>
        </form>
      </div>
    )

  }



  public   updateItem() {
    let date=   moment(this.state.DateReported).format("DD-MM-YYYY")
    let list = sp.web.lists.getByTitle("Issue tracker");
    console.log(this.state)
  this.web.lists.getByTitle("Issue tracker").items.getById(this.state.selectedId).update({
      Description: this.state.Description,
      Priority: this.state.Priority,
      Status:this.state.Status,
      DateReported: date,
      Amount: this.state.Amount,

    });
 
this.setState({hideTable:false});


  }


  private async getData() {
    // let web = Web(this.props.webURL);
    const data: ISPList[] = [];
    const items: any[] = await this.web.lists.getByTitle("Issue tracker").items.select("ID", "Description", "Priority", "Status", "Assignedto/Title", "DateReported", "IssueSource", "Images", "Issueloggedby/Title", "Amount").expand("Assignedto", "Issueloggedby").getAll();
    // console.log(items);
    await items.forEach(async item => {
      await data.push({
        Description: item.Description,
        Priority: item.Priority,
        Status: item.Status,
        Assignedto: item.Assignedto,
        DateReported: moment(item.DateReported).format("DD-MM-YYYY"),
        IssueSource: item.IssueSource,
        Images: item.Images,
        Issueloggedby: item.Issueloggedby,
        Amount: item.Amount,
        ID: item.ID
      });
    });
    // console.log(data);
    // console.log("Items",this.state.Items)
    this.setState({ Items: data });
    // console.log("Items",this.state.Items)

  }
  public async handleSubmit(event: any){
    event.preventDefault();
 let date=   moment(this.state.DateReported).format("DD-MM-YYYY")
 console.log("date",date)
    console.log(this.state.Description)
    console.log(this.state)
    await this.web.lists.getByTitle('Issue tracker').items.add({
      Description: this.state.Description,
      Priority: this.state.Priority,
      Status:this.state.Status,
      DateReported: date,
      Amount: this.state.Amount,
      // Assignedto: this.state.Assignedto.Title
    });
    console.log("added items")
    this.setState({hideTable:false})

    this.setState({ showForm: false })

  }
  public showForm = () => {
    return (

      <div>
        {/* {alert(this.state.showForm)} */}
        <form id="add-app" onSubmit={this.handleSubmit}>

          <Label>Description</Label>
          {/* <input type="text"> </input> */}
          <TextField value={this.state.Description}
          onChange={this.handleChange}
          >
          </TextField>
          <Label>Priority </Label>
          <Dropdown
            placeholder="Select an option"
            //  label="Dropdown with error message"
            defaultValue={this.state.Priority}
            options={PriorityOptions}
             onSelect={()=>this.handlePriorityChange}
             onChange={this.handlePriorityChange}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
            styles={dropdownStyles} />
          <Label>Status </Label>
          <Dropdown
            placeholder="Select an option"
            //  label="Dropdown with error message"
            defaultValue={this.state.Status}
            onSelect={()=>this.handleStatusChange}
            onChange={this.handleStatusChange}
            options={StatusOptions}
           
            styles={dropdownStyles} />
          <Label>Assigned To</Label>
          {/* {console.log("context",this.props.context)} */}
     <PeoplePicker
            context={this.props.context}
             principalTypes={[PrincipalType.User]} 
            // // titleText="People Picker"
            // personSelectionLimit={1}
            // ensureUser={true}    
            // showtooltip={true}
            // groupName={""}
            // //  isRequired={true}
            // disabled={false}
          //  selectedItems={this._getPeoplePickerItems} 
          />
          <Label>Date Reported</Label>
          <DatePicker
            // firstDayOfWeek={firstDayOfWeek}
            value={this.state.DateReported}
            // onChange={this.handleDateChange}
            placeholder="Select a date..."
            ariaLabel="Select a date"
     
            onSelectDate={ date => this.setState({ DateReported: date }) }
            // DatePicker uses English strings by default. For localized apps, you must override this prop.
            strings={defaultDatePickerStrings}
          />
          <Label>Issue Source</Label>
          <TextField ></TextField>
          <Label>IssueLoggedBy</Label>
          <PeoplePicker
            context={this.props.context}
            principalTypes={[PrincipalType.User]}
            // titleText="People Picker"
            personSelectionLimit={1}
            groupName={""} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            //  isRequired={true}
            disabled={false}
          //  selectedItems={this._getPeoplePickerItems} 
          />
          <Label>Images</Label>
          
          <Label>Amount</Label>
          <TextField type="number" defaultValue={this.state.Amount}
          onChange={this.handleAMountChange}
          ></TextField>

          <DefaultButton onClick={() => this.setState({ showForm: false ,hideTable:false})}> Cancel</DefaultButton>
          <PrimaryButton type="submit" >Create Item</PrimaryButton>
        </form>
      </div>
    );
  }
  public _onRenderItemColumn = (item: ISPList, index: number, column: IColumn): JSX.Element | any => {
    const src = item.Images;
    // console.log(column.key);
    // console.log("Heelo world");
    // console.log("item",item)
    if (column.key == "Amount") {
      return item.Amount
    }
    else if (column.key == "Priority") {
      return item.Priority
    }
    else if (column.key == "AssignedTo") {
      return item.Assignedto ? item.Assignedto.Title : ""
    }
    else if (column.key == "Status") {
      return item.Status ? item.Status : ""
    }
    else if (column.key == "Description") {
      return item.Description ? item.Description : ""
    }
    else if (column.key == "Images") {
      return item.Images ? item.Images : ""
    }
    else if (column.key == "DateReported") {
      return item.DateReported ? item.DateReported : ""
    }
    else if (column.key == "AssignedTo") {
      return item.Issueloggedby ? item.Issueloggedby.Title : ""
    }
    //  else if(column.key=="IssueSource")
    //  {
    //   return item.IssueSource?item.IssueSource:""
    //  }
    else {
      return ""
    }
  }
}

// function async(arg0: (item: any) => void): any {
//   throw new Error('Function not implemented.');
// }

