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
  DateReported: string;
  IssueSource: {

  };
  Images: any;
  Issueloggedby: {
    Title: string
  };
  Amount: string;
  ID: number;
  selectedId: number;
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
      Description: '',
      Priority: '',
      Status: '',
      Assignedto: {
        Title: ''
      },
      DateReported: '',
      IssueSource: {},
      Images: '',
      Issueloggedby: {
        Title: ''
      },
      Amount: '0',
      ID: 0,
      // value:'',
      selectionDetails: this._getSelectionDetails(),
      showForm: false,
      selectionMode: SelectionMode.single,
      IsEditEnabled: false,
      DisplaySelectedItem: false,
      getSelectedData: false,
      selectedId: 0
      //  selectionMode:false
    };
    // sp.setup({
    //   spfxContext: this.props.spcontext
    // });
  }


  handleSubmit = (event: any) => {
    event.preventDefault();

    this.setState({ showForm: false })

  }

  public handleChange = (event: any, value: any) => {
    // event.preventDefault();
    return this.setState({ Description: value });

  }
  public handlePriorityChange = (event: any) => {
    // event.preventDefault();
    this.setState({ Priority: event.target.value })
    console.log("Priority", this.state.Priority)
  }
  public handleStatusChange = (event: any) => {
    // event.preventDefault();
    this.setState({ Status: event.target.value })
    console.log("Priority", this.state.Priority)
  }
  // public handleAssignedToChange(event:any){

  //   this.setState({:event.target.value})
  // }
  public handleDateChange = (event: any) => {
    // event.preventDefault();
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
    //  this.state.getSelectedData?this.getSelectedItemData:""
  }


  public render(): React.ReactElement<IIssurTrackerProps> {

    const { selectionMode, selectionDetails } = this.state
    return (

      <div><h1>Display SharePoint list data using spfx</h1>

        <PrimaryButton text="New Item" onClick={() => this.setState({ showForm: true })} />
        <DefaultButton text="Display" onClick={() => this.setState({ DisplaySelectedItem: true })} />
        <PrimaryButton text="Edit" onClick={() => this.setState({ IsEditEnabled: true })} />"

        {this.state.showForm ? this.showForm() : ""}
        {this.state.DisplaySelectedItem ? this.getFormToDisplay() : console.log(this.state.DisplaySelectedItem)}
        <h1>{this.state.selectionDetails}</h1>
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
        <h3>Note : Image is clickable.</h3>




      </div>
    );

  }

  private _getSelectionDetails(): any {
    const selectionCount = this._selection.getSelectedCount();
    let SelectedItemId;
    if (selectionCount == 0) {
      return 'No items selected';

    }
    else {

      return (
        SelectedItemId = (this._selection.getSelection()[0] as IDetailsListState).ID,
        this.setState({ selectedId: SelectedItemId })




      )
    }

  }


  public async getSelectedItemData() {
    const data: ISPList[] = [];

    const id = this.state.selectedId
    console.log("Id", this.state.selectedId)

    const allListItems: ISPList[] = await this.web.lists.getByTitle("Issue tracker").items.getById(id).select("ID", "Description", "Priority", "Status", "Assignedto/Title", "DateReported", "IssueSource", "Images", "Issueloggedby/Title", "Amount").expand("Assignedto", "Issueloggedby").get();
    console.log("selecteditem", allListItems)
    console.table(allListItems)
    this.setState({ SelectedItems: { ...allListItems } })
    // this.setState({Description:allListItems.})
    //   // allListItems.
    // this.setState()


    // this.setState({SelectedItems:allListItems})
    // //  await allListItems.map(async (item:any)=> {
    //    await data.push({
    //       Description: allListItems.Description,
    //       Priority: allListItems.Priority,
    //       Status: item.Status,
    //       Assignedto:item.Assignedto,
    //       DateReported:moment(item.DateReported).format("DD-MM-YYYY") ,
    //       IssueSource: item.IssueSource,
    //       Images: item.Images,
    //       Issueloggedby:item.Issueloggedby,          
    //       Amount:item.Amount,
    //       ID:item.ID
    //     });
    // })

    console.log(data);
    console.log("seItems", this.state.SelectedItems)
    this.setState({ SelectedItems: data }),
      alert(this.state.SelectedItems)
  }


  private _onItemInvoked = (item: IDetailsListState): void => {
    alert(`Item invoked: ${item.ID}`);
    // this.geSelectedItemData()

  };


  getFormToDisplay() {

    this.getSelectedItemData(),
      console.log("id", this.state.selectedId);
    const items = this.state.SelectedItems;
    // items.map( item => {

    items.map((item, key) => {
      return (
        <div >
          <h1>{item.Description}</h1>
        </div>

      )
    }
    )
    // alert(items.map(item=>
    //   item[Description]

    // ))
    // <form>
    //   <Label>Description</Label>
    //   <p {item.Description?item.Description:''}

    //     <Label>Priority </Label>
    //     {item.Priority?item.Priority:''}

    //     <Label>Status </Label>
    //     {item.Status?item.Status:""} 

    //       <Label>Assigned To</Label>
    //     {item.Assignedto?item.Assignedto.Title:''}
    //       <Label>Date Reported</Label>
    //      {item.DateReported?item.DateReported:""}

    //     <Label>Issue Source</Label>
    //     {item.IssueSource?item.IssueSource:''} 
    //     <Label>IssueLoggedBy</Label>
    //     {item.Issueloggedby?item.Issueloggedby.Title:""} 
    //     <Label>Images</Label>
    //     {item.Images?item.Images:""} 
    //     <Label>Amount</Label>
    //     {item.Amount?item.Amount:""} 
    //     <DefaultButton onClick={()=>this.setState({DisplaySelectedItem:false,selectedId:0})}> Cancel</DefaultButton> 

    // </form>



    // })




  }

  getFormToEdit() {
    return (

      <div>

        <form id="add-app">

          <Label>Description</Label>
          {/* <input type="text"> </input> */}
          <TextField value={this.state.Description}></TextField>



          <Label>Priority </Label>
          <Dropdown
            placeholder="Select an option"
            label="Dropdown with error message"
            options={PriorityOptions}
            onSelect={() => { this.state.Priority }}
            //  value={this.state.}
            //  onChange={this.onChange}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
            styles={dropdownStyles} />

          <Label>Status </Label>
          <Dropdown
            placeholder="Select an option"
            label="Dropdown with error message"
            options={StatusOptions}
            onSelect={() => { this.state.Status }}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
            styles={dropdownStyles} />

          <Label>Assigned To</Label>
          <PeoplePicker
            context={this.props.context}
            titleText="People Picker"
            personSelectionLimit={1}
            groupName={"Team Site Owners"} // Leave this blank in case you want to filter from all users
            showtooltip={true}
            //  isRequired={true}
            disabled={false}
          //  selectedItems={this._getPeoplePickerItems} 
          />
          <Label>Date Reported</Label>
          <DatePicker
            //  value=this.state.DateReported}
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

          <DefaultButton onClick={() => this.setState({ IsEditEnabled: false })}> Cancel</DefaultButton>
          <PrimaryButton onClick={this.updateItem}>Create Item</PrimaryButton>
        </form>
      </div>
    )

  }

  public async createItem() {
    // return  alert("Added Items")


    await sp.web.lists.getByTitle('Issue tracker').items.add({
      Description: this.state.Description,
      Priority: this.state.Priority,
      DateReported: this.state.DateReported,
      Amount: this.state.Amount,
      Assignedto: this.state.Assignedto.Title
    });
    // alert(this.state.Description)
    // sp.web.lists.getByTitle("Issue Tracker").add({
  }
  // })

  public updateItem() {
    return console.log("Added Items")
    // sp.web.lists.getByTitle("Issue Tracker").add({

    // })
  }
  //   private _getSelectionDetails(): any {
  //     const selectionCount = this._selection.getSelectedCount();
  //     const selectedId = this._selection.getItems();
  //     console.log(selectionCount)
  //     console.log("selectedId",selectedId)


  // // let Items=[...this.state.Items]
  // //     this.setState({Items});

  // //   console.log("updates:",this.state.Items)
  //    }

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

  public showForm = () => {
    return (

      <div>
        {alert(this.state.showForm)}
        <form id="add-app" onSubmit={this.handleSubmit}>

          <Label>Description</Label>
          {/* <input type="text"> </input> */}
          <TextField defaultValue={this.state.Description}
          // onChange={(e)=>{this.handleChange(e,this.state.Description)}}
          >

          </TextField>

          {/* <Label>Title </Label>
          <TextField></TextField> */}

          <Label>Priority </Label>
          <Dropdown
            placeholder="Select an option"
            //  label="Dropdown with error message"
            defaultValue={this.state.Priority}
            options={PriorityOptions}
            onSelect={this.handlePriorityChange}
            //  onChange={this.handlePriorityChange}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
            styles={dropdownStyles} />

          <Label>Status </Label>
          <Dropdown
            placeholder="Select an option"
            //  label="Dropdown with error message"
            defaultValue={this.state.Status}
            //  onChange={this.handleStatusChange}
            options={StatusOptions}
            //  errorMessage={showError ? 'This dropdown has an error' : undefined}
            styles={dropdownStyles} />

          <Label>Assigned To</Label>
          {/* <PeoplePicker/> */}
          <Label>Date Reported</Label>
          <DatePicker
            // firstDayOfWeek={firstDayOfWeek}
            defaultValue={this.state.DateReported}
            // onChange={this.handleDateChange}
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
            src={this.state.Images}
            // onChange={this.handleImageChange}
            alt="Example with no image fit value and height or width is specified."
            width={100}
            height={100}
          />
          <Label>Amount</Label>
          <TextField type="number" value={this.state.Amount}
          // onChange={this.handleAMountChange}
          ></TextField>

          <DefaultButton onClick={() => this.setState({ showForm: false })}> Cancel</DefaultButton>
          <PrimaryButton onClick={() => this.createItem} >Create Item</PrimaryButton>
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

function async(arg0: (item: any) => void): any {
  throw new Error('Function not implemented.');
}

