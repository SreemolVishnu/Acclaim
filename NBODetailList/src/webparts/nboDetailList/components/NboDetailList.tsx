import * as React from 'react';
import styles from './NboDetailList.module.scss';
import { INboDetailListProps } from './INboDetailListProps';
import { DetailsList, Fabric, Selection, IColumn, DetailsListLayoutMode, Link, IconButton, IIconProps, TextField, Dropdown, IDropdownOption, DatePicker, PrimaryButton, DefaultButton, Dialog, DialogFooter, DialogType, Modal, Panel, Label, FontWeights, mergeStyleSets, getTheme, CommandBarButton, PanelType, Callout, IContextualMenuProps, CommandButton, MessageBar, SearchBox } from 'office-ui-fabric-react';
import { sp, Web, IAttachmentFileInfo, Items } from "@pnp/sp/presets/all";
import { forEach, isNumber } from 'lodash';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { MSGraphClient, IHttpClientOptions } from '@microsoft/sp-http';
import * as _ from 'lodash';
import { Pagination } from '@pnp/spfx-controls-react/lib/pagination';
import * as moment from 'moment';
import SimpleReactValidator from 'simple-react-validator';
import { ToastContainer, toast } from 'react-toastify';
import 'react-toastify/dist/ReactToastify.css';
import replaceString from 'replace-string';
//for msg bar
export interface IMessage {
  isShowMessage: boolean;
  messageType: number;
  message: string;
}
export interface INboDetailListState {
  statusMessage: IMessage;
  docRepositoryItems: any[];
  selectionDetails: string;
  Items: IDocument[];
  items: any[];
  AddNBO: string;
  hideDialog: boolean;
  currentItemID: any;
  showReviewModal: boolean;
  showReviewModalFromMailView: boolean;
  clientName: string;
  source: string;
  NB0StageText: string;
  complianceCleared: string;
  industry: string;
  industryKey: any;
  deleteMessageBar: string;
  classOfInsurance: any;
  estimatedPremium: any;
  brokerage: any;
  estimatedBrokerage: any;
  brokerageAmount: any;
  estimatedStartDate: any;
  feesIfAny: any;
  comments: any;
  nboItemWithoutFilter: any[];
  teamTypeArray: any[];
  paginatedItems: any[];
  //dropDownArrays
  sourceItems: any[];
  brokeragePercentageItems: any[];
  industryItems: any[];
  classOfInsuranceItems: any[];
  NBOStageItems: any[];
  brokerageKey: any;
  complianceClearedKey: any;
  classOfInsuranceKey: any;
  NB0StageKey: any;
  sourceKey: any;
  //additional
  tempArrayForExternalDocumentGrid: any[];
  externalArray: any[];
  externalArrayDiv: string;
  showDocInPanel: any;
  //message bar 
  messageBar: string;
  groupItems: any[];
  groupName: string;
  groupNamekey: any;
  departmentText: string;
  departmentkey: any;
  divForOtherDepts: string;
  divForSame: string;
  noItemErrorMsg: string;
  //from mail
  displayFromMail: string;
  displayWithOutQuery: string;
  oppurtunityDept: any[];
  divForDocumentUploadCompliance: string;
  divForDocumentUploadCompArrayDiv: string;
  editAuthorEmail: string;
  sameDepartmentItems: string;
  divForCurrentUser: string;
  confirmDialog: boolean;
  itemIDForDelete: number;
  isNBOAdmin: string;
  oppurtunityTypeKey: string;
  oppurtunityType: string;
  sortTypeAsc: string;
  sortTypeDesc: string;
  sortOppurtunityTypeAsc: string;
  sortOppurtunityTypeDesc: string;
  SourceTypeAsc: string;
  IndustryTypeAsc: string;
  ClientNameTypeAsc: string;
  ClassOfInsuranceTypeAsc: string;
  EstStartDateTypeAsc: string;
  CommentsTypeAsc: string;
  EstimatedPremiumTypeAsc: string;
  BrokerageTypeAsc: string;
  EstimatedBrokerageTypeAsc: string;
  NBOStageTypeAsc: string;
  WeightedBrokerageTypeAsc: string;
  FeesIfAnyTypeAsc: string;
  ComplianceClearedTypeAsc: string;
  DepartmentTypeAsc: string;
  CreatedByTypeAsc: string;
  SourceTypeDesc: string;
  IndustryTypeDesc: string;
  ClientNameTypeDesc: string;
  ClassOfInsuranceTypeDesc: string;
  EstStartDateTypeDesc: string;
  CommentsTypeDesc: string;
  EstimatedPremiumTypeDesc: string;
  BrokerageTypeDesc: string;
  EstimatedBrokerageTypeDesc: string;
  NBOStageTypeDesc: string;
  WeightedBrokerageTypeDesc: string;
  FeesIfAnyTypeDesc: string;
  ComplianceClearedTypeDesc: string;
  DepartmentTypeDesc: string;
  CreatedByTypeDesc: string;
  hideFilterDialog: boolean;
  currentFilterItem: string;
  filterItems: any[];
  currentPage: number;
  arrayForShowingPagination: any[];
  forOtherDeptFilter: string;
  selectedColumnKey: any;
  filterConditions: any[];
  textFiledForFilter: string;
  filterCondition: string;
  filterConditionKey: string;
  filterValue: any;
  dateForFilter: string;
  filterConditionDiv: string;
  divForShowingPagination: string;
  divForNoDataFound: string;
  estimatedFromStartDate: Date;
  estimatedToStartDate: Date;
}
export interface IDocument {
  Title: string;
  field_1: string;
  Edit: any;

}

const EditIcon: IIconProps = { iconName: 'Edit' };
const SortAcsIcon: IIconProps = { iconName: 'SortLinesAscending' };
const SortDescIcon: IIconProps = { iconName: 'SortLines' };
const FilterIcon: IIconProps = { iconName: 'Filter' };
const DeleteIcon: IIconProps = { iconName: 'Delete' };
const AddIcon: IIconProps = { iconName: 'Add' };
const CancelIcon: IIconProps = { iconName: 'Cancel' };
const theme = getTheme();
const contentStyles = mergeStyleSets({
  container: {
    display: 'flex',
    flexFlow: 'column nowrap',
    // alignItems: 'stretch',
  },



});
const iconButtonStyles = {
  root: {
    //color: theme.palette.neutralPrimary,
    color: "White",
    marginLeft: 'auto',
    marginTop: '4px',
    marginRight: '2px',
  },
  rootHovered: {
    color: theme.palette.neutralDark,
  },
};
export default class NboDetailList extends React.Component<INboDetailListProps, INboDetailListState, {}> {

  private validator: SimpleReactValidator;
  private _columns: IColumn[];
  private myfileadditional;
  private currentUserEmail;
  private currentUserName;
  private _selection: Selection;
  private team;
  private itemDepartment;
  private teamType = "Department Team";
  private nbolid;
  private content;
  private compliance;
  private sortedArray = [];
  private pageSize: number = 30;
  private forDeptCreatedBy;
  //private teamTypeArray: any[];
  private reqWeb = Web(window.location.protocol + "//" + window.location.hostname + "/sites/Acclaim");
  constructor(props: INboDetailListProps) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      docRepositoryItems: [],
      selectionDetails: "",
      Items: [],
      items: [],
      AddNBO: "none",
      hideDialog: false,
      currentItemID: "",
      deleteMessageBar: "none",
      showReviewModal: false,
      showReviewModalFromMailView: false,
      clientName: "",
      source: "",
      NB0StageText: "",
      complianceCleared: "",
      industry: "",
      classOfInsurance: "",
      estimatedPremium: "35000",
      brokerage: "15",
      estimatedBrokerage: "",
      brokerageAmount: "",
      estimatedStartDate: "",
      feesIfAny: "0",
      comments: "",
      nboItemWithoutFilter: [],
      teamTypeArray: [],
      paginatedItems: [],
      //dropDownArrays
      sourceItems: [],
      brokeragePercentageItems: [],
      industryItems: [],
      classOfInsuranceItems: [],
      NBOStageItems: [],
      brokerageKey: "",
      complianceClearedKey: "",
      classOfInsuranceKey: "",
      NB0StageKey: "",
      sourceKey: "",
      industryKey: "",
      //additional
      tempArrayForExternalDocumentGrid: [],
      externalArray: [],
      externalArrayDiv: "none",
      showDocInPanel: "none",
      //messageBar
      messageBar: "none",
      groupItems: [],
      groupName: "",
      groupNamekey: "",
      divForOtherDepts: "none",
      divForSame: "",
      noItemErrorMsg: "none",
      displayFromMail: "none",
      displayWithOutQuery: "",
      oppurtunityDept: [],
      divForDocumentUploadCompliance: "none",
      divForDocumentUploadCompArrayDiv: "none",
      editAuthorEmail: "none",
      sameDepartmentItems: "No",
      divForCurrentUser: "none",
      confirmDialog: true,
      itemIDForDelete: null,
      isNBOAdmin: "",
      departmentText: "",
      departmentkey: "",
      oppurtunityTypeKey: "",
      oppurtunityType: "",
      sortTypeAsc: "",
      sortTypeDesc: "none",
      sortOppurtunityTypeAsc: "",
      sortOppurtunityTypeDesc: "none",
      SourceTypeAsc: "",
      IndustryTypeAsc: "",
      ClientNameTypeAsc: "",
      ClassOfInsuranceTypeAsc: "",
      EstStartDateTypeAsc: "",
      CommentsTypeAsc: "",
      EstimatedPremiumTypeAsc: "",
      BrokerageTypeAsc: "",
      EstimatedBrokerageTypeAsc: "",
      NBOStageTypeAsc: "",
      WeightedBrokerageTypeAsc: "",
      FeesIfAnyTypeAsc: "",
      ComplianceClearedTypeAsc: "",
      DepartmentTypeAsc: "",
      CreatedByTypeAsc: "",
      SourceTypeDesc: "none",
      IndustryTypeDesc: "none",
      ClientNameTypeDesc: "none",
      ClassOfInsuranceTypeDesc: "none",
      EstStartDateTypeDesc: "none",
      CommentsTypeDesc: "none",
      EstimatedPremiumTypeDesc: "none",
      BrokerageTypeDesc: "none",
      EstimatedBrokerageTypeDesc: "none",
      NBOStageTypeDesc: "none",
      WeightedBrokerageTypeDesc: "none",
      FeesIfAnyTypeDesc: "none",
      ComplianceClearedTypeDesc: "none",
      DepartmentTypeDesc: "none",
      CreatedByTypeDesc: "none",
      hideFilterDialog: false,
      currentFilterItem: "",
      filterItems: [],
      currentPage: 1,
      arrayForShowingPagination: [],
      forOtherDeptFilter: "",
      selectedColumnKey: "",
      filterConditions: [],
      textFiledForFilter: "none",
      filterCondition: "",
      filterConditionKey: "",
      filterValue: "",
      dateForFilter: "none",
      filterConditionDiv: "",
      divForShowingPagination: "none",
      divForNoDataFound: "none",
      estimatedFromStartDate: null,
      estimatedToStartDate: null,
    };

    this._drpdwnChangeSource = this._drpdwnChangeSource.bind(this);
    this.clientNameChange = this.clientNameChange.bind(this);
    this._submitNBOPipeline = this._submitNBOPipeline.bind(this);
    this._updateNBOPipeline = this._updateNBOPipeline.bind(this);
    this._drpdwnNBOStage = this._drpdwnNBOStage.bind(this);
    this._drpdwnComplianceCleared = this._drpdwnComplianceCleared.bind(this);
    this.loadDocProfile = this.loadDocProfile.bind(this);
    this._drpdwnIndustry = this._drpdwnIndustry.bind(this);
    this._drpdwnBrokerage = this._drpdwnBrokerage.bind(this);
    this._drpdwnClassOfInsurance = this._drpdwnClassOfInsurance.bind(this);
    this._feesIfAnyChange = this._feesIfAnyChange.bind(this);
    this._commentsChange = this._commentsChange.bind(this);
    this._estimatedPremiumChange = this._estimatedPremiumChange.bind(this);
    this.checkingcurrentUserDept = this.checkingcurrentUserDept.bind(this);
    this._showExternalGrid = this._showExternalGrid.bind(this);
    this._getPage = this._getPage.bind(this);
    this.onEditClick = this.onEditClick.bind(this);
    this.onDeleteClick = this.onDeleteClick.bind(this);
    //dropDownBindings
    this._drpdwnSource = this._drpdwnSource.bind(this);
    // this._drpdwnBrokeragePercentage = this._drpdwnBrokeragePercentage.bind(this);
    this._drpdwnClassOfInsuranceBind = this._drpdwnClassOfInsuranceBind.bind(this);
    this._drpdwnIndustryBind = this._drpdwnIndustryBind.bind(this);
    this._drpdwnNBOStageBind = this._drpdwnNBOStageBind.bind(this);
    this.openCCSPopUp = this.openCCSPopUp.bind(this);
    this._drpdwnGroupName = this._drpdwnGroupName.bind(this);
    this._sendAnEmailUsingMSGraph = this._sendAnEmailUsingMSGraph.bind(this);
    this._sendAnEmailForComplianceGroup = this._sendAnEmailForComplianceGroup.bind(this);
    this.GetUserProperties = this.GetUserProperties.bind(this);
    this.others = this.others.bind(this);
    this.sameDept = this.sameDept.bind(this);
    this.updateComplianceFromMail = this.updateComplianceFromMail.bind(this);
    this._showExternalGridForComplianceUpload = this._showExternalGridForComplianceUpload.bind(this);
    this._selectDepartmentFromSameDepartmentTab = this._selectDepartmentFromSameDepartmentTab.bind(this);
    this._drpdwnOppurtunityType = this._drpdwnOppurtunityType.bind(this);
    this._onSortClickAscForMyNBO = this._onSortClickAscForMyNBO.bind(this);
    this._onSortClickDescForMyNBO = this._onSortClickDescForMyNBO.bind(this);
    this._onSortClickAscForSameDept = this._onSortClickAscForSameDept.bind(this);
    this._onSortClickDescForSameDept = this._onSortClickDescForSameDept.bind(this);
    this._onSortClickAscForOtherDept = this._onSortClickAscForOtherDept.bind(this);
    this._onSortClickDescForOtherDept = this._onSortClickDescForOtherDept.bind(this);
    this._filterPanelCloseButton = this._filterPanelCloseButton.bind(this);
    this._onFilter = this._onFilter.bind(this);
    this.filterColumnChange = this.filterColumnChange.bind(this);
    this.filterConditionColumnChange = this.filterConditionColumnChange.bind(this);
    this._onFilterButtonSubmit = this._onFilterButtonSubmit.bind(this);
    this.filterValueChange = this.filterValueChange.bind(this);
    this._onFilterForModal = this._onFilterForModal.bind(this);
    this._estimateFromDateChange = this._estimateFromDateChange.bind(this);
    this._estimateToDateChange = this._estimateToDateChange.bind(this);
  }
  public componentWillMount = async () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields"
      }
    });

  }
  public async componentDidMount() {
    this.checkingcurrentUserDept();
    //Query getting...
    let params = new URLSearchParams(window.location.search);
    this.nbolid = params.get('nbolid');
    this.compliance = params.get('ViewComp');
    console.log("transmittalID", this.nbolid);
    if (this.nbolid != "" && this.nbolid != null) {
      this.setState({
        displayWithOutQuery: "none",
        displayFromMail: "",
        showReviewModalFromMailView: true,
      });

      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
        select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("ID eq '" + this.nbolid + "'").get().then(docProfileItems => {
          this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
          this.setState({
            docRepositoryItems: this.sortedArray,
            items: this.sortedArray,
            paginatedItems: this.sortedArray.slice(0, this.pageSize),
            noItemErrorMsg: docProfileItems.length == 0 ? " " : "none",
          });
          console.log(this.state.docRepositoryItems);
          if (docProfileItems.length == 0) {
            this.setState({ noItemErrorMsg: "" });
          }
        });
      this.setState({
        divForSame: "",
        divForOtherDepts: "none",
      });

    }
    else {
      this.setState({
        displayWithOutQuery: "",
        displayFromMail: "none"
      });
      await this.loadDocProfile();
      await this.GetUserProperties();
    }

    await sp.web.currentUser.get().then(currentUser => {
      console.log(currentUser);
      this.currentUserEmail = currentUser.Email;
      this.currentUserName = currentUser.Title;
    });
    // await this.GetUserProperties();
    //dropdownbinding
    this._drpdwnSource();
    // this._drpdwnBrokeragePercentage();
    this._drpdwnClassOfInsuranceBind();
    this._drpdwnIndustryBind();

  }
  private async checkingcurrentUserDept() {

    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.teamList).items.get().then(teamItems => {
      console.log("teamItems", teamItems);
      //getting current user groups
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api(`me/transitiveMemberOf/microsoft.graph.group?$count=true`)
            .get((error, response: any, rawResponse?: any) => {
              console.log(JSON.stringify(response));
              console.log(response.value);
              for (let i = 0; i < response.value.length; i++) {

                if (response.value[i].displayName == "Security Group - NBO Admin") {
                  //alert(response.value[i].displayName);
                  this.setState({
                    isNBOAdmin: "true",
                  });
                }
              }

            });
        });
    });
  }
  //checkin current user groups and getting team Type and department 
  private async GetUserProperties() {
    let tempTeamTypeArray = [];
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.teamList).items.get().then(teamItems => {
      console.log("teamItems", teamItems);
      //getting current user groups
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api(`me/transitiveMemberOf/microsoft.graph.group?$count=true`)
            .get((error, response: any, rawResponse?: any) => {
              console.log(JSON.stringify(response));
              console.log(response.value);
              for (let i = 0; i < response.value.length; i++) {
                //groupName binding in dropDown
                let teamTypeArrayItems = {
                  key: i,
                  text: response.value[i].displayName,
                };
                tempTeamTypeArray.push(teamTypeArrayItems);
                this.setState({
                  groupItems: tempTeamTypeArray,
                });
                //checking from team list
                for (let j = 0; j < teamItems.length; j++) {
                  //  alert(response.value[i].displayName) 
                  //dropdown binding opputunity department                 
                  if (teamItems[j].Title == response.value[i].displayName) {
                    // alert("teams" + teamItems[j].Title);
                    // alert("response" + response.value[i].displayName);
                    if (response.value[i].displayName != "Security Group - NBO Admin") {
                      console.log(response.value[i].displayName, "--", response.value[i].id);
                      console.log(teamItems[j].TeamType);
                      let teamItemTitle = teamItems[j].Title.split('- ');
                      let tempOppurtunityItems = {
                        key: j,
                        text: teamItemTitle[1],
                      };
                      this.state.oppurtunityDept.push(tempOppurtunityItems);
                      this.state.teamTypeArray.push({ teamtype: teamItems[j].TeamType, team: teamItems[j].Title });
                    }


                  }
                }
              }
              console.log(this.state.teamTypeArray);
              for (let k = 0; k < this.state.teamTypeArray.length; k++) {
                if (this.state.teamTypeArray[k].teamtype == "NBO Admin Team" || this.state.teamTypeArray[k].teamtype == "Compliance Team") {
                  this.teamType = this.state.teamTypeArray[k].teamtype;
                  let teamItemTitle = this.state.teamTypeArray[k].team.split('- ');
                  this.team = teamItemTitle[1];
                  console.log("team", this.team);
                  console.log("teamType", this.teamType);
                  // alert(this.teamType);
                  break;
                }
                else {
                  let teamItemTitle = this.state.teamTypeArray[k].team.split('- ');
                  this.teamType = this.state.teamTypeArray[k].teamtype;
                  this.team = teamItemTitle[1];

                }
              }

              this.loadDocProfile();
              //alert(this.team);
            });
        });
    });
  }
  //user profiles 
  public _gettingUserProfiles() {
    //user profile items for manager email
    sp.profiles.myProperties.get().then(function (result) {
      var userProperties = result.UserProfileProperties;
      var userPropertyValues = "";
      let email = [];
      forEach(function (property) {
        userPropertyValues += property.Key + " - " + property.Value + "<br/>";
      });

      //document.getElementById("spUserProfileProperties").innerHTML = userPropertyValues;
      console.log("userProperties", userProperties);
      for (let k = 0; k < userProperties.length; k++) {
        if (userProperties[k].Key == "Manager") {
          console.log(userProperties[k].Key, userProperties[k].Value);
          email = userProperties[k].Value.split('i:0#.f|membership|');
          console.log(email[1]);
          console.log(email[1]);
          this.setState({ LineManagerEmail: email[1] });
          this.emailOfLineManager = email[1];
          console.log(this.emailOfLineManager);
        }

      }
    }).catch(function (error) {
      console.log("Error: " + error);
    });
  }
  //getting items from NBOPipeline List
  private loadDocProfile = () => {
    //alert(this.team)
    this.forDeptCreatedBy = "";
    let docProfileItems = [];
    this.setState({
      sameDepartmentItems: "no",
      currentItemID: "",
      forOtherDeptFilter: "MYNBOSame",
      filterCondition: "",
      filterValue: "",
      selectedColumnKey: "",
      filterConditions: [],
      divForNoDataFound: "none"
    });
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
      select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail eq '" + this.currentUserEmail + "'")
      .top(4000).getPaged()
      .then(async docItems => {
        this.setState({ arrayForShowingPagination: docItems.results });
        for (let i = 0; i < this.pageSize; i++) {
          docProfileItems.push({
            "ID": null,
            "Title": null,
            "BrokeragePercentage": null,
            "Source": { "ID": null, "Title": null },
            "Industry": { "ID": null, "Title": null },
            "ClassOfInsurance": { "ID": null, "Title": null },
            "NBOStage": { "ID": null, "Title": null },
            "EstimatedBrokerage": null,
            "FeesIfAny": null,
            "Comments": null,
            "EstimatedStartDate": null,
            "EstimatedPremium": null,
            "Department": null,
            "ComplianceCleared": null,
            "Author": { "EMail": null, "Title": null },
            "WeightedBrokerage": null,
            "OpportunityType": null,
          });
        }
        docProfileItems = docProfileItems.concat(docItems.results);
        console.log(docProfileItems);
        while (docItems.hasNext) {
          docItems = await docItems.getNext();
          docProfileItems.push(...(docItems.results));
        }

        this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
        this.setState({
          docRepositoryItems: this.sortedArray,
          items: this.sortedArray,
          paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
          divForShowingPagination: "",
        });
        console.log(this.state.docRepositoryItems);

      });
    this.setState({
      divForSame: "none",
      divForCurrentUser: "",
      divForOtherDepts: "none",
      divForDocumentUploadCompArrayDiv: "none",
    });
  }

  // sending Email for managers
  private async _sendAnEmailUsingMSGraph(nbolid): Promise<void> {
    let email = [];
    await sp.profiles.myProperties.get().then(function (result) {
      var userProperties = result.UserProfileProperties;
      var userPropertyValues = "";
      forEach(function (property) {
        userPropertyValues += property.Key + " - " + property.Value + "<br/>";
      });
      //document.getElementById("spUserProfileProperties").innerHTML = userPropertyValues;
      console.log("userProperties", userProperties);
      for (let k = 0; k < userProperties.length; k++) {
        if (userProperties[k].Key == "Manager") {
          console.log(userProperties[k].Key, userProperties[k].Value);
          email = userProperties[k].Value.split('i:0#.f|membership|');
          console.log(email[1]);
        }
      }
    }).catch(function (error) {
      console.log("Error: " + error);
    });
    console.log("email of manager inside mail function", email[1]);
    //alert("emailtoOwner");
    const user = await sp.web.siteUsers.getByEmail(email[1])();
    console.log("user", user.Title);
    let Subject;
    let Body;
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNotificationSettings).items.filter("Title eq 'NBO'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    let link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/NboPipeline.aspx?nbolid=" + nbolid}>NBOPipeline</a>`;
    //Replacing the email body with current values
    let replacedSubject1 = replaceString(Subject, '[NBOTitle]', this.state.clientName);
    let replaceBodyStaff = replaceString(Body, '[staff]', this.currentUserName);
    let replaceBody = replaceString(replaceBodyStaff, '[ProspectName]', this.state.clientName);
    let replaceBodyWithManagerName = replaceString(replaceBody, '[ManagerName]', user.Title);
    let replacelink = replaceString(replaceBodyWithManagerName, '[NBOPipeline]', link);
    let FinalBody1 = replacelink;
    //mail sending
    //  if (this.status == "Yes") {
    //Check if TextField value is empty or not  
    if (email[1]) {
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject1,
          "body": {
            "contentType": "HTML",
            "content": FinalBody1
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": email[1],

              }
            }
          ],
        }
      };
      //Send Email uisng MS Graph  
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody, (error, response: any, rawResponse?: any) => {
            });
        });
    }
    // }
  }
  // sending Email for owners
  private async _sendAnEmailForComplianceGroup(nbolid): Promise<void> {
    //alert("emailtoOwner");
    let Subject;
    let Body;
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNotificationSettings).items.filter("Title eq 'NBOCompliance'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    let link = `<a href=${window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/NboPipeline.aspx?nbolid=" + nbolid + "&ViewComp=Yes"}>NBOPipeline</a>`;
    //Replacing the email body with current values
    let replacedSubject1 = replaceString(Subject, '[NBOTitle]', this.state.clientName);
    let replaceBody = replaceString(Body, '[NBOTitle]', this.state.clientName);
    let replacelink = replaceString(replaceBody, '[NBOPipeline]', link);
    let FinalBody = replacelink;
    //mail sending
    //  if (this.status == "Yes") {
    //Check if TextField value is empty or not  
    if (this.props.complianceGroupEmail) {
      //Create Body for Email  
      let emailPostBody: any = {
        "message": {
          "subject": replacedSubject1,
          "body": {
            "contentType": "HTML",
            "content": FinalBody
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": this.props.complianceGroupEmail,
              }
            }
          ],

        }
      };
      //Send Email uisng MS Graph  
      this.props.context.msGraphClientFactory
        .getClient()
        .then((client: MSGraphClient): void => {
          client
            .api('/me/sendMail')
            .post(emailPostBody, (error, response: any, rawResponse?: any) => {
            });
        });
    }
    // }
  }
  // ---------------ItemInvoked---------------------
  private async onEditClick(item) {
    //alert(item.ID);
    console.log("edit Item", item.ID);
    console.log("edit Item", item.Author.EMail);
    //alert(this.teamType)
    this.setState({
      hideDialog: true,
      currentItemID: item.ID,
      externalArray: [],
      tempArrayForExternalDocumentGrid: [],
      editAuthorEmail: item.Author.EMail
    });
    //panel rebinding
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(parseInt(item.ID)).select("Title,Source/Title,Industry/Title,BrokeragePercentage,ClassOfInsurance/Title,NBOStage/Title,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared").expand("Source,Industry,ClassOfInsurance,NBOStage").get().then(docProfileItems => {
      this.setState({
        items: docProfileItems,
        clientName: docProfileItems.Title,
        source: docProfileItems.Source.Title,
        sourceKey: docProfileItems.Source.ID,
        NB0StageText: docProfileItems.NBOStage.Title,
        NB0StageKey: docProfileItems.NBOStage.ID,
        complianceCleared: docProfileItems.ComplianceCleared,
        industry: docProfileItems.Industry.Title,
        industryKey: docProfileItems.Industry.ID,
        classOfInsurance: docProfileItems.ClassOfInsurance.Title,
        classOfInsuranceKey: docProfileItems.ClassOfInsurance.ID,
        brokerageAmount: docProfileItems.BrokerageAmount,
        estimatedPremium: docProfileItems.EstimatedPremium,
        feesIfAny: docProfileItems.FeesIfAny,
        comments: docProfileItems.Comments,
        brokerage: docProfileItems.BrokeragePercentage,
        //brokerageKey: docProfileItems.BrokeragePercentage.ID,
        estimatedBrokerage: docProfileItems.EstimatedBrokerage,
        estimatedStartDate: new Date(docProfileItems.EstimatedStartDate),
      });
      console.log(this.state.docRepositoryItems);
      if (this.teamType == "Compliance Team" && docProfileItems.ComplianceCleared == "Yes") {
        //alert(docProfileItems.ComplianceCleared)
        this.setState({
          divForDocumentUploadCompliance: "",
        });
      }
    }).then(forDropDownbinding => {
      //alert(forDropDownbinding[0].ComplianceCleared);

      //checkin compliance cleared and setting nbo stage according to that.
      if (this.state.complianceCleared == "Yes") {
        //alert(this.teamType);
        this._drpdwnNBOStageBind();
      }
      else {
        let tempNBOStage = [];
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboStage).items.filter("Title ne '100% - won'").get().then(NBOStage => {
          console.log("NBOStage", NBOStage);
          for (let i = 0; i < NBOStage.length; i++) {
            // if(subcontractor[i].Active == true){
            let NBOStageItemdata = {
              key: NBOStage[i].ID,
              text: NBOStage[i].Title
            };
            tempNBOStage.push(NBOStageItemdata);
            //}       
          }
          this.setState({
            NBOStageItems: tempNBOStage
          });
        });
      }

    });
  }
  private async onDeleteClick(item) {
    //alert(item.ID);
    console.log("edit Item", item);
    console.log(item.ID);

    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(parseInt(item.ID)).select("Title,Source/Title,Industry/Title,BrokeragePercentage,ClassOfInsurance/Title,NBOStage/Title,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared").expand("Source,Industry,ClassOfInsurance,NBOStage").get().then(docProfileItems => {

    });
    this.setState({
      confirmDialog: false,
      itemIDForDelete: item.ID,
    });
  }
  //edit click from grid from mail link
  private onEditClickFromMail(item) {

    this.setState({
      hideDialog: true,
      currentItemID: item,
      externalArray: [],
    });
    //panel rebinding
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(parseInt(item)).select("Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared").expand("Source,Industry,ClassOfInsurance,NBOStage").get().then(docProfileItems => {
      this.setState({
        items: docProfileItems,
        clientName: docProfileItems.Title,
        source: docProfileItems.Source.Title,
        sourceKey: docProfileItems.Source.ID,
        NB0StageText: docProfileItems.NBOStage.Title,
        NB0StageKey: docProfileItems.NBOStage.ID,
        complianceCleared: docProfileItems.ComplianceCleared,
        industry: docProfileItems.Industry.Title,
        industryKey: docProfileItems.Industry.ID,
        classOfInsurance: docProfileItems.ClassOfInsurance.Title,
        classOfInsuranceKey: docProfileItems.ClassOfInsurance.ID,
        brokerageAmount: docProfileItems.BrokerageAmount,
        estimatedPremium: docProfileItems.EstimatedPremium,
        feesIfAny: docProfileItems.FeesIfAny,
        comments: docProfileItems.Comments,
        brokerage: docProfileItems.BrokeragePercentage,
        //brokerageKey: docProfileItems.BrokeragePercentage.ID,
        estimatedBrokerage: docProfileItems.EstimatedBrokerage,
        estimatedStartDate: new Date(docProfileItems.EstimatedStartDate),
      });
      console.log(this.state.docRepositoryItems);
    }).then(forDropDownbinding => {
      //checkin compliance cleared and setting nbo stage according to that.
      if (this.state.complianceCleared == "Yes") {
        this._drpdwnNBOStageBind();
      }
      else {
        let tempNBOStage = [];
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboStage).items.filter("Title ne '100% - won'").get().then(NBOStage => {
          console.log("NBOStage", NBOStage);
          for (let i = 0; i < NBOStage.length; i++) {
            // if(subcontractor[i].Active == true){
            let NBOStageItemdata = {
              key: NBOStage[i].ID,
              text: NBOStage[i].Title
            };
            tempNBOStage.push(NBOStageItemdata);
            //}       
          }
          this.setState({
            NBOStageItems: tempNBOStage
          });
        });
      }

    });
  }
  private _addNBO = (item) => {

    this.setState({
      tempArrayForExternalDocumentGrid: [],
      AddNBO: "", showReviewModal: true,
      clientName: "",
      source: "",
      NB0StageText: "",
      complianceCleared: "",
      industry: "",
      classOfInsurance: "",
      estimatedBrokerage: "",
      brokerageAmount: "",
      estimatedStartDate: "",
      comments: "",
    });

  }
  //temporary array for external documents grid.
  private async _showExternalGrid() {
    let fileInfos: IAttachmentFileInfo[] = [];
    let input = document.getElementById("newfile") as HTMLInputElement;
    var fileCount = input.files.length;
    var filesize = input.size;
    if ((document.querySelector("#newfile") as HTMLInputElement).files[0] != null) {
      let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
      var docname = myfile.name;
      console.log("UploadedFileDetails", myfile);
      //alert(myfile.size);
      if (myfile.size) {
        this.state.tempArrayForExternalDocumentGrid.push({
          documentName: myfile.name,
          content: myfile
        });
      }
      this.setState({
        externalArray: this.state.tempArrayForExternalDocumentGrid,
        externalArrayDiv: "",
      });
      (document.querySelector("#newfile") as HTMLInputElement).value = null;
    }



  }
  private async _showExternalGridForComplianceUpload() {
    let fileInfos: IAttachmentFileInfo[] = [];
    let input = document.getElementById("newfile") as HTMLInputElement;
    var fileCount = input.files.length;
    var filesize = input.size;
    if ((document.querySelector("#newfile") as HTMLInputElement).files[0] != null) {
      let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
      var docname = myfile.name;
      console.log("UploadedFileDetails", myfile);
      //alert(myfile.size);
      if (myfile.size) {
        this.state.tempArrayForExternalDocumentGrid.push({
          documentName: myfile.name,
          content: myfile
        });
      }
      this.setState({
        externalArray: this.state.tempArrayForExternalDocumentGrid,
        divForDocumentUploadCompArrayDiv: "",
      });
      (document.querySelector("#newfile") as HTMLInputElement).value = null;
    }



  }
  private _renderItemColumn(item: any, index: number, column: IColumn) {
    const fieldContent = item[column.fieldName];
    switch (column.key) {
      case 'Edit':
        return (<div onClick={() => this.onEditClick}><IconButton iconProps={EditIcon} aria-label="Edit" /></div>);
      default:
        return <><span>{item.Title}</span><span>{item.field_1}</span></>;
    }
  }
  private modalProps = {
    isBlocking: true,
  };
  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      hideDialog: false,
      hideFilterDialog: false,
      showReviewModal: false,
      sourceKey: "",
      industryKey: "",
      NB0StageKey: "",
      brokerageKey: "",
      classOfInsuranceKey: "",
      externalArray: [],
      clientName: "",
      complianceClearedKey: "",
      currentItemID: "",
      divForDocumentUploadCompliance: "none",
      divForDocumentUploadCompArrayDiv: "none",
      tempArrayForExternalDocumentGrid: [],
      confirmDialog: true,
      oppurtunityTypeKey: "",
      oppurtunityType: "",
      feesIfAny: "0",

    });

  }

  //For filter panel  box of cancel
  private _filterPanelCloseButton = () => {
    if (this.state.sameDepartmentItems == "no") {
      this.loadDocProfile();
    }
    else if (this.state.sameDepartmentItems == "Yes") {
      this.sameDept();
    }
    else if (this.state.forOtherDeptFilter == "Other") {
      this.others();
    }
    this.setState({
      hideFilterDialog: false,
      selectedColumnKey: "",
      filterCondition: "",
      estimatedStartDate: "",
      filterValue: "",
      filterConditionKey: "",
      dateForFilter: "none",
      textFiledForFilter: "none",
      filterConditionDiv: "",
      divForNoDataFound: "none",
      estimatedFromStartDate: null,
      estimatedToStartDate: null,
    });
  }
  // ---------------SubmitToNBOPipeline---------------------
  private _submitNBOPipeline = async () => {

    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboStage).items.filter("Title eq '10% - prospect identified'").get().then(NBOStageID => {
      this.setState({
        NB0StageKey: NBOStageID[0].ID,
      });
    });
    if (this.state.clientName != "" && this.state.sourceKey != "" && this.state.industryKey != "" && this.state.classOfInsuranceKey != "" && this.state.estimatedStartDate != "" && this.state.oppurtunityType != "") {
      this.validator.hideMessages();
      if (this.state.oppurtunityType == "New") {
        if (this.state.estimatedPremium == " ") {
          this.setState({
            estimatedPremium: "35000",
          });
        }
        toast("Nbo Pipeline added successfully");
        this.setState({
          messageBar: "",
        });
        // alert(this.state.estimatedPremium)
        let tempEstimatedBrokerage = (this.state.estimatedPremium * (this.state.brokerage / 100));
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.add({
          Title: this.state.clientName,
          SourceId: this.state.sourceKey,
          IndustryId: this.state.industryKey,
          ClassOfInsuranceId: this.state.classOfInsuranceKey,
          EstimatedPremium: (this.state.estimatedPremium == "" ? Number(35000) : Number(this.state.estimatedPremium)),
          BrokeragePercentage: (this.state.brokerage == "" ? Number(15) : Number(this.state.brokerage)),
          EstimatedStartDate: new Date(this.state.estimatedStartDate),
          FeesIfAny: Number(this.state.feesIfAny),
          Comments: this.state.comments,
          Department: this.state.groupName,
          // EstimatedBrokerage: ((this.state.estimatedPremium == "" ? Number(35000) : Number(this.state.estimatedPremium))) * Number(this.state.brokerage / 100),
          NBOStageId: parseInt(this.state.NB0StageKey),
          OpportunityType: this.state.oppurtunityType,

        }).then(async afterInsertion => {
          sp.web.getList(this.props.siteUrl + "/Lists/NBOLogList").items.add({
            Title: "New Item ADD",
            NBOPipelineID: Number(afterInsertion.data.ID),
          });
          // this.nbolid = afterInsertion.data.ID;
          if (this.state.externalArray.length > 0) {
            for (var i in this.state.externalArray) {
              var splitted = this.state.externalArray[i].documentName.split(".");
              let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + afterInsertion.data.ID + '.' + splitted[splitted.length - 1];
              let docName = documentNameExtension;
              await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/NBODocuments/").files.add(docName, this.state.externalArray[i].content, true).then(async fileUploaded => {
                const item = await fileUploaded.file.getItem();
                console.log(item);
                console.log(afterInsertion.data.ID);
                await sp.web.getList(this.props.siteUrl + "/NBODocuments/").items.getById(item["ID"]).update({
                  NBOPipelineIdId: afterInsertion.data.ID,
                  Title: this.state.externalArray[i].documentName,
                });


              });
            }
          }
          this._sendAnEmailUsingMSGraph(afterInsertion.data.ID);//mail for manager
          this._sendAnEmailForComplianceGroup(afterInsertion.data.ID);// mail for compliance group
          this.loadDocProfile();
        }).then(afterDocumentInsertion => {
          this.loadDocProfile();
          setTimeout(() => {
            this.setState({
              showReviewModal: false,
              groupNamekey: "",
              sourceKey: "",
              industryKey: "",
              classOfInsuranceKey: "",
              brokerageKey: "",
              externalArrayDiv: "none",
              oppurtunityTypeKey: "",
              oppurtunityType: ""
            });
          }, 7000);
        });
      }
      else {
        // alert(this.state.oppurtunityType)
        if (this.state.estimatedPremium == " ") {
          this.setState({
            estimatedPremium: "35000",
          });
        }
        toast("Nbo Pipeline added successfully");
        this.setState({
          messageBar: "",
        });
        // alert(this.state.estimatedPremium)
        let tempEstimatedBrokerage = (this.state.estimatedPremium * (this.state.brokerage / 100));
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.add({
          Title: this.state.clientName,
          SourceId: this.state.sourceKey,
          IndustryId: this.state.industryKey,
          ClassOfInsuranceId: this.state.classOfInsuranceKey,
          EstimatedPremium: (this.state.estimatedPremium == "" ? Number(35000) : Number(this.state.estimatedPremium)),
          BrokeragePercentage: (this.state.brokerage == "" ? Number(15) : Number(this.state.brokerage)),
          EstimatedStartDate: new Date(this.state.estimatedStartDate),
          FeesIfAny: Number(this.state.feesIfAny),
          Comments: this.state.comments,
          Department: this.state.groupName,
          // EstimatedBrokerage: ((this.state.estimatedPremium == "" ? Number(35000) : Number(this.state.estimatedPremium))) * Number(this.state.brokerage / 100),
          NBOStageId: parseInt(this.state.NB0StageKey),
          ComplianceCleared: "Yes",
          OpportunityType: this.state.oppurtunityType,
        }).then(async afterInsertion => {
          sp.web.getList(this.props.siteUrl + "/Lists/NBOLogList").items.add({
            Title: "New Item ADD",
            NBOPipelineID: Number(afterInsertion.data.ID),
          });
          // this.nbolid = afterInsertion.data.ID;
          if (this.state.externalArray.length > 0) {
            for (var i in this.state.externalArray) {
              var splitted = this.state.externalArray[i].documentName.split(".");
              let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + afterInsertion.data.ID + '.' + splitted[splitted.length - 1];
              let docName = documentNameExtension;
              await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/NBODocuments/").files.add(docName, this.state.externalArray[i].content, true).then(async fileUploaded => {
                const item = await fileUploaded.file.getItem();
                console.log(item);
                console.log(afterInsertion.data.ID);
                await sp.web.getList(this.props.siteUrl + "/NBODocuments/").items.getById(item["ID"]).update({
                  NBOPipelineIdId: afterInsertion.data.ID,
                  Title: this.state.externalArray[i].documentName,
                });

              });
            }
          }
          this._sendAnEmailUsingMSGraph(afterInsertion.data.ID);
          this.loadDocProfile();
        }).then(afterDocumentInsertion => {
          this.loadDocProfile();
          setTimeout(() => {
            this.setState({
              showReviewModal: false,
              groupNamekey: "",
              sourceKey: "",
              industryKey: "",
              classOfInsuranceKey: "",
              brokerageKey: "",
              externalArrayDiv: "none",
              oppurtunityTypeKey: "",
              oppurtunityType: ""
            });
          }, 7000);
        });
      }

    }
    else {
      this.validator.showMessages();
      this.forceUpdate();
    }
  }
  // ---------------UpdateNBOPipeline---------------------
  private _updateNBOPipeline = async () => {
    if (this.nbolid == "" || this.nbolid == null || this.compliance != "Yes") {
      //alert(this.state.currentItemID);
      if (this.state.clientName != "") {
        this.validator.hideMessages();
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(this.state.currentItemID).update({
          Title: this.state.clientName,
          SourceId: this.state.sourceKey,
          IndustryId: this.state.industryKey,
          ClassOfInsuranceId: this.state.classOfInsuranceKey,
          EstimatedPremium: Number(this.state.estimatedPremium),
          BrokeragePercentage: Number(this.state.brokerage),
          ComplianceCleared: this.state.complianceCleared,
          BrokerageAmount: this.state.brokerageAmount,
          NBOStageId: this.state.NB0StageKey,
          EstimatedStartDate: new Date(this.state.estimatedStartDate),
          FeesIfAny: Number(this.state.feesIfAny),
          Comments: this.state.comments,
          //EstimatedBrokerage: Number(this.state.estimatedPremium) * Number(this.state.brokerage / 100),
        }).then(async afterPipeLineUpdate => {
          //this.nbolid = this.state.currentItemID;
          if (this.state.externalArray.length > 0) {
            for (var i in this.state.externalArray) {
              var splitted = this.state.externalArray[i].documentName.split(".");
              let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + this.state.currentItemID + '.' + splitted[splitted.length - 1];
              let docName = documentNameExtension;
              await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/NBODocuments/").files.add(docName, this.state.externalArray[i].content, true).then(async fileUploaded => {
                const item = await fileUploaded.file.getItem();
                console.log(item);
                console.log(this.state.currentItemID);
                await sp.web.getList(this.props.siteUrl + "/NBODocuments/").items.getById(item["ID"]).update({
                  NBOPipelineIdId: this.state.currentItemID,
                  Title: this.state.externalArray[i].documentName,
                });
              });
            }
          }
          sp.web.getList(this.props.siteUrl + "/Lists/NBOLogList").items.add({
            Title: "Updation",
            NBOPipelineID: Number(this.state.currentItemID),
          });
        }).then(afterInsertion => {
          // alert("UpdatedSuccessfully");
          if (this.nbolid == "" || this.nbolid == null) {
            this.loadDocProfile();
          }
          toast("Nbo Pipeline updated successfully");
          this.setState({
            showDocInPanel: "",
            divForDocumentUploadCompArrayDiv: "none",
            // externalArray: [],
            // tempArrayForExternalDocumentGrid: [],

          });
          setTimeout(() => {
            this.setState({
              hideDialog: false,
              currentItemID: ""
            });
            this.nbolid = "";
            if (this.nbolid != "" || this.nbolid != null) {
              window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/NBOPipeline.aspx");
            }
          }, 6000);
        });


      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
    else {
      //update from mail link
      if (this.state.clientName != "" && this.compliance == "Yes") {
        //alert("from mail");
        this.validator.hideMessages();
        sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(Number(this.nbolid)).update({

          ComplianceCleared: this.state.complianceCleared,
        }).then(async afterPipeLineUpdate => {
          // this.nbolid = Number(this.nbolid);
          if (this.state.externalArray.length > 0) {
            for (var i in this.state.externalArray) {
              var splitted = this.state.externalArray[i].documentName.split(".");
              let documentNameExtension = splitted.slice(0, -1).join('.') + "_" + Number(this.nbolid) + '.' + splitted[splitted.length - 1];
              let docName = documentNameExtension;
              await sp.web.getFolderByServerRelativeUrl(this.props.siteUrl + "/NBODocuments/").files.add(docName, this.state.externalArray[i].content, true).then(async fileUploaded => {
                const item = await fileUploaded.file.getItem();
                console.log(item);
                console.log(this.nbolid);
                await sp.web.getList(this.props.siteUrl + "/NBODocuments/").items.getById(item["ID"]).update({
                  NBOPipelineIdId: Number(this.nbolid),
                  Title: this.state.externalArray[i].documentName,
                });
              });
            }
          }
          sp.web.getList(this.props.siteUrl + "/Lists/NBOLogList").items.add({
            Title: "Update from mail link",
            NBOPipelineID: Number(this.nbolid),
          });
        }).then(afterInsertion => {
          // alert("UpdatedSuccessfully");

          toast("Compliance Cleared successfully");
          this.setState({
            showDocInPanel: "",//from panel msg bar
            // divForDocumentUploadCompArrayDiv: "none",
            // externalArray: [],
            // tempArrayForExternalDocumentGrid: [],
          });
          setTimeout(() => {
            this.setState({
              hideDialog: false,

            });
            this.nbolid = "",
              window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/SitePages/NBOPipeline.aspx");
          }, 6000);

        });


      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
  }
  // ---------------EstimatedDatePickerChange---------------------
  private _estimateDateChange = (date?: Date): void => {
    this.setState({ estimatedStartDate: date, });
  }
  // ---------------Estimated From DatePickerChange---------------------
  private _estimateFromDateChange = (date?: Date): void => {
    this.setState({ estimatedFromStartDate: date, estimatedToStartDate: null });
  }
  // ---------------Estimated To DatePickerChange---------------------
  private _estimateToDateChange = (date?: Date): void => {
    this.setState({ estimatedToStartDate: date, });
  }
  // ---------------GroupName---------------------
  public _drpdwnGroupName(option: { key: any; text: any }) {
    this.setState({
      groupNamekey: option.key,
      groupName: option.text,
    });
  }
  // ---------------oppurtunityType---------------------
  public _drpdwnOppurtunityType(option: { key: any; text: any }) {
    this.setState({
      oppurtunityTypeKey: option.key,
      oppurtunityType: option.text,
    });
  }
  // ---------------oppurtunityType---------------------
  public async filterColumnChange(option: { key: any; text: any }) {
    this.setState({
      selectedColumnKey: option.key,
    });
    let oppTypeConditions = ["New", "Expanded"];
    let numberTypeConditions = ["<", ">", "=", "!=", "<=", ">="];
    let stringTypeConditions = ["equals", "not equal to", "Contains"];
    let complianceClearedConditions = ["Yes", "No", "Pending"];
    let tempConditionArray = [];
    if (option.key == "OpportunityType") {
      for (let i = 0; i < oppTypeConditions.length; i++) {
        // if(subcontractor[i].Active == true){
        let Itemdata = {
          key: oppTypeConditions[i],
          text: oppTypeConditions[i]
        };
        tempConditionArray.push(Itemdata);
      }
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: tempConditionArray, textFiledForFilter: "none" });
    }
    else if (option.key == "EstimatedBrokerage" || option.key == "FeesIfAny" || option.key == "EstimatedPremium" || option.key == "WeightedBrokerage" || option.key == "BrokeragePercentage") {
      for (let i = 0; i < numberTypeConditions.length; i++) {
        // if(subcontractor[i].Active == true){
        let Itemdata = {
          key: numberTypeConditions[i],
          text: numberTypeConditions[i]
        };
        tempConditionArray.push(Itemdata);
      }
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: tempConditionArray, textFiledForFilter: "" });
    }
    else if (option.key == "Title" || option.key == "Comments" || option.key == "Department" || option.key == "Author") {
      for (let i = 0; i < stringTypeConditions.length; i++) {
        // if(subcontractor[i].Active == true){
        let Itemdata = {
          key: stringTypeConditions[i],
          text: stringTypeConditions[i]
        };
        tempConditionArray.push(Itemdata);
      }
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: tempConditionArray, textFiledForFilter: "" });
    }
    else if (option.key == "ComplianceCleared") {
      for (let i = 0; i < complianceClearedConditions.length; i++) {
        // if(subcontractor[i].Active == true){
        let Itemdata = {
          key: complianceClearedConditions[i],
          text: complianceClearedConditions[i]
        };
        tempConditionArray.push(Itemdata);
      }
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: tempConditionArray, textFiledForFilter: "none" });
    }
    else if (option.key == "Industry") {
      await this._drpdwnIndustryBind();
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: this.state.industryItems, textFiledForFilter: "none" });

    }
    else if (option.key == "ClassOfInsurance") {
      await this._drpdwnClassOfInsuranceBind();
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: this.state.classOfInsuranceItems, textFiledForFilter: "none" });
    }
    else if (option.key == "Source") {
      await this._drpdwnSource();
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: this.state.sourceItems, textFiledForFilter: "none" });
    }
    else if (option.key == "NBOStage") {
      await this._drpdwnNBOStageBind();
      this.setState({ dateForFilter: "none", filterConditionDiv: "", filterConditions: this.state.NBOStageItems, textFiledForFilter: "none" });
    }
    else if (option.key == "EstimatedStartDate") {
      this.setState({
        dateForFilter: "",
        filterConditionDiv: "none",
        textFiledForFilter: "none",
      });
    }
  }
  public async filterConditionColumnChange(option: { key: any; text: any }) {
    this.setState({
      filterCondition: option.text,
      filterConditionKey: option.key,
    });
  }
  public async _selectDepartmentFromSameDepartmentTab(option: { key: any; text: any }) {
    this.setState({
      departmentkey: option.key,
      departmentText: option.text,
    });


    //binding with selected departments
    this.setState({
      sameDepartmentItems: "Yes",
      currentItemID: "",
    });

    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
      //select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')")
      select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Department eq  '" + option.text + "'")
      .get().then(docProfileItems => {
        console.log("SameDept", docProfileItems);
        this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
        this.setState({
          docRepositoryItems: this.sortedArray,
          items: this.sortedArray,
          paginatedItems: this.sortedArray.slice(0, this.pageSize),
          noItemErrorMsg: docProfileItems.length == 0 ? " " : "none",
        });
        console.log(this.state.docRepositoryItems);
        if (docProfileItems.length == 0) {
          this.setState({ noItemErrorMsg: "" });
        }
      });
    this.setState({
      divForSame: "",
      divForCurrentUser: "none",
      divForOtherDepts: "none",
    });


  }
  // ---------------Source---------------------
  public _drpdwnChangeSource(option: { key: any; text: any }) {
    this.setState({
      source: option.text,
      sourceKey: option.key
    });
  }
  // ---------------Industry---------------------
  public _drpdwnIndustry(option: { key: any; text: any }) {
    this.setState({
      industry: option.text,
      industryKey: option.key,
    });
  }
  // ---------------filter name change---------------------
  private filterValueChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState({ filterValue: newText || '' });
  }
  // ---------------ClientName---------------------
  private clientNameChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState({ clientName: newText || '' });
  }
  // ---------------FeesIfAny---------------------
  private _feesIfAnyChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState({ feesIfAny: newText || '' });
  }
  // ---------------EstimatedPremium---------------------
  private _estimatedPremiumChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {

    this.setState({ estimatedPremium: newText });
  }
  private _commentsChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState({ comments: newText || '' });
  }
  // ---------------ModalClose---------------------
  private _closeModal = (): void => {
    this.setState({
      tempArrayForExternalDocumentGrid: [],
      externalArrayDiv: "none",
      showReviewModal: false,
      clientName: "",
      source: "",
      NB0StageText: "",
      complianceCleared: "",
      industry: "",
      classOfInsurance: "",
      estimatedPremium: "",
      brokerage: "",
      estimatedBrokerage: "",
      brokerageAmount: "",
      estimatedStartDate: "",
      feesIfAny: "0",
      comments: "",
      divForDocumentUploadCompliance: "none",
      oppurtunityType: "",
      oppurtunityTypeKey: "",

    });
  }

  // ---------------NBOStage---------------------
  public _drpdwnNBOStage(option: { key: any; text: any }) {
    //console.log(option.key)      
    this.setState({ NB0StageText: option.text, NB0StageKey: option.key });
  }
  // ---------------ComplianceCleared---------------------
  public _drpdwnComplianceCleared(option: { key: any; text: any }) {
    //console.log(option.key)   
    this.setState({ complianceCleared: option.text });
    // alert(option.text);
    if (option.text == "Yes") {
      this.setState({
        divForDocumentUploadCompliance: "",
        externalArrayDiv: "none",
      });
    }
    else {
      this.setState({
        divForDocumentUploadCompliance: "none",
        externalArrayDiv: "none",
      });
    }
  }
  // ---------------Brokerage---------------------
  // public _drpdwnBrokerage(option: { key: any; text: any }) {
  //   //console.log(option.key)   
  //   this.setState({ brokerage: option.text, brokerageKey: option.key });
  // }
  private _drpdwnBrokerage = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState({ brokerage: newText || '' });
  }
  // ---------------ClassOFInsurance---------------------
  public _drpdwnClassOfInsurance(option: { key: any; text: any }) {
    //console.log(option.key)   
    this.setState({ classOfInsurance: option.text, classOfInsuranceKey: option.key });
  }
  // upload additional
  public _uploadadditional(e) {
    this.myfileadditional = e.target.value;
    let documentNameExtension;
    console.log(this.myfileadditional);
    console.log(e.target.value);
    console.log(e.currentTarget.value);
    let myfile = (document.querySelector("#newfile") as HTMLInputElement).files[0];
    var splitted = myfile.name.split(".");
  }
  //pagination onChange
  private _getPage(page: number) {
    // round a number up to the next largest integer.
    const roundupPage = Math.ceil(page);// Math.ceil(page);
    this.setState({
      currentPage: page,
      paginatedItems: this.sortedArray.slice(roundupPage * this.pageSize, (roundupPage * this.pageSize) + this.pageSize)
    });
  }
  //dropdown bindings--------------------------------------------------------------------------------->
  public _drpdwnSource() {
    let tempsourceItems = [];
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.sourceListName).items.get().then(source => {
      console.log("transmitFor", source);
      for (let i = 0; i < source.length; i++) {
        // if(subcontractor[i].Active == true){
        let sourceItemdata = {
          key: source[i].ID,
          text: source[i].Title
        };
        tempsourceItems.push(sourceItemdata);
        //}       
      }
      this.setState({
        sourceItems: tempsourceItems
      });
    });
  }
  // public _drpdwnBrokeragePercentage() {
  //   let tempBrokeragePercentageItems = [];
  //   sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.brokeragePercentage).items.get().then(BrokeragePercentage => {
  //     console.log("BrokeragePercentage", BrokeragePercentage);
  //     for (let i = 0; i < BrokeragePercentage.length; i++) {
  //       // if(subcontractor[i].Active == true){
  //       let BrokeragePercentageItemdata = {
  //         key: BrokeragePercentage[i].ID,
  //         text: BrokeragePercentage[i].Title
  //       };
  //       tempBrokeragePercentageItems.push(BrokeragePercentageItemdata);
  //       //}       
  //     }
  //     this.setState({
  //       brokeragePercentageItems: tempBrokeragePercentageItems
  //     });
  //   });
  // }
  public _drpdwnIndustryBind() {
    let tempIndustryItems = [];
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.industry).items.get().then(Industry => {
      console.log("Industry", Industry);
      for (let i = 0; i < Industry.length; i++) {
        // if(subcontractor[i].Active == true){
        let industryItemdata = {
          key: Industry[i].ID,
          text: Industry[i].Title
        };
        tempIndustryItems.push(industryItemdata);
        //}       
      }
      this.setState({
        industryItems: tempIndustryItems
      });
    });
  }
  public _drpdwnClassOfInsuranceBind() {
    let tempClassOfInsurance = [];
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.classOfInsurance).items.get().then(ClassOfInsurance => {
      console.log("ClassOfInsurance", ClassOfInsurance);
      for (let i = 0; i < ClassOfInsurance.length; i++) {
        // if(subcontractor[i].Active == true){
        let ClassOfInsuranceItemdata = {
          key: ClassOfInsurance[i].ID,
          text: ClassOfInsurance[i].Title
        };
        tempClassOfInsurance.push(ClassOfInsuranceItemdata);
        //}       
      }
      this.setState({
        classOfInsuranceItems: tempClassOfInsurance
      });
    });
  }
  public _drpdwnNBOStageBind() {
    let tempNBOStage = [];
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboStage).items.get().then(NBOStage => {
      console.log("NBOStage", NBOStage);
      for (let i = 0; i < NBOStage.length; i++) {
        // if(subcontractor[i].Active == true){
        let NBOStageItemdata = {
          key: NBOStage[i].ID,
          text: NBOStage[i].Title
        };
        tempNBOStage.push(NBOStageItemdata);
        //}       
      }
      this.setState({
        NBOStageItems: tempNBOStage
      });
    });
  }

  public _openDeleteConfirmation(items, key) {
    console.log(items);
    this.state.externalArray.splice(key, 1);
    console.log("after removal", this.state.externalArray);
    this.setState({
      externalArray: this.state.externalArray,
    });
    if (this.state.externalArray.length == 0) {
      this.setState({
        externalArrayDiv: "none"
      });
    }
  }
  public openCCSPopUp(items) {
    console.log(items.ID);

    window.open(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl + "/NBODocuments/Forms/AllItems.aspx?FilterField1=NBOPipelineId&FilterValue1=" + parseInt(items.ID) + "&FilterType1=Lookup&viewid=05f46dbe-8633-4e54-b0e3-b85ca3d8f235");
    sp.web.getList(this.props.siteUrl + "/NBODocuments").items.filter("NBOPipelineIdId eq '" + parseInt(items.ID) + "'").get().then(afterDocumentGettings => {
      console.log(afterDocumentGettings);
      this.setState({
        externalArray: afterDocumentGettings
      });
    });
  }

  //grid binding for same departments
  private async sameDept() {
    //alert("same dept");
    //  alert(this.team);
    let docProfileItems = [];
    this.forDeptCreatedBy = "ok";
    console.log("departments of current user", this.state.oppurtunityDept);
    this.setState({
      sameDepartmentItems: "Yes",
      currentItemID: "",
      forOtherDeptFilter: "MYNBOSame",
      filterCondition: "",
      filterValue: "",
      selectedColumnKey: "",
      filterConditions: [],
      divForNoDataFound: "none"
    });

    let tempArray = [];
    // await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
    //   //select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')")
    //   select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
    //   .get().then(docItems => {
    //     for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {
    //       for (let listItem = 0; listItem < docItems.length; listItem++) {
    //         if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text) {
    //           tempArray.push(docItems[listItem]);
    //         }
    //       }
    //     }
    //     for (let i = 0; i < this.pageSize; i++) {
    //       docProfileItems.push({
    //         "ID": null,
    //         "Title": null,
    //         "BrokeragePercentage": null,
    //         "Source": { "ID": null, "Title": null },
    //         "Industry": { "ID": null, "Title": null },
    //         "ClassOfInsurance": { "ID": null, "Title": null },
    //         "NBOStage": { "ID": null, "Title": null },
    //         "EstimatedBrokerage": null,
    //         "FeesIfAny": null,
    //         "Comments": null,
    //         "EstimatedStartDate": null,
    //         "EstimatedPremium": null,
    //         "Department": null,
    //         "ComplianceCleared": null,
    //         "Author": { "EMail": null, "Title": null },
    //         "WeightedBrokerage": null,
    //         "OpportunityType": null,
    //       });
    //     }
    //     console.log("SameDept", tempArray);
    //     docProfileItems = docProfileItems.concat(tempArray);
    //     this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
    //     this.setState({
    //       arrayForShowingPagination: tempArray,
    //       docRepositoryItems: this.sortedArray,
    //       items: this.sortedArray,
    //       paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
    //       noItemErrorMsg: tempArray.length == 0 ? " " : "none",
    //     });
    //     console.log(this.state.docRepositoryItems);
    //     if (tempArray.length == 0) {
    //       this.setState({ noItemErrorMsg: "" });
    //     }
    //     this.setState({
    //       divForSame: "",
    //       divForCurrentUser: "none",
    //       divForOtherDepts: "none",
    //       divForShowingPagination: "",
    //     });

    //   });
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
      //select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')")
      select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
      .expand("Source,Industry,ClassOfInsurance,NBOStage,Author").orderBy("ID", false)
      .get().then(docItems => {
        for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {
          for (let listItem = 0; listItem < docItems.length; listItem++) {
            if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text) {
              tempArray.push(docItems[listItem]);
            }
          }
        }
        for (let i = 0; i < this.pageSize; i++) {
          docProfileItems.push({
            "ID": null,
            "Title": null,
            "BrokeragePercentage": null,
            "Source": { "ID": null, "Title": null },
            "Industry": { "ID": null, "Title": null },
            "ClassOfInsurance": { "ID": null, "Title": null },
            "NBOStage": { "ID": null, "Title": null },
            "EstimatedBrokerage": null,
            "FeesIfAny": null,
            "Comments": null,
            "EstimatedStartDate": null,
            "EstimatedPremium": null,
            "Department": null,
            "ComplianceCleared": null,
            "Author": { "EMail": null, "Title": null },
            "WeightedBrokerage": null,
            "OpportunityType": null,
          });
        }
        console.log("SameDept", tempArray);
        docProfileItems = docProfileItems.concat(tempArray);
        this.sortedArray = docProfileItems;
        this.setState({
          arrayForShowingPagination: tempArray,
          docRepositoryItems: this.sortedArray,
          items: this.sortedArray,
          paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
          noItemErrorMsg: tempArray.length == 0 ? " " : "none",
        });
        console.log(this.state.docRepositoryItems);
        if (tempArray.length == 0) {
          this.setState({ noItemErrorMsg: "" });
        }
        this.setState({
          divForSame: "",
          divForCurrentUser: "none",
          divForOtherDepts: "none",
        });

      });




  }
  //grid binding for other departments tab and for NBO admin
  private async others() {
    this.setState({
      forOtherDeptFilter: "Other",
      filterConditions: [],
      sameDepartmentItems: "not",
      filterCondition: "",
      filterValue: "",
      selectedColumnKey: "",
      divForNoDataFound: "none"
    });
    this.forDeptCreatedBy = "ok";
    //alert("others");
    let docProfileItems = [];
    if (this.state.isNBOAdmin != "true") {
      //not an NBO Admin
      let tempArray = [];
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
        select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
        .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
        .filter("Author/EMail ne '" + this.currentUserEmail + "' and Department ne '" + this.team + "'")
        .top(4000).getPaged()
        .then(async docItems => {
          // for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {
          //   for (let listItem = 0; listItem < docProfileItems.length; listItem++) {
          //     if (this.state.oppurtunityDept[sd].text != docProfileItems[listItem].Department) {
          //       tempArray.push(docProfileItems[listItem]);
          //     }
          //   }
          // }
          for (let i = 0; i < this.pageSize; i++) {
            docProfileItems.push({
              "ID": null,
              "Title": null,
              "BrokeragePercentage": null,
              "Source": { "ID": null, "Title": null },
              "Industry": { "ID": null, "Title": null },
              "ClassOfInsurance": { "ID": null, "Title": null },
              "NBOStage": { "ID": null, "Title": null },
              "EstimatedBrokerage": null,
              "FeesIfAny": null,
              "Comments": null,
              "EstimatedStartDate": null,
              "EstimatedPremium": null,
              "Department": null,
              "ComplianceCleared": null,
              "Author": { "EMail": null, "Title": null },
              "WeightedBrokerage": null,
              "OpportunityType": null,
            });
          }
          docProfileItems = docProfileItems.concat(docItems.results);
          console.log(docProfileItems);
          while (docItems.hasNext) {
            docItems = await docItems.getNext();
            docProfileItems.push(...(docItems.results));

          }
          this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
          this.setState({
            arrayForShowingPagination: docItems.results,
            docRepositoryItems: this.sortedArray,
            items: this.sortedArray,
            paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
            noItemErrorMsg: this.sortedArray.length == 0 ? " " : "none"
          });
          console.log(this.state.docRepositoryItems);
          console.log("paginatedItems", this.state.paginatedItems);
          if (this.sortedArray.length == 0) {
            this.setState({ noItemErrorMsg: "" });
          }

          this.setState({
            divForSame: "none",
            divForOtherDepts: "",
            divForCurrentUser: "none",
            divForShowingPagination: "",
          });
        });
    }
    else {
      //alert(this.state.isNBOAdmin);
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
        select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
        .expand("Source,Industry,ClassOfInsurance,NBOStage,Author").top(4000).getPaged().then(async docItems => {
          // docProfileItems[this.pageSize] = docItems.results;
          for (let i = 0; i < this.pageSize; i++) {
            docProfileItems.push({
              "ID": null,
              "Title": null,
              "BrokeragePercentage": null,
              "Source": { "ID": null, "Title": null },
              "Industry": { "ID": null, "Title": null },
              "ClassOfInsurance": { "ID": null, "Title": null },
              "NBOStage": { "ID": null, "Title": null },
              "EstimatedBrokerage": null,
              "FeesIfAny": null,
              "Comments": null,
              "EstimatedStartDate": null,
              "EstimatedPremium": null,
              "Department": null,
              "ComplianceCleared": null,
              "Author": { "EMail": null, "Title": null },
              "WeightedBrokerage": null,
              "OpportunityType": null,
            });
          }
          docProfileItems = docProfileItems.concat(docItems.results);
          console.log(docProfileItems);
          while (docItems.hasNext) {
            docItems = await docItems.getNext();
            docProfileItems.push(...(docItems.results));
          }
          this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
          this.setState({
            docRepositoryItems: this.sortedArray,
            items: this.sortedArray,
            paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
            noItemErrorMsg: docProfileItems.length == 0 ? " " : "none",

          });
          console.log(this.state.docRepositoryItems);
          if (docProfileItems.length == 0) {
            this.setState({ noItemErrorMsg: "" });
          }
        });
      this.setState({
        divForSame: "none",
        divForOtherDepts: "",
        divForCurrentUser: "none",
        divForShowingPagination: "",
      });
    }

  }
  //updation from mail 
  private updateComplianceFromMail() {
    toast("Compliance Cleared updated successfully");
    this.setState({
      messageBar: "",
    });
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(parseInt(this.nbolid)).update({

      ComplianceCleared: this.state.complianceCleared,

    });

  }

  //confirm Delete button click
  private _confirmYesCancel = () => {
    //alert(this.state.itemIDForDelete);
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(Number(this.state.itemIDForDelete)).recycle().then(afterDelete => {
      sp.web.getList(this.props.siteUrl + "/Lists/NBOLogList").items.add({
        Title: "Deletion",
        NBOPipelineID: Number(this.state.itemIDForDelete),
      });
      sp.web.getList(this.props.siteUrl + "/NBODocuments/").items.select("ID").filter("NBOPipelineIdId eq '" + Number(this.state.itemIDForDelete) + "'").get().then(NBoIDs => {
        console.log(NBoIDs);
        if (NBoIDs.length > 0) {
          for (let h = 0; h < NBoIDs.length; h++) {
            console.log("NBODocumentID", NBoIDs[0].ID);
            sp.web.getList(this.props.siteUrl + "/NBODocuments/").items.getById(Number(NBoIDs[0].ID)).recycle();
          }
        }
      });

    }).then(afterDelete => {
      this.setState({ confirmDialog: true, deleteMessageBar: "", statusMessage: { isShowMessage: true, message: "Deleted Successfully", messageType: 4 } });
      setTimeout(() => {
        this.loadDocProfile();
      }, 1000);
      setTimeout(() => {
        this.setState({ deleteMessageBar: "none", });
      }, 6000);
    });

  }
  private _confirmNoCancel = () => {
    this.setState({
      confirmDialog: true,
    });
  }
  //sorting for MYNBO grid each header
  private _onSortClickAscForMyNBO = (sortBy, e) => {
    if (this.state.divForNoDataFound == "none") {
      let event = e.currentTarget.ariaLabel;
      let eventID = e.currentTarget.id
      let docProfileItems = [];
      console.log(event);
      this.setState({
        sameDepartmentItems: "no",
        currentItemID: "",
      });
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
        select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail eq '" + this.currentUserEmail + "'")
        .orderBy(sortBy)
        .top(4000).getPaged()
        .then(async docItems => {
          this.setState({ arrayForShowingPagination: docItems.results });
          for (let i = 0; i < this.pageSize; i++) {
            docProfileItems.push({
              "ID": null,
              "Title": null,
              "BrokeragePercentage": null,
              "Source": { "ID": null, "Title": null },
              "Industry": { "ID": null, "Title": null },
              "ClassOfInsurance": { "ID": null, "Title": null },
              "NBOStage": { "ID": null, "Title": null },
              "EstimatedBrokerage": null,
              "FeesIfAny": null,
              "Comments": null,
              "EstimatedStartDate": null,
              "EstimatedPremium": null,
              "Department": null,
              "ComplianceCleared": null,
              "Author": { "EMail": null, "Title": null },
              "WeightedBrokerage": null,
              "OpportunityType": null,
            });
          }
          docProfileItems = docProfileItems.concat(docItems.results);
          console.log(docProfileItems);
          while (docItems.hasNext) {
            docItems = await docItems.getNext();
            docProfileItems.push(...(docItems.results));
          }

          this.sortedArray = docProfileItems;
          this.setState({
            docRepositoryItems: this.sortedArray,
            items: this.sortedArray,
            paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
          });
          console.log(this.state.docRepositoryItems);

        });
      this.setState({
        divForSame: "none",
        divForCurrentUser: "",
        divForOtherDepts: "none",
        divForDocumentUploadCompArrayDiv: "none",
      });

      switch (this.sortedArray.length > 0) {
        case (e.currentTarget.ariaLabel == "OpportunityType" || e.currentTarget.id == "OpportunityType"):
          this.setState({
            sortOppurtunityTypeDesc: "",
            sortOppurtunityTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "Source" || e.currentTarget.id == "Source"):
          this.setState({
            SourceTypeDesc: "",
            SourceTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "Industry" || e.currentTarget.id == "Industry"):
          this.setState({
            IndustryTypeDesc: "",
            IndustryTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "Title" || e.currentTarget.id == "Title"):
          this.setState({
            ClientNameTypeDesc: "",
            ClientNameTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "ClassOfInsurance" || e.currentTarget.id == "ClassOfInsurance"):
          this.setState({
            ClassOfInsuranceTypeDesc: "",
            ClassOfInsuranceTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "EstimatedStartDate" || e.currentTarget.id == "EstimatedStartDate"):
          this.setState({
            EstStartDateTypeDesc: "",
            EstStartDateTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "Comments" || e.currentTarget.id == "Comments"):
          this.setState({
            CommentsTypeDesc: "",
            CommentsTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "EstimatedPremium" || e.currentTarget.id == "EstimatedPremium"):
          this.setState({
            EstimatedPremiumTypeDesc: "",
            EstimatedPremiumTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "BrokeragePercentage" || e.currentTarget.id == "BrokeragePercentage"):
          this.setState({
            BrokerageTypeDesc: "",
            BrokerageTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "EstimatedBrokerage" || e.currentTarget.id == "EstimatedBrokerage"):
          this.setState({
            EstimatedBrokerageTypeDesc: "",
            EstimatedBrokerageTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "NBOStage" || e.currentTarget.id == "NBOStage"):
          this.setState({
            NBOStageTypeDesc: "",
            NBOStageTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "WeightedBrokerage" || e.currentTarget.id == "WeightedBrokerage"):
          this.setState({
            WeightedBrokerageTypeDesc: "",
            WeightedBrokerageTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "FeesIfAny" || e.currentTarget.id == "FeesIfAny"):
          this.setState({
            FeesIfAnyTypeDesc: "",
            FeesIfAnyTypeAsc: "none",
          });
          break;
        case (e.currentTarget.ariaLabel == "ComplianceCleared" || e.currentTarget.id == "ComplianceCleared"):
          this.setState({
            ComplianceClearedTypeDesc: "",
            ComplianceClearedTypeAsc: "none",
          });
          break;

      }
    }
  }
  private _onSortClickDescForMyNBO = (sortBy, e) => {
    //alert("SortClicked");
    if (this.state.divForNoDataFound == "none") {
      let event = e.currentTarget.ariaLabel;
      let eventID = e.currentTarget.id
      let docProfileItems = [];
      this.setState({
        sameDepartmentItems: "no",
        currentItemID: "",

      });
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
        select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
        .filter("Author/EMail eq '" + this.currentUserEmail + "'")
        .top(4000).getPaged()
        .then(async docItems => {
          this.setState({ arrayForShowingPagination: docItems.results });
          for (let i = 0; i < this.pageSize; i++) {
            docProfileItems.push({
              "ID": null,
              "Title": null,
              "BrokeragePercentage": null,
              "Source": { "ID": null, "Title": null },
              "Industry": { "ID": null, "Title": null },
              "ClassOfInsurance": { "ID": null, "Title": null },
              "NBOStage": { "ID": null, "Title": null },
              "EstimatedBrokerage": null,
              "FeesIfAny": null,
              "Comments": null,
              "EstimatedStartDate": null,
              "EstimatedPremium": null,
              "Department": null,
              "ComplianceCleared": null,
              "Author": { "EMail": null, "Title": null },
              "WeightedBrokerage": null,
              "OpportunityType": null,
            });
          }
          docProfileItems = docProfileItems.concat(docItems.results);
          console.log(docProfileItems);
          while (docItems.hasNext) {
            docItems = await docItems.getNext();
            docProfileItems.push(...(docItems.results));
          }

          this.sortedArray = _.orderBy(docProfileItems, sortBy, ['desc']);
          this.setState({
            docRepositoryItems: this.sortedArray,
            items: this.sortedArray,
            paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
          });
          console.log(this.state.docRepositoryItems);

        });
      this.setState({
        divForSame: "none",
        divForCurrentUser: "",
        divForOtherDepts: "none",
        divForDocumentUploadCompArrayDiv: "none",

      });
      switch (this.sortedArray.length > 0) {
        case (e.currentTarget.ariaLabel == "OpportunityType" || e.currentTarget.id == "OpportunityType"):
          this.setState({
            sortOppurtunityTypeDesc: "none",
            sortOppurtunityTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "Source" || e.currentTarget.id == "Source"):
          this.setState({
            SourceTypeDesc: "none",
            SourceTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "Industry" || e.currentTarget.id == "Industry"):
          this.setState({
            IndustryTypeDesc: "none",
            IndustryTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "Title" || e.currentTarget.id == "Title"):
          this.setState({
            ClientNameTypeDesc: "none",
            ClientNameTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "ClassOfInsurance" || e.currentTarget.id == "ClassOfInsurance"):
          this.setState({
            ClassOfInsuranceTypeDesc: "none",
            ClassOfInsuranceTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "EstimatedStartDate" || e.currentTarget.id == "EstimatedStartDate"):
          this.setState({
            EstStartDateTypeDesc: "none",
            EstStartDateTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "Comments" || e.currentTarget.id == "Comments"):
          this.setState({
            CommentsTypeDesc: "none",
            CommentsTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "EstimatedPremium" || e.currentTarget.id == "EstimatedPremium"):
          this.setState({
            EstimatedPremiumTypeDesc: "none",
            EstimatedPremiumTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "Brokerage" || e.currentTarget.id == "Brokerage"):
          this.setState({
            BrokerageTypeDesc: "none",
            BrokerageTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "EstimatedBrokerage" || e.currentTarget.id == "EstimatedBrokerage"):
          this.setState({
            EstimatedBrokerageTypeDesc: "none",
            EstimatedBrokerageTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "NBOStage" || e.currentTarget.id == "NBOStage"):
          this.setState({
            NBOStageTypeDesc: "none",
            NBOStageTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "WeightedBrokerage" || e.currentTarget.id == "WeightedBrokerage"):
          this.setState({
            WeightedBrokerageTypeDesc: "none",
            WeightedBrokerageTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "FeesIfAny" || e.currentTarget.id == "FeesIfAny"):
          this.setState({
            FeesIfAnyTypeDesc: "none",
            FeesIfAnyTypeAsc: "",
          });
          break;
        case (e.currentTarget.ariaLabel == "ComplianceCleared" || e.currentTarget.id == "ComplianceCleared"):
          this.setState({
            ComplianceClearedTypeDesc: "none",
            ComplianceClearedTypeAsc: "",
          });
          break;

      }
    }
  }
  //sorting for MYNBO grid each header
  private _onSortClickAscForSameDept = async (sortBy, e) => {
    if (this.state.divForNoDataFound == "none") {
      let event = e.currentTarget.ariaLabel;
      let eventID = e.currentTarget.id;
      let docProfileItems = [];
      console.log(event);
      this.setState({
        sameDepartmentItems: "Yes",
        currentItemID: "",
      });
      let tempArray = [];
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
        //select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')")
        select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
        .expand("Source,Industry,ClassOfInsurance,NBOStage,Author").orderBy(sortBy)
        .get().then(docItems => {
          for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {
            for (let listItem = 0; listItem < docItems.length; listItem++) {
              if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text) {
                tempArray.push(docItems[listItem]);
              }
            }
          }
          for (let i = 0; i < this.pageSize; i++) {
            docProfileItems.push({
              "ID": null,
              "Title": null,
              "BrokeragePercentage": null,
              "Source": { "ID": null, "Title": null },
              "Industry": { "ID": null, "Title": null },
              "ClassOfInsurance": { "ID": null, "Title": null },
              "NBOStage": { "ID": null, "Title": null },
              "EstimatedBrokerage": null,
              "FeesIfAny": null,
              "Comments": null,
              "EstimatedStartDate": null,
              "EstimatedPremium": null,
              "Department": null,
              "ComplianceCleared": null,
              "Author": { "EMail": null, "Title": null },
              "WeightedBrokerage": null,
              "OpportunityType": null,
            });
          }
          console.log("SameDept", tempArray);
          docProfileItems = docProfileItems.concat(tempArray);
          this.sortedArray = docProfileItems;
          this.setState({
            arrayForShowingPagination: tempArray,
            docRepositoryItems: this.sortedArray,
            items: this.sortedArray,
            paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
            noItemErrorMsg: tempArray.length == 0 ? " " : "none",
          });
          console.log(this.state.docRepositoryItems);
          if (tempArray.length == 0) {
            this.setState({ noItemErrorMsg: "" });
          }
          this.setState({
            divForSame: "",
            divForCurrentUser: "none",
            divForOtherDepts: "none",
          });

        });



      console.log(event);
      switch (this.sortedArray.length > 0) {
        case (event == "OpportunityType" || eventID == "OpportunityType"):
          this.setState({
            sortOppurtunityTypeDesc: "",
            sortOppurtunityTypeAsc: "none",
          });
          break;
        case (event == "Source" || eventID == "Source"):
          this.setState({
            SourceTypeDesc: "",
            SourceTypeAsc: "none",
          });
          break;
        case (event == "Industry" || eventID == "Industry"):
          this.setState({
            IndustryTypeDesc: "",
            IndustryTypeAsc: "none",
          });
          break;
        case (event == "Title" || eventID == "Title"):
          this.setState({
            ClientNameTypeDesc: "",
            ClientNameTypeAsc: "none",
          });
          break;
        case (event == "ClassOfInsurance" || eventID == "ClassOfInsurance"):
          this.setState({
            ClassOfInsuranceTypeDesc: "",
            ClassOfInsuranceTypeAsc: "none",
          });
          break;
        case (event == "EstimatedStartDate" || eventID == "EstimatedStartDate"):
          this.setState({
            EstStartDateTypeDesc: "",
            EstStartDateTypeAsc: "none",
          });
          break;
        case (event == "Comments" || eventID == "Comments"):
          this.setState({
            CommentsTypeDesc: "",
            CommentsTypeAsc: "none",
          });
          break;
        case (event == "EstimatedPremium" || eventID == "EstimatedPremium"):
          this.setState({
            EstimatedPremiumTypeDesc: "",
            EstimatedPremiumTypeAsc: "none",
          });
          break;
        case (event == "BrokeragePercentage" || eventID == "BrokeragePercentage"):
          this.setState({
            BrokerageTypeDesc: "",
            BrokerageTypeAsc: "none",
          });
          break;
        case (event == "EstimatedBrokerage" || eventID == "EstimatedBrokerage"):
          this.setState({
            EstimatedBrokerageTypeDesc: "",
            EstimatedBrokerageTypeAsc: "none",
          });
          break;
        case (event == "NBOStage" || eventID == "NBOStage"):
          this.setState({
            NBOStageTypeDesc: "",
            NBOStageTypeAsc: "none",
          });
          break;
        case (event == "WeightedBrokerage" || eventID == "WeightedBrokerage"):
          this.setState({
            WeightedBrokerageTypeDesc: "",
            WeightedBrokerageTypeAsc: "none",
          });
          break;
        case (event == "FeesIfAny" || eventID == "FeesIfAny"):
          this.setState({
            FeesIfAnyTypeDesc: "",
            FeesIfAnyTypeAsc: "none",
          });
          break;
        case (event == "ComplianceCleared" || eventID == "ComplianceCleared"):
          this.setState({
            ComplianceClearedTypeDesc: "",
            ComplianceClearedTypeAsc: "none",
          });
          break;
        case (event == "Department" || eventID == "Department"):
          this.setState({
            DepartmentTypeDesc: "",
            DepartmentTypeAsc: "none",
          });
          break;
        case (event == "Author" || eventID == "Author"):
          this.setState({
            CreatedByTypeDesc: "",
            CreatedByTypeAsc: "none",
          });
          break;

      }
    }
  }
  private _onSortClickDescForSameDept = async (sortBy, e) => {
    if (this.state.divForNoDataFound == "none") {
      let event = e.currentTarget.ariaLabel;
      let eventID = e.currentTarget.id;
      let docProfileItems = [];
      this.setState({
        sameDepartmentItems: "Yes",
        currentItemID: "",
      });
      let tempArray = [];
      await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
        //select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')")
        select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
        .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
        .get().then(docItems => {
          for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {
            for (let listItem = 0; listItem < docItems.length; listItem++) {
              if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text) {
                tempArray.push(docItems[listItem]);
              }
            }
          }
          for (let i = 0; i < this.pageSize; i++) {
            docProfileItems.push({
              "ID": null,
              "Title": null,
              "BrokeragePercentage": null,
              "Source": { "ID": null, "Title": null },
              "Industry": { "ID": null, "Title": null },
              "ClassOfInsurance": { "ID": null, "Title": null },
              "NBOStage": { "ID": null, "Title": null },
              "EstimatedBrokerage": null,
              "FeesIfAny": null,
              "Comments": null,
              "EstimatedStartDate": null,
              "EstimatedPremium": null,
              "Department": null,
              "ComplianceCleared": null,
              "Author": { "EMail": null, "Title": null },
              "WeightedBrokerage": null,
              "OpportunityType": null,
            });
          }
          console.log("SameDept", tempArray);
          docProfileItems = docProfileItems.concat(tempArray);
          this.sortedArray = _.orderBy(docProfileItems, sortBy, ['desc']);
          this.setState({
            arrayForShowingPagination: tempArray,
            docRepositoryItems: this.sortedArray,
            items: this.sortedArray,
            paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
            noItemErrorMsg: tempArray.length == 0 ? " " : "none",
          });
          console.log(this.state.docRepositoryItems);
          if (tempArray.length == 0) {
            this.setState({ noItemErrorMsg: "" });
          }
          this.setState({
            divForSame: "",
            divForCurrentUser: "none",
            divForOtherDepts: "none",
          });

        });



      switch (this.sortedArray.length > 0) {
        case (event == "OpportunityType" || eventID == "OpportunityType"):
          this.setState({
            sortOppurtunityTypeDesc: "none",
            sortOppurtunityTypeAsc: "",
          });
          break;
        case (event == "Source" || eventID == "Source"):
          this.setState({
            SourceTypeDesc: "none",
            SourceTypeAsc: "",
          });
          break;
        case (event == "Industry" || eventID == "Industry"):
          this.setState({
            IndustryTypeDesc: "none",
            IndustryTypeAsc: "",
          });
          break;
        case (event == "Title" || eventID == "Title"):
          this.setState({
            ClientNameTypeDesc: "none",
            ClientNameTypeAsc: "",
          });
          break;
        case (event == "ClassOfInsurance" || eventID == "ClassOfInsurance"):
          this.setState({
            ClassOfInsuranceTypeDesc: "none",
            ClassOfInsuranceTypeAsc: "",
          });
          break;
        case (event == "EstimatedStartDate" || eventID == "EstimatedStartDate"):
          this.setState({
            EstStartDateTypeDesc: "none",
            EstStartDateTypeAsc: "",
          });
          break;
        case (event == "Comments" || eventID == "Comments"):
          this.setState({
            CommentsTypeDesc: "none",
            CommentsTypeAsc: "",
          });
          break;
        case (event == "EstimatedPremium" || eventID == "EstimatedPremium"):
          this.setState({
            EstimatedPremiumTypeDesc: "none",
            EstimatedPremiumTypeAsc: "",
          });
          break;
        case (event == "Brokerage" || eventID == "Brokerage"):
          this.setState({
            BrokerageTypeDesc: "none",
            BrokerageTypeAsc: "",
          });
          break;
        case (event == "EstimatedBrokerage" || eventID == "EstimatedBrokerage"):
          this.setState({
            EstimatedBrokerageTypeDesc: "none",
            EstimatedBrokerageTypeAsc: "",
          });
          break;
        case (event == "NBOStage" || eventID == "NBOStage"):
          this.setState({
            NBOStageTypeDesc: "none",
            NBOStageTypeAsc: "",
          });
          break;
        case (event == "WeightedBrokerage" || eventID == "WeightedBrokerage"):
          this.setState({
            WeightedBrokerageTypeDesc: "none",
            WeightedBrokerageTypeAsc: "",
          });
          break;
        case (event == "FeesIfAny" || eventID == "FeesIfAny"):
          this.setState({
            FeesIfAnyTypeDesc: "none",
            FeesIfAnyTypeAsc: "",
          });
          break;
        case (event == "ComplianceCleared" || eventID == "ComplianceCleared"):
          this.setState({
            ComplianceClearedTypeDesc: "none",
            ComplianceClearedTypeAsc: "",
          });
          break;
        case (event == "Department" || eventID == "Department"):
          this.setState({
            DepartmentTypeDesc: "none",
            DepartmentTypeAsc: "",
          });
          break;
        case (event == "Author" || eventID == "Author"):
          this.setState({
            CreatedByTypeDesc: "none",
            CreatedByTypeAsc: "",
          });
          break;

      }
    }
  }
  //sorting for Other Dept grid each header
  private _onSortClickAscForOtherDept = async (sortBy, e) => {
    if (this.state.divForNoDataFound == "none") {
      let event = e.currentTarget.ariaLabel;
      let eventID = e.currentTarget.id;
      console.log(event);
      let docProfileItems = [];
      if (this.state.isNBOAdmin != "true") {
        //not an NBO Admin
        let tempArray = [];
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
          select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
          .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
          .filter("Author/EMail ne '" + this.currentUserEmail + "' and Department ne '" + this.team + "'")
          .orderBy(sortBy)
          .top(4000).getPaged()
          .then(async docItems => {
            // for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {
            //   for (let listItem = 0; listItem < docProfileItems.length; listItem++) {
            //     if (this.state.oppurtunityDept[sd].text != docProfileItems[listItem].Department) {
            //       tempArray.push(docProfileItems[listItem]);
            //     }
            //   }
            // }
            for (let i = 0; i < this.pageSize; i++) {
              docProfileItems.push({
                "ID": null,
                "Title": null,
                "BrokeragePercentage": null,
                "Source": { "ID": null, "Title": null },
                "Industry": { "ID": null, "Title": null },
                "ClassOfInsurance": { "ID": null, "Title": null },
                "NBOStage": { "ID": null, "Title": null },
                "EstimatedBrokerage": null,
                "FeesIfAny": null,
                "Comments": null,
                "EstimatedStartDate": null,
                "EstimatedPremium": null,
                "Department": null,
                "ComplianceCleared": null,
                "Author": { "EMail": null, "Title": null },
                "WeightedBrokerage": null,
                "OpportunityType": null,
              });
            }
            docProfileItems = docProfileItems.concat(docItems.results);
            console.log(docProfileItems);
            while (docItems.hasNext) {
              docItems = await docItems.getNext();
              docProfileItems.push(...(docItems.results));

            }
            this.sortedArray = docProfileItems;
            this.setState({
              arrayForShowingPagination: docItems.results,
              docRepositoryItems: this.sortedArray,
              items: this.sortedArray,
              paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
              noItemErrorMsg: this.sortedArray.length == 0 ? " " : "none"
            });
            console.log(this.state.docRepositoryItems);
            console.log("paginatedItems", this.state.paginatedItems);
            if (this.sortedArray.length == 0) {
              this.setState({ noItemErrorMsg: "" });
            }

            this.setState({
              divForSame: "none",
              divForOtherDepts: "",
              divForCurrentUser: "none"
            });
          });
      }
      else {
        //alert(this.state.isNBOAdmin);
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
          select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
          .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
          .orderBy(sortBy)
          .top(4000).getPaged().then(async docItems => {
            // docProfileItems[this.pageSize] = docItems.results;
            for (let i = 0; i < this.pageSize; i++) {
              docProfileItems.push({
                "ID": null,
                "Title": null,
                "BrokeragePercentage": null,
                "Source": { "ID": null, "Title": null },
                "Industry": { "ID": null, "Title": null },
                "ClassOfInsurance": { "ID": null, "Title": null },
                "NBOStage": { "ID": null, "Title": null },
                "EstimatedBrokerage": null,
                "FeesIfAny": null,
                "Comments": null,
                "EstimatedStartDate": null,
                "EstimatedPremium": null,
                "Department": null,
                "ComplianceCleared": null,
                "Author": { "EMail": null, "Title": null },
                "WeightedBrokerage": null,
                "OpportunityType": null,
              });
            }
            docProfileItems = docProfileItems.concat(docItems.results);
            console.log(docProfileItems);
            while (docItems.hasNext) {
              docItems = await docItems.getNext();
              docProfileItems.push(...(docItems.results));
            }
            this.sortedArray = docProfileItems;
            this.setState({
              docRepositoryItems: this.sortedArray,
              items: this.sortedArray,
              paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
              noItemErrorMsg: docProfileItems.length == 0 ? " " : "none",

            });
            console.log(this.state.docRepositoryItems);
            if (docProfileItems.length == 0) {
              this.setState({ noItemErrorMsg: "" });
            }
          });
        this.setState({
          divForSame: "none",
          divForOtherDepts: "",
          divForCurrentUser: "none"
        });
      }

      console.log(event);
      switch (this.sortedArray.length > 0) {
        case (event == "OpportunityType" || eventID == "OpportunityType"):
          this.setState({
            sortOppurtunityTypeDesc: "",
            sortOppurtunityTypeAsc: "none",
          });
          break;
        case (event == "Source" || eventID == "Source"):
          this.setState({
            SourceTypeDesc: "",
            SourceTypeAsc: "none",
          });
          break;

        case (event == "Industry" || eventID == "Industry"):
          this.setState({
            IndustryTypeDesc: "",
            IndustryTypeAsc: "none",
          });
          break;
        case (event == "Title" || eventID == "Title"):
          this.setState({
            ClientNameTypeDesc: "",
            ClientNameTypeAsc: "none",
          });
          break;
        case (event == "ClassOfInsurance" || eventID == "ClassOfInsurance"):
          this.setState({
            ClassOfInsuranceTypeDesc: "",
            ClassOfInsuranceTypeAsc: "none",
          });
          break;
        case (event == "EstimatedStartDate" || eventID == "EstimatedStartDate"):
          this.setState({
            EstStartDateTypeDesc: "",
            EstStartDateTypeAsc: "none",
          });
          break;
        case (event == "Comments" || eventID == "Comments"):
          this.setState({
            CommentsTypeDesc: "",
            CommentsTypeAsc: "none",
          });
          break;
        case (event == "EstimatedPremium" || eventID == "EstimatedPremium"):
          this.setState({
            EstimatedPremiumTypeDesc: "",
            EstimatedPremiumTypeAsc: "none",
          });
          break;
        case (event == "BrokeragePercentage" || eventID == "BrokeragePercentage"):
          this.setState({
            BrokerageTypeDesc: "",
            BrokerageTypeAsc: "none",
          });
          break;
        case (event == "EstimatedBrokerage" || eventID == "EstimatedBrokerage"):
          this.setState({
            EstimatedBrokerageTypeDesc: "",
            EstimatedBrokerageTypeAsc: "none",
          });
          break;
        case (event == "NBOStage" || eventID == "NBOStage"):
          this.setState({
            NBOStageTypeDesc: "",
            NBOStageTypeAsc: "none",
          });
          break;
        case (event == "WeightedBrokerage" || eventID == "WeightedBrokerage"):
          this.setState({
            WeightedBrokerageTypeDesc: "",
            WeightedBrokerageTypeAsc: "none",
          });
          break;
        case (event == "FeesIfAny" || eventID == "FeesIfAny"):
          this.setState({
            FeesIfAnyTypeDesc: "",
            FeesIfAnyTypeAsc: "none",
          });
          break;
        case (event == "ComplianceCleared" || eventID == "ComplianceCleared"):
          this.setState({
            ComplianceClearedTypeDesc: "",
            ComplianceClearedTypeAsc: "none",
          });
          break;
        case (event == "Department" || eventID == "Department"):
          this.setState({
            DepartmentTypeDesc: "",
            DepartmentTypeAsc: "none",
          });
          break;
        case (event == "Author" || eventID == "Author"):
          this.setState({
            CreatedByTypeDesc: "",
            CreatedByTypeAsc: "none",
          });
          break;

      }
    }
  }
  private _onSortClickDescForOtherDept = async (sortBy, e) => {
    if (this.state.divForNoDataFound == "none") {
      let event = e.currentTarget.ariaLabel;
      let eventID = e.currentTarget.id;
      let docProfileItems = [];
      console.log(event);
      if (this.state.isNBOAdmin != "true") {
        //not an NBO Admin
        let tempArray = [];
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
          select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
          .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
          .filter("Author/EMail ne '" + this.currentUserEmail + "' and Department ne '" + this.team + "'")

          .top(4000).getPaged()
          .then(async docItems => {
            // for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {
            //   for (let listItem = 0; listItem < docProfileItems.length; listItem++) {
            //     if (this.state.oppurtunityDept[sd].text != docProfileItems[listItem].Department) {
            //       tempArray.push(docProfileItems[listItem]);
            //     }
            //   }
            // }
            for (let i = 0; i < this.pageSize; i++) {
              docProfileItems.push({
                "ID": null,
                "Title": null,
                "BrokeragePercentage": null,
                "Source": { "ID": null, "Title": null },
                "Industry": { "ID": null, "Title": null },
                "ClassOfInsurance": { "ID": null, "Title": null },
                "NBOStage": { "ID": null, "Title": null },
                "EstimatedBrokerage": null,
                "FeesIfAny": null,
                "Comments": null,
                "EstimatedStartDate": null,
                "EstimatedPremium": null,
                "Department": null,
                "ComplianceCleared": null,
                "Author": { "EMail": null, "Title": null },
                "WeightedBrokerage": null,
                "OpportunityType": null,
              });
            }
            docProfileItems = docProfileItems.concat(docItems.results);
            console.log(docProfileItems);
            while (docItems.hasNext) {
              docItems = await docItems.getNext();
              docProfileItems.push(...(docItems.results));

            }
            this.sortedArray = _.orderBy(docProfileItems, sortBy, ['desc']);
            this.setState({
              arrayForShowingPagination: docItems.results,
              docRepositoryItems: this.sortedArray,
              items: this.sortedArray,
              paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
              noItemErrorMsg: this.sortedArray.length == 0 ? " " : "none"
            });
            console.log(this.state.docRepositoryItems);
            console.log("paginatedItems", this.state.paginatedItems);
            if (this.sortedArray.length == 0) {
              this.setState({ noItemErrorMsg: "" });
            }

            this.setState({
              divForSame: "none",
              divForOtherDepts: "",
              divForCurrentUser: "none"
            });
          });
      }
      else {
        //alert(this.state.isNBOAdmin);
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
          select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
          .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
          .top(4000).getPaged().then(async docItems => {
            // docProfileItems[this.pageSize] = docItems.results;
            for (let i = 0; i < this.pageSize; i++) {
              docProfileItems.push({
                "ID": null,
                "Title": null,
                "BrokeragePercentage": null,
                "Source": { "ID": null, "Title": null },
                "Industry": { "ID": null, "Title": null },
                "ClassOfInsurance": { "ID": null, "Title": null },
                "NBOStage": { "ID": null, "Title": null },
                "EstimatedBrokerage": null,
                "FeesIfAny": null,
                "Comments": null,
                "EstimatedStartDate": null,
                "EstimatedPremium": null,
                "Department": null,
                "ComplianceCleared": null,
                "Author": { "EMail": null, "Title": null },
                "WeightedBrokerage": null,
                "OpportunityType": null,
              });
            }
            docProfileItems = docProfileItems.concat(docItems.results);
            console.log(docProfileItems);
            while (docItems.hasNext) {
              docItems = await docItems.getNext();
              docProfileItems.push(...(docItems.results));
            }
            this.sortedArray = _.orderBy(docProfileItems, sortBy, ['desc']);
            this.setState({
              docRepositoryItems: this.sortedArray,
              items: this.sortedArray,
              paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
              noItemErrorMsg: docProfileItems.length == 0 ? " " : "none",

            });
            console.log(this.state.docRepositoryItems);
            if (docProfileItems.length == 0) {
              this.setState({ noItemErrorMsg: "" });
            }
          });
        this.setState({
          divForSame: "none",
          divForOtherDepts: "",
          divForCurrentUser: "none"
        });
      }

      switch (this.sortedArray.length > 0) {
        case (event == "OpportunityType" || eventID == "OpportunityType"):
          this.setState({
            sortOppurtunityTypeDesc: "none",
            sortOppurtunityTypeAsc: "",
          });
          break;
        case (event == "Source" || eventID == "Source"):
          this.setState({
            SourceTypeDesc: "none",
            SourceTypeAsc: "",
          });
          break;
        case (event == "Industry" || eventID == "Industry"):
          this.setState({
            IndustryTypeDesc: "none",
            IndustryTypeAsc: "",
          });
          break;
        case (event == "Title" || eventID == "Title"):
          this.setState({
            ClientNameTypeDesc: "none",
            ClientNameTypeAsc: "",
          });
          break;
        case (event == "ClassOfInsurance" || eventID == "ClassOfInsurance"):
          this.setState({
            ClassOfInsuranceTypeDesc: "none",
            ClassOfInsuranceTypeAsc: "",
          });
          break;
        case (event == "EstimatedStartDate" || eventID == "EstimatedStartDate"):
          this.setState({
            EstStartDateTypeDesc: "none",
            EstStartDateTypeAsc: "",
          });
          break;
        case (event == "Comments" || eventID == "Comments"):
          this.setState({
            CommentsTypeDesc: "none",
            CommentsTypeAsc: "",
          });
          break;
        case (event == "EstimatedPremium" || eventID == "EstimatedPremium"):
          this.setState({
            EstimatedPremiumTypeDesc: "none",
            EstimatedPremiumTypeAsc: "",
          });
          break;
        case (event == "Brokerage" || eventID == "Brokerage"):
          this.setState({
            BrokerageTypeDesc: "none",
            BrokerageTypeAsc: "",
          });
          break;
        case (event == "EstimatedBrokerage" || eventID == "EstimatedBrokerage"):
          this.setState({
            EstimatedBrokerageTypeDesc: "none",
            EstimatedBrokerageTypeAsc: "",
          });
          break;
        case (event == "NBOStage" || eventID == "NBOStage"):
          this.setState({
            NBOStageTypeDesc: "none",
            NBOStageTypeAsc: "",
          });
          break;
        case (event == "WeightedBrokerage" || eventID == "WeightedBrokerage"):
          this.setState({
            WeightedBrokerageTypeDesc: "none",
            WeightedBrokerageTypeAsc: "",
          });
          break;
        case (event == "FeesIfAny" || eventID == "FeesIfAny"):
          this.setState({
            FeesIfAnyTypeDesc: "none",
            FeesIfAnyTypeAsc: "",
          });
          break;
        case (event == "ComplianceCleared" || eventID == "ComplianceCleared"):
          this.setState({
            ComplianceClearedTypeDesc: "none",
            ComplianceClearedTypeAsc: "",
          });
          break;
        case (event == "Department" || eventID == "Department"):
          this.setState({
            DepartmentTypeDesc: "none",
            DepartmentTypeAsc: "",
          });
          break;
        case (event == "Author" || eventID == "Author"):
          this.setState({
            CreatedByTypeDesc: "none",
            CreatedByTypeAsc: "",
          });
          break;

      }
    }
  }
  //filter
  private _onFilter = () => {
    let tempSortedItems = [];
    this.setState({
      hideFilterDialog: true,
      estimatedStartDate: "",
      filterCondition: "",
      filterConditions: [],
      dateForFilter: "none",
      selectedColumnKey: "",
      filterConditionKey: "",
      textFiledForFilter: "",
      divForNoDataFound: "none",
      filterValue: "",
      estimatedFromStartDate: null,
      estimatedToStartDate: null,
    });

  }
  private dialogStyles = { main: { maxWidth: 300 } };
  private dialogContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    title: 'Do you want to Delete?',
    //subText: '<b>Do you want to cancel? </b> ',
  };
  private dialogContentFilterProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'none',
    //subText: '<b>Do you want to cancel? </b> ',
  };

  //   private  _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
  //     let filteredItems = [];
  //     if (text != "") {
  //         this.sortedArray.map((item) => {
  //             let itemValue = false;
  //             itemValue = Object.keys(item).some((key) => {
  //                 let requiredfield = fieldInternalName.find(element => element.includes(key)); //fieldInternalName.includes(key);
  //                 if (requiredfield) {
  //                     return JSON.stringify(item[key]).toString().toLowerCase().indexOf(text.toLowerCase().trim()) != -1;
  //                 }
  //             });
  //             if (itemValue) {
  //                 filteredItems.push(item);
  //             }
  //         });
  //     }
  //     else {
  //         filteredItems = [...this.sortedArray];
  //     }

  // };

  private async _onFilterButtonSubmit() {

    if (this.state.selectedColumnKey != "EstimatedStartDate") {
      if (this.state.selectedColumnKey != "" && this.state.selectedColumnKey != "Select an option" && this.state.filterCondition != "") {
        this.validator.hideMessages();
        //same department grid
        if (this.state.sameDepartmentItems == "Yes") {
          let docProfileItems = [];
          this.forDeptCreatedBy = "ok";
          console.log("departments of current user", this.state.oppurtunityDept);
          this.setState({
            currentItemID: "",
            forOtherDeptFilter: "MYNBOSame",
          });

          let tempArray = [];
          await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
            //select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')")
            select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
            .expand("Source,Industry,ClassOfInsurance,NBOStage,Author").orderBy("ID", false)
            .get().then(docItems => {
              for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {

                for (let listItem = 0; listItem < docItems.length; listItem++) {
                  if (this.state.selectedColumnKey == "OpportunityType") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].OpportunityType == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                    console.log(this.state.filterValue);
                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "FeesIfAny") {
                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedPremium") {
                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }




                  }
                  else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                    let duedate = moment(docItems[listItem].EstimatedStartDate).toDate();
                    let toDate = moment(this.state.estimatedToStartDate).toDate();
                    let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                    duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                    toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                    fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                    if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Title") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Title.toLowerCase() == this.state.filterValue.toLowerCase()) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Title != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Comments") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Comments == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Comments != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Department") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Department == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Department != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Department.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Author") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Author.Title == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Author.Title != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Author.Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "ComplianceCleared") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].ComplianceCleared == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Industry") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Industry.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }

                  }
                  else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Source") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Source.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "NBOStage") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].NBOStage.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                }
              }
              for (let i = 0; i < this.pageSize; i++) {
                docProfileItems.push({
                  "ID": null,
                  "Title": null,
                  "BrokeragePercentage": null,
                  "Source": { "ID": null, "Title": null },
                  "Industry": { "ID": null, "Title": null },
                  "ClassOfInsurance": { "ID": null, "Title": null },
                  "NBOStage": { "ID": null, "Title": null },
                  "EstimatedBrokerage": null,
                  "FeesIfAny": null,
                  "Comments": null,
                  "EstimatedStartDate": null,
                  "EstimatedPremium": null,
                  "Department": null,
                  "ComplianceCleared": null,
                  "Author": { "EMail": null, "Title": null },
                  "WeightedBrokerage": null,
                  "OpportunityType": null,
                });
              }
              console.log("SameDept", tempArray);
              docProfileItems = docProfileItems.concat(tempArray);
              this.sortedArray = docProfileItems;
              if (tempArray.length == 0) {
                this.setState({
                  divForShowingPagination: "none",
                  divForNoDataFound: ""
                });
              }
              this.setState({
                arrayForShowingPagination: tempArray,
                docRepositoryItems: this.sortedArray,
                items: this.sortedArray,
                paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
                noItemErrorMsg: tempArray.length == 0 ? " " : "none",
              });
              console.log(this.state.docRepositoryItems);
              if (tempArray.length == 0) {
                this.setState({ noItemErrorMsg: "" });
              }
              this.setState({
                divForSame: "",
                divForCurrentUser: "none",
                divForOtherDepts: "none",
                hideFilterDialog: false,
                dateForFilter: "none",
              });

            });

        }
        else if (this.state.forOtherDeptFilter == "Other") {
          let docProfileItems = [];
          if (this.state.isNBOAdmin != "true") {
            //not an NBO Admin
            let tempArray = [];
            let tempArraydocItems = [];
            await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
              select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
              .expand("Source,Industry,ClassOfInsurance,NBOStage,Author").orderBy("ID", false)
              .filter("Author/EMail ne '" + this.currentUserEmail + "' and Department ne '" + this.team + "'")
              .top(4000).getPaged()
              .then(async docItems => {

                tempArraydocItems.push(docItems);
                for (let listItem = 0; listItem < docItems.results.length; listItem++) {
                  if (this.state.selectedColumnKey == "OpportunityType") {
                    if (docItems.results[listItem].OpportunityType == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                    console.log(this.state.filterValue);
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems.results[listItem].EstimatedBrokerage != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "FeesIfAny") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].FeesIfAny > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].FeesIfAny < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].FeesIfAny <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].FeesIfAny >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedPremium") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedPremium > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedPremium < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedPremium <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedPremium >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }




                  }
                  else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].WeightedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].WeightedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].WeightedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].WeightedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].BrokeragePercentage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].BrokeragePercentage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].BrokeragePercentage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].BrokeragePercentage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                    let duedate = moment(docItems.results[listItem].EstimatedStartDate).toDate();
                    let toDate = moment(this.state.estimatedToStartDate).toDate();
                    let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                    duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                    toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                    fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                    if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Title") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Title.toLowerCase() == this.state.filterValue.toLowerCase()) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Comments") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Comments == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Comments != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Department") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Department == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Department != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Department.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Author") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Author.Title == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Author.Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Author.Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "ComplianceCleared") {
                    if (docItems.results[listItem].ComplianceCleared == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Industry") {
                    if (docItems.results[listItem].Industry.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }

                  }
                  else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                    if (docItems.results[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Source") {
                    if (docItems.results[listItem].Source.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "NBOStage") {
                    if (docItems.results[listItem].NBOStage.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                for (let i = 0; i < this.pageSize; i++) {
                  docProfileItems.push({
                    "ID": null,
                    "Title": null,
                    "BrokeragePercentage": null,
                    "Source": { "ID": null, "Title": null },
                    "Industry": { "ID": null, "Title": null },
                    "ClassOfInsurance": { "ID": null, "Title": null },
                    "NBOStage": { "ID": null, "Title": null },
                    "EstimatedBrokerage": null,
                    "FeesIfAny": null,
                    "Comments": null,
                    "EstimatedStartDate": null,
                    "EstimatedPremium": null,
                    "Department": null,
                    "ComplianceCleared": null,
                    "Author": { "EMail": null, "Title": null },
                    "WeightedBrokerage": null,
                    "OpportunityType": null,
                  });
                }
                docProfileItems = docProfileItems.concat(tempArray);
                console.log(docProfileItems);
                while (docItems.hasNext) {
                  docItems = await docItems.getNext();
                  docProfileItems.push(...(docItems.results));

                }
                this.sortedArray = docProfileItems;
                if (tempArray.length == 0) {
                  this.setState({
                    divForShowingPagination: "none",
                    divForNoDataFound: "",
                  });
                }
                this.setState({
                  arrayForShowingPagination: docItems.results,
                  docRepositoryItems: this.sortedArray,
                  items: this.sortedArray,
                  paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
                  noItemErrorMsg: this.sortedArray.length == 0 ? " " : "none"
                });
                console.log(this.state.docRepositoryItems);
                console.log("paginatedItems", this.state.paginatedItems);
                if (this.sortedArray.length == 0) {
                  this.setState({ noItemErrorMsg: "" });
                }

                this.setState({
                  divForSame: "none",
                  divForOtherDepts: "",
                  divForCurrentUser: "none",
                  hideFilterDialog: false,
                  dateForFilter: "none",
                });
              });
          }
          else {
            //alert(this.state.isNBOAdmin);
            let tempArray = [];
            let tempArraydocItems = [];
            await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
              select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
              .expand("Source,Industry,ClassOfInsurance,NBOStage,Author").orderBy("ID", false)
              .top(4000).getPaged().then(async docItems => {
                // docProfileItems[this.pageSize] = docItems.results;
                for (let listItem = 0; listItem < docItems.results.length; listItem++) {
                  if (this.state.selectedColumnKey == "OpportunityType") {
                    if (docItems.results[listItem].OpportunityType == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                    console.log(this.state.filterValue);
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems.results[listItem].EstimatedBrokerage != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "FeesIfAny") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].FeesIfAny > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].FeesIfAny < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].FeesIfAny <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].FeesIfAny >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedPremium") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedPremium > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedPremium < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedPremium <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedPremium >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }




                  }
                  else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].WeightedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].WeightedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].WeightedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].WeightedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].BrokeragePercentage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].BrokeragePercentage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].BrokeragePercentage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].BrokeragePercentage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                    let duedate = moment(docItems.results[listItem].EstimatedStartDate).toDate();
                    let toDate = moment(this.state.estimatedToStartDate).toDate();
                    let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                    duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                    toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                    fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                    if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Title") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Title.toLowerCase() == this.state.filterValue.toLowerCase()) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "Comments") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Comments == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Comments != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Department") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Department.toLowerCase() == this.state.filterValue.toLowerCase) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Department != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Department.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "Author") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Author.Title.toLowerCase() == this.state.filterValue.toLowerCase()) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Author.Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Author.Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "ComplianceCleared") {
                    if (docItems.results[listItem].ComplianceCleared == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Industry") {
                    if (docItems.results[listItem].Industry.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }

                  }
                  else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                    if (docItems.results[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Source") {
                    if (docItems.results[listItem].Source.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "NBOStage") {
                    if (docItems.results[listItem].NBOStage.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                for (let i = 0; i < this.pageSize; i++) {
                  docProfileItems.push({
                    "ID": null,
                    "Title": null,
                    "BrokeragePercentage": null,
                    "Source": { "ID": null, "Title": null },
                    "Industry": { "ID": null, "Title": null },
                    "ClassOfInsurance": { "ID": null, "Title": null },
                    "NBOStage": { "ID": null, "Title": null },
                    "EstimatedBrokerage": null,
                    "FeesIfAny": null,
                    "Comments": null,
                    "EstimatedStartDate": null,
                    "EstimatedPremium": null,
                    "Department": null,
                    "ComplianceCleared": null,
                    "Author": { "EMail": null, "Title": null },
                    "WeightedBrokerage": null,
                    "OpportunityType": null,
                  });
                }
                docProfileItems = docProfileItems.concat(tempArray);
                console.log(docProfileItems);
                while (docItems.hasNext) {
                  docItems = await docItems.getNext();
                  docProfileItems.push(...(docItems.results));
                }
                this.sortedArray = docProfileItems;
                if (tempArray.length == 0) {
                  this.setState({
                    divForShowingPagination: "none",
                    divForNoDataFound: "",
                  });
                }
                this.setState({
                  docRepositoryItems: this.sortedArray,
                  items: this.sortedArray,
                  paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
                  noItemErrorMsg: docProfileItems.length == 0 ? " " : "none",

                });
                console.log(this.state.docRepositoryItems);
                if (docProfileItems.length == 0) {
                  this.setState({ noItemErrorMsg: "" });
                }
              });
            this.setState({
              divForSame: "none",
              divForOtherDepts: "",
              divForCurrentUser: "none",
              hideFilterDialog: false,
              dateForFilter: "none",
            });
          }
        }
        else {
          //my nbo   

          let docProfileItems = [];
          let tempArray = [];
          let tempArraydocItems = [];
          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
            select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail eq '" + this.currentUserEmail + "'")
            .top(4000).getPaged()
            .then(async docItems => {
              this.setState({ arrayForShowingPagination: docItems.results, hideFilterDialog: false, });
              for (let listItem = 0; listItem < docItems.results.length; listItem++) {
                if (this.state.selectedColumnKey == "OpportunityType") {
                  if (docItems.results[listItem].OpportunityType == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                  console.log(this.state.filterValue);
                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].EstimatedBrokerage > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].EstimatedBrokerage < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (docItems.results[listItem].EstimatedBrokerage != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].EstimatedBrokerage <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].EstimatedBrokerage >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }

                }
                else if (this.state.selectedColumnKey == "FeesIfAny") {
                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].FeesIfAny > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].FeesIfAny < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].FeesIfAny <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].FeesIfAny >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "EstimatedPremium") {
                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].EstimatedPremium > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].EstimatedPremium < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].EstimatedPremium <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].EstimatedPremium >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }




                }
                else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].WeightedBrokerage > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].WeightedBrokerage < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].WeightedBrokerage <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].WeightedBrokerage >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }

                }
                else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].BrokeragePercentage > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].BrokeragePercentage < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].BrokeragePercentage <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].BrokeragePercentage >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                  let duedate = moment(docItems.results[listItem].EstimatedStartDate).toDate();
                  let toDate = moment(this.state.estimatedToStartDate).toDate();
                  let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                  duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                  toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                  fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                  if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                    tempArray.push(docItems.results[listItem]);
                  }

                }
                else if (this.state.selectedColumnKey == "Title") {

                  if (this.state.filterCondition == "equals") {

                    if (docItems.results[listItem].Title.toLowerCase() == this.state.filterValue.toLowerCase()) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "not equal to") {
                    if (docItems.results[listItem].Title != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "Contains") {
                    if (docItems.results[listItem].Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }

                }
                else if (this.state.selectedColumnKey == "Comments") {

                  if (this.state.filterCondition == "equals") {
                    if (docItems.results[listItem].Comments == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition != "not equal to") {
                    if (docItems.results[listItem].Comments == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "Department") {

                  if (this.state.filterCondition == "equals") {
                    if (docItems.results[listItem].Department.toLowerCase() == this.state.filterValue.toLowerCase()) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "not equal to") {
                    if (docItems.results[listItem].Department != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "Contains") {
                    if (docItems.results[listItem].Department.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "Author") {

                  if (this.state.filterCondition == "equals") {
                    if (docItems.results[listItem].Author.Title == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "not equal to") {
                    if (docItems.results[listItem].Author.Title != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "Contains") {
                    if (docItems.results[listItem].Author.Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "ComplianceCleared") {
                  if (docItems.results[listItem].ComplianceCleared == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "Industry") {
                  if (docItems.results[listItem].Industry.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }

                }
                else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                  if (docItems.results[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "Source") {
                  if (docItems.results[listItem].Source.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "NBOStage") {
                  if (docItems.results[listItem].NBOStage.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }

              }
              for (let i = 0; i < this.pageSize; i++) {
                docProfileItems.push({
                  "ID": null,
                  "Title": null,
                  "BrokeragePercentage": null,
                  "Source": { "ID": null, "Title": null },
                  "Industry": { "ID": null, "Title": null },
                  "ClassOfInsurance": { "ID": null, "Title": null },
                  "NBOStage": { "ID": null, "Title": null },
                  "EstimatedBrokerage": null,
                  "FeesIfAny": null,
                  "Comments": null,
                  "EstimatedStartDate": null,
                  "EstimatedPremium": null,
                  "Department": null,
                  "ComplianceCleared": null,
                  "Author": { "EMail": null, "Title": null },
                  "WeightedBrokerage": null,
                  "OpportunityType": null,
                });
              }
              docProfileItems = docProfileItems.concat(tempArray);
              console.log(docProfileItems);
              while (docItems.hasNext) {
                docItems = await docItems.getNext();
                docProfileItems.push(...(tempArray));
              }
              this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
              if (tempArray.length == 0) {
                this.setState({
                  divForShowingPagination: "none",
                  divForNoDataFound: "",
                });
              }
              this.setState({
                docRepositoryItems: this.sortedArray,
                items: this.sortedArray,
                paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),

              });
              console.log(this.state.docRepositoryItems);
            });
          this.setState({
            divForSame: "none",
            divForCurrentUser: "",
            divForOtherDepts: "none",
            divForDocumentUploadCompArrayDiv: "none",
            //hideFilterDialog: false,
            dateForFilter: "none",
          });
        }
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
    else {
      if (this.state.selectedColumnKey != "" && this.state.selectedColumnKey != "Select an option") {
        this.validator.hideMessages();
        //same department grid
        if (this.state.sameDepartmentItems == "Yes") {
          let docProfileItems = [];
          this.forDeptCreatedBy = "ok";
          console.log("departments of current user", this.state.oppurtunityDept);
          this.setState({
            currentItemID: "",
            forOtherDeptFilter: "MYNBOSame",
          });

          let tempArray = [];
          await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
            //select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')")
            select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
            .get().then(docItems => {
              for (let sd = 0; sd < this.state.oppurtunityDept.length; sd++) {

                for (let listItem = 0; listItem < docItems.length; listItem++) {
                  if (this.state.selectedColumnKey == "OpportunityType") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].OpportunityType == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                    console.log(this.state.filterValue);
                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "FeesIfAny") {
                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].FeesIfAny >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedPremium") {
                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].EstimatedPremium >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }




                  }
                  else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].WeightedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage > this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage < this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && Number(docItems[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage <= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].BrokeragePercentage >= this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                    let duedate = moment(docItems[listItem].EstimatedStartDate).toDate();
                    let toDate = moment(this.state.estimatedToStartDate).toDate();
                    let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                    duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                    toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                    fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                    if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Title") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Title == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Title != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Comments") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Comments == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Comments != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Department") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Department == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Department != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Author") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Author.Title == this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Author.Title != this.state.filterValue) {
                        tempArray.push(docItems[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "ComplianceCleared") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].ComplianceCleared == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Industry") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Industry.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }

                  }
                  else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Source") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].Source.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "NBOStage") {
                    if (docItems[listItem].Department == this.state.oppurtunityDept[sd].text && docItems[listItem].NBOStage.Title == this.state.filterCondition) {
                      tempArray.push(docItems[listItem]);
                    }
                  }
                }
              }
              for (let i = 0; i < this.pageSize; i++) {
                docProfileItems.push({
                  "ID": null,
                  "Title": null,
                  "BrokeragePercentage": null,
                  "Source": { "ID": null, "Title": null },
                  "Industry": { "ID": null, "Title": null },
                  "ClassOfInsurance": { "ID": null, "Title": null },
                  "NBOStage": { "ID": null, "Title": null },
                  "EstimatedBrokerage": null,
                  "FeesIfAny": null,
                  "Comments": null,
                  "EstimatedStartDate": null,
                  "EstimatedPremium": null,
                  "Department": null,
                  "ComplianceCleared": null,
                  "Author": { "EMail": null, "Title": null },
                  "WeightedBrokerage": null,
                  "OpportunityType": null,
                });
              }
              console.log("SameDept", tempArray);
              docProfileItems = docProfileItems.concat(tempArray);
              this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
              if (tempArray.length == 0) {
                this.setState({
                  divForShowingPagination: "none",
                  divForNoDataFound: "",
                });
              }
              this.setState({
                arrayForShowingPagination: tempArray,
                docRepositoryItems: this.sortedArray,
                items: this.sortedArray,
                paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
                noItemErrorMsg: tempArray.length == 0 ? " " : "none",
              });
              console.log(this.state.docRepositoryItems);
              if (tempArray.length == 0) {
                this.setState({ noItemErrorMsg: "" });
              }
              this.setState({
                divForSame: "",
                divForCurrentUser: "none",
                divForOtherDepts: "none",
                hideFilterDialog: false,
                dateForFilter: "none",
              });

            });

        }
        else if (this.state.forOtherDeptFilter == "Other") {
          let docProfileItems = [];
          if (this.state.isNBOAdmin != "true") {
            //not an NBO Admin
            let tempArray = [];
            let tempArraydocItems = [];
            await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
              select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
              .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
              .filter("Author/EMail ne '" + this.currentUserEmail + "' and Department ne '" + this.team + "'")
              .top(4000).getPaged()
              .then(async docItems => {

                tempArraydocItems.push(docItems);
                for (let listItem = 0; listItem < docItems.results.length; listItem++) {
                  if (this.state.selectedColumnKey == "OpportunityType") {
                    if (docItems.results[listItem].OpportunityType == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                    console.log(this.state.filterValue);
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems.results[listItem].EstimatedBrokerage != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "FeesIfAny") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].FeesIfAny > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].FeesIfAny < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].FeesIfAny <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].FeesIfAny >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedPremium") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedPremium > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedPremium < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedPremium <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedPremium >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }




                  }
                  else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].WeightedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].WeightedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].WeightedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].WeightedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].BrokeragePercentage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].BrokeragePercentage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].BrokeragePercentage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].BrokeragePercentage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                    let duedate = moment(docItems.results[listItem].EstimatedStartDate).toDate();
                    let toDate = moment(this.state.estimatedToStartDate).toDate();
                    let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                    duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                    toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                    fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                    if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Title") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Title == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Comments") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Comments == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Comments != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Department") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Department == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Department != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "Author") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Author.Title == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Author.Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "ComplianceCleared") {
                    if (docItems.results[listItem].ComplianceCleared == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Industry") {
                    if (docItems.results[listItem].Industry.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }

                  }
                  else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                    if (docItems.results[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Source") {
                    if (docItems.results[listItem].Source.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "NBOStage") {
                    if (docItems.results[listItem].NBOStage.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                for (let i = 0; i < this.pageSize; i++) {
                  docProfileItems.push({
                    "ID": null,
                    "Title": null,
                    "BrokeragePercentage": null,
                    "Source": { "ID": null, "Title": null },
                    "Industry": { "ID": null, "Title": null },
                    "ClassOfInsurance": { "ID": null, "Title": null },
                    "NBOStage": { "ID": null, "Title": null },
                    "EstimatedBrokerage": null,
                    "FeesIfAny": null,
                    "Comments": null,
                    "EstimatedStartDate": null,
                    "EstimatedPremium": null,
                    "Department": null,
                    "ComplianceCleared": null,
                    "Author": { "EMail": null, "Title": null },
                    "WeightedBrokerage": null,
                    "OpportunityType": null,
                  });
                }
                docProfileItems = docProfileItems.concat(tempArray);
                console.log(docProfileItems);
                while (docItems.hasNext) {
                  docItems = await docItems.getNext();
                  docProfileItems.push(...(docItems.results));

                }
                this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
                if (tempArray.length == 0) {
                  this.setState({
                    divForShowingPagination: "none",
                    divForNoDataFound: "",
                  });
                }
                this.setState({
                  arrayForShowingPagination: docItems.results,
                  docRepositoryItems: this.sortedArray,
                  items: this.sortedArray,
                  paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
                  noItemErrorMsg: this.sortedArray.length == 0 ? " " : "none"
                });
                console.log(this.state.docRepositoryItems);
                console.log("paginatedItems", this.state.paginatedItems);
                if (this.sortedArray.length == 0) {
                  this.setState({ noItemErrorMsg: "" });
                }

                this.setState({
                  divForSame: "none",
                  divForOtherDepts: "",
                  divForCurrentUser: "none",
                  hideFilterDialog: false,
                  dateForFilter: "none",
                });
              });
          }
          else {
            //alert(this.state.isNBOAdmin);
            let tempArray = [];
            let tempArraydocItems = [];
            await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
              select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,Author/Title,WeightedBrokerage,OpportunityType")
              .expand("Source,Industry,ClassOfInsurance,NBOStage,Author")
              .top(4000).getPaged().then(async docItems => {
                // docProfileItems[this.pageSize] = docItems.results;
                for (let listItem = 0; listItem < docItems.results.length; listItem++) {
                  if (this.state.selectedColumnKey == "OpportunityType") {
                    if (docItems.results[listItem].OpportunityType == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                    console.log(this.state.filterValue);
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (docItems.results[listItem].EstimatedBrokerage != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "FeesIfAny") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].FeesIfAny > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].FeesIfAny < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].FeesIfAny <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].FeesIfAny >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedPremium") {
                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].EstimatedPremium > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].EstimatedPremium < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].EstimatedPremium <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].EstimatedPremium >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }




                  }
                  else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].WeightedBrokerage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].WeightedBrokerage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].WeightedBrokerage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].WeightedBrokerage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                    if (this.state.filterCondition == ">") {
                      if (docItems.results[listItem].BrokeragePercentage > this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<") {
                      if (docItems.results[listItem].BrokeragePercentage < this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "!=") {
                      if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "<=") {
                      if (docItems.results[listItem].BrokeragePercentage <= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == ">=") {
                      if (docItems.results[listItem].BrokeragePercentage >= this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                    let duedate = moment(docItems.results[listItem].EstimatedStartDate).toDate();
                    let toDate = moment(this.state.estimatedToStartDate).toDate();
                    let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                    duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                    toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                    fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                    if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Title") {
                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Title == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "Contains") {
                      if (docItems.results[listItem].Title.toLowerCase().includes(this.state.filterValue.toLowerCase())) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }

                  }
                  else if (this.state.selectedColumnKey == "Comments") {
                    if (docItems.results[listItem].Comments == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Department") {
                    if (docItems.results[listItem].Department == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Author") {

                    if (this.state.filterCondition == "equals") {
                      if (docItems.results[listItem].Author.Title == this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                    else if (this.state.filterCondition == "not equal to") {
                      if (docItems.results[listItem].Author.Title != this.state.filterValue) {
                        tempArray.push(docItems.results[listItem]);
                      }
                    }
                  }
                  else if (this.state.selectedColumnKey == "ComplianceCleared") {
                    if (docItems.results[listItem].ComplianceCleared == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Industry") {
                    if (docItems.results[listItem].Industry.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }

                  }
                  else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                    if (docItems.results[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "Source") {
                    if (docItems.results[listItem].Source.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.selectedColumnKey == "NBOStage") {
                    if (docItems.results[listItem].NBOStage.Title == this.state.filterCondition) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                for (let i = 0; i < this.pageSize; i++) {
                  docProfileItems.push({
                    "ID": null,
                    "Title": null,
                    "BrokeragePercentage": null,
                    "Source": { "ID": null, "Title": null },
                    "Industry": { "ID": null, "Title": null },
                    "ClassOfInsurance": { "ID": null, "Title": null },
                    "NBOStage": { "ID": null, "Title": null },
                    "EstimatedBrokerage": null,
                    "FeesIfAny": null,
                    "Comments": null,
                    "EstimatedStartDate": null,
                    "EstimatedPremium": null,
                    "Department": null,
                    "ComplianceCleared": null,
                    "Author": { "EMail": null, "Title": null },
                    "WeightedBrokerage": null,
                    "OpportunityType": null,
                  });
                }
                docProfileItems = docProfileItems.concat(tempArray);
                console.log(docProfileItems);
                while (docItems.hasNext) {
                  docItems = await docItems.getNext();
                  docProfileItems.push(...(docItems.results));
                }
                this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
                if (tempArray.length == 0) {
                  this.setState({
                    divForShowingPagination: "none",
                    divForNoDataFound: "",
                  });
                }
                this.setState({
                  docRepositoryItems: this.sortedArray,
                  items: this.sortedArray,
                  paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
                  noItemErrorMsg: docProfileItems.length == 0 ? " " : "none",
                });
                console.log(this.state.docRepositoryItems);
                if (docProfileItems.length == 0) {
                  this.setState({ noItemErrorMsg: "" });
                }
              });
            this.setState({
              divForSame: "none",
              divForOtherDepts: "",
              divForCurrentUser: "none",
              hideFilterDialog: false,
              dateForFilter: "none",
            });
          }
        }
        else {
          //my nbo  

          let docProfileItems = [];
          let tempArray = [];
          let tempArraydocItems = [];
          sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
            select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail,WeightedBrokerage,OpportunityType").expand("Source,Industry,ClassOfInsurance,NBOStage,Author").filter("Author/EMail eq '" + this.currentUserEmail + "'")
            .top(4000).getPaged()
            .then(async docItems => {
              this.setState({ arrayForShowingPagination: docItems.results, hideFilterDialog: false, });
              for (let listItem = 0; listItem < docItems.results.length; listItem++) {
                if (this.state.selectedColumnKey == "OpportunityType") {
                  if (docItems.results[listItem].OpportunityType == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "EstimatedBrokerage") {
                  console.log(this.state.filterValue);
                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].EstimatedBrokerage > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].EstimatedBrokerage < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].EstimatedBrokerage).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (docItems.results[listItem].EstimatedBrokerage != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].EstimatedBrokerage <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].EstimatedBrokerage >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }

                }
                else if (this.state.selectedColumnKey == "FeesIfAny") {
                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].FeesIfAny > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].FeesIfAny < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].FeesIfAny).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].FeesIfAny <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].FeesIfAny >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "EstimatedPremium") {
                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].EstimatedPremium > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].EstimatedPremium < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].EstimatedPremium).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].EstimatedPremium <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].EstimatedPremium >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }




                }
                else if (this.state.selectedColumnKey == "WeightedBrokerage") {

                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].WeightedBrokerage > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].WeightedBrokerage < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].WeightedBrokerage).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].WeightedBrokerage <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].WeightedBrokerage >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }

                }
                else if (this.state.selectedColumnKey == "BrokeragePercentage") {

                  if (this.state.filterCondition == ">") {
                    if (docItems.results[listItem].BrokeragePercentage > this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<") {
                    if (docItems.results[listItem].BrokeragePercentage < this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "=") {
                    if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "!=") {
                    if (Number(docItems.results[listItem].BrokeragePercentage).toFixed(0) != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "<=") {
                    if (docItems.results[listItem].BrokeragePercentage <= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == ">=") {
                    if (docItems.results[listItem].BrokeragePercentage >= this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "EstimatedStartDate") {
                  let duedate = moment(docItems.results[listItem].EstimatedStartDate).toDate();
                  let toDate = moment(this.state.estimatedToStartDate).toDate();
                  let fromDate = moment(this.state.estimatedFromStartDate).toDate();
                  duedate = new Date(duedate.getFullYear(), duedate.getMonth(), duedate.getDate());
                  toDate = new Date(toDate.getFullYear(), toDate.getMonth(), toDate.getDate());
                  fromDate = new Date(fromDate.getFullYear(), fromDate.getMonth(), fromDate.getDate());
                  if (moment(duedate).isBetween(fromDate, toDate) || moment(duedate).isSame(fromDate) || moment(duedate).isSame(toDate)) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "Title") {

                  if (this.state.filterCondition == "equals") {
                    if (docItems.results[listItem].Title == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "not equal to") {
                    if (docItems.results[listItem].Title != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }

                }
                else if (this.state.selectedColumnKey == "Comments") {

                  if (this.state.filterCondition == "equals") {
                    if (docItems.results[listItem].Comments == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition != "not equal to") {
                    if (docItems.results[listItem].Comments == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "Department") {

                  if (this.state.filterCondition == "equals") {
                    if (docItems.results[listItem].Department == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "not equal to") {
                    if (docItems.results[listItem].Department != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "Author") {

                  if (this.state.filterCondition == "equals") {
                    if (docItems.results[listItem].Author.Title == this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                  else if (this.state.filterCondition == "not equal to") {
                    if (docItems.results[listItem].Author.Title != this.state.filterValue) {
                      tempArray.push(docItems.results[listItem]);
                    }
                  }
                }
                else if (this.state.selectedColumnKey == "ComplianceCleared") {
                  if (docItems.results[listItem].ComplianceCleared == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "Industry") {
                  if (docItems.results[listItem].Industry.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }

                }
                else if (this.state.selectedColumnKey == "ClassOfInsurance") {
                  if (docItems.results[listItem].ClassOfInsurance.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "Source") {
                  if (docItems.results[listItem].Source.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }
                else if (this.state.selectedColumnKey == "NBOStage") {
                  if (docItems.results[listItem].NBOStage.Title == this.state.filterCondition) {
                    tempArray.push(docItems.results[listItem]);
                  }
                }

              }

              // alert(tempArray.length);

              for (let i = 0; i < this.pageSize; i++) {
                docProfileItems.push({
                  "ID": null,
                  "Title": null,
                  "BrokeragePercentage": null,
                  "Source": { "ID": null, "Title": null },
                  "Industry": { "ID": null, "Title": null },
                  "ClassOfInsurance": { "ID": null, "Title": null },
                  "NBOStage": { "ID": null, "Title": null },
                  "EstimatedBrokerage": null,
                  "FeesIfAny": null,
                  "Comments": null,
                  "EstimatedStartDate": null,
                  "EstimatedPremium": null,
                  "Department": null,
                  "ComplianceCleared": null,
                  "Author": { "EMail": null, "Title": null },
                  "WeightedBrokerage": null,
                  "OpportunityType": null,
                });
              }
              docProfileItems = docProfileItems.concat(tempArray);
              console.log(docProfileItems);
              while (docItems.hasNext) {
                docItems = await docItems.getNext();
                docProfileItems.push(...(tempArray));
              }

              this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
              if (tempArray.length == 0) {
                this.setState({
                  divForShowingPagination: "none",
                  divForNoDataFound: "",
                });
              }
              this.setState({
                docRepositoryItems: this.sortedArray,
                items: this.sortedArray,
                paginatedItems: this.sortedArray.slice(this.pageSize, this.pageSize + this.pageSize),
              });
              console.log(this.state.docRepositoryItems);

            });

          this.setState({
            divForSame: "none",
            divForCurrentUser: "",
            divForOtherDepts: "none",
            divForDocumentUploadCompArrayDiv: "none",
            hideFilterDialog: false,
            dateForFilter: "none",
          });
        }
      }
      else {
        this.validator.showMessages();
        this.forceUpdate();
      }
    }
  }
  private _onFilterForModal = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    if (text == "") {
      this.loadDocProfile();
    }
    else {
      let dataToFilter = this.state.selectedColumnKey;
      this.setState({
        paginatedItems: text ? this.state.paginatedItems.filter(i => i.Title.toLowerCase().indexOf(text.toString().toLowerCase()) > -1) : this.state.paginatedItems,
      });
    }
  }
  public render(): React.ReactElement<INboDetailListProps> {


    const menuPropsFilter: IContextualMenuProps = {
      items: [
        {
          key: 'ColumnFilter',
          text: 'Column Filter',
          iconProps: { iconName: 'Filter' },
          onClick: this._onFilter,
        },
        {
          key: 'ClearFilters',
          text: 'Clear Filters',
          iconProps: { iconName: 'ClearFilter' },
          onClick: this._filterPanelCloseButton,
        },
      ],
    };
    const menuProps: IContextualMenuProps = {
      items: [
        {
          key: 'MyNBO',
          text: 'My NBO',
          iconProps: { iconName: 'AccountManagement' },
          onClick: this.loadDocProfile,
        },
        {
          key: 'sameDept',
          text: 'My Departments',
          iconProps: { iconName: 'People' },
          onClick: this.sameDept,
        },
        {
          key: 'others',
          text: 'Others',
          iconProps: { iconName: 'AddGroup' },
          onClick: this.others,
        },
      ],
    };
    const ComplianceCleared: IDropdownOption[] = [
      { key: 'No', text: 'No' },
      { key: 'Yes', text: 'Yes' },
      { key: 'Pending', text: 'Pending' },
    ];
    const OpportunityType: IDropdownOption[] = [
      { key: 'New', text: 'New' },
      { key: 'Expanded', text: 'Expanded' },
    ];

    const MyNBOFilterColumns: IDropdownOption[] = [
      { key: 'Select an option', text: 'Select an option' },
      { key: 'OpportunityType', text: 'OpportunityType', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'Source', text: 'Source' },
      { key: 'Industry', text: 'Industry' },
      { key: 'Title', text: 'Client Name' },
      { key: 'ClassOfInsurance', text: 'Class Of Insurance' },
      { key: 'EstimatedStartDate', text: 'Estimated Start Date', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      // { key: 'Comments', text: 'Comments', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'EstimatedPremium', text: 'Estimated Premium', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'BrokeragePercentage', text: 'Brokerage Percentage', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'EstimatedBrokerage', text: 'Estimated Brokerage', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'NBOStage', text: 'NBO Stage', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'WeightedBrokerage', text: 'Weighted Brokerage', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'FeesIfAny', text: 'Fees If Any', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'ComplianceCleared', text: 'Compliance Cleared', hidden: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" || this.state.forOtherDeptFilter == "MYNBOSame") ? false : true },
      { key: 'Department', text: 'Department', hidden: (this.forDeptCreatedBy == "ok") ? false : true },
      { key: 'Author', text: 'Created By', hidden: (this.forDeptCreatedBy == "ok") ? false : true },
    ];
    const DeleteIcon: IIconProps = { iconName: 'Delete' };
    const ShowDocuments: IIconProps = { iconName: 'DocumentSet' };
    const NBODetails: IIconProps = { iconName: 'BulletedListMirrored' };
    const midBar: IIconProps = { iconName: 'BulletedListBulletMirrored' };
    return (
      <div className={styles.nboDetailList}>

        <><div className={styles.nboDetailList} style={{ display: this.state.displayWithOutQuery }}>

          <div style={{ display: "flex" }}>
            <div>
              <CommandButton
                iconProps={AddIcon} onClick={this._addNBO}
                text="New"
                primary={true}
                style={{ color: "#25ddd0" }}
                split />

            </div>
            {/* <CommandButton
      iconProps={midBar}
      style={{ color: "#25ddd0" }}
      split
    /> */}
            <div>
              <CommandButton
                iconProps={NBODetails}
                text="NBO Details"
                primary
                split
                splitButtonAriaLabel="See 2 options"
                aria-roledescription="split button"
                menuProps={menuProps}
                style={{ color: "#25ddd0" }} />
            </div>
            <div>
              <CommandButton
                iconProps={FilterIcon}
                text="Filter"
                primary
                split
                splitButtonAriaLabel="See 2 options"
                aria-roledescription="split button"
                menuProps={menuPropsFilter}
                style={{ color: "#25ddd0" }} />
            </div>
          </div>

          <div style={{ display: this.state.deleteMessageBar }}>
            {/* Show Message bar for Notification*/}
            {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''}
          </div>
          {/* currenUserdiv */}
          <div style={{ display: this.state.divForCurrentUser, marginTop: "10px", overflowX: "auto" }}>
            {/* <SearchBox placeholder="Type application name" className={styles['ms-SearchBox']} onSearch={newValue => console.log('value is ' + newValue)} onChange={this._onFilterForModal} /> */}
            <div className={styles.Heading} style={{ display: this.state.sameDepartmentItems == "Yes" ? "none" : "" }}> My NBO Pipeline</div>
            <div style={{ display: this.state.docRepositoryItems.length == 0 ? "" : "none", color: "#f4f4f4", textAlign: "center" }}> <h1>No items</h1></div>

            <table style={{ overflowX: "scroll", display: this.state.docRepositoryItems.length == 0 ? "none" : "" }}>
              <tr style={{ background: "#f4f4f4" }}>
                <th style={{ padding: "5px 10px", }}>Edit</th>
                {/* <th style={{ padding: "5px 10px", }}>Delete</th> */}
                <th style={{ padding: "5px 10px", }}>View Documents </th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>
                    Opportunity Type
                    <IconButton id="OpportunityType" style={{ color: "Black", display: this.state.sortOppurtunityTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="OpportunityType" onClick={(e) => this._onSortClickAscForMyNBO('OpportunityType', e)} />
                    <IconButton id="OpportunityType" style={{ color: "Black", display: this.state.sortOppurtunityTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="OpportunityType" onClick={(e) => this._onSortClickDescForMyNBO('OpportunityType', e)} />

                  </div>
                </th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Source
                    <IconButton id="Source" style={{ color: "Black", display: this.state.SourceTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Source" onClick={(e) => this._onSortClickAscForMyNBO('Source', e)} />
                    <IconButton id="Source" style={{ color: "Black", display: this.state.SourceTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Source" onClick={(e) => this._onSortClickDescForMyNBO('Source', e)} />

                  </div></th>

                <th style={{ padding: "5px 10px" }}>
                  <div style={{ display: "flex" }}> Industry
                    <IconButton id="Industry" style={{ color: "Black", display: this.state.IndustryTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Industry" onClick={(e) => this._onSortClickAscForMyNBO('Industry', e)} />
                    <IconButton id="Industry" style={{ color: "Black", display: this.state.IndustryTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Industry" onClick={(e) => this._onSortClickDescForMyNBO('Industry', e)} />
                  </div>
                </th>
                <th style={{ padding: "5px 10px" }}>
                  <div style={{ display: "flex" }}>Client Name
                    <IconButton id="ClientName" style={{ color: "Black", display: this.state.ClientNameTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Title" onClick={(e) => this._onSortClickAscForMyNBO('Title', e)} />
                    <IconButton id="ClientName" style={{ color: "Black", display: this.state.ClientNameTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Title" onClick={(e) => this._onSortClickDescForMyNBO('Title', e)} />
                  </div> </th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Class of Insurance
                    <IconButton id="ClassOfInsurance" style={{ color: "Black", display: this.state.ClassOfInsuranceTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="ClassOfInsurance" onClick={(e) => this._onSortClickAscForMyNBO('ClassOfInsurance', e)} />
                    <IconButton id="ClassOfInsurance" style={{ color: "Black", display: this.state.ClassOfInsuranceTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="ClassOfInsurance" onClick={(e) => this._onSortClickDescForMyNBO('ClassOfInsurance', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Est Start Date
                    <IconButton id="EstimatedStartDate" style={{ color: "Black", display: this.state.EstStartDateTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedStartDate" onClick={(e) => this._onSortClickAscForMyNBO('EstimatedStartDate', e)} />
                    <IconButton id="EstimatedStartDate" style={{ color: "Black", display: this.state.EstStartDateTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedStartDate" onClick={(e) => this._onSortClickDescForMyNBO('EstimatedStartDate', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Comments
                    <IconButton id="Comments" style={{ color: "Black", display: this.state.CommentsTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Comments" onClick={(e) => this._onSortClickAscForMyNBO('Comments', e)} />
                    <IconButton id="Comments" style={{ color: "Black", display: this.state.CommentsTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Comments" onClick={(e) => this._onSortClickDescForMyNBO('Comments', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Estimated Premium
                    <IconButton id="EstimatedPremium" style={{ color: "Black", display: this.state.EstimatedPremiumTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedPremium" onClick={(e) => this._onSortClickAscForMyNBO('EstimatedPremium', e)} />
                    <IconButton id="EstimatedPremium" style={{ color: "Black", display: this.state.EstimatedPremiumTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedPremium" onClick={(e) => this._onSortClickDescForMyNBO('EstimatedPremium', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px" }}>
                  <div style={{ display: "flex" }}>Brokerage %
                    <IconButton id="Brokerage" style={{ color: "Black", display: this.state.BrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="BrokeragePercentage" onClick={(e) => this._onSortClickAscForMyNBO('BrokeragePercentage', e)} />
                    <IconButton id="Brokerage" style={{ color: "Black", display: this.state.BrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="BrokeragePercentage" onClick={(e) => this._onSortClickDescForMyNBO('BrokeragePercentage', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px" }}>
                  <div style={{ display: "flex" }}>Estimated Brokerage
                    <IconButton id="EstimatedBrokerage" style={{ color: "Black", display: this.state.EstimatedBrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedBrokerage" onClick={(e) => this._onSortClickAscForMyNBO('EstimatedBrokerage', e)} />
                    <IconButton id="EstimatedBrokerage" style={{ color: "Black", display: this.state.EstimatedBrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedBrokerage" onClick={(e) => this._onSortClickDescForMyNBO('EstimatedBrokerage', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>NBO Stage
                    <IconButton id="NBOStage" style={{ color: "Black", display: this.state.NBOStageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="NBOStage" onClick={(e) => this._onSortClickAscForMyNBO('NBOStage', e)} />
                    <IconButton id="NBOStage" style={{ color: "Black", display: this.state.NBOStageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="NBOStage" onClick={(e) => this._onSortClickDescForMyNBO('NBOStage', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Weighted Brokerage
                    <IconButton id="WeightedBrokerage" style={{ color: "Black", display: this.state.WeightedBrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="WeightedBrokerage" onClick={(e) => this._onSortClickAscForMyNBO('WeightedBrokerage', e)} />
                    <IconButton id="WeightedBrokerage" style={{ color: "Black", display: this.state.WeightedBrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="WeightedBrokerage" onClick={(e) => this._onSortClickDescForMyNBO('WeightedBrokerage', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Fees If Any
                    <IconButton id="FeesIfAny" style={{ color: "Black", display: this.state.FeesIfAnyTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="FeesIfAny" onClick={(e) => this._onSortClickAscForMyNBO('FeesIfAny', e)} />
                    <IconButton id="FeesIfAny" style={{ color: "Black", display: this.state.FeesIfAnyTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="FeesIfAny" onClick={(e) => this._onSortClickDescForMyNBO('FeesIfAny', e)} />
                  </div></th>
                <th style={{ padding: "5px 10px", }}>
                  <div style={{ display: "flex" }}>Compliance Cleared
                    <IconButton id="ComplianceCleared" style={{ color: "Black", display: this.state.ComplianceClearedTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="ComplianceCleared" onClick={(e) => this._onSortClickAscForMyNBO('ComplianceCleared', e)} />
                    <IconButton id="ComplianceCleared" style={{ color: "Black", display: this.state.ComplianceClearedTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="ComplianceCleared" onClick={(e) => this._onSortClickDescForMyNBO('ComplianceCleared', e)} />
                  </div></th>

              </tr>
              {this.state.paginatedItems.map((items, key) => {
                return (
                  <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                    <td style={{ padding: "5px 10px", }}><IconButton iconProps={EditIcon} title="Edit" ariaLabel="Delete" onClick={() => this.onEditClick(items)} disabled={items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? false : true} /></td>
                    {/* <td style={{ padding: "5px 10px", }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this.onDeleteClick(items)} disabled={items.Author.EMail == this.currentUserEmail ? false : true} /></td> */}
                    <td style={{ padding: "5px 10px", }}><IconButton
                      iconProps={ShowDocuments} onClick={() => this.openCCSPopUp(items)}
                      text="View Documents"
                      disabled={items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? false : true} /></td>
                    <td style={{ padding: "5px 10px" }}> {items.OpportunityType}  </td>
                    <td style={{ padding: "5px 10px" }}> {items.Source.Title}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Industry.Title}  </td>
                    <td style={{ padding: "5px 10px" }}> {items.Title} </td>
                    <td style={{ padding: "5px 10px", }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? items.ClassOfInsurance.Title : " "}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? moment(items.EstimatedStartDate).format("DD/MM/YYYY") : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? items.Comments : " "} </td>
                    <td style={{ padding: "5px 10px", }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? Number(items.EstimatedPremium).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Author.EMail != null && items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? items.BrokeragePercentage + " %" : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? Number(items.EstimatedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? items.NBOStage.Title : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? Number(items.WeightedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? Number(items.FeesIfAny).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px", }}>{items.Author.EMail != null && items.Author.EMail == this.currentUserEmail ? items.ComplianceCleared : " "} </td>
                  </tr>
                );
              })}
            </table>
            <div className={styles.NoDataFound} style={{ display: this.state.divForNoDataFound }}> No Record Found</div>
            <div style={{ display: this.state.divForShowingPagination }}>
              <Pagination
                currentPage={0}
                totalPages={(this.sortedArray.length / this.pageSize) - 1}
                onChange={(page) => this._getPage(page)}
                limiterIcon={"Emoji12"} // Optional
              />
            </div>
          </div>
          {/* My Departments */}
          <div style={{ display: this.state.divForSame, marginTop: "10px", overflowX: "auto" }}>
            <div className={styles.Heading} style={{ display: this.state.sameDepartmentItems == "Yes" ? "" : "none" }}> My Departments NBO Pipeline</div>
            <div style={{ display: this.state.docRepositoryItems.length == 0 ? "" : "none", color: "#f4f4f4", textAlign: "center" }}> <h1>No items</h1></div>
            <table style={{ display: this.state.docRepositoryItems.length == 0 ? "none" : "", }}>
              <tr style={{ background: "#f4f4f4" }}>
                <th style={{ padding: "5px 10px", }}>Edit</th>
                <th style={{ padding: "5px 10px", }}>View Documents</th>
                <th style={{ padding: "5px 10px" }}> <div style={{ display: "flex" }}>
                  Opportunity Type
                  <IconButton id="OpportunityType" style={{ color: "Black", display: this.state.sortOppurtunityTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="OpportunityType" onClick={(e) => this._onSortClickAscForSameDept('OpportunityType', e)} />
                  <IconButton id="OpportunityType" style={{ color: "Black", display: this.state.sortOppurtunityTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="OpportunityType" onClick={(e) => this._onSortClickDescForSameDept('OpportunityType', e)} />

                </div></th>
                <th style={{ padding: "5px 10px" }}>  <div style={{ display: "flex" }}>Source
                  <IconButton id="Source" style={{ color: "Black", display: this.state.SourceTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Source" onClick={(e) => this._onSortClickAscForSameDept('Source', e)} />
                  <IconButton id="Source" style={{ color: "Black", display: this.state.SourceTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Source" onClick={(e) => this._onSortClickDescForSameDept('Source', e)} />

                </div></th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}> Industry
                  <IconButton id="Industry" style={{ color: "Black", display: this.state.IndustryTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Industry" onClick={(e) => this._onSortClickAscForSameDept('Industry', e)} />
                  <IconButton id="Industry" style={{ color: "Black", display: this.state.IndustryTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Industry" onClick={(e) => this._onSortClickDescForSameDept('Industry', e)} />
                </div></th>
                <th style={{ padding: "5px 10px" }}> <div style={{ display: "flex" }}>Client Name
                  <IconButton id="ClientName" style={{ color: "Black", display: this.state.ClientNameTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Title" onClick={(e) => this._onSortClickAscForSameDept('Title', e)} />
                  <IconButton id="ClientName" style={{ color: "Black", display: this.state.ClientNameTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Title" onClick={(e) => this._onSortClickDescForSameDept('Title', e)} />
                </div> </th>
                <th style={{ padding: "5px 10px", }}><div style={{ display: "flex" }}>Class of Insurance
                  <IconButton id="ClassOfInsurance" style={{ color: "Black", display: this.state.ClassOfInsuranceTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="ClassOfInsurance" onClick={(e) => this._onSortClickAscForSameDept('ClassOfInsurance', e)} />
                  <IconButton id="ClassOfInsurance" style={{ color: "Black", display: this.state.ClassOfInsuranceTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="ClassOfInsurance" onClick={(e) => this._onSortClickDescForSameDept('ClassOfInsurance', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}> <div style={{ display: "flex" }}>Est Start Date
                  <IconButton id="EstimatedStartDate" style={{ color: "Black", display: this.state.EstStartDateTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedStartDate" onClick={(e) => this._onSortClickAscForSameDept('EstimatedStartDate', e)} />
                  <IconButton id="EstimatedStartDate" style={{ color: "Black", display: this.state.EstStartDateTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedStartDate" onClick={(e) => this._onSortClickDescForSameDept('EstimatedStartDate', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}> <div style={{ display: "flex" }}>Comments
                  <IconButton id="Comments" style={{ color: "Black", display: this.state.CommentsTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Comments" onClick={(e) => this._onSortClickAscForSameDept('Comments', e)} />
                  <IconButton id="Comments" style={{ color: "Black", display: this.state.CommentsTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Comments" onClick={(e) => this._onSortClickDescForSameDept('Comments', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}><div style={{ display: "flex" }}>Estimated Premium
                  <IconButton id="EstimatedPremium" style={{ color: "Black", display: this.state.EstimatedPremiumTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedPremium" onClick={(e) => this._onSortClickAscForSameDept('EstimatedPremium', e)} />
                  <IconButton id="EstimatedPremium" style={{ color: "Black", display: this.state.EstimatedPremiumTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedPremium" onClick={(e) => this._onSortClickDescForSameDept('EstimatedPremium', e)} />
                </div></th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}>Brokerage %
                  <IconButton id="Brokerage" style={{ color: "Black", display: this.state.BrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="BrokeragePercentage" onClick={(e) => this._onSortClickAscForSameDept('BrokeragePercentage', e)} />
                  <IconButton id="Brokerage" style={{ color: "Black", display: this.state.BrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="BrokeragePercentage" onClick={(e) => this._onSortClickDescForSameDept('BrokeragePercentage', e)} />
                </div></th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}>Estimated Brokerage
                  <IconButton id="EstimatedBrokerage" style={{ color: "Black", display: this.state.EstimatedBrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedBrokerage" onClick={(e) => this._onSortClickAscForSameDept('EstimatedBrokerage', e)} />
                  <IconButton id="EstimatedBrokerage" style={{ color: "Black", display: this.state.EstimatedBrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedBrokerage" onClick={(e) => this._onSortClickDescForSameDept('EstimatedBrokerage', e)} />
                </div></th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}>NBO Stage
                  <IconButton id="NBOStage" style={{ color: "Black", display: this.state.NBOStageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="NBOStage" onClick={(e) => this._onSortClickAscForSameDept('NBOStage', e)} />
                  <IconButton id="NBOStage" style={{ color: "Black", display: this.state.NBOStageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="NBOStage" onClick={(e) => this._onSortClickDescForSameDept('NBOStage', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}> <div style={{ display: "flex" }}>Weighted Brokerage
                  <IconButton id="WeightedBrokerage" style={{ color: "Black", display: this.state.WeightedBrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="WeightedBrokerage" onClick={(e) => this._onSortClickAscForSameDept('WeightedBrokerage', e)} />
                  <IconButton id="WeightedBrokerage" style={{ color: "Black", display: this.state.WeightedBrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="WeightedBrokerage" onClick={(e) => this._onSortClickDescForSameDept('WeightedBrokerage', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}> <div style={{ display: "flex" }}>Fees If Any
                  <IconButton id="FeesIfAny" style={{ color: "Black", display: this.state.FeesIfAnyTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="FeesIfAny" onClick={(e) => this._onSortClickAscForSameDept('FeesIfAny', e)} />
                  <IconButton id="FeesIfAny" style={{ color: "Black", display: this.state.FeesIfAnyTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="FeesIfAny" onClick={(e) => this._onSortClickDescForSameDept('FeesIfAny', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}><div style={{ display: "flex" }}>Compliance Cleared
                  <IconButton id="ComplianceCleared" style={{ color: "Black", display: this.state.ComplianceClearedTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="ComplianceCleared" onClick={(e) => this._onSortClickAscForSameDept('ComplianceCleared', e)} />
                  <IconButton id="ComplianceCleared" style={{ color: "Black", display: this.state.ComplianceClearedTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="ComplianceCleared" onClick={(e) => this._onSortClickDescForSameDept('ComplianceCleared', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}> <div style={{ display: "flex" }}>Department
                  <IconButton id="Department" style={{ color: "Black", display: this.state.DepartmentTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Department" onClick={(e) => this._onSortClickAscForSameDept('Department', e)} />
                  <IconButton id="Department" style={{ color: "Black", display: this.state.DepartmentTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Department" onClick={(e) => this._onSortClickDescForSameDept('Department', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}><div style={{ display: "flex" }}>Created By
                  <IconButton id="CreatedBy" style={{ color: "Black", display: this.state.CreatedByTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Author" onClick={(e) => this._onSortClickAscForSameDept('Author', e)} />
                  <IconButton id="CreatedBy" style={{ color: "Black", display: this.state.CreatedByTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Author" onClick={(e) => this._onSortClickDescForSameDept('Author', e)} />
                </div></th>

              </tr>
              {this.state.paginatedItems.map((items, key) => {
                return (
                  <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                    <td style={{ padding: "5px 10px", }}><IconButton iconProps={EditIcon} title="Edit" ariaLabel="Delete" onClick={() => this.onEditClick(items)} disabled={items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? false : true} /></td>
                    <td style={{ padding: "5px 10px", }}><IconButton
                      iconProps={ShowDocuments} onClick={() => this.openCCSPopUp(items)}
                      text="View Documents"
                      disabled={items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? false : true} /></td>
                    <td style={{ padding: "5px 10px" }}> {items.OpportunityType}  </td>
                    <td style={{ padding: "5px 10px" }}> {items.Source.Title}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Industry.Title}  </td>
                    <td style={{ padding: "5px 10px" }}> {items.Title} </td>
                    <td style={{ padding: "5px 10px", }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? items.ClassOfInsurance.Title : " "}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? moment(items.EstimatedStartDate).format("DD/MM/YYYY") : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? items.Comments : " "} </td>
                    <td style={{ padding: "5px 10px", }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? Number(items.EstimatedPremium).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? items.BrokeragePercentage + " %" : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? Number(items.EstimatedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? items.NBOStage.Title : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? Number(items.WeightedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? Number(items.FeesIfAny).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px", }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? items.ComplianceCleared : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? items.Department : " "} </td>
                    <td style={{ padding: "5px 10px", }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "NBO Admin Team" || this.teamType == "Compliance Team" ? items.Author.Title : " "} </td>

                  </tr>
                );
              })}
            </table>
            <div className={styles.NoDataFound} style={{ display: this.state.divForNoDataFound }}> No Record Found</div>
            <div style={{ display: this.state.divForShowingPagination }}>
              <Pagination
                currentPage={0}
                totalPages={(this.sortedArray.length / this.pageSize) - 1}
                onChange={(page) => this._getPage(page)}
                limiter={10}
                limiterIcon={"Emoji12"} // Optional
              />
            </div>
          </div>
          {/* divForOtherDepts */}

          {/* <div style={{ display: this.state.divForOtherDepts, marginTop: "10px", overflowX: "auto" }}>
            <div className={styles.Heading}> Other Departments NBO Pipeline</div>
            <div style={{ display: this.state.docRepositoryItems.length == 0 ? "" : "none", color: "#f4f4f4", textAlign: "center" }}> <h1>No items</h1></div>
            <table style={{ overflowX: "scroll", display: this.state.docRepositoryItems.length == 0 ? "none" : "", marginLeft: "auto", marginRight: "auto" }}>
              <tr style={{ background: "#f4f4f4" }}>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Edit</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>View Documents</th>
                <th style={{ padding: "5px 10px" }}>Client Name</th>
                <th style={{ padding: "5px 10px" }}>Source</th>
                <th style={{ padding: "5px 10px" }}>Industry</th>
                <th style={{ padding: "5px 10px" }}>Department</th>
                <th style={{ padding: "5px 10px", }}>Class of Insurance</th>
                <th style={{ padding: "5px 10px", }}>Created By</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Comments</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Class of Insurance</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Estimated Premium</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Brokerage %</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Estimated Brokerage</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Est Start Date</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Fees If Any</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>NBO stage</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Compliance Cleared</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Created By</th>

              </tr>
              {this.state.paginatedItems.map((items, key) => {
                return (
                  <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><IconButton iconProps={EditIcon} title="Edit" ariaLabel="Delete" onClick={() => this.onEditClick(items)}
                      disabled={this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? false : true} /></td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><IconButton
                      iconProps={ShowDocuments} onClick={() => this.openCCSPopUp(items)}
                      text="View Documents"
                      disabled={this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? false : true} /></td>
                    <td style={{ padding: "5px 10px" }}> {items.Title} </td>
                    <td style={{ padding: "5px 10px" }}> {items.Source.Title}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Industry.Title}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department}  </td>
                    <td style={{ padding: "5px 10px", }}>{items.ClassOfInsurance.Title}  </td>
                    <td style={{ padding: "5px 10px", }}>{items.Author.Title} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.Comments : " "}  </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.ClassOfInsurance.Title : " "}  </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.EstimatedPremium : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.BrokeragePercentage + " %" : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? Number(items.EstimatedBrokerage).toFixed(2) : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? moment(items.EstimatedStartDate).format("DD/MM/YYYY") : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.FeesIfAny : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.NBOStage.Title : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.ComplianceCleared : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.Author.Title : " "} </td>
                  </tr>
                );
              })}
            </table>
            <div style={{ display: this.state.docRepositoryItems.length >= this.pageSize ? "" : "none" }}>
              <Pagination
                currentPage={1}
                totalPages={(this.sortedArray.length / this.pageSize) - 1}
                onChange={(page) => this._getPage(page)}
                limiter={100}
                limiterIcon={"Emoji12"} // Optional
              />
            </div>
          </div> */}
          <div style={{ display: this.state.divForOtherDepts, marginTop: "10px", overflowX: "auto" }}>
            <div className={styles.Heading}> Other Departments NBO Pipeline</div>
            <div style={{ display: this.state.docRepositoryItems.length == 0 ? "" : "none", color: "#f4f4f4", textAlign: "center" }}> <h1>No items</h1></div>
            <table style={{ overflowX: "scroll", display: this.state.docRepositoryItems.length == 0 ? "none" : "", marginLeft: "auto", marginRight: "auto" }}>
              <tr style={{ background: "#f4f4f4" }}>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>Edit</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" ? " " : "none") }}>Delete</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>View Documents</th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><div style={{ display: "flex" }}>
                  Opportunity Type
                  <IconButton id="OpportunityType" style={{ color: "Black", display: this.state.sortOppurtunityTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="OpportunityType" onClick={(e) => this._onSortClickAscForOtherDept('OpportunityType', e)} />
                  <IconButton id="OpportunityType" style={{ color: "Black", display: this.state.sortOppurtunityTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="OpportunityType" onClick={(e) => this._onSortClickDescForOtherDept('OpportunityType', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}>Source
                  <IconButton id="Source" style={{ color: "Black", display: this.state.SourceTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Source" onClick={(e) => this._onSortClickAscForOtherDept('Source', e)} />
                  <IconButton id="Source" style={{ color: "Black", display: this.state.SourceTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Source" onClick={(e) => this._onSortClickDescForOtherDept('Source', e)} />

                </div>
                </th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}> Industry
                  <IconButton id="Industry" style={{ color: "Black", display: this.state.IndustryTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Industry" onClick={(e) => this._onSortClickAscForOtherDept('Industry', e)} />
                  <IconButton id="Industry" style={{ color: "Black", display: this.state.IndustryTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Industry" onClick={(e) => this._onSortClickDescForOtherDept('Industry', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}>Client Name
                  <IconButton id="Title" style={{ color: "Black", display: this.state.ClientNameTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Title" onClick={(e) => this._onSortClickAscForOtherDept('Title', e)} />
                  <IconButton id="Title" style={{ color: "Black", display: this.state.ClientNameTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Title" onClick={(e) => this._onSortClickDescForOtherDept('Title', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", }}><div style={{ display: "flex" }}>Class of Insurance
                  <IconButton id="ClassOfInsurance" style={{ color: "Black", display: this.state.ClassOfInsuranceTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="ClassOfInsurance" onClick={(e) => this._onSortClickAscForOtherDept('ClassOfInsurance', e)} />
                  <IconButton id="ClassOfInsurance" style={{ color: "Black", display: this.state.ClassOfInsuranceTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="ClassOfInsurance" onClick={(e) => this._onSortClickDescForOtherDept('ClassOfInsurance', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><div style={{ display: "flex" }}>Est Start Date
                  <IconButton id="EstimatedStartDate" style={{ color: "Black", display: this.state.EstStartDateTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedStartDate" onClick={(e) => this._onSortClickAscForOtherDept('EstimatedStartDate', e)} />
                  <IconButton id="EstimatedStartDate" style={{ color: "Black", display: this.state.EstStartDateTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedStartDate" onClick={(e) => this._onSortClickDescForOtherDept('EstimatedStartDate', e)} />
                </div>

                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}> <div style={{ display: "flex" }}>Comments
                  <IconButton id="Comments" style={{ color: "Black", display: this.state.CommentsTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Comments" onClick={(e) => this._onSortClickAscForOtherDept('Comments', e)} />
                  <IconButton id="Comments" style={{ color: "Black", display: this.state.CommentsTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Comments" onClick={(e) => this._onSortClickDescForOtherDept('Comments', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><div style={{ display: "flex" }}>Estimated Premium
                  <IconButton id="EstimatedPremium" style={{ color: "Black", display: this.state.EstimatedPremiumTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedPremium" onClick={(e) => this._onSortClickAscForOtherDept('EstimatedPremium', e)} />
                  <IconButton id="EstimatedPremium" style={{ color: "Black", display: this.state.EstimatedPremiumTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedPremium" onClick={(e) => this._onSortClickDescForOtherDept('EstimatedPremium', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><div style={{ display: "flex" }}>Brokerage %
                  <IconButton id="Brokerage" style={{ color: "Black", display: this.state.BrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="BrokeragePercentage" onClick={(e) => this._onSortClickAscForOtherDept('BrokeragePercentage', e)} />
                  <IconButton id="Brokerage" style={{ color: "Black", display: this.state.BrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="BrokeragePercentage" onClick={(e) => this._onSortClickDescForOtherDept('BrokeragePercentage', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><div style={{ display: "flex" }}>Estimated Brokerage
                  <IconButton id="EstimatedBrokerage" style={{ color: "Black", display: this.state.EstimatedBrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="EstimatedBrokerage" onClick={(e) => this._onSortClickAscForOtherDept('EstimatedBrokerage', e)} />
                  <IconButton id="EstimatedBrokerage" style={{ color: "Black", display: this.state.EstimatedBrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="EstimatedBrokerage" onClick={(e) => this._onSortClickDescForOtherDept('EstimatedBrokerage', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><div style={{ display: "flex" }}>NBO Stage
                  <IconButton id="NBOStage" style={{ color: "Black", display: this.state.NBOStageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="NBOStage" onClick={(e) => this._onSortClickAscForOtherDept('NBOStage', e)} />
                  <IconButton id="NBOStage" style={{ color: "Black", display: this.state.NBOStageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="NBOStage" onClick={(e) => this._onSortClickDescForOtherDept('NBOStage', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><div style={{ display: "flex" }}>Weighted Brokerage
                  <IconButton id="WeightedBrokerage" style={{ color: "Black", display: this.state.WeightedBrokerageTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="WeightedBrokerage" onClick={(e) => this._onSortClickAscForOtherDept('WeightedBrokerage', e)} />
                  <IconButton id="WeightedBrokerage" style={{ color: "Black", display: this.state.WeightedBrokerageTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="WeightedBrokerage" onClick={(e) => this._onSortClickDescForOtherDept('WeightedBrokerage', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}> <div style={{ display: "flex" }}>Fees If Any
                  <IconButton id="FeesIfAny" style={{ color: "Black", display: this.state.FeesIfAnyTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="FeesIfAny" onClick={(e) => this._onSortClickAscForOtherDept('FeesIfAny', e)} />
                  <IconButton id="FeesIfAny" style={{ color: "Black", display: this.state.FeesIfAnyTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="FeesIfAny" onClick={(e) => this._onSortClickDescForOtherDept('FeesIfAny', e)} />
                </div>
                </th>
                <th style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}> <div style={{ display: "flex" }}>Compliance Cleared
                  <IconButton id="ComplianceCleared" style={{ color: "Black", display: this.state.ComplianceClearedTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="ComplianceCleared" onClick={(e) => this._onSortClickAscForOtherDept('ComplianceCleared', e)} />
                  <IconButton id="ComplianceCleared" style={{ color: "Black", display: this.state.ComplianceClearedTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="ComplianceCleared" onClick={(e) => this._onSortClickDescForOtherDept('ComplianceCleared', e)} />
                </div></th>
                <th style={{ padding: "5px 10px" }}><div style={{ display: "flex" }}>Department
                  <IconButton id="Department" style={{ color: "Black", display: this.state.DepartmentTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Department" onClick={(e) => this._onSortClickAscForOtherDept('Department', e)} />
                  <IconButton id="Department" style={{ color: "Black", display: this.state.DepartmentTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Department" onClick={(e) => this._onSortClickDescForOtherDept('Department', e)} />
                </div></th>
                <th style={{ padding: "5px 10px", }}><div style={{ display: "flex" }}>Created By
                  <IconButton id="Author" style={{ color: "Black", display: this.state.CreatedByTypeAsc }} iconProps={SortAcsIcon} title="sort" ariaLabel="Author" onClick={(e) => this._onSortClickAscForOtherDept('Author', e)} />
                  <IconButton id="Author" style={{ color: "Black", display: this.state.CreatedByTypeDesc }} iconProps={SortDescIcon} title="sort" ariaLabel="Author" onClick={(e) => this._onSortClickDescForOtherDept('Author', e)} />
                </div></th>
              </tr>
              {this.state.paginatedItems.map((items, key) => {
                return (
                  <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><IconButton iconProps={EditIcon} title="Edit" ariaLabel="Delete" onClick={() => this.onEditClick(items)}
                      disabled={this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? false : true} /></td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" ? " " : "none") }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this.onDeleteClick(items)} /></td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}><IconButton
                      iconProps={ShowDocuments} onClick={() => this.openCCSPopUp(items)}
                      text="View Documents"
                      disabled={this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? false : true} /></td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.OpportunityType : " "} </td>
                    <td style={{ padding: "5px 10px" }}> {items.Source.Title}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Industry.Title}  </td>
                    <td style={{ padding: "5px 10px" }}> {items.Title} </td>
                    <td style={{ padding: "5px 10px", }}>{items.ClassOfInsurance.Title}  </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? moment(items.EstimatedStartDate).format("DD/MM/YYYY") : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.Comments : " "}  </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? Number(items.EstimatedPremium).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.BrokeragePercentage + " %" : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? Number(items.EstimatedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.NBOStage.Title : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? Number(items.WeightedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? Number(items.FeesIfAny).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,') : " "} </td>
                    <td style={{ padding: "5px 10px", display: (this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.state.isNBOAdmin == "true" || this.teamType == "Compliance Team" ? items.ComplianceCleared : " "} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department}  </td>
                    <td style={{ padding: "5px 10px", }}>{items.Author.Title} </td>
                  </tr>
                );
              })}
            </table>
            <div className={styles.NoDataFound} style={{ display: this.state.divForNoDataFound }}> No Record Found</div>
            <div style={{ display: this.state.divForShowingPagination }}>
              <Pagination
                currentPage={0}
                totalPages={(this.sortedArray.length / this.pageSize) - 1}
                onChange={(page) => this._getPage(page)}
                limiter={100}
                limiterIcon={"Emoji12"} // Optional
              />

            </div>
          </div>
        </div><div style={{ display: this.state.AddNBO }} className={styles.nboAddDiv}>
            <Modal
              isOpen={this.state.showReviewModal}
              onDismiss={this._closeModal}
              containerClassName={contentStyles.container}>
              {/* header */}
              <div style={{ display: "flex", backgroundColor: "#008f85", }}>
                <h1 style={{ marginLeft: "410px", color: "white" }}>New Business Opportunity Form</h1>
                <div style={{ marginLeft: "25%" }}>
                  <IconButton
                    iconProps={CancelIcon}
                    ariaLabel="Close popup modal"
                    onClick={this._closeModal}
                    styles={iconButtonStyles} />
                </div>
              </div>
              {/* body */}
              <div style={{ padding: "17px 7px 11px 25px", border: "1px solid #25ddd0" }}>
                <div style={{ marginLeft: "10px", marginRight: "10px", display: "flex" }}>
                  <div style={{ marginRight: "16px", width: "87%" }}>
                    <TextField autoComplete="off" label="Prospect Legal Name " required={true} value={this.state.clientName} onChange={this.clientNameChange} ></TextField>
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("subContractor", this.state.clientName, "required")}{" "}</div>
                  </div>
                  <div><Dropdown id="Opportunity Type"
                    required={true}
                    selectedKey={this.state.oppurtunityTypeKey}
                    placeholder="Select Opportunity Type"
                    options={OpportunityType}
                    onChanged={this._drpdwnOppurtunityType}
                    label="Opportunity Type"
                    style={{ width: "105%" }} />
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("oppurtunityType", this.state.oppurtunityType, "required")}{" "}</div></div>
                </div>

                <div style={{ marginLeft: "10px", marginRight: "10px", display: "flex", marginTop: "27px", marginBottom: "24px" }}>
                  <div>
                    <Dropdown id="t3"
                      required={true}
                      selectedKey={this.state.groupNamekey}
                      placeholder="Select Opportunity Department?"
                      options={this.state.oppurtunityDept}
                      onChanged={this._drpdwnGroupName}
                      label="Opportunity Department?"
                      style={{ width: "100%" }} />
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("groupName", this.state.groupName, "required")}{" "}</div>
                  </div>
                  <div style={{ marginLeft: "20px", }}>
                    <Dropdown id="t3"
                      required={true}
                      selectedKey={this.state.sourceKey}
                      placeholder="Select an option"
                      options={this.state.sourceItems}
                      onChanged={this._drpdwnChangeSource}
                      label="How did we get this prospect?"
                      style={{ width: "100%" }} />
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("sourceKey", this.state.sourceKey, "required")}{" "}</div>
                  </div>
                  <div style={{ marginLeft: "10px", }}>
                    <Dropdown id="t3"
                      required={true}
                      selectedKey={this.state.industryKey}
                      placeholder="Select an option"
                      options={this.state.industryItems}
                      onChanged={this._drpdwnIndustry} style={{ width: "100%" }}
                      label="Which SICC industry does this prospect belong to?" />
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("industryKey", this.state.industryKey, "required")}{" "}</div>
                  </div>
                  <div style={{ marginLeft: "10px", }}><Dropdown id="t3"
                    required={true}
                    selectedKey={this.state.classOfInsuranceKey}
                    placeholder="Select an option"
                    options={this.state.classOfInsuranceItems}
                    onChanged={this._drpdwnClassOfInsurance} style={{ width: "100%" }}
                    label="Which class of insurance is the prospect enquiring?" />
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("classOfInsuranceKey", this.state.classOfInsuranceKey, "required")}{" "}</div>
                  </div>

                </div>
                <hr style={{ backgroundColor: "#008f85", height: "1.5px" }} />
                <div style={{ marginLeft: "10px", marginRight: "10px", display: "flex", marginTop: "27px", marginBottom: "24px" }}>
                  <div style={{}}> <TextField autoComplete="off"
                    label="What is the estimated gross premium amount in SGD?" value={this.state.estimatedPremium} defaultValue={this.state.estimatedPremium} type='number' required={true} onChange={this._estimatedPremiumChange}>
                  </TextField>
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("estimatedPremium", this.state.estimatedPremium, "required|numeric")}{" "}</div></div>

                  <div style={{ marginLeft: "10px", }}>
                    <TextField autoComplete="off" label="What is the brokerage %?" type='number' value={this.state.brokerage} defaultValue={this.state.brokerage} onChange={this._drpdwnBrokerage}></TextField>
                    {/* <Dropdown id="t3"
                      required={true}
                      selectedKey={this.state.brokerageKey}
                      placeholder="Select an option"
                      options={this.state.brokeragePercentageItems}
                      onChanged={this._drpdwnBrokerage} style={{ width: "100%" }}
                      label="What is the brokerage %?" /> */}
                    {/* <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("brokerageKey", this.state.brokerageKey, "required")}{" "}</div> */}
                  </div>
                  <div style={{ marginLeft: "10px", }}>
                    <TextField autoComplete="off" label="What are the fees amount?" type='number' value={this.state.feesIfAny} defaultValue={this.state.feesIfAny} onChange={this._feesIfAnyChange}>
                    </TextField></div>
                  <div style={{ marginLeft: "10px", marginRight: "10px" }}>
                    <DatePicker label="When is the projected policy renewal date?"
                      value={this.state.estimatedStartDate}
                      //  hidden={this.state.hideDueDate}
                      onSelectDate={this._estimateDateChange}
                      // minDate={this.state.dueDateForBindingApprovalLifeCycle}
                      placeholder="Select a date..."
                      ariaLabel="Select a date" />
                    <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("estimatedStartDate", this.state.estimatedStartDate, "required")}{" "}</div>
                  </div>
                </div>
                <hr style={{ backgroundColor: "#008f85", marginTop: "7px", height: "1.5px" }} />
                <div>
                  <div style={{ width: "50%", display: "flex", marginTop: "17px" }}>
                    <input type="file" name="myFile" id="newfile" style={{ marginRight: "-13px", marginLeft: "12px" }} onChange={(e) => this._uploadadditional(e)}></input>
                    <PrimaryButton onClick={this._showExternalGrid} style={{ backgroundColor: "#008f85", color: "White", marginLeft: "48px" }}>Upload</PrimaryButton>
                  </div>
                  <div style={{ display: this.state.externalArrayDiv, marginLeft: "13px", marginTop: "7px" }}>
                    <table className={styles.tableModal}>
                      <tr style={{ background: "#008f85" }}>
                        <th style={{ padding: "5px 10px" }}>Slno</th>
                        <th style={{ padding: "5px 10px" }}>Document Name</th>
                        <th style={{ padding: "5px 10px" }}>Delete</th>
                      </tr>
                      {this.state.externalArray.map((items, key) => {
                        return (
                          <tr style={{ borderBottom: "1px solid #b6ede83b", backgroundColor: "#daebea" }}>
                            <td style={{ padding: "5px 10px" }}>{key + 1}</td>
                            <td style={{ padding: "5px 10px" }}>{items.documentName}</td>
                            <td style={{ padding: "5px 10px" }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key)} /></td>
                          </tr>
                        );
                      })}
                    </table>
                  </div>
                </div>
                <hr style={{ marginTop: "7px", backgroundColor: "#008f85", height: "1.5px" }} />
                <div style={{ marginLeft: "10px", marginRight: "10px", marginBottom: " 10px", marginTop: "27px" }}>
                  <TextField autoComplete="off" label="Please indicate your comments on the deal" multiline placeholder="" value={this.state.comments} onChange={this._commentsChange}></TextField></div>
                <div style={{ marginLeft: "80%" }}><PrimaryButton text='Submit' onClick={this._submitNBOPipeline} style={{ backgroundColor: "#008f85", color: "White" }} />
                  <PrimaryButton onClick={this._dialogCloseButton} style={{ marginLeft: "7px", backgroundColor: "#008f85", color: "White" }}>Cancel</PrimaryButton></div>
              </div>
              {/* footer */}
              <div style={{ display: this.state.messageBar }}>
                <ToastContainer
                  position="bottom-center"
                  autoClose={10000}
                  hideProgressBar={false}
                  newestOnTop={false}
                  closeOnClick={false}
                  rtl={false}
                  pauseOnFocusLoss
                  draggable={false}
                  pauseOnHover={false} />
              </div>
            </Modal>
          </div>
          <div>
            <Panel
              isOpen={this.state.hideDialog}
              onDismiss={this._dialogCloseButton}
              headerText="Edit NBO PineLine"
              closeButtonAriaLabel="Close"
              isFooterAtBottom={true}
              type={PanelType.medium}
            >

              <TextField autoComplete="off" required={true} label="Prospect Legal Name " value={this.state.clientName} onChange={this.clientNameChange} disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false}></TextField>
              <Dropdown id="t3"
                required={true}
                selectedKey={this.state.sourceKey}
                placeholder="Select an option"
                options={this.state.sourceItems}
                onChanged={this._drpdwnChangeSource} style={{ width: "100%" }}
                label="How did we get this prospect?"
                disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false} />
              <Dropdown id="t3"
                required={true}
                selectedKey={this.state.industryKey}
                placeholder="Select an option"
                options={this.state.industryItems}
                onChanged={this._drpdwnIndustry}
                label="Which industry does this prospect belong to?"
                disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false} />

              <Dropdown id="t3"
                required={true}
                selectedKey={this.state.classOfInsuranceKey}
                placeholder="Select an option"
                options={this.state.classOfInsuranceItems}
                onChanged={this._drpdwnClassOfInsurance}
                label="Which class of insurance is the prospect enquiring?"
                disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false} />
              <TextField autoComplete="off" label="What is the estimated premium amount in SGD?" type="number" value={this.state.estimatedPremium} onChange={this._estimatedPremiumChange} disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false}></TextField>
              <TextField autoComplete="off" label="What is the brokerage %?" type='number' value={this.state.brokerage} onChange={this._drpdwnBrokerage} disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false}></TextField>
              {/* <Dropdown id="t3"
                required={true}
                selectedKey={this.state.brokerageKey}
                placeholder="Select an option"
                options={this.state.brokeragePercentageItems}
                onChanged={this._drpdwnBrokerage} style={{ width: "100%" }}
                label="What is the brokerage %?"
                disabled={this.teamType == "Compliance Team" || this.compliance == "Yes" ? true : false} /> */}
              <TextField autoComplete="off" label="What are the fees amount?" type="number" value={this.state.feesIfAny} onChange={this._feesIfAnyChange} disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false}></TextField>
              <DatePicker label="When is the projected policy renewal date?"
                value={this.state.estimatedStartDate}
                //  hidden={this.state.hideDueDate}
                onSelectDate={this._estimateDateChange}
                // minDate={this.state.dueDateForBindingApprovalLifeCycle}
                placeholder="Select a date..."
                ariaLabel="Select a date"
                disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false} />
              <TextField autoComplete="off" label="Please indicate your comments on the deal" multiline placeholder="" value={this.state.comments}
                onChange={this._commentsChange} disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false}></TextField>
              <Dropdown id="t3"
                required={true}
                selectedKey={this.state.NB0StageKey}
                placeholder="Select an option"
                options={this.state.NBOStageItems}
                onChanged={this._drpdwnNBOStage}
                label="NBO stage"
                disabled={(this.teamType == "Compliance Team" && this.state.editAuthorEmail != this.currentUserEmail) || this.compliance == "Yes" ? true : false} />
              <Dropdown id="t3"
                required={true}
                selectedKey={this.state.complianceCleared}
                placeholder="Select an option"
                options={ComplianceCleared}
                onChanged={this._drpdwnComplianceCleared}
                label="Compliance Cleared"
                disabled={this.teamType == "Compliance Team" || this.compliance == "Yes" ? false : true} />
              <div style={{ display: this.state.divForDocumentUploadCompliance }}>
                <div style={{ width: "50%", display: "flex", marginTop: "17px" }}>
                  <input type="file" name="myFile" id="newfile" style={{ marginRight: "-13px", marginLeft: "12px" }} onChange={(e) => this._uploadadditional(e)}></input>
                  <PrimaryButton onClick={this._showExternalGridForComplianceUpload} style={{ backgroundColor: "#008f85", color: "White", marginLeft: "48px" }}>Upload</PrimaryButton>
                </div>

              </div>
              <div style={{ display: this.state.divForDocumentUploadCompArrayDiv, marginLeft: "13px", marginTop: "7px" }}>
                <table className={styles.tableModal}>
                  <tr style={{ background: "#008f85" }}>
                    <th style={{ padding: "5px 10px" }}>Slno</th>
                    <th style={{ padding: "5px 10px" }}>Document Name</th>
                    <th style={{ padding: "5px 10px" }}>Delete</th>
                  </tr>
                  {this.state.externalArray.map((items, key) => {
                    return (
                      <tr style={{ borderBottom: "1px solid #b6ede83b", backgroundColor: "#daebea" }}>
                        <td style={{ padding: "5px 10px" }}>{key + 1}</td>
                        <td style={{ padding: "5px 10px" }}>{items.documentName}</td>
                        <td style={{ padding: "5px 10px" }}><IconButton iconProps={DeleteIcon} title="Delete" ariaLabel="Delete" onClick={() => this._openDeleteConfirmation(items, key)} /></td>
                      </tr>
                    );
                  })}
                </table>
              </div>
              <div style={{ marginTop: "10%" }}>
                <div style={{ display: this.state.showDocInPanel }}>
                  <div>
                    <ToastContainer />
                  </div>
                </div>
              </div>
              <div style={{ marginTop: "10%" }}>
                <PrimaryButton onClick={this._updateNBOPipeline} style={{ backgroundColor: "#008f85", color: "White" }}>Update</PrimaryButton>
                <PrimaryButton onClick={this._dialogCloseButton} style={{ marginLeft: "5px", backgroundColor: "#008f85", color: "White" }}>Cancel</PrimaryButton>
              </div>
            </Panel>
          </div></>
        <div style={{ display: this.state.displayFromMail }}>
          <div style={{ display: this.state.divForSame, marginTop: "10px", overflowX: "auto" }}>
            <table style={{ overflowX: "scroll", display: this.state.docRepositoryItems.length == 0 ? "none" : "" }}>
              <tr style={{ background: "#f4f4f4" }}>
                {/* <th style={{ padding: "5px 10px" }} >Slno</th> */}
                {/* <th style={{ padding: "5px 10px" }}>Doc Id</th> */}
                <th style={{ padding: "5px 10px", }}>Edit</th>
                <th style={{ padding: "5px 10px", }}>View Documents</th>
                <th style={{ padding: "5px 10px", }}>Opportunity Type</th>
                <th style={{ padding: "5px 10px" }}>Source</th>
                <th style={{ padding: "5px 10px" }}>Industry</th>
                <th style={{ padding: "5px 10px" }}>Client Name</th>
                <th style={{ padding: "5px 10px", }}>Class of Insurance</th>
                <th style={{ padding: "5px 10px", }}>Est Start Date</th>
                <th style={{ padding: "5px 10px", }}>Comments</th>
                <th style={{ padding: "5px 10px", }}>Estimated Premium</th>
                <th style={{ padding: "5px 10px" }}>Brokerage %</th>
                <th style={{ padding: "5px 10px" }}>Estimated Brokerage</th>
                <th style={{ padding: "5px 10px", }}>NBO Stage</th>
                <th style={{ padding: "5px 10px", }}>Weighted Brokerage</th>
                <th style={{ padding: "5px 10px", }}>Fees If Any</th>
                <th style={{ padding: "5px 10px", }}>Compliance Cleared</th>
                <th style={{ padding: "5px 10px" }}>Department</th>
              </tr>

              {this.state.paginatedItems.map((items, key) => {
                return (
                  <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                    {/* <td style={{ padding: "5px 10px" }}>{key + 1}</td> */}
                    {/* <td style={{ padding: "5px 10px" }}>{items.documentID} </td> */}
                    <td style={{ padding: "5px 10px", }}><IconButton iconProps={EditIcon} title="Edit" ariaLabel="Delete" onClick={() => this.onEditClickFromMail(this.nbolid)} /></td>
                    <td style={{ padding: "5px 10px", }}><IconButton
                      iconProps={ShowDocuments} onClick={() => this.openCCSPopUp(items)}
                      text="View Documents"
                    /></td>
                    <td style={{ padding: "5px 10px" }}> {items.OpportunityType}  </td>
                    <td style={{ padding: "5px 10px" }}> {items.Source.Title}  </td>
                    <td style={{ padding: "5px 10px" }}>{items.Industry.Title}  </td>
                    <td style={{ padding: "5px 10px" }}> {items.Title} </td>
                    <td style={{ padding: "5px 10px", }}>{items.ClassOfInsurance.Title}  </td>
                    <td style={{ padding: "5px 10px" }}>{moment(items.EstimatedStartDate).format("DD/MM/YYYY")} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Comments} </td>
                    <td style={{ padding: "5px 10px", }}>{Number(items.EstimatedPremium).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,')}</td>
                    <td style={{ padding: "5px 10px" }}>{items.BrokeragePercentage + " %"} </td>
                    <td style={{ padding: "5px 10px" }}>{Number(items.EstimatedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,')} </td>
                    <td style={{ padding: "5px 10px" }}>{items.NBOStage.Title} </td>
                    <td style={{ padding: "5px 10px" }}>{Number(items.WeightedBrokerage).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,')} </td>
                    <td style={{ padding: "5px 10px" }}>{Number(items.FeesIfAny).toFixed(2).replace(/(\d)(?=(\d{3})+(?!\d))/g, '$1,')} </td>
                    <td style={{ padding: "5px 10px", }}>{items.ComplianceCleared} </td>
                    <td style={{ padding: "5px 10px" }}>{items.Department} </td>

                  </tr>
                );
              })}
            </table>

          </div>

        </div >
        <Dialog
          hidden={this.state.confirmDialog}
          dialogContentProps={this.dialogContentProps}
          onDismiss={this._dialogCloseButton}
          styles={this.dialogStyles}
          modalProps={this.modalProps}>
          <DialogFooter>
            <PrimaryButton onClick={this._confirmYesCancel} text="Yes" style={{ backgroundColor: "#008f85", color: "White" }} />
            <DefaultButton onClick={this._confirmNoCancel} text="No" />
          </DialogFooter>
        </Dialog>
        <div>
          <Modal
            isOpen={this.state.hideFilterDialog}
            onDismiss={this._closeModal}
            containerClassName={contentStyles.container}
          >

            <div >
              <p>
                <div style={{ marginLeft: "8px", marginRight: "8%" }} ><Dropdown label='Select one column' placeholder='Select one column' selectedKey={this.state.selectedColumnKey} options={MyNBOFilterColumns} onChanged={this.filterColumnChange}></Dropdown>
                  <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("selectedColumnKey", this.state.selectedColumnKey, "required")}{" "}</div></div>
                <div style={{ marginLeft: "8px", marginRight: "8%", marginBottom: "20px", display: this.state.filterConditionDiv }} >
                  <Dropdown label='Select condition' placeholder='Select condition'
                    selectedKey={this.state.filterConditionKey} options={this.state.filterConditions} onChanged={this.filterConditionColumnChange}></Dropdown>
                  <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px", marginTop: "5px" }}>{this.validator.message("filterCondition", this.state.filterCondition, "required")}{" "}</div></div>
                <div style={{ width: "90%", marginLeft: "8px", display: this.state.textFiledForFilter }}><TextField title='Enter Value' placeholder='Enter Value' value={this.state.filterValue} onChange={this.filterValueChange}></TextField></div>
                <div style={{ width: "88%", marginLeft: "10px", marginRight: "10px", display: this.state.dateForFilter }}>
                  <DatePicker label="From date"
                    value={this.state.estimatedFromStartDate}
                    //  hidden={this.state.hideDueDate}
                    onSelectDate={this._estimateFromDateChange}
                    // minDate={this.state.dueDateForBindingApprovalLifeCycle}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                  />
                  <DatePicker label="To date"
                    value={this.state.estimatedToStartDate}
                    //  hidden={this.state.hideDueDate}
                    onSelectDate={this._estimateToDateChange}
                    minDate={this.state.estimatedFromStartDate}
                    placeholder="Select a date..."
                    ariaLabel="Select a date"
                  />
                  <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("estimatedStartDate", this.state.estimatedStartDate, "required")}{" "}</div>
                </div>
                {/* <SearchBox placeholder="Type application name" className={styles['ms-SearchBox']} onSearch={newValue => console.log('value is ' + newValue)} onChange={this._onFilterForModal} /> */}
                <div style={{
                  float: "right",
                  marginTop: "15%",
                  marginBottom: "8%",
                  marginRight: "5%"
                }}>
                  <PrimaryButton text="Filter" onClick={this._onFilterButtonSubmit} style={{ backgroundColor: "#008f85", color: "White" }} />
                  <DefaultButton text="Cancel" onClick={this._filterPanelCloseButton} style={{ marginLeft: "8px" }} />
                </div>
              </p>
            </div>


          </Modal>
        </div >


      </div >

    );
  }
}
