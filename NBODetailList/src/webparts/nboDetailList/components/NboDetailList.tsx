import * as React from 'react';
import styles from './NboDetailList.module.scss';
import { INboDetailListProps } from './INboDetailListProps';
import { DetailsList, Fabric, Selection, IColumn, DetailsListLayoutMode, Link, IconButton, IIconProps, TextField, Dropdown, IDropdownOption, DatePicker, PrimaryButton, DefaultButton, Dialog, DialogFooter, DialogType, Modal, Panel, Label, FontWeights, mergeStyleSets, getTheme, CommandBarButton, PanelType, Callout, IContextualMenuProps, CommandButton } from 'office-ui-fabric-react';
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
export interface INboDetailListState {
  docRepositoryItems: any[];
  selectionDetails: string;
  Items: IDocument[];
  items: any[];
  AddNBO: string;
  hideDialog: boolean;
  currentItemID: any;
  showReviewModal: boolean;
  clientName: string;
  source: string;
  NB0StageText: string;
  complianceCleared: string;
  industry: string;
  industryKey: any;
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
  divForOtherDepts: string;
  divForSame: string;
  noItemErrorMsg: string;

}
export interface IDocument {
  Title: string;
  field_1: string;
  Edit: any;

}

const EditIcon: IIconProps = { iconName: 'Edit' };
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
  private _selection: Selection;
  private team;
  private itemDepartment;
  private teamType = "Department Team";
  private documentName;
  private content;
  private sortedArray = [];
  private pageSize = 100;
  //private teamTypeArray: any[];
  private reqWeb = Web(window.location.protocol + "//" + window.location.hostname + "/sites/Acclaim");
  constructor(props: INboDetailListProps) {
    super(props);
    this.state = {
      docRepositoryItems: [],
      selectionDetails: "",
      Items: [],
      items: [],
      AddNBO: "none",
      hideDialog: false,
      currentItemID: "",
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
      feesIfAny: "",
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
    //dropDownBindings
    this._drpdwnSource = this._drpdwnSource.bind(this);
    this._drpdwnBrokeragePercentage = this._drpdwnBrokeragePercentage.bind(this);
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
  }
  public componentWillMount = async () => {
    this.validator = new SimpleReactValidator({
      messages: {
        required: "Please enter mandatory fields"
      }
    });

  }
  public async componentDidMount() {

    await this.GetUserProperties();
    await this.loadDocProfile();
    await sp.web.currentUser.get().then(currentUser => {
      console.log(currentUser);
      this.currentUserEmail = currentUser.Email;
    });
    // await this.GetUserProperties();
    //dropdownbinding
    this._drpdwnSource();
    this._drpdwnBrokeragePercentage();
    this._drpdwnClassOfInsuranceBind();
    this._drpdwnIndustryBind();

    // await this.loadDocProfile();
    // sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage/Title,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,BrokeragePercentage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage").expand("Source,Industry,ClassOfInsurance,NBOStage,BrokeragePercentage").get().then(docProfileItems => {
    //   this.setState({
    //     docRepositoryItems: docProfileItems,
    //     items: docProfileItems,
    //     paginatedItems: docProfileItems.slice(0, this.pageSize)
    //   });
    //   console.log(this.state.docRepositoryItems);
    // });
  }
  private async checkingcurrentUserDept() {


    //this._sendAnEmailUsingMSGraph();

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
                  if (teamItems[j].Title == response.value[i].displayName) {
                    console.log(response.value[i].displayName, "--", response.value[i].id);
                    console.log(teamItems[j].TeamType);
                    this.state.teamTypeArray.push({ teamtype: teamItems[j].TeamType, team: teamItems[j].Title });
                  }
                }
              }
              console.log(this.state.teamTypeArray);
              for (let k = 0; k < this.state.teamTypeArray.length; k++) {
                if (this.state.teamTypeArray[k].teamtype == "Management Team" || this.state.teamTypeArray[k].teamtype == "Compliance Team") {
                  this.teamType = this.state.teamTypeArray[k].teamtype;
                  this.team = this.state.teamTypeArray[k].team;
                  console.log("team", this.team);
                  console.log("teamType", this.teamType);
                  // alert(this.teamType);
                  break;
                }
                else {
                  this.teamType = this.state.teamTypeArray[k].teamtype;
                  this.team = this.state.teamTypeArray[k].team;

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
          this.setState({ LineManagerEmail: email[1] })
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
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
      select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage/Title,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,BrokeragePercentage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,BrokeragePercentage,Author").filter("Author/EMail eq '" + this.currentUserEmail + "'").get().then(docProfileItems => {
        this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
        this.setState({
          docRepositoryItems: this.sortedArray,
          items: this.sortedArray,
          paginatedItems: this.sortedArray.slice(0, this.pageSize),

        });
        console.log(this.state.docRepositoryItems);

      });
    this.setState({
      divForSame: "",
      divForOtherDepts: "none",

    });

  }
  // sending Email for owners
  private async _sendAnEmailUsingMSGraph(): Promise<void> {
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
    let Subject;
    let Body;
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNotificationSettings).items.filter("Title eq 'NBO'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    //Replacing the email body with current values
    let replacedSubject1 = replaceString(Subject, '[NBOTitle]', this.state.clientName);
    let replacelink = replaceString(Body, '[NBOTitle]', this.state.clientName);
    let FinalBody = replacelink;
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
            "content": FinalBody
          },
          "toRecipients": [
            {
              "emailAddress": {
                "address": email[1],
              }
            }
          ],
          // "ccRecipients": [
          //   {
          //     "emailAddress": {
          //       "address": "dev14@ccsdev01.onmicrosoft.com"
          //     }
          //   }
          // ],
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
  private async _sendAnEmailForComplianceGroup(): Promise<void> {
    //alert("emailtoOwner");
    let Subject;
    let Body;
    //Email Notification Settings.
    const emailNoficationSettings: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.emailNotificationSettings).items.filter("Title eq 'NBOCompliance'").get();
    //console.log(emailNoficationSettings);
    Subject = emailNoficationSettings[0].Subject;
    Body = emailNoficationSettings[0].Body;
    //Replacing the email body with current values
    let replacedSubject1 = replaceString(Subject, '[NBOTitle]', this.state.clientName);
    let replacelink = replaceString(Body, '[NBOTitle]', this.state.clientName);
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

  private dialogContentProps = {
    type: DialogType.normal,
    title: "Select Customer Contacts",
    closeButtonAriaLabel: 'Close',
    //subText: 'Do you want to send this message without a subject?',

  };
  private onEditClick(item) {
    // alert(item.Department);
    this.setState({
      hideDialog: true,
      currentItemID: item.ID,
    });
    //panel rebinding
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(parseInt(item.ID)).select("Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage/Title,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,BrokeragePercentage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared").expand("Source,Industry,ClassOfInsurance,NBOStage,BrokeragePercentage").get().then(docProfileItems => {
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
        brokerage: docProfileItems.BrokeragePercentage.Title,
        brokerageKey: docProfileItems.BrokeragePercentage.ID,
        estimatedBrokerage: docProfileItems.EstimatedBrokerage,
        estimatedStartDate: new Date(docProfileItems.EstimatedStartDate),
      });
      console.log(this.state.docRepositoryItems);
    }).then(forDropDownbinding => {
      //checkin complaince cleared and setting nbo stage according to that.
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
      AddNBO: "", showReviewModal: true,
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
      feesIfAny: "",
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
    });

  }
  // ---------------SubmitToNBOPipeline---------------------
  private _submitNBOPipeline = async () => {
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboStage).items.filter("Title eq '10% - prospect identified'").get().then(NBOStageID => {
      this.setState({
        NB0StageKey: NBOStageID[0].ID,
      });
    });
    if (this.state.clientName != "" && this.state.sourceKey != "" && this.state.industryKey != "" && this.state.classOfInsuranceKey != "" && this.state.estimatedPremium != "" && this.state.brokerageKey != "" && this.state.estimatedStartDate != "") {
      this.validator.hideMessages();
      toast("Nbo Pipeline added successfully");
      this.setState({
        messageBar: "",
      });
      let tempEstimatedBrokerage = (this.state.estimatedPremium * (this.state.brokerage / 100));
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.add({
        Title: this.state.clientName,
        SourceId: this.state.sourceKey,
        IndustryId: this.state.industryKey,
        ClassOfInsuranceId: this.state.classOfInsuranceKey,
        EstimatedPremium: this.state.estimatedPremium,
        BrokeragePercentageId: this.state.brokerageKey,
        EstimatedStartDate: new Date(this.state.estimatedStartDate),
        FeesIfAny: this.state.feesIfAny,
        Comments: this.state.comments,
        Department: this.state.groupName,
        EstimatedBrokerage: Number(this.state.estimatedPremium) * Number(this.state.brokerage / 100),
        NBOStageId: parseInt(this.state.NB0StageKey),
      }).then(async afterInsertion => {
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
        this._sendAnEmailUsingMSGraph();
        this._sendAnEmailForComplianceGroup();
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
            externalArrayDiv: "none"
          });
        }, 7000);
      });
    }
    else {
      this.validator.showMessages();
      this.forceUpdate();
    }
  }
  // ---------------UpdateNBOPipeline---------------------
  private _updateNBOPipeline = () => {
    if (this.state.clientName != "") {
      this.validator.hideMessages();
      sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.getById(this.state.currentItemID).update({
        Title: this.state.clientName,
        SourceId: this.state.sourceKey,
        IndustryId: this.state.industryKey,
        ClassOfInsuranceId: this.state.classOfInsuranceKey,
        EstimatedPremium: this.state.estimatedPremium,
        BrokeragePercentageId: this.state.brokerageKey,
        ComplianceCleared: this.state.complianceCleared,
        BrokerageAmount: this.state.brokerageAmount,
        NBOStageId: this.state.NB0StageKey,
        EstimatedStartDate: new Date(this.state.estimatedStartDate),
        FeesIfAny: this.state.feesIfAny,
        Comments: this.state.comments,
        EstimatedBrokerage: Number(this.state.estimatedPremium) * Number(this.state.brokerage / 100),
      }).then(afterInsertion => {
        // alert("UpdatedSuccessfully");
        this.loadDocProfile();
        toast("Nbo Pipeline list updated successfully");
        this.setState({
          showDocInPanel: "",
        });
        setTimeout(() => {
          this.setState({
            hideDialog: false,
          });
        }, 6000);
      });


    }
    else {
      this.validator.showMessages();
      this.forceUpdate();
    }

  }
  // ---------------EstimatedDatePickerChange---------------------
  private _estimateDateChange = (date?: Date): void => {
    this.setState({ estimatedStartDate: date, });
  }
  // ---------------GroupName---------------------
  public _drpdwnGroupName(option: { key: any; text: any }) {
    this.setState({
      groupNamekey: option.key,
      groupName: option.text,
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

    this.setState({ estimatedPremium: newText || '' });
  }
  private _commentsChange = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string): void => {
    this.setState({ comments: newText || '' });
  }
  // ---------------ModalClose---------------------
  private _closeModal = (): void => {
    this.setState({
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
      feesIfAny: "",
      comments: "",
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
  }
  // ---------------Brokerage---------------------
  public _drpdwnBrokerage(option: { key: any; text: any }) {
    //console.log(option.key)   
    this.setState({ brokerage: option.text, brokerageKey: option.key });
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
  private _getPage(page: number) {
    // round a number up to the next largest integer.
    const roundupPage = Math.ceil(page);
    this.setState({
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
  public _drpdwnBrokeragePercentage() {
    let tempBrokeragePercentageItems = [];
    sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.brokeragePercentage).items.get().then(BrokeragePercentage => {
      console.log("BrokeragePercentage", BrokeragePercentage);
      for (let i = 0; i < BrokeragePercentage.length; i++) {
        // if(subcontractor[i].Active == true){
        let BrokeragePercentageItemdata = {
          key: BrokeragePercentage[i].ID,
          text: BrokeragePercentage[i].Title
        };
        tempBrokeragePercentageItems.push(BrokeragePercentageItemdata);
        //}       
      }
      this.setState({
        brokeragePercentageItems: tempBrokeragePercentageItems
      });
    });
  }
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
  public notify() {
    toast("Wow so easy!");
  }
  private async sameDept() {
    //alert("same dept");
    // alert(this.team);
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
      select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage/Title,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,BrokeragePercentage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,BrokeragePercentage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department eq  '" + this.team + "')").get().then(docProfileItems => {
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
  private async others() {
    //alert("others");
    await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.nboListName).items.
      select("ID,Title,Source/Title,Industry/Title,ClassOfInsurance/Title,NBOStage/Title,BrokeragePercentage/Title,Source/ID,Industry/ID,ClassOfInsurance/ID,NBOStage/ID,BrokeragePercentage/ID,EstimatedBrokerage,FeesIfAny,Comments,EstimatedStartDate,EstimatedPremium,Department,ComplianceCleared,EstimatedBrokerage,Author/EMail").expand("Source,Industry,ClassOfInsurance,NBOStage,BrokeragePercentage,Author").filter("Author/EMail ne '" + this.currentUserEmail + "' and (Department ne  '" + this.team + "')").get().then(docProfileItems => {
        this.sortedArray = _.orderBy(docProfileItems, 'ID', ['desc']);
        this.setState({
          docRepositoryItems: this.sortedArray,
          items: this.sortedArray,
          paginatedItems: this.sortedArray.slice(0, this.pageSize),
          noItemErrorMsg: docProfileItems.length == 0 ? " " : "none"
        });
        console.log(this.state.docRepositoryItems);
        if (docProfileItems.length == 0) {
          this.setState({ noItemErrorMsg: "" });
        }
      });
    this.setState({
      divForSame: "none",
      divForOtherDepts: "",
    });

  }
  public render(): React.ReactElement<INboDetailListProps> {
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
    ];
    const DeleteIcon: IIconProps = { iconName: 'Delete' };
    const ShowDocuments: IIconProps = { iconName: 'DocumentSet' };
    const NBODetails: IIconProps = { iconName: 'BulletedListMirrored' };
    const midBar: IIconProps = { iconName: 'BulletedListBulletMirrored' };
    return (
      <><div className={styles.nboDetailList}>
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
        </div>




        <div style={{ display: this.state.divForSame, marginTop: "10px" }}>
          <table style={{ overflowX: "scroll", display: this.state.docRepositoryItems.length == 0 ? "none" : "" }}>
            <tr style={{ background: "#f4f4f4" }}>
              {/* <th style={{ padding: "5px 10px" }} >Slno</th> */}
              {/* <th style={{ padding: "5px 10px" }}>Doc Id</th> */}
              <th style={{ padding: "5px 10px", }}>Action</th>
              <th style={{ padding: "5px 10px", }}>View Documents</th>
              <th style={{ padding: "5px 10px" }}>Client Name</th>
              <th style={{ padding: "5px 10px" }}>Source</th>
              <th style={{ padding: "5px 10px" }}>Insustry</th>
              <th style={{ padding: "5px 10px", }}>Class of Insurance</th>
              {/* <th style={{ padding: "5px 10px", }}>Currency</th> */}
              <th style={{ padding: "5px 10px", }}>Estimated Premium</th>
              {/* <th style={{ padding: "5px 10px" }}>Brokerage %</th> */}
              <th style={{ padding: "5px 10px" }}>Brokerage %</th>
              <th style={{ padding: "5px 10px" }}>Estimated Brokerage</th>
              <th style={{ padding: "5px 10px", }}>Est Start Date</th>
              <th style={{ padding: "5px 10px", }}>Fees If Any</th>
              <th style={{ padding: "5px 10px", }}>NBO stage</th>
              <th style={{ padding: "5px 10px", }}>Complaince Cleared</th>

            </tr>
            {this.state.paginatedItems.map((items, key) => {
              return (
                <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                  {/* <td style={{ padding: "5px 10px" }}>{key + 1}</td> */}
                  {/* <td style={{ padding: "5px 10px" }}>{items.documentID} </td> */}
                  <td style={{ padding: "5px 10px", }}><IconButton iconProps={EditIcon} title="Edit" ariaLabel="Delete" onClick={() => this.onEditClick(items)} disabled={items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? false : true} /></td>
                  <td style={{ padding: "5px 10px", }}><IconButton
                    iconProps={ShowDocuments} onClick={() => this.openCCSPopUp(items)}
                    text="View Documents"
                    disabled={items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? false : true} /></td>
                  <td style={{ padding: "5px 10px" }}> {items.Title} </td>
                  <td style={{ padding: "5px 10px" }}> {items.Source.Title}  </td>
                  <td style={{ padding: "5px 10px" }}>{items.Industry.Title}  </td>
                  <td style={{ padding: "5px 10px", }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.ClassOfInsurance.Title : " "}  </td>
                  <td style={{ padding: "5px 10px", }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.EstimatedPremium : " "} </td>
                  <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.BrokeragePercentage.Title + " %" : " "} </td>
                  <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.EstimatedBrokerage : " "} </td>
                  {/* <td style={{ padding: "5px 10px" }}>{items.BrokerageAmount} </td> */}
                  <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? moment(items.EstimatedStartDate).format("DD/MM/YYYY") : " "} </td>
                  <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.FeesIfAny : " "} </td>
                  <td style={{ padding: "5px 10px" }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.NBOStage.Title : " "} </td>
                  <td style={{ padding: "5px 10px", }}>{items.Department == this.team || items.Author.EMail == this.currentUserEmail || this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.ComplianceCleared : " "} </td>

                </tr>
              );
            })}
          </table>
          <div style={{ display: this.state.docRepositoryItems.length >= this.pageSize ? "" : "none" }}>
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

        <div style={{ display: this.state.divForOtherDepts, marginTop: "10px" }}>
          <div style={{ display: this.state.docRepositoryItems.length == 0 ? "" : "none", color: "#f4f4f4" }}> <h1>No items</h1></div>
          <table style={{ overflowX: "scroll", display: this.state.docRepositoryItems.length == 0 ? "none" : "" }}>
            <tr style={{ background: "#f4f4f4" }}>

              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Action</th>
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>View Documents</th>
              <th style={{ padding: "5px 10px" }}>Client Name</th>
              <th style={{ padding: "5px 10px" }}>Source</th>
              <th style={{ padding: "5px 10px" }}>Insustry</th>
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Class of Insurance</th>
              {/* <th style={{ padding: "5px 10px", }}>Currency</th> */}
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Estimated Premium</th>
              {/* <th style={{ padding: "5px 10px" }}>Brokerage %</th> */}
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Brokerage %</th>
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Estimated Brokerage</th>
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Est Start Date</th>
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Fees If Any</th>
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>NBO stage</th>
              <th style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>Complaince Cleared</th>

            </tr>
            {this.state.paginatedItems.map((items, key) => {
              return (
                <tr style={{ borderBottom: "1px solid #f4f4f4" }}>
                  {/* <td style={{ padding: "5px 10px" }}>{key + 1}</td> */}
                  {/* <td style={{ padding: "5px 10px" }}>{items.documentID} </td> */}
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}><IconButton iconProps={EditIcon} title="Edit" ariaLabel="Delete" onClick={() => this.onEditClick(items)} disabled={this.teamType == "Management Team" || this.teamType == "Compliance Team" ? false : true} /></td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}><IconButton
                    iconProps={ShowDocuments} onClick={() => this.openCCSPopUp(items)}
                    text="View Documents"
                    disabled={this.teamType == "Management Team" || this.teamType == "Compliance Team" ? false : true} /></td>
                  <td style={{ padding: "5px 10px" }}> {items.Title} </td>
                  <td style={{ padding: "5px 10px" }}> {items.Source.Title}  </td>
                  <td style={{ padding: "5px 10px" }}>{items.Industry.Title}  </td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.ClassOfInsurance.Title : " "}  </td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.EstimatedPremium : " "} </td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.BrokeragePercentage.Title + " %" : " "} </td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.EstimatedBrokerage : " "} </td>
                  {/* <td style={{ padding: "5px 10px" }}>{items.BrokerageAmount} </td> */}
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? moment(items.EstimatedStartDate).format("DD/MM/YYYY") : " "} </td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.FeesIfAny : " "} </td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.NBOStage.Title : " "} </td>
                  <td style={{ padding: "5px 10px", display: (this.teamType == "Management Team" || this.teamType == "Compliance Team" ? " " : "none") }}>{this.teamType == "Management Team" || this.teamType == "Compliance Team" ? items.ComplianceCleared : " "} </td>

                </tr>
              );
            })}
          </table>
          <div style={{ display: this.state.docRepositoryItems.length >= this.pageSize ? "" : "none" }}>
            <Pagination
              currentPage={0}
              totalPages={(this.sortedArray.length / this.pageSize) - 1}
              onChange={(page) => this._getPage(page)}
              limiter={10}
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
              <h1 style={{ marginLeft: "35%", color: "white" }}>NB Oppurtunity Form</h1>
              <div style={{ marginLeft: "35%" }}>
                <IconButton
                  iconProps={CancelIcon}
                  ariaLabel="Close popup modal"
                  onClick={this._closeModal}
                  styles={iconButtonStyles} />
              </div>
            </div>
            {/* body */}
            <div style={{ padding: "17px 7px 11px 25px", border: "1px solid #25ddd0" }}>
              <div style={{ marginLeft: "10px", marginRight: "10px" }}> <TextField autoComplete="off" label="Prospect Legal Name " required={true} value={this.state.clientName} onChange={this.clientNameChange}></TextField></div>
              <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("subContractor", this.state.clientName, "required")}{" "}</div>
              <div style={{ marginLeft: "10px", marginRight: "10px", display: "flex", marginTop: "27px", marginBottom: "24px" }}>
                <div>
                  <Dropdown id="t3"
                    required={true}
                    selectedKey={this.state.groupNamekey}
                    placeholder="Select a group"
                    options={this.state.groupItems}
                    onChanged={this._drpdwnGroupName}
                    label="Select the group?"
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
                    label="Which industry does this prospect belong to?" />
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
                <div style={{}}> <TextField autoComplete="off" label="What is the estimated premium amount in SGD?" value={this.state.estimatedPremium} type='number' required={true} onChange={this._estimatedPremiumChange}>
                </TextField><div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("estimatedPremium", this.state.estimatedPremium, "required|numeric")}{" "}</div></div>

                <div style={{ marginLeft: "10px", }}>
                  <Dropdown id="t3"
                    required={true}
                    selectedKey={this.state.brokerageKey}
                    placeholder="Select an option"
                    options={this.state.brokeragePercentageItems}
                    onChanged={this._drpdwnBrokerage} style={{ width: "100%" }}
                    label="What is the brokerage %?" />
                  <div style={{ color: "#dc3545", marginLeft: "10px", marginRight: "10px" }}>{this.validator.message("brokerageKey", this.state.brokerageKey, "required")}{" "}</div>
                </div>
                <div style={{ marginLeft: "10px", }}>
                  <TextField autoComplete="off" label="What are the fees amount?" type='number' value={this.state.feesIfAny} onChange={this._feesIfAnyChange}>
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
                <TextField autoComplete="off" label="Comments on the deal" multiline placeholder="" value={this.state.comments} onChange={this._commentsChange}></TextField></div>
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

            <TextField autoComplete="off" required={true} label="Prospect Legal Name " value={this.state.clientName} onChange={this.clientNameChange} disabled={this.teamType == "Compliance Team" ? true : false}></TextField>
            <Dropdown id="t3"
              required={true}
              selectedKey={this.state.sourceKey}
              placeholder="Select an option"
              options={this.state.sourceItems}
              onChanged={this._drpdwnChangeSource} style={{ width: "100%" }}
              label="How did we get this prospect?"
              disabled={this.teamType == "Compliance Team" ? true : false} />
            <Dropdown id="t3"
              required={true}
              selectedKey={this.state.industryKey}
              placeholder="Select an option"
              options={this.state.industryItems}
              onChanged={this._drpdwnIndustry}
              label="Which industry does this prospect belong to?"
              disabled={this.teamType == "Compliance Team" ? true : false} />

            <Dropdown id="t3"
              required={true}
              selectedKey={this.state.classOfInsuranceKey}
              placeholder="Select an option"
              options={this.state.classOfInsuranceItems}
              onChanged={this._drpdwnClassOfInsurance}
              label="Which class of insurance is the prospect enquiring?"
              disabled={this.teamType == "Compliance Team" ? true : false} />
            <TextField autoComplete="off" label="What is the estimated premium amount in SGD?" type="number" value={this.state.estimatedPremium} onChange={this._estimatedPremiumChange} disabled={this.teamType == "Compliance Team" ? true : false}></TextField>
            <Dropdown id="t3"
              required={true}
              selectedKey={this.state.brokerageKey}
              placeholder="Select an option"
              options={this.state.brokeragePercentageItems}
              onChanged={this._drpdwnBrokerage} style={{ width: "100%" }}
              label="What is the brokerage %?"
              disabled={this.teamType == "Compliance Team" ? true : false} />
            <TextField autoComplete="off" label="What are the fees amount?" type="number" value={this.state.feesIfAny} onChange={this._feesIfAnyChange} disabled={this.teamType == "Compliance Team" ? true : false}></TextField>
            <DatePicker label="When is the projected policy renewal date?"
              value={this.state.estimatedStartDate}
              //  hidden={this.state.hideDueDate}
              onSelectDate={this._estimateDateChange}
              // minDate={this.state.dueDateForBindingApprovalLifeCycle}
              placeholder="Select a date..."
              ariaLabel="Select a date"
              disabled={this.teamType == "Compliance Team" ? true : false} />
            <TextField autoComplete="off" label="Comments on the deal" multiline placeholder="" value={this.state.comments}
              onChange={this._commentsChange} disabled={this.teamType == "Compliance Team" ? true : false}></TextField>
            <Dropdown id="t3"
              required={true}
              selectedKey={this.state.NB0StageKey}
              placeholder="Select an option"
              options={this.state.NBOStageItems}
              onChanged={this._drpdwnNBOStage}
              label="NBO stage"
              disabled={this.teamType == "Compliance Team" ? true : false} />
            <Dropdown id="t3"
              required={true}
              selectedKey={this.state.complianceCleared}
              placeholder="Select an option"
              options={ComplianceCleared}
              onChanged={this._drpdwnComplianceCleared}
              label="Complaince Cleared"
              disabled={this.teamType == "Compliance Team" ? false : true} />
            <div style={{ marginTop: "10%" }}>
              <div style={{ display: this.state.showDocInPanel }}>
                <div>
                  <ToastContainer />
                </div>
              </div>
            </div>
            <div style={{ marginTop: "10%" }}>
              <PrimaryButton onClick={this._updateNBOPipeline}>Update</PrimaryButton>
              <PrimaryButton onClick={this._dialogCloseButton} style={{ marginLeft: "5px" }}>Cancel</PrimaryButton>
            </div>
          </Panel>
        </div></>
    );
  }
}
