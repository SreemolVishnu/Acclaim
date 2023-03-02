import * as React from 'react';
import styles from './HrTrainingWp.module.scss';
import { IHrTrainingWpProps } from './IHrTrainingWpProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Item, sp } from '@pnp/sp/presets/all';
import { Checkbox, Dialog, DialogFooter, DialogType, HoverCard, HoverCardType, IconButton, IIconProps, IToggleProps, ITooltipHostStyles, ITooltipStyles, Label, MessageBar, PrimaryButton, TooltipHost } from 'office-ui-fabric-react';


export interface IHrTrainingWpState {
  statusMessage: IMessage;
  hrTrainingListItems: any;
  checkedState: any;
  currentUserDepartment: string;
  currentUserId: Number;
  hrTrainingItems: any;
  LineManagerEmail: string;
  temArray: any;
  temArrayVideos: any;
  itemDiv: string;
  confirmDeleteDialog: boolean;
  mainFolder: string;
  errorMsg: string;
  noItemsDiv: string;
  trainingTileName: any;
  buttonHide: boolean;
}
//for msg bar
export interface IMessage {
  isShowMessage: boolean;
  messageType: number;
  message: string;
}
const contactSearch: IIconProps = { iconName: 'DoubleChevronLeftMedMirrored' };
const calloutProps = { gapSpace: 0 };
const hostStyles: Partial<ITooltipStyles> = { root: { display: 'inline-block', color: 'b;ack', fontWeight: 'bold' }, };

export default class HrTrainingWp extends React.Component<IHrTrainingWpProps, IHrTrainingWpState, {}> {

  private accessedTrainingId: any[];
  private currentUserEmail;
  private currentlyRun;
  constructor(props: IHrTrainingWpProps) {
    super(props);
    this.state = {
      statusMessage: {
        isShowMessage: false,
        message: "",
        messageType: 90000,
      },
      hrTrainingListItems: [],
      checkedState: false,
      currentUserDepartment: "",
      currentUserId: null,
      hrTrainingItems: [],
      LineManagerEmail: "",
      itemDiv: "none",
      temArray: [],
      temArrayVideos: [],
      confirmDeleteDialog: true,
      mainFolder: "",
      errorMsg: "none",
      noItemsDiv: "none",
      trainingTileName: [],
      buttonHide: false

    };
    this.handleOnChange = this.handleOnChange.bind(this);
    this.addToHRList = this.addToHRList.bind(this);
    this.GetUserProperties = this.GetUserProperties.bind(this);
    this.cancel = this.cancel.bind(this);
    this.checkingPreviouslySelected = this.checkingPreviouslySelected.bind(this);
    this.appendIds = this.appendIds.bind(this);
    this.cancel = this.cancel.bind(this);
  }
  public async componentDidMount(): Promise<void> {
    await sp.web.currentUser.get().then(currentUser => {
      console.log(currentUser);
      this.currentUserEmail = currentUser.Email;
      this.setState({
        currentUserId: currentUser.Id,
      });
    });
    this.GetUserProperties();
    let items = [];
    let oneArray = [];
    let hrTrainingNameDuplicates = [];
    //this.checkingPreviouslySelected();
    await sp.web.getList(this.props.siteUrl + "/Lists/HRDuplicates").items.filter("User eq '" + this.currentUserEmail + "' ").get()
      .then(async HRDuplicates => {
        console.log("HRDuplicates", HRDuplicates);
        if (HRDuplicates.length == 0 || HRDuplicates[0].Title != "Pending") {
          this.currentlyRun = "No";
          //alert("No Duplicates")
          await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.hrTrainingTitles).items.get().then(async hrTrainingListItems => {
            // ResourceId: { results: this.internalResourceID, }
            console.log("hrTrainingListItems", hrTrainingListItems);
            await sp.web.getList(this.props.siteUrl + "/" + this.props.TrainingModuleLibrary).items.filter("Title eq '" + this.currentUserEmail + "'").get()
              .then(async folder => {
                console.log("folder", folder);
                for (let i = 0; i < hrTrainingListItems.length; i++) {
                  if (folder.length != 0) {
                    await sp.web.getList(this.props.siteUrl + "/" + this.props.TrainingModuleLibrary).items.filter("TrainingTitle eq '" + hrTrainingListItems[i].Title + "'  and (CurrentUser eq '" + this.currentUserEmail + "')").get()
                      .then(checkingPreviouslySelected => {
                        console.log("PreviouslySelected", checkingPreviouslySelected);
                        if (checkingPreviouslySelected.length == 0) {
                          items.push(
                            { id: hrTrainingListItems[i].ID, value: hrTrainingListItems[i].Title, description: hrTrainingListItems[i].TrainingDescription, isChecked: false },
                          );
                          this.setState({
                            hrTrainingListItems: items, temArray: oneArray
                          });
                        }
                      });
                  }
                  else {
                    items.push(
                      { id: hrTrainingListItems[i].ID, value: hrTrainingListItems[i].Title, description: hrTrainingListItems[i].TrainingDescription, isChecked: false },
                    );
                    this.setState({
                      hrTrainingListItems: items, temArray: oneArray
                    });
                  }
                }
              });
          });
          if (this.state.hrTrainingListItems.length == 0) {
            this.setState({
              noItemsDiv: " ",
            });
          }
        }
        else if (HRDuplicates[0].Title == "Pending") {
          let splitted = HRDuplicates[0].TrainingTitle.split(',');
          hrTrainingNameDuplicates = HRDuplicates[0].TrainingTitle.split(',');
          console.log("SplittedNames", hrTrainingNameDuplicates);
          this.setState({
            trainingTileName: hrTrainingNameDuplicates,
          });
          await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.hrTrainingTitles).items.get().then(async hrTrainingListItems => {
            console.log("hrTrainingListItems", hrTrainingListItems);
            hrTrainingListItems = hrTrainingListItems.filter(val => {
              return !hrTrainingNameDuplicates.some((val2) => {
                //  console.log({valueID:val.id+":"+val2.id});
                return val.Title === val2
              })
            });
            console.log("withOut", hrTrainingListItems);
            await sp.web.getList(this.props.siteUrl + "/" + this.props.TrainingModuleLibrary).items.filter("Title eq '" + this.currentUserEmail + "'").get().then(async folder => {
              console.log("folder", folder);
              for (let i = 0; i < hrTrainingListItems.length; i++) {
                if (folder.length != 0) {
                  await sp.web.getList(this.props.siteUrl + "/" + this.props.TrainingModuleLibrary).items.filter("TrainingTitle eq '" + hrTrainingListItems[i].Title + "'  and (CurrentUser eq '" + this.currentUserEmail + "')").get().then(checkingPreviouslySelected => {
                    console.log("PreviouslySelected", checkingPreviouslySelected);
                    if (checkingPreviouslySelected.length == 0) {
                      items.push(
                        { id: hrTrainingListItems[i].ID, value: hrTrainingListItems[i].Title, description: hrTrainingListItems[i].TrainingDescription, isChecked: false },
                      );
                      this.setState({
                        hrTrainingListItems: items, temArray: oneArray
                      });
                    }
                  });
                }
                else {
                  items.push(
                    { id: hrTrainingListItems[i].ID, value: hrTrainingListItems[i].Title, description: hrTrainingListItems[i].TrainingDescription, isChecked: false },
                  );
                  this.setState({
                    hrTrainingListItems: items, temArray: oneArray
                  });
                }
              }

            });

            // for (let i = 0; i < hrTrainingListItems.length; i++) {
            //   items.push({ id: hrTrainingListItems[i].ID, value: hrTrainingListItems[i].Title, description: hrTrainingListItems[i].TrainingDescription, isChecked: false });
            // }
            // this.setState({
            //   hrTrainingListItems: items, temArray: oneArray
            // });

          });


          if (this.state.hrTrainingListItems.length == 0) {
            this.setState({
              noItemsDiv: " ",
            });
          }


          // await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.hrTrainingTitles).items.get().then(async hrTrainingListItems => {
          //   // ResourceId: { results: this.internalResourceID, }
          //   console.log("hrTrainingListItems", hrTrainingListItems);
          //   await sp.web.getList(this.props.siteUrl + "/" + this.props.TrainingModuleLibrary).items.filter("Title eq '" + this.currentUserEmail + "'").get().then(async folder => {
          //     console.log("folder", folder);
          //     for (let i = 0; i < hrTrainingListItems.length; i++) {

          //       if (folder.length != 0) {
          //         await sp.web.getList(this.props.siteUrl + "/" + this.props.TrainingModuleLibrary).items.filter("TrainingTitle eq '" + hrTrainingListItems[i].Title + "'  and (CurrentUser eq '" + this.currentUserEmail + "')").get().then(checkingPreviouslySelected => {
          //           console.log("PreviouslySelected", checkingPreviouslySelected);
          //           if (checkingPreviouslySelected.length == 0) {
          //             items.push(
          //               { id: hrTrainingListItems[i].ID, value: hrTrainingListItems[i].Title, description: hrTrainingListItems[i].TrainingDescription, isChecked: false },
          //             );
          //             this.setState({
          //               hrTrainingListItems: items, temArray: oneArray
          //             });
          //           }
          //         });
          //       }
          //       else {
          //         items.push(
          //           { id: hrTrainingListItems[i].ID, value: hrTrainingListItems[i].Title, description: hrTrainingListItems[i].TrainingDescription, isChecked: false },
          //         );
          //         this.setState({
          //           hrTrainingListItems: items, temArray: oneArray
          //         });
          //       }
          //     }

          //   });
          // });

        }
      });
  }
  private checkingPreviouslySelected() {
    sp.web.getList(this.props.siteUrl + "/List/HRDuplicates").items.filter("User eq '" + this.currentUserEmail + "'").get().then(checkingPreviouslySelected => {
      console.log("HRDuplicates", checkingPreviouslySelected);
    });

  }
  private GetUserProperties(): void {
    sp.profiles.myProperties.get().then(result => {
      var userProperties = result.UserProfileProperties;
      var userPropertyValues = "";
      let email = [];
      console.log("userProperties--", userProperties);
      for (var k in userProperties) {
        if (userProperties[k].Key == "Department") {
          this.setState({ currentUserDepartment: userProperties[k].Value });
          console.log("Department --", userProperties[k].Value);
        }
        if (userProperties[k].Key == "Manager") {
          console.log(userProperties[k].Key, userProperties[k].Value);
          email = userProperties[k].Value.split('i:0#.f|membership|');
          console.log(email[1]);
          console.log(email[1]);
          this.setState({ LineManagerEmail: email[1] });

        }
      }
    });
  }
  public async appendIds(trainingTitleId) {
    //alert(trainingTitleId)
    let oneArray = [];
    let videoArray = [];
    // sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.hrTrainingTitles).items.select("DocumentFiles/DocName,VideoFiles/DocName").expand("DocumentFiles,VideoFiles").filter("ID eq '" + trainingTitleId + "'")
    //   .get().then(itemFromDoc => {
    //     console.log("itemsDF", itemFromDoc[0].DocumentFiles);
    //     console.log("itemsVideos", itemFromDoc[0].VideoFiles);
    //     oneArray.push(itemFromDoc[0].DocumentFiles);
    //     this.setState({
    //       itemDiv: "",
    //       errorMsg: "none",
    //       mainFolder: "none",
    //       confirmDeleteDialog: false,
    //       temArray: itemFromDoc[0].DocumentFiles,
    //       temArrayVideos: itemFromDoc[0].VideoFiles,
    //     });
    //   });
    let documents;
    let videos;
    const itemFromDoc: any[] = await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.hrTrainingTitles).items.select("Documents,VideoName").filter("ID eq '" + trainingTitleId + "'").get();
    // .get().then(async itemFromDoc => {
    documents = itemFromDoc[0].Documents == null ? "" : itemFromDoc[0].Documents.split(',');
    videos = itemFromDoc[0].VideoName == null ? "" : itemFromDoc[0].VideoName.split(',');
    if (itemFromDoc) {
      console.log("itemsDF", itemFromDoc);
      console.log("itemsDF", itemFromDoc[0].Documents);
      console.log("itemsVideos", itemFromDoc[0].VideoName);
      console.log("documents", documents);
      if (documents.length > 0) {
        for (let k = 0; k < documents.length - 1; k++) {
          oneArray.push(documents[k]);
        }
      }

      for (let k = 0; k < videos.length - 1; k++) {
        videoArray.push(videos[k]);
      }
      //oneArray.push(itemFromDoc[0].Documents);        
      //}).then(afterSpliting => {

    }
    this.setState({
      itemDiv: "",
      errorMsg: "none",
      mainFolder: "none",
      confirmDeleteDialog: false,
      temArray: oneArray,
      temArrayVideos: videoArray,
    });

    // });


  }
  public handleOnChange(ev, isChecked) {
    console.log(isChecked);
    let modules = this.state.hrTrainingListItems;
    let tempTrainingItems = [];
    let trainingTitleId;
    if (isChecked) {
      modules.forEach(async selectedModule => {
        if (selectedModule.value === ev.target.name) {
          let trainingTitleId = selectedModule.id;
          selectedModule.isChecked = ev.target.checked;
          tempTrainingItems.push(trainingTitleId);
          this.state.hrTrainingItems.push(selectedModule.id);
          this.state.trainingTileName.push(selectedModule.value);
        }
      });
      console.log(this.accessedTrainingId);
      console.log("checkedItems", this.state.hrTrainingItems);
    }
    else {
      modules.forEach(selectedModule => {
        if (selectedModule.value === ev.target.name) {
          let newarray = this.state.hrTrainingItems.filter(element => element !== selectedModule.id);
          let newarrayName = this.state.trainingTileName.filter(element => element !== selectedModule.value);
          console.log(newarray);
          this.setState({ hrTrainingItems: newarray, itemDiv: "none", trainingTileName: newarrayName });
          console.log("checkedItems", newarray);
        }
      });
    }
  }
  public addToHRList = async () => {
    console.log("AccessedID", this.accessedTrainingId);
    let names = (this.state.trainingTileName).toString();
    if (this.state.hrTrainingItems.length == 0) { //if no module selected 
      this.setState({
        confirmDeleteDialog: false,
        errorMsg: "",
      });
    }
    else {
      this.setState({ buttonHide: true });
      for (let i = 0; i < this.state.hrTrainingItems.length; i++) {
        await sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.hrTrainingReports).items
          .add({
            Title: this.state.currentUserDepartment,
            UserId: this.state.currentUserId,
            TrainingTitlesId: { results: [this.state.hrTrainingItems[i]] },
            LineManagerEmail: this.state.LineManagerEmail,
          });
      }
      sp.web.getList(this.props.siteUrl + "/Lists/HRTrainingItems").items
        .add({
          Title: this.state.currentUserDepartment,
          UserId: this.state.currentUserId,
          TrainingTitlesId: { results: this.state.hrTrainingItems },
          LineManagerEmail: this.state.LineManagerEmail,
        })
        .then(afterSave => {
          sp.web.getList(this.props.siteUrl + "/Lists/HRDuplicates").items.filter("User eq '" + this.currentUserEmail + "'").get()
            .then(HRDuplicates => {
              console.log("HRDuplicates", HRDuplicates);
              if (HRDuplicates.length > 0) {
                sp.web.getList(this.props.siteUrl + "/Lists/HRDuplicates").items.getById(HRDuplicates[0].ID).update({
                  Title: "Pending",
                  User: this.currentUserEmail,
                  TrainingTitle: names
                });
              }
              else {
                sp.web.getList(this.props.siteUrl + "/Lists/HRDuplicates").items.add({
                  Title: "Pending",
                  User: this.currentUserEmail,
                  TrainingTitle: names,
                });
              }
            })
            .then(afterSave => {
              this.setState({ statusMessage: { isShowMessage: true, message: this.props.messageBar, messageType: 4 }, });
              setTimeout(() => {
                window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
              }, 20000);
            });
        });
      //insertion as array in training reports title(old code)
      // sp.web.getList(this.props.siteUrl + "/Lists/" + this.props.hrTrainingReports).items
      //   .add({
      //     Title: this.state.currentUserDepartment,
      //     UserId: this.state.currentUserId,
      //     TrainingTitlesId: { results: this.state.hrTrainingItems },
      //     LineManagerEmail: this.state.LineManagerEmail,
      //   })
      //   .then(afterSave => {
      //     sp.web.getList(this.props.siteUrl + "/Lists/HRDuplicates").items.filter("User eq '" + this.currentUserEmail + "'").get()
      //       .then(HRDuplicates => {
      //         console.log("HRDuplicates", HRDuplicates);
      //         if (HRDuplicates.length > 0) {
      //           sp.web.getList(this.props.siteUrl + "/Lists/HRDuplicates").items.getById(HRDuplicates[0].ID).update({
      //             Title: "Pending",
      //             User: this.currentUserEmail,
      //             TrainingTitle: names
      //           });
      //         }
      //         else {
      //           sp.web.getList(this.props.siteUrl + "/Lists/HRDuplicates").items.add({
      //             Title: "Pending",
      //             User: this.currentUserEmail,
      //             TrainingTitle: names,
      //           });
      //         }
      //       });
      //     this.setState({ statusMessage: { isShowMessage: true, message: this.props.messageBar, messageType: 4 }, });
      //     setTimeout(() => {
      //       window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
      //     }, 20000);
      //   });
    }

  }
  public cancel() {
    setTimeout(() => {
      window.location.replace(window.location.protocol + "//" + window.location.hostname + this.props.siteUrl);
    }, 2000);
  }
  private modalProps = {
    isBlocking: false,
  };
  private dialogCancelContentProps = {
    type: DialogType.normal,
    closeButtonAriaLabel: 'Close',
    title: 'Module Contains',
  };
  private dialogStyles = { main: { maxWidth: 500, scrollX: "hidden", } };
  //For dialog box of cancel
  private _dialogCloseButton = () => {
    this.setState({
      itemDiv: "none",
      confirmDeleteDialog: true,
    });

  }
  public render(): React.ReactElement<IHrTrainingWpProps> {
    return (
      <div>
        <div className={styles.hrTrainingWp}>
          <Label style={{ textAlign: "center", fontSize: "20px" }}>{this.props.webPartTitle}</Label>
          <Label style={{ textAlign: "center", fontSize: "13px" }}>{this.props.labelForInstructions}</Label>
          <div style={{ display: (this.state.hrTrainingListItems.length == 0 ? "none" : " ") }}>
            <div className={styles.container}>
              {this.state.hrTrainingListItems.map(items => {
                return (
                  <div>
                    <div className={styles.folder} >
                      <div>
                        <TooltipHost
                          content={items.description}
                          //id={tooltipId}
                          calloutProps={calloutProps}
                          styles={hostStyles}
                        >
                          <div className={styles.tooltip}>
                            <div className={styles.name} style={{ height: "75px" }}>
                              <Checkbox label={items.value} name={items.value} key={items.id} onChange={this.handleOnChange} style={{ color: "#498205" }} />
                              {/* <div dangerouslySetInnerHTML={{ __html: items.description }}></div> */}
                            </div>
                          </div>
                          <i style={{ cursor: "pointer" }}> <IconButton iconProps={contactSearch} title="View Details" ariaLabel="Delete" onClick={() => this.appendIds(items.id)} style={{ padding: "0px 0px 0px 129px", }} /></i>
                        </TooltipHost>
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
            {this.state.buttonHide === false &&
              <div className={styles.buttonStyle}>
                <PrimaryButton onClick={this.addToHRList} text='Submit' disabled={this.currentlyRun == "Pending" ? true : false}></PrimaryButton>
                <PrimaryButton onClick={this.cancel} text='Cancel' style={{ marginLeft: "8px" }}></PrimaryButton>
              </div>
            }
            <Dialog
              hidden={this.state.confirmDeleteDialog}
              //dialogContentProps={this.dialogCancelContentProps}
              onDismiss={this._dialogCloseButton}
              modalProps={this.modalProps}
              styles={this.dialogStyles}
            >
              <div className={'ms-Dialog-content ' + styles.container}>
                <div className={'ms-Dialog-content ' + styles.folder} style={{ display: this.state.itemDiv }}>
                  <h1>Module Contains</h1>
                  <div >
                    <div style={{ display: (this.state.temArray.length == 0 && this.state.temArrayVideos.length == 0) ? "" : "none" }}>No Documents and Videos</div>
                    <div style={{ fontWeight: "bold", display: (this.state.temArray.length == 0 ? "none" : "") }} >Documents </div>
                    {this.state.temArray.map((doc, key) => {
                      return (
                        <div>
                          <div style={{ display: "flex" }}>
                            {key + 1}. <div style={{ marginLeft: "11px" }}>{doc}</div>
                          </div>
                        </div>
                      );
                    })}

                    <div style={{ fontWeight: "bold", display: (this.state.temArrayVideos.length == 0 ? "none" : "") }}>Videos </div>
                    {this.state.temArrayVideos.map((doc, key) => {
                      return (
                        <div>
                          <div style={{ display: "flex" }}>
                            {key + 1}. <div style={{ marginLeft: "11px" }}>{doc}</div>
                          </div>
                        </div>
                      );
                    })}

                  </div>
                </div>

              </div>
              <div style={{ display: this.state.errorMsg }}>
                <div><h2>{this.props.errorMessage}</h2></div>
              </div>
            </Dialog>
            {/* Show Message bar for Notification*/}
            {this.state.statusMessage.isShowMessage ?
              <MessageBar
                messageBarType={this.state.statusMessage.messageType}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >{this.state.statusMessage.message}</MessageBar>
              : ''}
          </div>
          <div style={{ display: this.state.noItemsDiv, textAlign: "center", marginTop: "45px", fontSize: "17px" }}>{this.props.statementIfNoItems}</div>
        </div>
      </div>
    );
  }
}
