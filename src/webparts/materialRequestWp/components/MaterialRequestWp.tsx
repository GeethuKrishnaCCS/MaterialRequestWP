import * as React from 'react';
import styles from './MaterialRequestWp.module.scss';
import { IMaterialRequestWpProps, IMaterialRequestWpState } from '../interfaces/IMaterialRequestWpProps';
import { DefaultButton, Dropdown, FocusTrapZone, IDropdownOption, IIconProps, IconButton, Label, Layer, Overlay, Popup, PrimaryButton, TextField } from '@fluentui/react';
import { MaterialRequestWpService } from '../services';
// import { escape } from '@microsoft/sp-lodash-subset';
import * as moment from 'moment';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import MaterialRequestAdminApprovalForm from './MaterialRequestAdminApprovalForm';
import MaterialRequestHOSApprovalForm from './MaterialRequestHOSApprovalForm';
import MaterialRequestViewTable from './MaterialRequestViewTable';
import Toast from './Toast';
// import MaterialRequestViewTable from './MaterialRequestViewTable';

export default class MaterialRequestWp extends React.Component<IMaterialRequestWpProps, IMaterialRequestWpState, {}> {
  private _service: any;

  public constructor(props: IMaterialRequestWpProps) {
    super(props);
    this._service = new MaterialRequestWpService(this.props.context);

    this.state = {
      listClient: [],
      listMaterial: [],
      client: [],
      material: [],
      listProgram: [],
      listProject: [],
      selectedClient: { key: "", text: "" },
      selectedProgram: { key: "", text: "" },
      program: [],
      project: [],
      getProject: { key: "", text: "" },
      getMaterial: { key: "", text: "" },
      comment: "",
      quantity: "",
      isQuantityEntered: false,
      isFirstRowSelected: false,
      currentDate: "",
      rows: [],
      HOSName: null,
      Departmentslist: [],
      department: '',
      departmentName: '',
      adminApproverName: '',
      navigateToList: false,
      isPopupVisible: false,
      quantityError: '',
      statusMessage: '',
      MasterMaterialRequestId: null,
      taskListItemId: null,
      // isLoading: false,
      isOkButtonDisabled: false,
      selectedMaterials: [],
      materialSelectionError: '',
    }
    this.getClientList = this.getClientList.bind(this);
    this.getProgramList = this.getProgramList.bind(this);
    this.getProjectList = this.getProjectList.bind(this);
    this.getProjectChange = this.getProjectChange.bind(this);
    this.getMaterialListItem = this.getMaterialListItem.bind(this);
    this.onMaterialChange = this.onMaterialChange.bind(this);
    this.onChangeQuantity = this.onChangeQuantity.bind(this);
    this.checkQuantityEntered = this.checkQuantityEntered.bind(this);
    this.onChangeComment = this.onChangeComment.bind(this);
    this.getCurrentDate = this.getCurrentDate.bind(this);
    this.onSubmitClick = this.onSubmitClick.bind(this);
    this.deleteRow = this.deleteRow.bind(this);
    this.getCurrentUser = this.getCurrentUser.bind(this);
    this.UserProfiles = this.UserProfiles.bind(this);
    this.getDepartmentsList = this.getDepartmentsList.bind(this);
    this.sendEmailNotificationToHOS = this.sendEmailNotificationToHOS.bind(this);
    this.getAdminApprover = this.getAdminApprover.bind(this);
    this.addRow = this.addRow.bind(this);
    this.hidePopup = this.hidePopup.bind(this);
    this.onPopOk = this.onPopOk.bind(this);
    this.onClickCancel = this.onClickCancel.bind(this);
    this.firstRowChecking = this.firstRowChecking.bind(this);

  }


  public componentDidMount() {
    this.getClientList();
    this.getMaterialListItem();
    this.getCurrentDate();
    this.getCurrentUser();
    this.UserProfiles();
    this.getDepartmentsList();
    this.getAdminApprover();
    // const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    // console.log('url: ', url);
  }

  public getCurrentDate() {
    const date = moment(new Date).format("DD-MM-YYYY");
    this.setState({ currentDate: date })
  }
  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    console.log('getcurrentuser: ', getcurrentuser);
  }

  public UserProfiles() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl
    const getUsers: string = url + `/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
    // console.log('getUsers: ', getUsers);
    fetch(getUsers, {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-Type': 'application/json'
      },
      credentials: 'include'
    })
      .then(response => response.json())
      .then(data => {
        console.log('data: ', data);
        // this.setState({ userProfileData: data });

        this.setState({
          //   displayName: data.DisplayName,
          //   mail: data.Email,
          //   photoUrl: data.PictureUrl,
          department: data.UserProfileProperties.filter((p: any) => p.Key === 'Department')[0].Value,
        });

      })
      .catch(error => {
        console.error('Error:', error);
      });
  }

  public async getClientList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listClient = await this._service.getClientListItems("ClientList", url)
    this.setState({ listClient: listClient })

    const ClientList: any[] = [];
    listClient.forEach((client: any) => {
      ClientList.push({ key: client.ID, text: client.Client });
    });
    this.setState({ client: ClientList });
  }

  public getProgramList = async (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption) => {
    this.setState({ selectedClient: item });
    // console.log('selectedClient: ', this.state.selectedClient);

    // console.log('this.state.selectedClient.key: ', this.state.selectedClient.key);
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listProgram = await this._service.getProgramListItems("ProgramList", item.key, url)
    this.setState({ listProgram: listProgram })

    const Program: any[] = [];
    listProgram.forEach((programItem: any) => {
      Program.push({ key: programItem.ID, text: programItem.Program });
    });
    this.setState({ program: Program });
  }

  public getProjectList = async (event: React.FormEvent<HTMLDivElement>, data: IDropdownOption) => {
    this.setState({ selectedProgram: data });

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listProject = await this._service.getProjectListItems("ProjectList", data.key, url)
    this.setState({ listProject: listProject })

    const ProjectList: any[] = [];
    listProject.forEach((project: any) => {
      ProjectList.push({ key: project.ID, text: project.Project });
    });
    this.setState({ project: ProjectList });
  }

  public getProjectChange(event: React.FormEvent<HTMLDivElement>, getProject: IDropdownOption) {
    this.setState({ getProject: getProject });
  }

  public onChangeComment(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, Comment: string) {
    this.setState({ comment: Comment });
  }

  public async getMaterialListItem() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const listMaterial = await this._service.getClientListItems("MaterialsMasterList", url)
    this.setState({ listMaterial: listMaterial })

    const MaterialList: any[] = [];
    listMaterial.forEach((materialItem: any) => {
      MaterialList.push({ key: materialItem.ID, text: materialItem.Materials });
    });
    this.setState({ material: MaterialList });
  }


  // public onMaterialChange(event: React.FormEvent<HTMLDivElement>, getMaterial: IDropdownOption) {
  //   this.setState({ getMaterial: getMaterial }, this.checkQuantityEntered);
  // }

  

  // public onMaterialChange = (event: React.FormEvent<HTMLDivElement>, getMaterial: IDropdownOption) => {
  //   const { selectedMaterials } = this.state;

  //   if (selectedMaterials.some((material: any) => material.key === getMaterial.key)) {
  //     this.setState({
  //       materialSelectionError: 'Material already selected!',
  //       getMaterial: { key: "", text: "" },
  //     });
  //   } else {
  //     this.setState({
  //       getMaterial,
  //       materialSelectionError: '',
  //     }, this.checkQuantityEntered);
  //   }
  // }

  // public onChangeQuantity = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, quantity: string) => {
  //   const isNumber = /^\d+$/.test(quantity); // Check if the input is a number
  //   if (!isNumber && quantity !== '') {
  //     this.setState({ quantityError: 'Please enter a valid number', isQuantityEntered: false });
  //   } else {
  //     this.setState({ quantity, quantityError: '' }, this.checkQuantityEntered);
  //   }
  // }

  public onMaterialChange = (event: React.FormEvent<HTMLDivElement>, getMaterial: IDropdownOption) => {
    const { selectedMaterials } = this.state;
  
    if (selectedMaterials.some((material: any) => material.key === getMaterial.key)) {
      this.setState({
        materialSelectionError: 'Material already selected!',
        getMaterial: { key: "", text: "" },
      });
    } else {
      this.setState({
        getMaterial,
        materialSelectionError: '',
      }, () => {
        this.checkQuantityEntered();
        this.firstRowChecking();
      });
    }
  }
  
  public onChangeQuantity = (event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, quantity: string) => {
    const isNumber = /^\d+$/.test(quantity); // Check if the input is a number
    if (!isNumber && quantity !== '') {
      this.setState({ quantityError: 'Please enter a valid number', isQuantityEntered: false });
    } else {
      this.setState({ quantity, quantityError: '' }, () => {
        this.checkQuantityEntered();
        this.firstRowChecking();
      });
    }
  }
  

  public checkQuantityEntered = () => {
    const isMaterialSelected = this.state.getMaterial && this.state.getMaterial.key !== "";
    const isQuantityEntered = this.state.quantity !== "" && this.state.quantityError === "";

    this.setState({ isQuantityEntered: isMaterialSelected && isQuantityEntered });
  }

  public firstRowChecking = () => {
    const isFirstMaterialSelected = this.state.getMaterial && this.state.getMaterial.key !== "";
    const isFirstQuantityEntered = this.state.quantity !== "" && this.state.quantityError === "";

    this.setState({ isFirstRowSelected: isFirstMaterialSelected && isFirstQuantityEntered });
  }



  public async getDepartmentsList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const DepartmentslistItem = await this._service.getListItems("Departments", url)
    // console.log('DepartmentslistItem: ', DepartmentslistItem);
    this.setState({ Departmentslist: DepartmentslistItem });

    // const Departmentslist: any[] = [];
    DepartmentslistItem.map((Item: any) => {
      // console.log('Item: ', Item);
      const departmentName = Item.Title;
      const HOSName = Item.HOSNameId;

      this.setState({
        departmentName: departmentName,
        HOSName: HOSName,
      });
    })
    // console.log('Departmentslist: ', this.state.Departmentslist);
  }

  public async getAdminApprover() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const AdminApproverlistItem = await this._service.getListItems("AdminApprover", url)
    const ApproverIdUserInfo = await this._service.getUser(AdminApproverlistItem[0].AdminApproverId);
    // const AdminApprover = ApproverIdUserInfo.Title; 
    this.setState({ adminApproverName: ApproverIdUserInfo.Id });
  }

  // public addRow() {
  //   const { getMaterial, quantity } = this.state;
  //   const newRow = { getMaterial, quantity };
  //   const rows = [...this.state.rows, newRow];

  //   this.setState({
  //     rows,
  //     getMaterial: { key: "", text: "" },
  //     quantity: "",
  //     isQuantityEntered: false,
  //   });
  // }


  public addRow() {
    const { getMaterial, quantity, selectedMaterials } = this.state;
    const newRow = { getMaterial, quantity };
    const rows = [...this.state.rows, newRow];

    this.setState({
      rows,
      getMaterial: { key: "", text: "" },
      quantity: "",
      isQuantityEntered: false,
      selectedMaterials: [...selectedMaterials, getMaterial],
    });
  }



  // public deleteRow(index: number) {
  //   const rows = [...this.state.rows];
  //   rows.splice(index, 1);
  //   this.setState({ rows });
  // }

  public deleteRow(index: number) {
    const rows = [...this.state.rows];
    const deletedRow = rows[index];

    // Remove the deleted material from selectedMaterials
    const updatedSelectedMaterials = this.state.selectedMaterials.filter(
      (material: any) => material.key !== deletedRow.getMaterial.key
    );

    rows.splice(index, 1);
    this.setState({ rows, selectedMaterials: updatedSelectedMaterials });
  }


  public async onSubmitClick(): Promise<void> {
    this.setState({
      isPopupVisible: true,
      //  isLoading: true
    });
  }

  public async onPopOk(): Promise<void> {
    await this.setState({ isOkButtonDisabled: true });
    const filteredDepartment = this.state.Departmentslist.find((dept: any) => dept.Title === this.state.department);
    // console.log('this.state.department: ', this.state.department);

    if (filteredDepartment) {
      const HOSName = filteredDepartment.HOSNameId;

      const dataItem = {
        //  MaterialRequestCode: ,
        ClientId: this.state.selectedClient.key,
        ProgramId: this.state.selectedProgram.key,
        ProjectId: this.state.getProject.key,
        RequestorComments: this.state.comment,
        // HOSApprovalComments
        Status: "Pending",
        // AdminComments
        HOSApproverId: HOSName,
        AdminApproverId: this.state.adminApproverName,
        // HOSApprovedDate
        // AdminApprovedDate
        // Created: this.state.currentDate,
        // Author
      };

      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
      await this._service.addMaterialRequestForm(dataItem, "MasterMaterialRequestList", url).then(async (item: any) => {
        console.log('item: ', item);

        // const taskURL = url + "/SitePages/" + "MaterialRequestApprovalForm" + ".aspx?did=" + item.data.Id + "&formType=HOSApproval";
        // console.log('taskURL: ', taskURL);

        const itemId = item.data.Id;
        this.setState({ MasterMaterialRequestId: itemId });
        const dataItem = {
          MaterialRequestCode: "00" + itemId
        }
        await this._service.updateMaterialRequestForm("MasterMaterialRequestList", dataItem, itemId, url);

        const materialItemsData = this.state.rows.map(row => ({
          MasterMaterialRequestIDId: itemId,
          MaterialsIDId: row.getMaterial.key,
          Quantity: row.quantity,
        }));

        for (const materialItemData of materialItemsData) {
          await this._service.addMaterialRequestForm(materialItemData, "MaterialItemsList", url);
        }

        if (!this.state.rows.some(row => row.getMaterial.key === this.state.getMaterial.key) && this.state.getMaterial.key && this.state.quantity) {
          const lastMaterialItemData = {
            MasterMaterialRequestIDId: itemId,
            MaterialsIDId: this.state.getMaterial.key,
            Quantity: this.state.quantity,
          };
          await this._service.addMaterialRequestForm(lastMaterialItemData, "MaterialItemsList", url);
        }

        const taskItem = {
          MasterMaterialRequestIDId: itemId,
          AssignedToId: HOSName,
          // TaskTitleWithLink: {
          //   Description: "-- HOS Approval",
          //   Url: taskURL,
          // }
        }
        await this._service.addMaterialRequestForm(taskItem, "TasksList", url).then(async (task: any) => {

          this.setState({ taskListItemId: task.data.Id });
          const taskURL = url + "/SitePages/" + "MaterialRequestApprovalForm" + ".aspx?did=" + item.data.Id + "&itemid=" + task.data.Id + "&formType=HOSApproval";
          const taskItemtoupdate = {
            TaskTitleWithLink: {
              Description: "-- HOS Approval",
              Url: taskURL,
            }
          }
          await this._service.updateMaterialRequestForm("TasksList", taskItemtoupdate, task.data.Id, url);


          await this.sendEmailNotificationToHOS(this.props.context);

          this.hidePopup();

          // this.setState({ statusMessage: 'Successfully applied' });
          Toast("success", "Successfully Submitted");
          setTimeout(() => {
            window.location.href = url;
          }, 3000);
        });
      });
    } else {
      console.error('Department not found');
    }
  }

  public async sendEmailNotificationToHOS(context: any): Promise<void> {
    const filteredDepartment = this.state.Departmentslist.find((dept: any) => dept.Title === this.state.department);

    if (filteredDepartment && filteredDepartment.HOSNameId) {
      const hosApproverId = filteredDepartment.HOSNameId;
      const hosApproverIdUserInfo = await this._service.getUser(hosApproverId);
      const HOSApproverEmail = hosApproverIdUserInfo.Email;
      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
      const evaluationURL = "https://ccsdev01.sharepoint.com/" + url + "/SitePages/" + "MaterialRequestApprovalForm" + ".aspx?did=" + this.state.MasterMaterialRequestId + "&itemid=" + this.state.taskListItemId + "&formType=HOSApproval";

      // Get the Project Name
      const project = this.state.project.find((proj: any) => proj.key === this.state.getProject.key)?.text || 'Project';

      // Get the Employee Name
      const getcurrentuser = await this._service.getCurrentUser();
      const getcurrentUserInfo = await this._service.getUser(getcurrentuser.Id);
      const employeeName = getcurrentUserInfo.Title;

      // Get the Requested Date
      const requestedDate = this.state.currentDate;

      const emailPostBody: any = {
        message: {
          subject: `Material Request for ${project}`,
          body: {
            contentType: 'HTML',
            content: `Hi ${hosApproverIdUserInfo.Title},<br><br>
            ${employeeName} has submitted a material request for the ${project} on ${requestedDate}.<br>
            Please click on the<a href="${evaluationURL}" target="_blank">link</a> to review the details and kindly approve the request at your earliest convenience.<br><br>
            `
          },
          toRecipients: [
            {
              emailAddress: {
                address: HOSApproverEmail,
              },
            },
          ],
        },
      };

      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }


  public hidePopup = () => {
    this.setState({ isPopupVisible: false });
  };

  public onClickCancel() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    window.location.href = url;
  };



  public render(): React.ReactElement<IMaterialRequestWpProps> {
    const currentUrl = window.location.href;

    if (currentUrl === 'https://ccsdev01.sharepoint.com/sites/MaterialRequest/SitePages/ViewSubmittedRequests.aspx?debug=true&noredir=true&debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fmanifests.js') {
      // if (currentUrl === 'https://ccsdev01.sharepoint.com/sites/MaterialRequest/SitePages/ViewSubmittedRequests.aspx') {
      return <MaterialRequestViewTable
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
      />;
    }


    else if (new URLSearchParams(window.location.search).get("formType") === "HOSApproval") {
      return <MaterialRequestHOSApprovalForm
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
      />;
    } else if (new URLSearchParams(window.location.search).get("formType") === "AdminApproval") {
      return <MaterialRequestAdminApprovalForm
        description={''}
        isDarkTheme={false}
        environmentMessage={''}
        hasTeamsContext={false}
        userDisplayName={''}
        context={this.props.context}
      />;
    }
    else {

      const deleteIcon: IIconProps = { iconName: 'Delete' };
      const addIcon: IIconProps = { iconName: 'CircleAdditionSolid' };
      const {
        hasTeamsContext,
      } = this.props;


      return (
        <section className={`${styles.materialRequestWp} ${hasTeamsContext ? styles.teams : ''}`}>

          <div className={styles.borderBox}>
            <div className={styles.MaterialRequestHeading}>{"Material Request"}</div>

            <div className={styles.onediv}>

              <div className={styles.fieldWrapper}>
                {/* <label className={styles.fieldLabel}>Request Date</label> */}
                <Label className={styles.fieldLabel} required={true} >Request Date</Label>
                <TextField
                  className={styles.fieldInput}
                  value={this.state.currentDate}
                />
              </div>

              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true} >Client</Label>
                <Dropdown
                  className={styles.fieldInput}
                  placeholder="Select One"
                  options={this.state.client}
                  onChange={this.getProgramList}
                  selectedKey={this.state.selectedClient.key}
                />
              </div>

              <div className={styles.fieldWrapper}>
                <Label className={styles.fieldLabel} required={true}> Program</Label>
                <Dropdown
                  className={styles.fieldInput}
                  placeholder="Select One"
                  options={this.state.program}
                  onChange={this.getProjectList}
                  selectedKey={this.state.selectedProgram.key}
                />
              </div>

              <div className={styles.fieldWrapper}>
                <Label required={true}> Project</Label>
                <Dropdown
                  className={styles.fieldInput}
                  placeholder="Select One"
                  options={this.state.project}
                  onChange={this.getProjectChange}
                  selectedKey={this.state.getProject.key}
                />
              </div>
            </div>

            <div>
              <table className={styles.table}>
                <thead>
                  <tr>
                    <th className={styles.tablediv}>SL No</th>
                    <th className={styles.tablediv}>Material Request</th>
                    <th className={styles.tablediv}>Quantity</th>
                    <th className={styles.iconButton}></th>
                  </tr>
                </thead>
                <tbody>
                  {this.state.rows.map((row, index) => (
                    <tr key={index}>
                      <td className={styles.tablediv}>{index + 1}</td>
                      <td className={styles.tablediv}>
                        <Dropdown
                          placeholder="Material Request"
                          options={this.state.material}
                          onChange={(event, option) => {
                            const rows = [...this.state.rows];
                            rows[index].getMaterial = option;
                            this.setState({ rows });
                          }}
                          selectedKey={row.getMaterial?.key}
                        />
                      </td>
                      <td className={styles.tablediv}>
                        <TextField
                          placeholder="Quantity"
                          onChange={(event, newValue) => {
                            const rows = [...this.state.rows];
                            rows[index].quantity = newValue || "";
                            this.setState({ rows }, this.checkQuantityEntered);
                          }}
                          value={row.quantity}
                        />
                      </td>
                      <td>
                        <IconButton
                          iconProps={deleteIcon}
                          ariaLabel="Delete item"
                          onClick={() => this.deleteRow(index)}
                          className={styles.iconButton}
                        />
                      </td>
                    </tr>
                  ))}
                  <tr>
                    <td className={styles.tablediv}>{this.state.rows.length + 1}</td>
                    <td className={`${styles.tablediv} ${styles.dropdownstyles}`}>
                      <Dropdown
                        placeholder="Material Request"
                        required={true}
                        options={this.state.material}
                        onChange={this.onMaterialChange}
                        selectedKey={this.state.getMaterial.key}
                        className={styles.dropdownpadding}
                      />
                      {/* Material Selection Error Message */}
                      {this.state.materialSelectionError && (
                        <div className={styles.error}>{this.state.materialSelectionError}</div>
                      )}

                    </td>
                    <td className={styles.tablediv}>
                      <TextField
                        required={true}
                        placeholder="0"
                        onChange={this.onChangeQuantity}
                        value={this.state.quantity}
                        className={styles.dropdownpadding}
                      />
                      {this.state.quantityError && (
                        <div className={styles.error}>{this.state.quantityError}</div>
                      )}
                    </td>
                    <td>
                      <IconButton
                        iconProps={addIcon}
                        ariaLabel="Add item"
                        disabled={!this.state.isQuantityEntered}
                        onClick={this.addRow}
                        className={styles.iconButton}
                      />
                    </td>
                  </tr>
                </tbody>

              </table>
            </div>

            <div>
              <TextField
                label="Comment"
                multiline rows={3}
                onChange={this.onChangeComment}
                value={this.state.comment}
                className={styles.commentArea}
              />
            </div>

            <div className={styles.reuired} >
              {"* All fields are required"}
            </div>

            <div className={styles.btndiv}>
              <PrimaryButton
                text="Submit"
                onClick={this.onSubmitClick}
                disabled={
                  !this.state.selectedClient.key ||
                  !this.state.selectedProgram.key ||
                  !this.state.getProject.key ||
                  !this.state.isFirstRowSelected
                }
                
              />

              <DefaultButton
                text="Cancel"
                onClick={this.onClickCancel}
              />
            </div>

            {/* status message */}
            <div className={styles.statusMessage}>
              {this.state.statusMessage && <span>{this.state.statusMessage}</span>}
            </div>
          </div>


          {/* pop up */}
          <div>
            {this.state.isPopupVisible && (
              <Layer>
                <Popup
                  className={styles.root}
                  role="dialog"
                  aria-modal="true"
                  onDismiss={this.hidePopup}
                >
                  <Overlay
                    onClick={this.hidePopup}
                  />
                  <FocusTrapZone>
                    <div
                      role="document"
                      className={styles.content}
                    >
                      <div>
                        Did you want to apply?
                      </div>

                      <div className={styles.popbtndiv}>
                        <PrimaryButton
                          onClick={this.onPopOk}
                          text="Yes"
                          disabled={this.state.isOkButtonDisabled}

                        />
                        <DefaultButton onClick={this.hidePopup} >No </DefaultButton>
                      </div>

                    </div>
                  </FocusTrapZone>
                </Popup>
              </Layer>
            )}
          </div>

        </section >
      );
    }
  }
}