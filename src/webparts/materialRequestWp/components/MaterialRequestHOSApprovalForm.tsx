import * as React from 'react';
import styles from './MaterialRequestHOSApprovalForm.module.scss';
import { IMaterialRequestHOSApprovalFormProps, IMaterialRequestHOSApprovalFormState } from '../interfaces';
import { MaterialRequestHOSApprovalFormsService } from '../services';
import * as moment from 'moment';
import { MessageBar, MessageBarType, PrimaryButton, TextField } from '@fluentui/react';
import { MSGraphClientV3 } from '@microsoft/sp-http';
import Toast from './Toast';
import * as strings from 'MaterialRequestWpWebPartStrings';
import replaceString from 'replace-string';

export default class MaterialMaterialRequestHOSApprovalForm extends React.Component<IMaterialRequestHOSApprovalFormProps, IMaterialRequestHOSApprovalFormState, {}> {
  private _service: any;

  public constructor(props: IMaterialRequestHOSApprovalFormProps) {
    super(props);
    this._service = new MaterialRequestHOSApprovalFormsService(this.props.context);

    this.state = {
      isTaskIdPresent: "",
      noAccessId: "",
      statusMessageTAskIdNull: "",
      materialRequestData: "",
      RequestedBy: "",
      RequestedDate: "",
      client: "",
      program: "",
      project: "",
      materialRequestDataId: null,
      materialId: null,
      materialQuantity: null,
      masterMaterial: [],
      materialDataArray: [],
      getItemId: "",
      comment: "",
      taskListItemId: null,
      isPopupVisibleForApprove: false,
      isPopupVisibleForReject: false,
      successfullStatusMessage: '',
      rejectStatusMessage: '',
      getcurrentuserId: null,
      isOkButtonDisabled: false,

    }
    this.getMasterMaterialRequestListData = this.getMasterMaterialRequestListData.bind(this);
    this.getmaterialList = this.getmaterialList.bind(this);
    this.onChangeComment = this.onChangeComment.bind(this);
    this.OnClickApprove = this.OnClickApprove.bind(this);
    this.sendApprovedEmailNotificationToAdminFromHOS = this.sendApprovedEmailNotificationToAdminFromHOS.bind(this);
    this.sendApprovedEmailNotificationToRequestorFromHOS = this.sendApprovedEmailNotificationToRequestorFromHOS.bind(this);
    this.deleteTaskListItem = this.deleteTaskListItem.bind(this);
    this.OnClickReject = this.OnClickReject.bind(this);
    this.sendRejectEmailNotificationToRequestorFromHOS = this.sendRejectEmailNotificationToRequestorFromHOS.bind(this);
    this.getTaskList = this.getTaskList.bind(this);
    this.checkHOS = this.checkHOS.bind(this);
    this.getCurrentUser = this.getCurrentUser.bind(this);

  }


  public async componentDidMount() {
    await this.getCurrentUser();
    await this.getMasterMaterialRequestListData();
    await this.checkHOS();
    this.getmaterialList();
    this.getTaskList();


  }
  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    this.setState({ getcurrentuserId: getcurrentuser.Id });
  }

  public checkHOS() {
    if (this.state.getcurrentuserId !== this.state.materialRequestData.HOSApproverId) {
      this.setState({
        noAccessId: "false",
        statusMessageTAskIdNull: strings.AccessDenied
      });
    } else {
      this.setState({ noAccessId: "true" });
    }
  }

  public onChangeComment(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, Comment: string) {
    this.setState({ comment: Comment });
  }

  public async getTaskList() {
    const taskItemid = new URLSearchParams(window.location.search).get('itemid');

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const taskListData: any[] = await this._service.getItemSelectExpandFilter(
      url,
      this.props.TasksList,
      "ID, TaskTitleWithLink, MasterMaterialRequestID/ID",
      "MasterMaterialRequestID",
      `ID eq ${taskItemid}`
    );

    if (taskListData.length === 0) {
      this.setState({
        isTaskIdPresent: "false",
        statusMessageTAskIdNull: strings.Alreadycheckedtherequest
      });
    } else {
      this.setState({ isTaskIdPresent: "true" });
    }
  }


  public async getMasterMaterialRequestListData() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const itemId = new URLSearchParams(window.location.search).get('did');
    this.setState({ getItemId: itemId });
    const materialRequestData = await this._service.getItemById(url, this.props.MasterMaterialRequestList, itemId);
    this.setState({ materialRequestData: materialRequestData });

    const requestedBy = await this._service.getUser(this.state.materialRequestData.AuthorId);
    const RequestedBy = requestedBy.Title;

    const date = this.state.materialRequestData.Created
    const dateformatted = moment(date).format("DD-MM-YYYY");

    const clientListData = await this._service.getItemById(url, this.props.ClientList, this.state.materialRequestData.ClientId);
    const programListData = await this._service.getItemById(url, this.props.ProgramList, this.state.materialRequestData.ProgramId);
    const projectListData = await this._service.getItemById(url, this.props.ProjectList, this.state.materialRequestData.ProjectId);

    const materialItemId = this.state.materialRequestData.Id

    this.setState({
      materialRequestDataId: materialItemId,
      RequestedBy: RequestedBy,
      RequestedDate: dateformatted,
      client: clientListData.Client,
      program: programListData.Program,
      project: projectListData.Project

    });

  }

  public async getmaterialList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const MaterialListData = await this._service.getItemSelectExpandFilter(
      url,
      this.props.MaterialItemsList,
      "MasterMaterialRequestID/ID,MasterMaterialRequestID/Title,MaterialsID/ID,MaterialsID/Title,Quantity",
      "MasterMaterialRequestID, MaterialsID",
      `MasterMaterialRequestID/ID eq ${this.state.materialRequestDataId}`
    );

    const materialDataArray: any[] = [];
    MaterialListData.map((material: any) => {

      const MaterialId = material.MaterialsID.ID;
      const MaterialQuantity = material.Quantity;

      materialDataArray.push({
        materialId: MaterialId,
        materialQuantity: MaterialQuantity,
      });
    });

    this.setState({
      materialDataArray: materialDataArray,
    });

    const getmasterMaterials = materialDataArray.map(async (item) => {
      const MaterialMasterListData = await this._service.getItemById(url, this.props.MaterialsMasterList, item.materialId);
      return MaterialMasterListData.Materials;
    });

    const masterMaterials = await Promise.all(getmasterMaterials);

    this.setState({
      masterMaterial: masterMaterials,
    });
  }

  public async OnClickApprove() {
    await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: strings.HOSApproved,
      HOSApprovalComments: this.state.comment,
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateEvaluation(this.props.MasterMaterialRequestList, itemsForUpdate, this.state.getItemId, url);

    await this.deleteTaskListItem();


    const dataItem = {
      MasterMaterialRequestIDId: this.state.getItemId,
      AssignedToId: this.state.materialRequestData.AdminApproverId,
    };
    this._service.addListItem(dataItem, this.props.TasksList, url).then(async (task: any) => {

      this.setState({ taskListItemId: task.data.Id });
      const taskURL = url + "/SitePages/" + "MaterialRequestApprovalForm" + ".aspx?did=" + this.state.getItemId + "&itemid=" + task.data.Id + "&formType=AdminApproval";
      const taskItemtoupdate = {
        TaskTitleWithLink: {
          Description: "-- Admin Approval",
          Url: taskURL,
        }
      }
      await this._service.updateEvaluation(this.props.TasksList, taskItemtoupdate, task.data.Id, url);

      await this.sendApprovedEmailNotificationToRequestorFromHOS(this.props.context);
      await this.sendApprovedEmailNotificationToAdminFromHOS(this.props.context);

      Toast("success", strings.SuccessfullySubmitted);
      setTimeout(() => {
        window.location.href = url;
      }, 3000);
    });
  }

  public async OnClickReject() {
    await this.setState({ isOkButtonDisabled: true });
    const itemsForUpdate = {
      Status: strings.HOSRejected,
      HOSApprovalComments: this.state.comment,
    };

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    await this._service.updateEvaluation(this.props.MasterMaterialRequestList, itemsForUpdate, this.state.getItemId, url);

    await this.deleteTaskListItem();

    await this.sendRejectEmailNotificationToRequestorFromHOS(this.props.context);
    Toast("warning", strings.SuccessfullySubmitted);
    setTimeout(() => {
      window.location.href = url;
    }, 3000);
  }

  public async sendApprovedEmailNotificationToAdminFromHOS(context: any): Promise<void> {
    const AdminApproverIdUserInfo = await this._service.getUser(this.state.materialRequestData.AdminApproverId);
    const AdminApproverEmail = AdminApproverIdUserInfo.Email;

    // const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const url: string = this.props.context.pageContext.web.absoluteUrl;
    const evaluationURL = url + "/SitePages/" + "MaterialRequestApprovalForm" + ".aspx?did=" + this.state.getItemId + "&itemid=" + this.state.taskListItemId + "&formType=AdminApproval";

    const HOSInfo = await this._service.getUser(this.state.materialRequestData.HOSApproverId);
    const HOSName = HOSInfo.Title;


    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.MaterialRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToAdminFromHOS");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceAdminApprover = replaceString(bodyTemplate, '[AdminApprover]', AdminApproverIdUserInfo.Title)
      const replaceRequestedBy = replaceString(replaceAdminApprover, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceRequestedDate = replaceString(replaceProject, '[RequestedDate]', this.state.RequestedDate)
      const replaceHOSName = replaceString(replaceRequestedDate, '[HOSName]', HOSName)
      const replacedBodyWithLink = replaceString(replaceHOSName, '[Link]', `<a href="${evaluationURL}" target="_blank">Click here</a>`);

      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replacedBodyWithLink
          },
          toRecipients: [
            {
              emailAddress: {
                address: AdminApproverEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async sendApprovedEmailNotificationToRequestorFromHOS(context: any): Promise<void> {
    const requestedBy = await this._service.getUser(this.state.materialRequestData.AuthorId);
    const requestedEmail = requestedBy.Email;

    const HOSInfo = await this._service.getUser(this.state.materialRequestData.HOSApproverId);
    const HOSName = HOSInfo.Title;
    const date = moment(new Date).format("DD-MM-YYYY");

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.MaterialRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendApprovedEmailNotificationToRequestorFromHOS");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceHOSName = replaceString(replaceProject, '[HOSName]', HOSName)
      const replaceDate = replaceString(replaceHOSName, '[Date]', date)


      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replaceDate
          },
          toRecipients: [
            {
              emailAddress: {
                address: requestedEmail,
              },
            },
          ]
        },
      };
      return context.msGraphClientFactory
        .getClient('3')
        .then((client: MSGraphClientV3): void => {
          client.api('/me/sendMail').post(emailPostBody);
        });
    }
  }

  public async sendRejectEmailNotificationToRequestorFromHOS(context: any): Promise<void> {
    const requestedBy = await this._service.getUser(this.state.materialRequestData.AuthorId);
    const requestedEmail = requestedBy.Email;

    const HOSInfo = await this._service.getUser(this.state.materialRequestData.HOSApproverId);
    const HOSName = HOSInfo.Title;

    const date = moment(new Date).format("DD-MM-YYYY");

    const serverurl: string = this.props.context.pageContext.web.serverRelativeUrl;
    const emailNoficationSettings = await this._service.getListItems(this.props.MaterialRequestSettingsList, serverurl);
    const emailNotificationSetting = emailNoficationSettings.find((item: any) => item.Title === "SendRejectEmailNotificationToRequestorFromHOS");

    if (emailNotificationSetting) {
      const subjectTemplate = emailNotificationSetting.Subject;
      const bodyTemplate = emailNotificationSetting.Body;

      const replaceSubject = replaceString(subjectTemplate, '[Project]', this.state.project)

      const replaceRequestedBy = replaceString(bodyTemplate, '[RequestedBy]', this.state.RequestedBy)
      const replaceProject = replaceString(replaceRequestedBy, '[Project]', this.state.project)
      const replaceHOSName = replaceString(replaceProject, '[HOSName]', HOSName)
      const replaceDate = replaceString(replaceHOSName, '[Date]', date)

      const emailPostBody: any = {
        message: {
          subject: replaceSubject,
          body: {
            contentType: 'HTML',
            content: replaceDate
          },
          toRecipients: [
            {
              emailAddress: {
                address: requestedEmail,
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

  public async deleteTaskListItem() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const taskListData = await this._service.getItemSelectExpandFilter(
      url,
      this.props.TasksList,
      "ID, MasterMaterialRequestID/ID",
      "MasterMaterialRequestID",
      `MasterMaterialRequestID/ID eq ${this.state.getItemId}`
    );

    const taskIdItem = taskListData[0].ID;
    await this._service.deleteItemById(url, this.props.TasksList, taskIdItem);
  }


  public render(): React.ReactElement<IMaterialRequestHOSApprovalFormProps> {

    const {

    } = this.props;
    return (
      <section>
        <div className={styles.borderBox}>
          <div>

            {this.state.noAccessId === "false" &&
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.statusMessageTAskIdNull}
              </MessageBar>
            }

          </div>

          <div>

            {this.state.isTaskIdPresent === "false" && this.state.noAccessId === "true" &&
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.statusMessageTAskIdNull}
              </MessageBar>
            }

          </div>

          <div>
            {this.state.isTaskIdPresent === "true" && this.state.noAccessId === "true" &&
              <><div>
                <div className={styles.MaterialRequestHeading}>{"Material Request"}</div>

                <div className={styles.onediv}>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Requested By</div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.RequestedBy}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Requested Date </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.RequestedDate}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Client </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.client}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Program </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.program}</div>
                  </div>
                  <div className={styles.fieldwrapper}>
                    <div className={styles.fieldlabel}>Project </div>
                    <div className={styles.colon}>:</div>
                    <div className={styles.fieldoutput}>{this.state.project}</div>
                  </div>

                </div>

                <div>
                  <table className={`${styles.table} ${styles.tablethtddiv}`}>
                    <thead>
                      <tr>
                        <th className={styles.tablethtddiv}>SL No</th>
                        <th className={styles.tablethtddiv}>Material Request</th>
                        <th className={styles.tablethtddiv}>Quantity</th>

                      </tr>
                    </thead>
                    <tbody>
                      {this.state.masterMaterial.map((material: any, index: any) => (
                        <tr key={index}>
                          <td className={styles.tablethtddiv}>{index + 1}</td>
                          <td className={styles.tablethtddiv}>{material}</td>
                          <td className={styles.tablethtddiv}>{this.state.materialDataArray[index].materialQuantity}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>

                <div>
                  <TextField
                    label="Comment"
                    multiline rows={3}
                    onChange={this.onChangeComment}
                    value={this.state.comment} />
                </div>

                <div className={styles.btndiv}>
                  <PrimaryButton
                    text="Approve"
                    onClick={this.OnClickApprove}
                    disabled={this.state.isOkButtonDisabled}
                  />

                  <PrimaryButton
                    text="Reject"
                    onClick={this.OnClickReject}
                    disabled={this.state.isOkButtonDisabled}
                  />
                </div>
              </div>
              </>
            }

          </div>
        </div>
      </section>
    );
  }
}

