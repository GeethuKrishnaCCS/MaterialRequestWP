import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMaterialRequestAdminApprovalFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;

  AdminApprover : string;
  ClientList : string;
  Departments : string;
  MasterMaterialRequestList : string;
  MaterialItemsList : string;
  MaterialsMasterList : string;
  ProgramList : string;
  ProjectList : string;
  TasksList : string;
  MaterialRequestSettingsList : string;
}


export interface IMaterialRequestAdminApprovalFormState {
  materialRequestData: any;
  RequestedBy: string;
  RequestedDate: string;
  client: string;
  program: string;
  project: string;
  getItemId: string;
  materialRequestDataId: Number;
  ApprovedBy: string;
  RequestorComments: string;
  ApproverComments: string;
  masterMaterial: any;
  materialDataArray: any;
  comment: string;
  successfullStatusMessage: string;
  rejectStatusMessage: string;
  isTaskIdPresent: any; 
  noAccessId: any; 
  statusMessageTAskIdNull: string,
  getcurrentuserId: Number;
  isOkButtonDisabled: boolean; 

}

