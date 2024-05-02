import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMaterialRequestAdminApprovalFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
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

