import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMaterialRequestHOSApprovalFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}


export interface IMaterialRequestHOSApprovalFormState {
  
  materialRequestData: any;
  RequestedBy: string;
  RequestedDate: string;
  client: string;
  program: string;
  project: string;
  materialRequestDataId: Number;
  materialId: Number;
  materialQuantity: Number;
  masterMaterial: any;
  materialDataArray: any;
  getItemId: string;
  comment: string;
  taskListItemId: Number;
  isTaskIdPresent: any;
  noAccessId: any;
  statusMessageTAskIdNull: string;
  isPopupVisibleForApprove: boolean;
  isPopupVisibleForReject: boolean;
  successfullStatusMessage: string;
  rejectStatusMessage: string;
  getcurrentuserId: Number;
  isOkButtonDisabled: boolean; 

}

