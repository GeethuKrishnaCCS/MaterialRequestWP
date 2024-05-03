import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDropdownOption } from "office-ui-fabric-react";

export interface IMaterialRequestWpProps {
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
export interface IMaterialRequestWpWebPartProps {
  description: string;
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

export interface IMaterialRequestWpState {

  listClient: any[];
  listMaterial: any[];
  client: any;
  material: any;
  selectedClient: IDropdownOption;
  selectedProgram: IDropdownOption;
  getProject: IDropdownOption;
  getMaterial: IDropdownOption;
  listProgram: any[];
  listProject: any[];
  program: any;
  project: any;
  comment: string;
  quantity: string;
  isQuantityEntered: boolean;
  currentDate: string;
  rows: any[];
  HOSName: Number;
  Departmentslist: any;
  department: string;
  departmentName: string;
  navigateToList: boolean;
  adminApproverName: string;
  quantityError: string;
  isPopupVisible: boolean;
  statusMessage:string;
  MasterMaterialRequestId: Number;
  taskListItemId: Number;
  // isLoading: boolean;
  isOkButtonDisabled: boolean;
  selectedMaterials: any;
  materialSelectionError: string;
  isFirstRowSelected: boolean;
  
  

}

