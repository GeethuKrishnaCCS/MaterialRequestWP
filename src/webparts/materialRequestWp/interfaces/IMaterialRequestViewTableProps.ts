import { IGroup } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMaterialRequestViewTableProps {
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


export interface IMaterialRequestViewTableState {
  materialDatas: any;
  // materialRequestData: any;
  groups?: IGroup[];
  expandedItems: any[];
  statusGroups: any[];
  combinedGroups: any[];
  adminApproverName: number;
  departmentName: '';
  HOSName: null;
  Departmentslist: [];
  getcurrentuserId: number;
  noItems: any;
  statusMessageNoItems: string;
  currentPage: number;
  itemsPerPage: number;


}

