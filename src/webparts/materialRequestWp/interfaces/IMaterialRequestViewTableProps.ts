import { IGroup } from "@fluentui/react";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMaterialRequestViewTableProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
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

