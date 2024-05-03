declare interface IMaterialRequestWpWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  Pending : string;
  HOSApproved : string;
  HOSRejected : string;
  AdminApproved : string;
  AdminRejected : string;

  AccessDenied : string;
  Alreadycheckedtherequest : string;
  SuccessfullySubmitted : string;

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

declare module 'MaterialRequestWpWebPartStrings' {
  const strings: IMaterialRequestWpWebPartStrings;
  export = strings;
}
