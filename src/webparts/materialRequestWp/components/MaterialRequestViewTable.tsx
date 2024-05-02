import * as React from 'react';
import styles from './MaterialRequestViewTable.module.scss';
import { IMaterialRequestViewTableProps, IMaterialRequestViewTableState } from '../interfaces';
import { DetailsList, DetailsListLayoutMode, SelectionMode, IColumn, } from '@fluentui/react/lib/DetailsList';
//IGroup
import { MaterialRequestViewTableService } from '../services';
import * as moment from 'moment';
import { Pagination } from '@pnp/spfx-controls-react/lib/pagination';
import * as _ from 'lodash';
import { MessageBar, MessageBarType } from '@fluentui/react';


export default class MaterialRequestViewTable extends React.Component<IMaterialRequestViewTableProps, IMaterialRequestViewTableState, {}> {
  private _service: any;

  public constructor(props: IMaterialRequestViewTableProps) {
    super(props);
    this._service = new MaterialRequestViewTableService(this.props.context);

    this.state = {
      noItems: "",
      statusMessageNoItems: "",
      materialDatas: [],
      groups: [],
      // materialRequestData: [],
      expandedItems: [],
      statusGroups: [],
      combinedGroups: [],
      adminApproverName: null,
      departmentName: '',
      HOSName: null,
      Departmentslist: [],
      getcurrentuserId: null,
      currentPage: 1,
      itemsPerPage: 5,



    }
    //this.getMasterMaterialRequestListData = this.getMasterMaterialRequestListData.bind(this);

    this.getDocumentIndexItems = this.getDocumentIndexItems.bind(this);;
    this.getAdminApprover = this.getAdminApprover.bind(this);
    this.getCurrentUser = this.getCurrentUser.bind(this);
    // this.getTotalPages = this.getTotalPages.bind(this);
    this.getDepartmentsList = this.getDepartmentsList.bind(this);
    //   this.getCurrentPageItems = this.getCurrentPageItems.bind(this);
    this.handlePageChange = this.handlePageChange.bind(this);
  }


  public async componentDidMount() {
    await this.getAdminApprover();
    await this.getCurrentUser();
    await this.getDepartmentsList();
    // this.getMasterMaterialRequestListData();

    this.getDocumentIndexItems();
  }

  // public async getDocumentIndexItems() {

  //      const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const currentUserId = this.state.getcurrentuserId;
  //   const isAdmin = currentUserId === this.state.adminApproverName;
  //   console.log('isAdmin: ', isAdmin);
  //   const isHOS = this.state.Departmentslist.some((dept: any) => dept.HOSNameId === currentUserId);
  //   console.log('isHOS: ', isHOS);



  //   //let materialRequestData = await this._service.getListItems("MasterMaterialRequestList", url);
  //   let materialRequestData = await this._service.getItemSelectExpand(url, "MasterMaterialRequestList",
  //     "*,Project/ID,Project/Project,Client/ID,Client/Client,Program/ID,Program/Program", "Project,Client,Program");
  //   console.log(materialRequestData);
  //   const materialItemListVal = await this._service.getItemSelectExpand(
  //     url,
  //     "MaterialItemsList",
  //     "MasterMaterialRequestID/ID,MasterMaterialRequestID/Title,MaterialsID/ID,MaterialsID/Materials,Quantity",
  //     "MasterMaterialRequestID, MaterialsID"
  //     //`MasterMaterialRequestID/ID eq ${data.Id}`
  //   );
  //   if (isAdmin) {
  //     // Display all items for admin user
  //   } else if (isHOS) {
  //     // Display items submitted by the current user and those users' department HOS is current user
  //     materialRequestData = materialRequestData.filter((data: any) => data.AuthorId === currentUserId ||
  //       materialRequestData.filter((item: any) => item.HOSApproverId === currentUserId));
  //   } else {
  //     // Display items submitted by the current user
  //     materialRequestData = materialRequestData.filter((data: any) => data.AuthorId === currentUserId);
  //   }

  //   //const materialDatas: any[] = [];
  //   const expandedItems: any[] = [];

  //   for (const data of materialRequestData) {
  //     console.log('data: ', data);
  //     // const Project = (await this._service.getItemById(url, "ProjectList", data.ProjectId)).Project;
  //     // const Client = (await this._service.getItemById(url, "ClientList", data.ClientId)).Client;
  //     // const Program = (await this._service.getItemById(url, "ProgramList", data.ProgramId)).Program;

  //     const Project = data.Project.Project
  //     const Client = data.Client.Client
  //     const Program = data.Program.Program
  //     const materialItemList = materialItemListVal.filter((d:any)=> d.MasterMaterialRequestID.ID===data.Id);

  //     // const materialItemList = await this._service.getItemSelectExpandFilter(
  //     //   url,
  //     //   "MaterialItemsList",
  //     //   "MasterMaterialRequestID/ID,MasterMaterialRequestID/Title,MaterialsID/ID,MaterialsID/Materials,Quantity",
  //     //   "MasterMaterialRequestID, MaterialsID",
  //     //   `MasterMaterialRequestID/ID eq ${data.Id}`
  //     // );
  //     console.log('materialItemList: ', materialItemList);

  //     //const materialsItemData: any[] = [];

  //     for (const materialItem of materialItemList) {
  //     //   const material = await this._service.getItemById(url, "MaterialsMasterList", materialItem.MaterialsID.ID);
  //     //   materialsItemData.push({
  //     //     MaterialTitle: material.Materials,
  //     //     Quantity: materialItem.Quantity
  //     //   });

  //       expandedItems.push({
  //         MaterialRequestCode: data.MaterialRequestCode,
  //         RequestedDate: moment(data.Created).format("DD-MM-YYYY"),
  //         Project: Project,
  //         Client: Client,
  //         Program: Program,
  //         material: materialItem.MaterialsID.Materials,
  //         quantities: materialItem.Quantity,
  //         status: data.Status,
  //         HOSComment: data.HOSApprovalComments,
  //         AdminComment: data.AdminComments
  //       });
  //     }
  //   }


  //       const groups = expandedItems.reduce((acc, cur) => {
  //         const { MaterialRequestCode } = cur
  //         const group = {
  //           key: MaterialRequestCode,
  //           name: `${MaterialRequestCode}`,
  //           startIndex: 0,
  //           count: 1,
  //         }
  //         if (acc.length === 0) {
  //           acc.push(group)
  //           return acc
  //         } else if (acc[acc.length - 1].key !== cur.MaterialRequestCode) {
  //           const { count, startIndex } = acc[acc.length - 1]
  //           acc.push({
  //             ...group,
  //             startIndex: count + startIndex,
  //           })
  //           return acc
  //         }
  //         acc[acc.length - 1].count++
  //         return acc
  //       }, []);
  //       // const groups = expandedItems.reduce((acc, cur) => {
  //       //   const { status } = cur
  //       //   const group = {
  //       //     key: status,
  //       //     name: `${status}`,
  //       //     startIndex: 0,
  //       //     count: 1,
  //       //   }
  //       //   if (acc.length === 0) {
  //       //     acc.push(group)
  //       //     return acc
  //       //   } else if (acc[acc.length - 1].key !== cur.status) {
  //       //     const { count, startIndex } = acc[acc.length - 1]
  //       //     acc.push({
  //       //       ...group,
  //       //       startIndex: count + startIndex,
  //       //     })
  //       //     return acc
  //       //   }
  //       //   acc[acc.length - 1].count++
  //       //   return acc
  //       // }, []);
  //       this.setState({
  //         expandedItems: expandedItems,
  //         groups: groups,

  //       })

  // }

  public async getDocumentIndexItems() {

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const currentUserId = this.state.getcurrentuserId;
    const isAdmin = currentUserId === this.state.adminApproverName;
    console.log('isAdmin: ', isAdmin);
    const isHOS = this.state.Departmentslist.some((dept: any) => dept.HOSNameId === currentUserId);
    console.log('isHOS: ', isHOS);

    let materialRequestData = await this._service.getItemSelectExpand(url, "MasterMaterialRequestList",
      "*,Project/ID,Project/Project,Client/ID,Client/Client,Program/ID,Program/Program", "Project,Client,Program");
    console.log(materialRequestData);

    const materialItemListVal = await this._service.getItemSelectExpand(
      url,
      "MaterialItemsList",
      "MasterMaterialRequestID/ID,MasterMaterialRequestID/Title,MaterialsID/ID,MaterialsID/Materials,Quantity",
      "MasterMaterialRequestID, MaterialsID"
    );

    if (isAdmin) {
      // Display all items for admin user
    } else if (isHOS) {
      // Display items submitted by the current user and those users' department HOS is current user
      materialRequestData = materialRequestData.filter((data: any) => data.AuthorId === currentUserId ||
        materialRequestData.filter((item: any) => item.HOSApproverId === currentUserId));
    } else {
      // Display items submitted by the current user
      materialRequestData = materialRequestData.filter((data: any) => data.AuthorId === currentUserId);
    }

    const expandedItems: any[] = [];

    for (const data of materialRequestData) {
      const Project = data.Project.Project;
      const Client = data.Client.Client;
      const Program = data.Program.Program;
      const materialItemList = materialItemListVal.filter((d: any) => d.MasterMaterialRequestID.ID === data.Id);

      for (const materialItem of materialItemList) {
        expandedItems.push({
          MaterialRequestCode: data.MaterialRequestCode,
          RequestedDate: moment(data.Created).format("DD-MM-YYYY"),
          Project: Project,
          Client: Client,
          Program: Program,
          material: materialItem.MaterialsID.Materials,
          quantities: materialItem.Quantity,
          status: data.Status,
          HOSComment: data.HOSApprovalComments,
          AdminComment: data.AdminComments
        });
      }
    }

    const groupedByMaterialRequest = _.groupBy(expandedItems, 'MaterialRequestCode');

    const groups = await Promise.all(_.map(groupedByMaterialRequest, async (materialItems: any, materialRequestCode: any) => {
      let cumulativeCount = 0;
      const groupedByStatus = _.groupBy(materialItems, 'status');
      const statusGroups = _.map(groupedByStatus, (statusItems: any, status: any) => {
        const statusCount = statusItems.length;
        const startIndex = cumulativeCount;
        cumulativeCount += statusCount;
        return {
          key: `${materialRequestCode}_${status}`,
          name: status,
          startIndex: startIndex,
          count: statusCount,
          level: 1,
        };
      });

      return {
        key: materialRequestCode,
        name: materialRequestCode,
        startIndex: cumulativeCount - materialItems.length,
        count: materialItems.length,
        children: statusGroups,
      };
    }));

    this.setState({
      expandedItems: expandedItems,
      groups: groups,
    });
    if (this.state.expandedItems.length === 0) {
      this.setState({
        noItems: "false",
        statusMessageNoItems: 'No items to display'
      });
    } else {
      this.setState({ noItems: "true" });
    }
  }


  public async getAdminApprover() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const AdminApproverlistItem = await this._service.getListItems("AdminApprover", url)
    const ApproverIdUserInfo = await this._service.getUser(AdminApproverlistItem[0].AdminApproverId);
    console.log('ApproverIdUserInfo: ', ApproverIdUserInfo);
    // const AdminApprover = ApproverIdUserInfo.Title; 
    this.setState({ adminApproverName: ApproverIdUserInfo.Id });
  }

  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    this.setState({ getcurrentuserId: getcurrentuser.Id });
    console.log('getcurrentuser: ', getcurrentuser);
  }

  public async getDepartmentsList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const DepartmentslistItem = await this._service.getListItems("Departments", url)
    console.log('DepartmentslistItem: ', DepartmentslistItem);
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

  // public async getMasterMaterialRequestListData() {
  //   const url: string = this.props.context.pageContext.web.serverRelativeUrl;
  //   const currentUserId = this.state.getcurrentuserId;
  //   const isAdmin = currentUserId === this.state.adminApproverName;
  //   console.log('isAdmin: ', isAdmin);
  //   const isHOS = this.state.Departmentslist.some((dept: any) => dept.HOSNameId === currentUserId);
  //   console.log('isHOS: ', isHOS);



  //   //let materialRequestData = await this._service.getListItems("MasterMaterialRequestList", url);
  //   let materialRequestData = await this._service.getItemSelectExpand(url, "MasterMaterialRequestList",
  //     "*,Project/ID,Project/Project,Client/ID,Client/Client,Program/ID,Program/Program", "Project,Client,Program");
  //   console.log(materialRequestData);
  //   if (isAdmin) {
  //     // Display all items for admin user
  //   } else if (isHOS) {
  //     // Display items submitted by the current user and those users' department HOS is current user
  //     materialRequestData = materialRequestData.filter((data: any) => data.AuthorId === currentUserId ||
  //       materialRequestData.filter((item: any) => item.HOSApproverId === currentUserId));
  //   } else {
  //     // Display items submitted by the current user
  //     materialRequestData = materialRequestData.filter((data: any) => data.AuthorId === currentUserId);
  //   }

  //   const materialDatas: any[] = [];
  //   const expandedItems: any[] = [];

  //   for (const data of materialRequestData) {
  //     console.log('data: ', data);
  //     // const Project = (await this._service.getItemById(url, "ProjectList", data.ProjectId)).Project;
  //     // const Client = (await this._service.getItemById(url, "ClientList", data.ClientId)).Client;
  //     // const Program = (await this._service.getItemById(url, "ProgramList", data.ProgramId)).Program;

  //     const Project = data.Project.Project
  //     const Client = data.Client.Client
  //     const Program = data.Program.Program


  //     const materialItemList = await this._service.getItemSelectExpandFilter(
  //       url,
  //       "MaterialItemsList",
  //       "MasterMaterialRequestID/ID,MasterMaterialRequestID/Title,MaterialsID/ID,MaterialsID/Materials,Quantity",
  //       "MasterMaterialRequestID, MaterialsID",
  //       `MasterMaterialRequestID/ID eq ${data.Id}`
  //     );
  //     console.log('materialItemList: ', materialItemList);

  //     const materialsItemData: any[] = [];

  //     for (const materialItem of materialItemList) {
  //       const material = await this._service.getItemById(url, "MaterialsMasterList", materialItem.MaterialsID.ID);
  //       materialsItemData.push({
  //         MaterialTitle: material.Materials,
  //         Quantity: materialItem.Quantity
  //       });

  //       expandedItems.push({
  //         MaterialRequestCode: data.MaterialRequestCode,
  //         RequestedDate: moment(data.Created).format("DD-MM-YYYY"),
  //         Project: Project,
  //         Client: Client,
  //         Program: Program,
  //         material: material.Materials,
  //         quantities: materialItem.Quantity,
  //         status: data.Status,
  //         HOSComment: data.HOSApprovalComments,
  //         AdminComment: data.AdminComments
  //       });
  //     }

  //     materialDatas.push({
  //       MaterialRequestCode: data.MaterialRequestCode,
  //       RequestedDate: moment(data.Created).format("DD-MM-YYYY"),
  //       Project: Project,
  //       Client: Client,
  //       Program: Program,
  //       status: data.Status,
  //       materials: materialsItemData,
  //       HOSComment: data.HOSApprovalComments,
  //       AdminComment: data.AdminComments
  //     });
  //   }

  //   const groups: IGroup[] = materialDatas.map((data) => {
  //     const groupItems = expandedItems.filter(item => item.MaterialRequestCode === data.MaterialRequestCode);

  //     // Create an object to store status groups
  //     const statusGroupsObj: { [key: string]: IGroup } = {};

  //     // Populate status groups
  //     groupItems.forEach(item => {
  //       const statusKey = `${data.MaterialRequestCode}-${item.status}`;
  //       if (!statusGroupsObj[statusKey]) {
  //         statusGroupsObj[statusKey] = {
  //           key: statusKey,
  //           name: item.status,
  //           startIndex: expandedItems.indexOf(item),
  //           count: 0,
  //           level: 1,
  //         };
  //       }
  //       statusGroupsObj[statusKey].count++;
  //     });

  //     // Convert statusGroupsObj to an array
  //     const statusGroups: IGroup[] = [];
  //     for (const key in statusGroupsObj) {
  //       if (statusGroupsObj.hasOwnProperty(key)) {
  //         statusGroups.push(statusGroupsObj[key]);
  //       }
  //     }

  //     return {
  //       key: data.MaterialRequestCode,
  //       name: data.MaterialRequestCode,
  //       startIndex: expandedItems.indexOf(groupItems[0]),
  //       count: groupItems.length,
  //       children: statusGroups,  // Add status groups as children
  //     };
  //   }).filter(group => group.count > 0);  // Filter out groups with no items

  //   this.setState({
  //     materialDatas: materialDatas, groups: groups,
  //     expandedItems: expandedItems
  //   });

  //   if (this.state.expandedItems.length === 0) {
  //     this.setState({
  //       noItems: "false",
  //       statusMessageNoItems: 'No items to display'
  //     });
  //   } else {
  //     this.setState({ noItems: "true" });
  //   }
  // }


  public handlePageChange = (pageNumber: number) => {
    this.setState({
      currentPage: pageNumber
    });
  };


  public render(): React.ReactElement<IMaterialRequestViewTableProps> {


    const {
    } = this.props;

    const columns: IColumn[] = [
      // {
      //   key: 'column1',
      //   name: 'Code',
      //   fieldName: 'MaterialRequestCode',
      //   minWidth: 70,
      //   maxWidth: 90,
      //   isResizable: true,
      //   isCollapsible: true,
      //   data: 'number',
      //   // onColumnClick: this._onColumnClick,
      //   // onRender: (item: IDocument) => {
      //   // return <span>{item.fileSize}</span>;
      //   // },
      // },
      {
        key: 'column2',
        name: 'Request Date',
        fieldName: 'RequestedDate',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        sortAscendingAriaLabel: 'Sorted A to Z',
        sortDescendingAriaLabel: 'Sorted Z to A',
        // onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column3',
        name: 'Project',
        fieldName: 'Project',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        // onColumnClick: this._onColumnClick,
        data: 'string',
        // onRender: (item: IDocument) => {
        //   return <span>{item.dateModified}</span>;
        // },
        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Client',
        fieldName: 'Client',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        // onColumnClick: this._onColumnClick,
        data: 'string',
        // onRender: (item: IDocument) => {
        //   return <span>{item.dateModified}</span>;
        // },
        isPadded: true,
      },
      {
        key: 'column5',
        name: 'Program',
        fieldName: 'Program',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        // onColumnClick: this._onColumnClick,
        data: 'string',
        // onRender: (item: IDocument) => {
        //   return <span>{item.dateModified}</span>;
        // },
        isPadded: true,
      },
      {
        key: 'column6',
        name: 'Material',
        fieldName: 'material',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'string',
        // onColumnClick: this._onColumnClick,
        // onRender: (item: IDocument) => {
        //   return <span>{item.modifiedBy}</span>;
        // },
        isPadded: true,
      },
      {
        key: 'column7',
        name: 'Quantity',
        fieldName: 'quantities',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsible: true,
        data: 'number',
        // onColumnClick: this._onColumnClick,
        // onRender: (item: IDocument) => {
        // return <span>{item.fileSize}</span>;
        // },
      },
      // {
      //   key: 'column8',
      //   name: 'Status',
      //   fieldName: 'status',
      //   minWidth: 70,
      //   maxWidth: 90,
      //   isResizable: true,
      //   isCollapsible: true,
      //   data: 'string',
      //   // onColumnClick: this._onColumnClick,
      //   // onRender: (item: IDocument) => {
      //   // return <span>{item.fileSize}</span>;
      //   // },
      // },
    ];

    const indexOfLastItem = this.state.currentPage * this.state.itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - this.state.itemsPerPage;
    const currentGroups = this.state.groups.slice(indexOfFirstItem, indexOfLastItem);

    // Calculate the total number of pages
    const totalPages = Math.ceil(this.state.groups.length / this.state.itemsPerPage);



    return (
      <section>
        <div className={styles.borderBox}>
          <div className={styles.MaterialRequestHeading}>{"Material Request"}</div>

          <div>


            {this.state.noItems === "false" &&
              <MessageBar
                messageBarType={MessageBarType.error}
                isMultiline={false}
                dismissButtonAriaLabel="Close"
              >
                {this.state.statusMessageNoItems}
              </MessageBar>
            }
          </div>
          <div>
            {this.state.noItems === "true" &&

              <>
                <DetailsList
                  items={this.state.expandedItems}
                  columns={columns}
                  setKey='set'
                  layoutMode={DetailsListLayoutMode.justified}
                  isHeaderVisible={true}
                  selectionMode={SelectionMode.none}
                  // groups={this.state.groups}
                  groups={currentGroups}
                />

                <Pagination
                  totalPages={totalPages}
                  currentPage={this.state.currentPage}
                  onChange={this.handlePageChange}
                />
              </>

            }

          </div>
        </div>
      </section >
    );
  }

}

