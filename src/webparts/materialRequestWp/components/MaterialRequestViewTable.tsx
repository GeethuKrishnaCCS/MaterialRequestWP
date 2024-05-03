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

    this.getDocumentIndexItems = this.getDocumentIndexItems.bind(this);;
    this.getAdminApprover = this.getAdminApprover.bind(this);
    this.getCurrentUser = this.getCurrentUser.bind(this);
    this.getDepartmentsList = this.getDepartmentsList.bind(this);
    this.handlePageChange = this.handlePageChange.bind(this);
  }


  public async componentDidMount() {
    await this.getAdminApprover();
    await this.getCurrentUser();
    await this.getDepartmentsList();
    this.getDocumentIndexItems();
  }


  public async getDocumentIndexItems() {

    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const currentUserId = this.state.getcurrentuserId;
    const isAdmin = currentUserId === this.state.adminApproverName;
   
    const isHOS = this.state.Departmentslist.some((dept: any) => dept.HOSNameId === currentUserId);    

    let materialRequestData = await this._service.getItemSelectExpand(url, this.props.MasterMaterialRequestList,
      "*,Project/ID,Project/Project,Client/ID,Client/Client,Program/ID,Program/Program", "Project,Client,Program");    

    const materialItemListVal = await this._service.getItemSelectExpand(
      url,
      this.props.MaterialItemsList,
      "MasterMaterialRequestID/ID,MasterMaterialRequestID/Title,MaterialsID/ID,MaterialsID/Materials,Quantity",
      "MasterMaterialRequestID, MaterialsID"
    );

    if (isAdmin) {
    } else if (isHOS) {
      materialRequestData = materialRequestData.filter((data: any) => data.AuthorId === currentUserId ||
        materialRequestData.filter((item: any) => item.HOSApproverId === currentUserId));
    } else {
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
    const AdminApproverlistItem = await this._service.getListItems(this.props.AdminApprover, url)
    const ApproverIdUserInfo = await this._service.getUser(AdminApproverlistItem[0].AdminApproverId);
   
    this.setState({ adminApproverName: ApproverIdUserInfo.Id });
  }

  public async getCurrentUser() {
    const getcurrentuser = await this._service.getCurrentUser();
    this.setState({ getcurrentuserId: getcurrentuser.Id });
  }

  public async getDepartmentsList() {
    const url: string = this.props.context.pageContext.web.serverRelativeUrl;
    const DepartmentslistItem = await this._service.getListItems(this.props.Departments, url)
    this.setState({ Departmentslist: DepartmentslistItem });

    DepartmentslistItem.map((Item: any) => {
      const departmentName = Item.Title;
      const HOSName = Item.HOSNameId;

      this.setState({
        departmentName: departmentName,
        HOSName: HOSName,
      });
    })
  }


  public handlePageChange = (pageNumber: number) => {
    this.setState({
      currentPage: pageNumber
    });
  };


  public render(): React.ReactElement<IMaterialRequestViewTableProps> {


    const {
    } = this.props;

    const columns: IColumn[] = [
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
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column4',
        name: 'Client',
        fieldName: 'Client',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
        isPadded: true,
      },
      {
        key: 'column5',
        name: 'Program',
        fieldName: 'Program',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        data: 'string',
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
      },
     
    ];

    const indexOfLastItem = this.state.currentPage * this.state.itemsPerPage;
    const indexOfFirstItem = indexOfLastItem - this.state.itemsPerPage;
    const currentGroups = this.state.groups.slice(indexOfFirstItem, indexOfLastItem);

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

