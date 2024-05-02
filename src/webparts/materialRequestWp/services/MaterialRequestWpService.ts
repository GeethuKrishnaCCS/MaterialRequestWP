import { BaseService } from "./BaseService";
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { getSP } from "../shared/Pnp/pnpjsConfig";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/attachments";

export class MaterialRequestWpService extends BaseService {
    private _spfi: SPFI;
    constructor(context: WebPartContext) {
        super(context);
        this._spfi = getSP(context);
    }
    
    public getListfilter(listname: string, id: number): Promise<any> {
        return this._spfi.web.lists.getByTitle(listname).items.filter("Id eq '" + id + "'").getAll();
    }
    public async getUser(userId: number): Promise<any> {
        return this._spfi.web.getUserById(userId)();
    }
    public async getCurrentUser(): Promise<any> {
        return this._spfi.web.currentUser();
    }



    public addMaterialRequestForm(data: any, listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.add(data);
    }    
    public getListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }
    public getClientListItems(listname: string, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items();
    }
    public getProgramListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Program,ID,Client/Title,Client/ID").expand("Client").filter("Client/ID eq '" + id + "'")();
    }
    public getProjectListItems(listname: string, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.select("Project,ID,Program/Title,Program/ID").expand("Program").filter("Program/ID eq '" + id + "'")();
    }
    public updateMaterialRequestForm(listname: string, data: any, id: number, url: string): Promise<any> {
        return this._spfi.web.getList(url + "/Lists/" + listname).items.getById(id).update(data);
    }

}