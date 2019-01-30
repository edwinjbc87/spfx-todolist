import { BaseComponentContext } from '@microsoft/sp-component-base';
import { sp } from '@pnp/sp';

export default class SPService {
    protected webUrl: string;
    protected currentUserId: number;
    protected currentUserEmail: string;
    
    public constructor(ctx:BaseComponentContext) {
        this.webUrl = ctx.pageContext.web.absoluteUrl;
        this.currentUserEmail = ctx.pageContext.user.email;

        sp.setup({
            spfxContext: ctx,
            sp: {
                headers: {
                    'Accept': 'application/json;odata=nometadata'
                }
            }
        });
    }

    public async init():Promise<void>{
        this.currentUserId = await this.getUserId(this.currentUserEmail);
    }

    public async getUserId(email: string): Promise<number> {
        return await sp.site.rootWeb.ensureUser(email).then(result => result.data.Id);
    }

    public async getItems(list:any, filters?:any, expand?:any, fields?:any, top?:any): Promise<any> {
        let res = [];
        
        try {
            let filter = "Id gt 0";
            if (!fields) { fields = "ID,Id,Title"; }
            if (filters) { filter = filters; }
            let pnpReq = sp.web.lists.getByTitle(list).items.filter(filter).select(fields);
            if (expand) { pnpReq = pnpReq.expand(expand); }
            if (top) { pnpReq = pnpReq.top(top); } 
            else { pnpReq = pnpReq.top(5000); }

            res = await pnpReq.get();
        }
        catch (err) {
            res = [];
        }

        return res;
    }

    public async getUserInfo(userId:number): Promise<any> {
        let user = {
            Id: 0,
            Title: "",
            LoginName: ""
        };
      
        try {
            let u = await sp.web.siteUsers.getById(userId).get();
            user = {
                Id: u.Id,
                Title: u.Title,
                LoginName: u.LoginName
            };
        } catch (e) {
            user = {
                Id: 0,
                Title: "",
                LoginName: ""
            };
        }

        return user;
    }

    public async saveItem(lista, item): Promise<void> {
        let it = null;
        try {
          if (item.Id){
            let id = item.Id;
            delete item.Id;
            it = await sp.web.lists.getByTitle(lista).items.getById(id).update(item);
          } else {
            it = await sp.web.lists.getByTitle(lista).items.add(item);
          }
        } catch (err) {
          it = null;
        }
        return it;
    }

    public async saveItems(lista, items): Promise<boolean> {
        let res = false;
        let batch = sp.createBatch();

        try{
            for(let i=0; i<items.length; i++){
                let item = items[i];
    
                if(item.Id){
                    sp.web.lists.getByTitle(lista).items.getById(item.Id).inBatch(batch).update(item);
                } else {
                    sp.web.lists.getByTitle(lista).items.inBatch(batch).add(item);
                }
            }
    
            await batch.execute();
            res = true;
        } catch(ex){
            res = false;
        }
        return res;
    }

    public async deleteItem(lista, id): Promise<boolean> {
        let it = false;
        try {
            await sp.web.lists.getByTitle(lista).items.getById(id).delete();
            it = true;
        } catch (err) {
            it = false;
        }
        return it;
    }

    public async deleteItems(lista:string, ids:number[]): Promise<boolean>{
        let batch = sp.web.createBatch();
        let result = false;

        try {
            for(let i=0; i<ids.length; i++){
                sp.web.lists.getByTitle(lista).items.getById(ids[i]).inBatch(batch).delete();
            }                    
            await batch.execute();
            result = true;
        } catch (err) {
            result = false;
        }
        
		return result;
    }

    public async deleteItemsWhere (lista:string, query:string): Promise<boolean>{
        let batch = sp.web.createBatch();
        let result = false;
	
        try {
            let items = await this.getItems(lista, query);       
                
            for(let i=0; i<items.length; i++){
                sp.web.lists.getByTitle(lista).items.getById(items[i].Id).inBatch(batch).delete();
            }
                        
            await batch.execute();
            result = true;
        } catch (err) {
            result = false;
        }
        
		return result;
    }
      
}