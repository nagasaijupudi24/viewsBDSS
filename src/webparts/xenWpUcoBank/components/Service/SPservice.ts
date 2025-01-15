import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/fields";
import "@pnp/sp/views";
import "@pnp/sp/files";
import "@pnp/sp/folders";


export default class SPService {
    private _sp;
    constructor(private context: WebPartContext) {
        this._sp = spfi().using(SPFx(this.context))
    }

    public getcolumnInfo = async (listName: any): Promise<{key:string,text:string}[]> => {
    
        const listObj=listName
        const listTitle=listObj?.title
        const temp: { key: string; text: string }[] = []
        await this._sp.web.lists.getByTitle(listTitle).fields.filter(" Hidden eq false and ReadOnlyField eq false")().then(field =>{
            field.filter((value: {
                InternalName: string; Title: string; TypeDisplayName?: string; Choices?: string[], TypeAsString?: string, SchemaXml?: string
            }) => {
                if (!( value.InternalName === "Attachments" || value.InternalName === "ContentType")) {
                    temp.push({
                                key: value.InternalName,
                                text: value.Title})
    
                }
    
            })
        });
        console.log(temp)
        return temp;
       
    }
}