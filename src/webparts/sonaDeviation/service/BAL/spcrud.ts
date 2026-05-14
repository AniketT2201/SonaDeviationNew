import "@pnp/sp/lists";
import "@pnp/sp/items";
// import { IPatelEngProps } from "../../components/IPatelEngProps";
import { ISonaDeviationProps } from "../../components/ISonaDeviationProps";
import SPCRUDOPS from "../../service/DAL/spcrudops";
 
export interface ISPCRUD {
    [x: string]: any;
    getData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: ISonaDeviationProps): Promise<any>;
    getRootData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean },top:number, props: ISonaDeviationProps): Promise<any>;
    insertData(listName: string, data: any, props: ISonaDeviationProps): Promise<any>;
    updateData(listName: string, itemId: number, data: any, props: ISonaDeviationProps): Promise<any>;
    deleteData(listName: string, itemId: number, props: ISonaDeviationProps): Promise<any>;
    getListInfo(listName: string, props: ISonaDeviationProps): Promise<any>;
    getListData(listName: string, columnsToRetrieve: string, props: ISonaDeviationProps): Promise<any>;
    batchInsert(listName: string, data: any, props: ISonaDeviationProps): Promise<any>;
    batchUpdate(listName: string, data: any, props: ISonaDeviationProps): Promise<any>;
    batchDelete(listName: string, data: any, props: ISonaDeviationProps): Promise<any>;
    createFolder(listName: string, folderName: string, props: ISonaDeviationProps):Promise<any>;
    uploadFile(folderServerRelativeUrl: string, file: File, props: ISonaDeviationProps): Promise<any>;
    deleteFile(fileServerRelativeUrl: string, props: ISonaDeviationProps): Promise<any>;
    currentProfile(props: ISonaDeviationProps): Promise<any>;
    //currentUserProfile(props: IDeviationuatProps): Promise<any>;
    getLoggedInSiteGroups(props: ISonaDeviationProps): Promise<any>;
    getAllSiteGroups(props: ISonaDeviationProps): Promise<any>;
    getTopData(listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string, orderby: { column: string, isAscending: boolean }, top: number, props: ISonaDeviationProps): Promise<any>;
    addAttchmentInList(attFiles: File, listName: string, itemId: number, fileName: string, props: ISonaDeviationProps): Promise<any>;
}

export default async function USESPCRUD(): Promise<ISPCRUD> {
    const spCrudOps = await SPCRUDOPS();
    return {
        getData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: ISonaDeviationProps) => {
            return await spCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, props);
        },
        getRootData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean },top:number, props: ISonaDeviationProps) => {
            return await spCrudOps.getData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, props);
        },
        insertData: async (listName: string, data: any, props: ISonaDeviationProps) => {
            return await spCrudOps.insertData(listName, data, props);
        },
        updateData: async (listName: string, itemId: number, data: any, props: ISonaDeviationProps) => {
            return await spCrudOps.updateData(listName, itemId, data, props);
        },
        deleteData: async (listName: string, itemId: number, props: ISonaDeviationProps) => {
            return await spCrudOps.deleteData(listName, itemId, props);
        },
        getListInfo: async (listName: string, props: ISonaDeviationProps) => {
            return await spCrudOps.getListInfo(listName, props);
        },
        getListData: async (listName: string, columnsToRetrieve: string, props: ISonaDeviationProps) => {
            return await spCrudOps.getListData(listName, columnsToRetrieve, props);
        },
        batchInsert: async (listName: string, data: any, props: ISonaDeviationProps) => {
            return await spCrudOps.batchInsert(listName, data, props);
        },
        batchUpdate: async (listName: string, data: any, props: ISonaDeviationProps) => {
            return await spCrudOps.batchUpdate(listName, data, props);
        },
        batchDelete: async (listName: string, data: any, props: ISonaDeviationProps) => {
            return await spCrudOps.batchDelete(listName, data, props);
        },
        createFolder: async (listName: string, folderName: string, props: ISonaDeviationProps) => {
            return await spCrudOps.createFolder(listName, folderName, props);
        },
        uploadFile: async (folderServerRelativeUrl: string, file: File, props: ISonaDeviationProps) => {
            return await spCrudOps.uploadFile(folderServerRelativeUrl, file, props);
        },
        deleteFile: async (fileServerRelativeUrl: string, props: ISonaDeviationProps) => {
            return await spCrudOps.deleteFile(fileServerRelativeUrl, props);
        },
        currentProfile: async (props: ISonaDeviationProps) => {
            return await spCrudOps.currentProfile(props);
        },
        // const currentUserProfile = async (props: IDeviationuatProps) => {
          
        //    // const queryUrl = "https://etgworld.sharepoint.com/sites/UAT_BPM/_api/web/currentuser/groups";
            
        //     const result: any = await (await spCrudOps).currentUserProfile( props);
        //     return result;
        // };
        getLoggedInSiteGroups: async (props: ISonaDeviationProps) => {
            return await spCrudOps.getLoggedInSiteGroups(props);
        },
        getAllSiteGroups: async (props: ISonaDeviationProps) => {
            return await spCrudOps.getAllSiteGroups(props);
        },
        getTopData: async (listName: string, columnsToRetrieve: string, columnsToExpand: string, filters: string
            , orderby: { column: string, isAscending: boolean }, top: number, props: ISonaDeviationProps) => {
            return await spCrudOps.getTopData(listName, columnsToRetrieve, columnsToExpand, filters, orderby, top, props);
        },
        addAttchmentInList: async (attFiles: File, listName: string, itemId: number, fileName: string, props: ISonaDeviationProps) => {
            return await spCrudOps.addAttchmentInList(attFiles, listName, itemId, fileName, props);
        }
    };
}