
import '@pnp/sp/webs';
import "@pnp/sp/sites";
import '@pnp/sp/site-groups/web';
import "@pnp/sp/profiles";

//import { sp } from '@pnp/pnpjs';
//import { SPFx } from '@pnp/sp/behaviors/spfx';
import { spfi, SPFx } from '@pnp/sp';

import pnp, {
  SearchQuery,
  SearchResults,
  Web,
  CamlQuery,
  //EmailProperties,
  setup,
  Item,
  Site
} from 'sp-pnp-js';


  import { WebPartContext } from '@microsoft/sp-webpart-base';
  import { Environment, EnvironmentType } from '@microsoft/sp-core-library';


  export const fieldsTrend =
    'Id,Title,Created,Author/Title,Author/Id';

export const proxyUrl = 'http://localhost:4323';
// Should be updated with environment web relative URL
export const webRelativeUrl = '/sites/GD_arquitectura_empresarial';



  export interface TypedHash<T> {
    [key: string]: T;
  }
  export interface EmailProperties {
    To: string[];
    CC?: string[];
    BCC?: string[];
    Subject: string;
    Body: string;
    From?: string;
  }

  export class PNP {
    public context: WebPartContext;
    public siteRelativeUrl: string;
    public repository: string = '';
    public web: Web;
    public site: Site;
    public sp:any;

    public fieldsSearch = {
      fields: [
        'Title',
        'DMSDocTitle',
        'ListItemID',
        'CreatedBy'
      ]
    };

    constructor(context: WebPartContext) {
      this.context = context;
      this.siteRelativeUrl = this.context.pageContext.web.serverRelativeUrl;

      this.sp = spfi().using(SPFx(this.context));

      if (Environment.type === EnvironmentType.Local) {
        this.web = new Web(`${proxyUrl}${webRelativeUrl}`);
        this.site = new Site(`${proxyUrl}${webRelativeUrl}`);

      } else {
        // On SharePoint page sp-pnp-js should be configured with
        setup({ spfxContext: this.context });
        // or a Web object should be created with explicit web URL
        this.web = new Web(this.context.pageContext.web.absoluteUrl);
        this.site = new Site(this.context.pageContext.web.absoluteUrl);
      }
    }


    public getChoiceFieldValues(
      listName: string,
      fieldName: string
    ): Promise<any> {
      let filters: string = `(InternalName eq '${fieldName}')`;
      return this.web.lists
        .getByTitle(listName)
        .fields.filter(filters)
        .get();
    }

    public senEmail(emailProps: EmailProperties): Promise<any> {
      return pnp.sp.utility.sendEmail(emailProps);
    }

    public getListItems(
      listName: string,
      fields: any,
      filters: string,
      expand: string,
      sortid?: any,
      topItem?: number
    ): Promise<any> {
      let top = topItem ? topItem : 9999;
      let sort = sortid ? sortid :{property : "ID", asc:true};
      return new Promise((resolve, reject) => {
        let list = this.web.lists.getByTitle(listName);
        if (list) {
          list.items
            .filter(filters)
            .select(fields)
            .expand(expand)
            .orderBy(sort.property, sort.asc)
            .top(top)
            .get()
            .then((items: any[]) => {
              resolve(items);
            })
            .catch(() => {
              reject(null);
            });
        }
      });
    }

    public getFeaturedItems(
      listName: string,
      fields: any,
      filters: string,
      expand: string,
      sortOne: any,
      sortTwo?: any,
      topItem?: number
    ): Promise<any> {
      let top = topItem ? topItem : 9999;
      return new Promise((resolve, reject) => {
        let list = this.web.lists.getByTitle(listName);
        if (list) {
          if (sortTwo) {
            list.items
              .filter(filters)
              .select(fields)
              .expand(expand)
              .orderBy(sortOne.property, sortOne.asc)
              .orderBy(sortTwo.property, sortTwo.asc)
              .top(top)
              .get()
              .then((items: any[]) => {
                resolve(items);
              })
              .catch(() => {
                reject(null);
              });
          } else {
            list.items
              .filter(filters)
              .select(fields)
              .expand(expand)
              .orderBy(sortOne.property, sortOne.asc)
              .top(top)
              .get()
              .then((items: any[]) => {
                resolve(items);
              })
              .catch(() => {
                reject(null);
              });
          }
        }
      });
    }

    public getAttchmentsFiles(listName: string, id: number): Promise<any> {
      return new Promise((resolve, reject) => {
        let item = this.web.lists.getByTitle(listName).items.getById(id);
        item.attachmentFiles.get().then((adjunct: any) => {
          resolve(adjunct);
        });
      });
    }

    public deleteAttachment(listname: string, itemId:number, nameItem: string): Promise<any> {
      return new Promise((resolve, reject) => {
        let item = this.web.lists.getByTitle(listname).items.getById(itemId);

        item.attachmentFiles
          .getByName(nameItem)
          .delete()
          .then((att: any) => {
            resolve(att);
          });
      });
    }

    public deleteAllAttachments(listname: string, itemId:number): Promise<any> {
      return new Promise((resolve, reject) => {
        let item = this.web.lists.getByTitle(listname).items.getById(itemId);
        item.attachmentFiles
        .get().then((items: any[])=>{
          items.forEach((item: { FileName: any; })=>{
            let itemDelete = this.web.lists.getByTitle(listname).items.getById(itemId);
            itemDelete.attachmentFiles
            .getByName(item.FileName)
            .delete()
            .then((att: any) => {
              resolve(att);
            });
          })
        })
      });
    }

    public getItemsByCAMLQuery(
      listName: string,
      query: string,
      expand: string
    ): Promise<any> {
      let q: CamlQuery = {
        ViewXml: query
      };
      let list = this.web.lists.getByTitle(listName);
      return list.getItemsByCAMLQuery(q, expand);
    }

    public getListItemsPagedFeaturedTwo(
      listname: any,
      fields: any,
      filters: string,
      expand: string,
      topItem: number = 9999
    ): Promise<any> {
      return this.web.lists
        .getByTitle(listname)
        .items.filter(filters)
        .select(fields)
        .expand(expand)
        .top(topItem)
        .getPaged();
    }

    public getListItemsPaged(
      listname: any,
      fields: any,
      filters: string,
      expand: string,
      sort: any,
      topItem: number = 9999
    ): Promise<any> {
      return this.web.lists
        .getByTitle(listname)
        .items.filter(filters)
        .select(fields)
        .expand(expand)
        .orderBy(sort.property, sort.asc)
        .top(topItem)
        .getPaged();
    }

    public insertItem(
      listName: string,
      properties: any,
      attachment?: any
    ): Promise<any> {
      return new Promise((resolve, reject) => {
        let list = this.web.lists.getByTitle(listName);
        list.items
          .add(properties)
          .then((res: { item: { attachmentFiles: { add: (arg0: any, arg1: any) => Promise<any>; }; }; data: any; }) => {
            if (attachment) {
              res.item.attachmentFiles
                .add(attachment.name, attachment)
                .then((_: any) => {
                  resolve(res.data);
                });
            }
            else {
              resolve(res.data);
            }
          })
          .catch((err: any) => {
            reject(err);
          });
      });
    }

    //Enviar siempre como array para usar la funcion
    public insertItemArrayFiles(
      listName: string,
      properties: any,
      attachment?: any,
      attachmentName?: string,
    ): Promise<any> {
      return new Promise((resolve, reject) => {
        let list = this.web.lists.getByTitle(listName);
        list.items
          .add(properties)
          .then((res: { item: any; data: any; }) => {
            if(attachment && attachment.length > 0){
              this.insertAttachments(res.item, 0, attachment[0], attachment, () => {
                  resolve(res.data);
                })
            }
            else{
              resolve(res.data)
            }
          })
      })
    }

    public insertAttachments(item: { attachmentFiles: { add: (arg0: string, arg1: File) => Promise<any>; }; }, pos: number, fileItem:any, attachmentArray:any, functionsuccess: { (item: any): void; (): any; (arg0: any): void; }){
      if(fileItem!=undefined){
        let file: File=fileItem.file;

        item.attachmentFiles
        .add(file.name, file)
        .then((att: any) => {
          if(pos< attachmentArray.length-1){
            this.insertAttachments(item, pos+1, attachmentArray[pos+1], attachmentArray, functionsuccess)
          }
          else if(pos == attachmentArray.length-1){
            functionsuccess(item)
          }
          //resolve(att);
        });
      }
      else{
        functionsuccess(item);
      }
    }

    public deleteItem(listName: string, id: number): Promise<any> {
      let list = this.web.lists.getByTitle(listName);
      return list.items.getById(id).delete();
    }

    async insertTrend(listName: string,
      props: any, attachment: File) {
      try {
        const lst = this.web.lists.getByTitle(listName);
        const data = await lst.items.add(props);
        const item = await this.finishSaveTrend(data.item, attachment);
        return item;
      } catch (error) {
        return error;
      }
    }


    public verifyFolderExistance(libraryName:string, folderName:string): Promise<any>{
      return new Promise(async (resolve, reject) => {

        try {
          const carpeta = await this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}/${folderName}`).get();
          if(carpeta.Exists){
            resolve(carpeta);
          }
          else{
            resolve(false);
          }
        } catch (error) {
          console.error(`La carpeta no existe: ${error}`);
          resolve(false)
        }

      })
    }

    public createNewFolder(libraryName:string, newFolderName:string): Promise<any> {
      return this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}`).folders.add(newFolderName)
    }

    public validateCreateSpFolder(libraryName:string, folderName:string): Promise<any>{
      return new Promise((resolve, reject) => {

        this.verifyFolderExistance(libraryName, folderName)
        .then(folderExists => {
          if(folderExists){
            resolve(folderExists)
          }
          else{
            try {
              this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}`).folders.add(folderName)
              .then((folderResponse: { data: any; }) => {
                this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}/${folderName}`).listItemAllFields.get()
                .then(async (response: { ID: any; }) => {
                  //await this.web.lists.getByTitle(libraryName).items.getById(response.ID).update({IDForma7CRId:lookUpID})
                  resolve(folderResponse.data);
                })
                .catch((onReject: any) => {
                  console.error(onReject);
                  resolve(false);
                })
              })
              .catch((onRejected: any) => {
                console.error(onRejected);
                resolve(false);
              })
            }
            catch (error) {
              console.error(`Error al crear la carpeta: ${error}`);
              resolve(false);
            }
          }
        })
        .catch(error => {
          resolve(false)
        })

      })
    }

    public getFiles(listName: string): Promise<any> {
      return this.web
        .getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${listName}/`)
        .files.get();
    }

    public getArrayFolderFiles(libraryName: string, folderName:string): Promise<any> {
      return new Promise((resolve, reject) => {
        this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}/${folderName}`).files.get()
        .then((arrayFolderFiles: any[]) => {
          let arrayFilesReturn: { FileName: any; ServerRelativeUrl: any; file: any; }[] = []
          if(arrayFolderFiles && arrayFolderFiles.length > 0){
            arrayFolderFiles.forEach((folderFile: { Name: any; ServerRelativeUrl: any; }) => {
              arrayFilesReturn.push({
                FileName:folderFile.Name,
                ServerRelativeUrl: folderFile.ServerRelativeUrl,
                file:folderFile
              })
            })
          }
          resolve(arrayFilesReturn)
        })
      })
    }

    public uploadFile(repository: string, file: File, name: string): Promise<any> {
      return this.web
        .getFolderByServerRelativeUrl(
          `${this.siteRelativeUrl}/${repository}/`
        )
        .files.add(name, file, true);
    }


    public uploadArrayFilesToFolder(library: string, folderName: string, attachmentsArray: File[]): Promise<any> {
      return new Promise((resolve, reject) => {
        const folderURL = `${this.siteRelativeUrl}/${library}/${folderName}`
        let folder = this.web.getFolderByServerRelativeUrl(folderURL);

        if(attachmentsArray && attachmentsArray.length > 0){
          this.uploadFilesArray(folder, 0, attachmentsArray[0], attachmentsArray, function(result: any){
            resolve(result)
          }, library)
        }
        else{
          resolve(false)
        }
      })
    }
    // Function used complementary with uploadArrayFiles, to upload files from an Array
    public uploadFilesArray(currentFolder:any, pos:number, fileItem:any, attachmentArray:File[], functionsuccess: { (result: any): void; (arg0: boolean): void; }, libraryName:string){
      if(fileItem!=undefined){
        let file: File=fileItem.file;

        currentFolder.files
        .add(file.name, file, true)
        .then((folderInfo: { data: { ServerRelativeUrl: any; }; }) => {
          this.web.getFolderByServerRelativeUrl(folderInfo.data.ServerRelativeUrl).listItemAllFields.get()
          .then(async (response: { ID: any; }) => {
            //await this.web.lists.getByTitle(libraryName).items.getById(response.ID).update({IDForma7CRId:lookUpID})
            if(pos < attachmentArray.length-1){
              this.uploadFilesArray(currentFolder, pos+1, attachmentArray[pos+1], attachmentArray, functionsuccess, libraryName)
            }
            else if(pos == attachmentArray.length-1){
              functionsuccess(folderInfo)
            }
          })
          .catch((onReject: any) => {
            console.error(onReject);
          })

        });
      }
      else{
        functionsuccess(false);
      }
    }


    public deleteFileByPath(url: string): Promise<any> {
      return this.web.getFileByServerRelativeUrl(url).delete();
    }
    public deleteFileByRelativeURL(relativeFileUrl: string): Promise<any> {
      return this.web.getFileByServerRelativeUrl(`${this.siteRelativeUrl}/${relativeFileUrl}`).delete();
    }

    public getFileByName(libraryRelativeUrl: string, fileNameWithExtension: string): Promise<any> {     
      return this.web
        .getFileByServerRelativeUrl(
          `${libraryRelativeUrl}/${fileNameWithExtension}`
        )
        .get();
    }

    public deleteFilesByPath(deletedocs: any): Promise<any> {
      return new Promise((resolve, reject) => {
        let docs = Object.keys(deletedocs);
        if (docs.length > 0) {
          docs.forEach(key => {
            let filename = encodeURIComponent(deletedocs[key]);
            this.web
              .getFileByServerRelativeUrl(
                `${this.siteRelativeUrl}/${this.repository}/${filename}`
              )
              .recycle();
          });
          resolve(true);
        }
      });
    }

    public getAdjunt(listname: string, id: number): Promise<any> {
      return new Promise((resolve, reject) => {
        let item = this.web.lists.getByTitle(listname).items.getById(id);
        item.attachmentFiles.get().then((v: any) => {
          resolve(v);
        });
      });
    }

    public deleteAdjunt(
      listname: string,
      id: number,
      name: string
    ): Promise<any> {
      return new Promise((resolve, reject) => {
        let item = this.web.lists.getByTitle(listname).items.getById(id);
        item.attachmentFiles
          .getByName(name)
          .delete()
          .then((v: any) => {
            resolve(v);
          });
      });
    }

    public updateTrendById(
      listname: string,
      id: number,
      properties: any,
      attachment?: File
    ): Promise<any> {
      let list = this.web.lists.getByTitle(listname);
      return list.items
        .getById(id)
        .update(properties)
        .then((res: { item: any; }) => {
          return this.finishSaveTrend(res.item, attachment).then(item => {
            return item;
          });
        });
    }

    async finishSaveTrend(item: Item, attachment?: File) {
      if (attachment) {
        await item.attachmentFiles.add(attachment.name, attachment);
        const itemUpdated = await item
          .select(fieldsTrend)
          .expand('Author')
          .get();
        return itemUpdated;
      } else {
        const itemUpdated = await item
          .select(fieldsTrend)
          .expand('Author')
          .get();
        return itemUpdated;
      }
    }

    public updateById(
      listname: string,
      id: number,
      properties: any,
      attachment?: any,
      attachmentName?: string
    ): Promise<any> {
      let list = this.web.lists.getByTitle(listname);
      return list.items
        .getById(id)
        .update(properties)
        .then((res: { item: any; }) => {
          if(attachment!=undefined) {
            this.insertAttachments(res.item, 0, attachment[0], attachment, function(){
              return res
            })
          }
          else
            return res
        });
    }

    async finishSave(item: Item, attachment: any, attachmentName:string) {
      if (attachment) {

        attachment.forEach((data: { file: File; }, i: number)=>{
          let file: File = data.file;
          item.attachmentFiles
          .add(file.name, file)
          .then((_: any) => {
            if(attachment.length-1 == i){
              return(item);
            }
          });
        })

        /* await item.attachmentFiles.add(attachmentName, attachment);
        const itemUpdated = await item
          .get();
        return itemUpdated; */
      } else {
        const itemUpdated = await item
          .get();
        return itemUpdated;
      }
    }

    public getCurrentUser(): Promise<any> {
      return this.web.currentUser.get();
    }

    public getUserInfoOffice(): Promise<any> {
      return pnp.graph.v1.get();
    }

    public getGroups(): Promise<any> {
      return this.web.siteGroups.get();
    }
    public getGroupsByName(nombreGrupo: any): Promise<any> {
      return this.web.siteGroups.getByName(nombreGrupo).users.get();
    }

    public getGroupsByUser(id: number): Promise<any> {
      return this.web.siteUsers.getById(id).groups.get();
    }

    public getByIdUser(id: number): Promise<any> {
      return this.web.siteUsers.getById(id).get();
    }


    public getByUser(Title: string): Promise<any> {
      return this.web.siteUsers.filter(`Title eq '${Title}'`).get();
    }

    public getUserInGroup(groupName: string, userID: number): Promise<any> {
      return this.web.siteGroups
        .getByName(groupName)
        .users.getById(userID)
        .get();
    }

    public getItemsByCamlquery(
      listName: string,
      camlquery: string,
      expand: string
    ): Promise<any> {
      const q: any = {
        ViewXml: camlquery
      };

      return this.web.lists.getByTitle(listName).getItemsByCAMLQuery(q);
    }

    public async searchInLibrary(
      siteUrl: string,
      refiners: any,
      // rows: number,
      listName: string,
      sort: any,
      text: string = ''
    ): Promise<any> {
      try {
        let filters: any = [];
        if (refiners) {
          filters = [`${refiners}`];
        }
        let query: string = `path:${siteUrl}/${listName}/ (ListItemID:${text}) OR ${text}*`;

        const searchQuery: SearchQuery = {
          EnableInterleaving: true,
          Querytext: query,
          RefinementFilters: filters,
          RowLimit: 20,
          SelectProperties: this.fieldsSearch.fields,
          SortList: sort,
          TrimDuplicates: false
        };

        const data: SearchResults = await pnp.sp.search(searchQuery);
        return data.PrimarySearchResults;
      } catch (error) {
        console.log(`Error consultando Señales: ${error}`);
        return null;
      }
    }


    public convertImageToBase64FromUrl(urlImg: string):string{
      let base64Image: string = '';

      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      const image = new Image();

      image.onload = () => {
        canvas.width = image.width;
        canvas.height = image.height;

        ctx && ctx.drawImage(image, 0, 0);

        base64Image = canvas.toDataURL('image/jpeg');
        return base64Image;
      };

      image.src = urlImg;

      if(base64Image)
        return base64Image

      return '';
    }



    public getImageFile(name: string) {
      if (name !== undefined && name !== "") {
        var val = name.split('.')
        var ext = val[val.length - 1].toLocaleLowerCase()

        var ico = 'exe.svg'

        if (ext == 'txt') {
          ico = "txt.svg";
        } else
          if (ext == 'xls' || ext == 'xlsx' || ext == "csv" || ext == "xlsm" || ext == "xlsb") {
            ico = "xlsx.svg";
          } else
            if (ext == 'doc' || ext == 'docx') {
              ico = "docx.svg";
            } else
              if (ext == 'pdf') {
                ico = "pdf.svg";
              } else
                if (ext == 'ppt' || ext == 'pptm' || ext == 'pptx') {
                  ico = "pptx.svg";
                } else
                  if (ext == 'png' || ext == 'jpg' || ext == 'gif' || ext == 'jpeg' || ext == 'svg') {
                    ico = "photo.svg";
                    // } else if (ext == 'zip' || ext == 'rar') {
                    //   ico = "photo.svg";
                  } else if (ext == 'js' || ext == "css") {
                    ico = "code.svg";
                  } else if (ext == "TTF") {
                    ico = "font.svg";
                  } else if (ext == "mp4" || ext == "mov" || ext == 'mpg') {
                    ico = "video.svg";
                  } else if (ext == "html") {
                    ico = "html.svg";
                  } else if (ext == 'one') {
                    ico = "one.svg"
                  } else if (ext == 'vsdx') {
                    ico = "vsdx.svg"
                  } else if (ext == 'aspx') {
                    ico = "spo.svg"
                  } else if (ext == 'msg') {
                    ico = "email.svg"
                  } else if (ext == 'fig') {
                    ico = "vector.svg"
                  } else if (ext == 'url') {
                    ico = "link.svg"
                  } else if (ext == 'zip' || ext == 'rar') {
                    ico = "zip.svg"
                  } else if (ext == 'bpm' || ext == 'bpmx' || ext == 'bpmn') {
                    ico = "bpm.svg"
                  } else {
                    ico = "genericfile.svg"
                  }

        return "https://res-1.cdn.office.net/files/fabric-cdn-prod_20220628.003/assets/item-types/20/" + ico
      }
    }

    public genericFile() {
      return "https://res-1.cdn.office.net/files/fabric-cdn-prod_20220628.003/assets/item-types/20/genericfile.svg"
    }

    public async getListId(listTitle:string): Promise<string> {
      const list = await this.web.lists.getByTitle(listTitle).select('Id').get();
      return list.Id;
    }

    public async listenListChanges(ListName:string):Promise<any>{
      let previousItems: string | any[] = [];
      const result = await this.web.lists.getById(ListName).getItemsByCAMLQuery({
        ViewXml: `<View><Query><Where><Geq><FieldRef Name='Modified' /><Value Type='DateTime' IncludeTimeValue='TRUE'>${previousItems.length > 0 ? previousItems[0].Modified.toISOString() : ''}</Value></Geq></Where></Query></View>`
      });
      const modifiedItems = result.items;
      previousItems = modifiedItems;
      return modifiedItems;
    }



    public async copyFolderFiles(libraryName:string, sourceFolderName: string, destinationFolderName: string, IdLookUp:number): Promise<void> {
      try {
        //const sourceFolder = await this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}/${sourceFolderName}`).get();
        const destinationFolder = await this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}/${destinationFolderName}`).get();

        const files = await this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}/${sourceFolderName}`).files.select('ServerRelativeUrl').get();
        for (const file of files) {
          const fileUrl = file.ServerRelativeUrl;
          const destinationFileUrl = `${this.siteRelativeUrl}/${libraryName}/${destinationFolderName}/${fileUrl.substring(fileUrl.lastIndexOf('/') + 1)}`;
          await this.web.getFileByServerRelativeUrl(fileUrl).copyTo(destinationFileUrl, true)
          .then((fileInfo: any) => {
            this.web.getFolderByServerRelativeUrl(destinationFolder.ServerRelativeUrl).listItemAllFields.get()
            .then(async (response: { ID: any; }) => {
              await this.web.lists.getByTitle(libraryName).items.getById(response.ID).update({IDForma7CRId:IdLookUp});
            })
          })
        }
      }
      catch (error) {
        console.error("Error:", error);
      }
    }

    /**
    * Función que elimina una carpeta dentro de una libreria y todos los documentos que ésta carpeta contiene
    *
    * @param {string} libraryName Nombre de la libreria que contiene la carpeta
    * @param {string} folderNameToDelete Nombre de la carpeta a eliminar
    * @return {boolean} Retorna un valor booleano, teniendo verdadero para cuando la eliminación fue satisfactoria, y false, para cuando ocurrio algun error
    */
    public deleteFolderFromLibrary = async (libraryName:string, folderNameToDelete:string) => {
      let deletionState = {successfullDeletion:false};
      try {
        const folder = this.web.getFolderByServerRelativeUrl(`${this.siteRelativeUrl}/${libraryName}/${folderNameToDelete}`);
        await folder.recycle();
        deletionState.successfullDeletion = true
        return deletionState;
      } catch (error) {
        console.error("Error al eliminar la carpeta:", error);
        deletionState.successfullDeletion = false;
        return deletionState;
      }
    }

    /**
    * Función que obtiene toda la información de un usuario que ha iniciado sesion, incluyendo la foto de perfil
    *
    * @param {string} loginName El loginName se obtiene de la informacion previa sobre el usuario es parecido a i:0#.f|membership|testuser@mytenant.onmicrosoft.com
    * @return {object} Retorna un objeto que contiene toda la informacion detallada del usuario que inicio sesion
    */
    public async detailedUserInfo(loginName:string) {
      try{
        const userProfileInfo = await this.sp.profiles.getPropertiesFor(loginName)
        return userProfileInfo;
      }
      catch (error) {
        console.error('Error obteniendo la imagen del perfil:', error);
        return null;
      }
    }


    /**
    * Función para validar existencia de una lista en el entorno de sharepoint, en caso de que no exista llama a la función createList para su creación, y en caso de existir, retorna el objeto con informacion de la lista
    *
    * @param {string} listNameValidate Nombre de la lista a crear para validar que exista en el entorno
    * @return {object} Retorna un objeto que contiene toda la información de la lista
    */
    public validateCreateList = async (listNameValidate:string, groupPermissions:any[]): Promise<object> => {
      try{
        const list = await this.web.lists.getByTitle(listNameValidate).get();

        /* Crear un apartado para validar que los permisos en la lista sean los establecidos, en caso de que no, establecerlos 
        const ensureListExistance = await this.web.lists.ensure(`${listNameValidate}`);
        const listPermissions = await ensureListExistance.list.;

        console.log(listPermissions);
        * usando => await this.addListPermissionsToGroup(listNameValidate, groupPermissions);
        */
        await this.addListPermissionsToGroup(listNameValidate, groupPermissions);
        
        return list;
      }
      catch (error) {
        await this.createList(listNameValidate);
        await this.addListPermissionsToGroup(listNameValidate, groupPermissions);
        const list = await this.web.lists.getByTitle(listNameValidate).get();
        return list;
      }
    }

    /**
    * Función para crear una lista
    *
    * @param {string} listNameCreate Nombre de la lista a crear
    * @return {object} Retorna un objeto que contiene toda la informacion de la lista
    */
    private createList = async (listNameCreate:string): Promise<object> => {
      const listCreate = await this.web.lists.add(listNameCreate);
      return listCreate;
    }

    public addListPermissionsToGroup = async(listName:string, groupPermissions:any[]): Promise<any> => {
      const ensureListExistance = await this.web.lists.ensure(`${listName}`);
      const list = ensureListExistance.list;

      await list.breakRoleInheritance();

      groupPermissions.forEach(async (group:{groupName:string, permission:string}) => {
        const groupPermission = await this.web.siteGroups.getByName(`${group.groupName}`).get();
        const roleToAsign = await this.web.roleDefinitions.getByName(`${group.permission}`).get();
        const listWithRoles = await list.roleAssignments.add(groupPermission.Id, roleToAsign.Id);

        return listWithRoles;
      });
    }
  }