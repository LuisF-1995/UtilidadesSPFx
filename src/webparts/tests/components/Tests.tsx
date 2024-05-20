import * as React from 'react';
//import styles from './Tests.module.scss';
import type { ITestsProps } from './ITestsProps';
import { PNP } from '../../services/utilities';
import { DocumentLibraryInformation } from 'sp-pnp-js';


export default class Tests extends React.Component<ITestsProps, any> {
  public pnp:PNP;

  constructor(props: ITestsProps) {
    super(props);
    this.pnp = new PNP(this.props.context);

    this.state = {
      siteLibraries: [],
      libraryFiles: [],
      documentInformation: null,
      unknownLibrary: true
    };

    this.getSiteLibraries();
  }

  private getSiteLibraries = async () => {
    try {
      const libraries:DocumentLibraryInformation[] = await this.pnp.site.getDocumentLibraries(this.props.context.pageContext.site.absoluteUrl);
      this.setState({
        siteLibraries: libraries
      });
    } catch (error) {
      console.error("Error al obtener datos del sitio:", error);
    }
  };

  private searchFile = (formData:React.FormEvent<HTMLFormElement>) => {
    formData.preventDefault();
    const fileName:string = formData.currentTarget['fileNameInput'].value;

    if(!this.state.siteLibraries || this.state.siteLibraries.length === 0){
      alert("No se han encontrado librerias en el sitio actual")
    }
    else{
      if(this.state.unknownLibrary && this.state.siteLibraries.length > 0){
        for (const library of this.state.siteLibraries) {
          if(library && library.ServerRelativeUrl && library.ServerRelativeUrl.length > 0)
            this.getLibraryFile(library.ServerRelativeUrl, fileName);
        }

        if(!this.state.documentInformation || this.state.documentInformation.length === 0){
          alert("No se encontró el archivo buscado");
        }
      }
      else{
        const selectedLibraryRelativeUrl:string = formData.currentTarget['librarySelector'].value;
        this.getLibraryFile(selectedLibraryRelativeUrl, fileName);
      }
    }
  }

  private async getLibraryFile(libraryRelativeUrl:string, fileName:string): Promise<void> {
    this.setState({documentInformation: null});
    try {
      const documentInfo = await this.pnp.getFileByName(
        libraryRelativeUrl,
        fileName
      );
      this.setState({documentInformation: documentInfo});
      console.clear();
      console.log("Informacion del archivo:", documentInfo);
      alert(`La información del archivo fue consultada con exito. Para ver el documento, seguir el Link y para ver mayor informacion del documento, dirigirse a la consola.`);
    } catch (error) {
      console.error("Error al obtener datos del documento:", error);
    }
  }


  public render(): React.ReactElement<ITestsProps> {
    return (
      <main>
        <form onSubmit={(form:React.FormEvent<HTMLFormElement>) => this.searchFile(form)}>
          <label htmlFor='unknownLibraryCheck'>Libreria desconocida</label>
          <input type='checkbox' id='unknownLibraryCheck' value={this.state.unknownLibrary} checked={this.state.unknownLibrary} onChange={() => this.setState({unknownLibrary: !this.state.unknownLibrary})} />

          <div>
            {
              !this.state.unknownLibrary && 
              <select name="librarySelector">
                <option value="" defaultChecked >Seleccionar libreria...</option>
                {
                  this.state.siteLibraries.length > 0 && this.state.siteLibraries.map((library:DocumentLibraryInformation) => (
                    <option key={library.ServerRelativeUrl} value={library.ServerRelativeUrl}>
                      {library.Title}
                    </option>
                  ))
                }
              </select>
            }
            <input type="text" name='fileNameInput' placeholder='Ingresar Nombre del documento a consultar' />
          </div>
          <button type='submit'>
            Obtener informacion al archivo
          </button>
        </form>
        {
          this.state.documentInformation && this.state.documentInformation.LinkingUrl.length > 0 &&
          <section>
            <br/>
            <h4>Se encontró la siguiente información del documento buscado: </h4>
            <p><b>Nombre del archivo: </b>{this.state.documentInformation.Title}</p>
            <p><b>Ubicación del archivo: </b>{this.state.documentInformation.ServerRelativeUrl}</p>
            <a href={this.state.documentInformation.LinkingUrl} target='_blank'>Abrir documento</a>
          </section>
        }
      </main>
    );
  }
}
