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
      documentLink: "",
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

  private async getLibraryFile(formData:React.FormEvent<HTMLFormElement>): Promise<void> {
    formData.preventDefault();

    const selectedLibrary:string = formData.currentTarget['librarySelector'].value;
    const fileName:string = formData.currentTarget['fileNameInput'].value;
    try {
      const documentInfo = await this.pnp.getFileByName(
        selectedLibrary,
        fileName
      );
      this.setState({documentLink: documentInfo.LinkingUrl});
      console.clear();
      console.log("Informacion del archivo:", documentInfo);
      alert(`La información del archivo fue consultada con exito. Para ver el documento, seguir el Link y para ver mayor informacion del documento, dirigirse a la consola.`);
    } catch (error) {
      this.setState({documentLink: ""});
      console.error("Error al obtener datos de SharePoint:", error);
      alert("Error al obtener informacion del archivo o archivo no encontrado");
    }
  }


  public render(): React.ReactElement<ITestsProps> {
    return (
      <main>
        <form onSubmit={(form:React.FormEvent<HTMLFormElement>) => this.getLibraryFile(form)}>
          <div>
            <select name="librarySelector">
              <option value="" defaultChecked disabled >Seleccionar libreria ...</option>
              {
                this.state.siteLibraries.length > 0 && this.state.siteLibraries.map((library:DocumentLibraryInformation) => (
                  <option key={library.ServerRelativeUrl} value={library.ServerRelativeUrl}>
                    {library.Title}
                  </option>
                ))
              }
            </select>
            <input type="text" name='fileNameInput' placeholder='Ingresar Nombre del documento a consultar' />
          </div>
          <button type='submit'>
            Obtener informacion al archivo
          </button>
        </form>
        {
          this.state.documentLink && this.state.documentLink.length > 0 &&
          <>
            <br/>
            <section>
              <a href={this.state.documentLink} target='_blank'>Abrir documento</a>
            </section>
          </>
        }
      </main>
    );
  }
}
