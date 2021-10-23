import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'SimpleHtmlViewerWebPartStrings';
import styles from './SimpleHtmlViewerWebPart.module.scss';

/**
 * Component props
 */
export interface ISimpleHtmlViewerWebPartProps {
  html: string;
}

export default class SimpleHtmlViewerWebPart extends BaseClientSideWebPart<ISimpleHtmlViewerWebPartProps> {

  /**
   * Disable reactive property change.
   * @returns true: disable
   */
  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

  public render(): void {
    const html = this.properties.html ? this.properties.html : `<div class="${styles.simpleHtmlViewer}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <i class="ms-Icon ms-Icon--IncidentTriangle" aria-hidden="true"></i>  
              <p>Source is not setted.</p>
            </div>
          </div>
        </div>
      </div>`;


    // contextを渡す
    (window as any).___spContext___ = this.context;

    this.loadHtml(this.domElement, html);
  }

  private loadHtml(elElement: HTMLElement, html: string) {

    if (!html) {
      console.error("Require HTML");
    }

    elElement.innerHTML = html || "";

    const scripts = elElement.querySelectorAll("script");
    /*
    for (let i = 0; i < scripts.length; i++) {

      

    }
    */
  }

  protected onDispose(): void {
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('html', {
                  label: strings.HTMLLabel,
                  multiline: true,
                  placeholder: 'Input markup',
                  onGetErrorMessage: this.validateHTMLField.bind(this)
                })
              ]
            }
          ]
        }
      ]
    };
  }


  /**
   * Validate HTML field
   * @param value 
   * @returns 
   */
  private validateHTMLField(value: string): string {
    if (!value || !value.trim()) {
      return 'HTML field must be inputed';
    }
    return '';
  }
}
