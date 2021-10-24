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
  /** WebPartに描画するHTML */
  html: string;
}

/**
 * HTML描画webpart
 */
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

    // HTMLを読み込んで画面に反映する
    this.loadHtml(this.domElement, html);
  }

  /**
   * HTMLの読み込み
   * @param parentElement HTMLを差し込む親HTML
   * @param innerHtml 挿入するHTML
   */
  private loadHtml(parentElement: HTMLElement, innerHtml: string) {

    if (!innerHtml) {
      console.error("Require HTML");
    }

    // HTMLを親に差し込む
    parentElement.innerHTML = innerHtml || "";

    // innerHTMLへの挿入だとscriptタグの内容が動かないため
    // scriptをタグとして再度挿入することで動かす
    const elScripts = parentElement.querySelectorAll("script");
    for (let i = 0; i < elScripts.length; i++) {
      const elOldScript = elScripts[i];
      const elRenewScript = document.createElement("script");
      if (elOldScript.src) {
        elRenewScript.src = elOldScript.src;
      }
      else {
        elRenewScript.text = elOldScript.text;
      }

      // scriptタグの差し替え
      parentElement.removeChild(elOldScript);
      parentElement.appendChild(elRenewScript);
    }
  }

  /**
   * Dispose
   */
  protected onDispose(): void {
  }

  /**
   * Component version
   */
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  /**
   * Property configuration
   * @returns Property configuration
   */
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
