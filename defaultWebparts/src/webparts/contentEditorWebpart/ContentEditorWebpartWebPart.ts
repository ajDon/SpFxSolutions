import { Version } from '@microsoft/sp-core-library';
import { DigestCache, IDigestCache, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { SPComponentLoader } from '@microsoft/sp-loader';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneLink, PropertyPaneTextField, PropertyPaneToggle } from '@microsoft/sp-webpart-base';
import * as strings from 'ContentEditorWebpartWebPartStrings';
import styles from './ContentEditorWebpartWebPart.module.scss';


export interface IContentEditorWebpartWebPartProps {
  spPageContextInfo: boolean;
  htmlUrl: URL;
  addHtmlDirectly: boolean;
  addHtmlScript: string;
  enableRequestDigest: boolean;
}

export default class ContentEditorWebpartWebPart extends BaseClientSideWebPart<IContentEditorWebpartWebPartProps> {

  public render(): void {
    this.addSpContextInfo();
    this.addRequestDigest();
    if (this.properties.addHtmlDirectly) {
      if (this.properties.addHtmlScript.trim() != '') {
        let responseHTML = this.convertStringToHTML(this.properties.addHtmlScript);
        this.addAllScripts(responseHTML).then((response) => {
          this.domElement.innerHTML = responseHTML.outerHTML;
        });
      }
      else
        this.defaultHTML();
    }
    else if (this.properties.htmlUrl != undefined || this.properties.htmlUrl != null) {
      let htmlLink: string = this.properties.htmlUrl.toString();
      if (htmlLink != "") {
        this.loadHTML(htmlLink).then((response) => {
          if (response.trim() != '') {
            let responseHTML = this.convertStringToHTML(response);
            this.addAllScripts(responseHTML).then(() => {
              this.domElement.innerHTML = responseHTML.outerHTML;
            });

          }
          else
            this.noHTMLFound();
        })
          .catch((error) => {
            this.domElement.innerHTML = `
            <div class="${ styles.contentEditorWebpart}">
              <div class="${ styles.container}">
                <div class="${ styles.row}">
                  <div class="${ styles.column}">
                    <span class="${ styles.title}">Error Occurred while loading HTML File</span>
                    <p class="${ styles.subTitle}">${error}</p>
                  </div>
                </div>
              </div>
            </div>
          `;
          });
      }
      else
        this.defaultHTML();
    }
    else {
      this.defaultHTML();
    }
  }

  /**
   * Determine Type of the code
   * @param elem HTMLElement
   * @param name Nodename
   */
  private nodeName(elem: HTMLElement, name: string) {
    return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
  }

  /**
   * Add Script from HTML Element
   * @param element HTMLElemt
   */
  protected async addAllScripts(element: HTMLElement) {
    (<any>window).ScriptGlobal = {};
    const scripts = [];
    const children_nodes = element.childNodes;
    for (let i = 0; children_nodes[i]; i++) {
      const child: any = children_nodes[i];
      if (this.nodeName(child, "script") &&
        (!child.type || child.type.toLowerCase() === "text/javascript")) {
        scripts.push(child);
      }
    }
    const urls = [];
    const onLoads = [];
    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.src && scriptTag.src.length > 0) {
        urls.push(scriptTag.src);
      }
      if (scriptTag.onload && scriptTag.onload.length > 0) {
        onLoads.push(scriptTag.onload);
      }
    }
    let oldamd = null;
    if (window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for (let i = 0; i < urls.length; i++) {
      try {
        await SPComponentLoader.loadScript(urls[i], { globalExportsName: "ScriptGlobal" });
      } catch (error) {
        console.error(error);
      }
    }
    if (oldamd) {
      window["define"].amd = oldamd;
    }
    for (let i = 0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if (scriptTag.parentNode) { scriptTag.parentNode.removeChild(scriptTag); }
      this.evalScript(scripts[i]);
    }
    // execute any onload people have added
    for (let i = 0; onLoads[i]; i++) {
      onLoads[i]();
    }
  }

  /**
   * Evaluate Scripts
   * @param elem Script without Source
   */
  private evalScript(elem) {
    const data = (elem.text || elem.textContent || elem.innerHTML || "");
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag = document.createElement("script");

    scriptTag.type = "text/javascript";
    if (elem.src && elem.src.length > 0) {
      return;
    }
    if (elem.onload && elem.onload.length > 0) {
      scriptTag.onload = elem.onload;
    }

    try {
      // doesn't work on ie...
      scriptTag.appendChild(document.createTextNode(data));
    } catch (e) {
      // IE has funky script nodes
      scriptTag.text = data;
    }
    setTimeout(() => {
      headTag.insertBefore(scriptTag, headTag.lastChild);
      headTag.removeChild(scriptTag);
    }, 1000);
  }

  /**
   * Convert HTML from String
   * @param htmlString HTML string
   * @returns Div Element
   */
  protected convertStringToHTML(htmlString: string): HTMLDivElement {
    const div = document.createElement('div');
    div.innerHTML = htmlString;
    return div;
  }
  
  /**
   * Default HTML to the Page
   */
  protected defaultHTML(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.contentEditorWebpart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Content and Script Editor</span>
              <p class="${ styles.subTitle}">Add HTML directly or add HTML link to webpart.</p>
            </div>
          </div>
        </div>
      </div>
      `;
  }

  /**
   * No HTML Found
   */
  protected noHTMLFound(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.contentEditorWebpart}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Content and Script Editor</span>
              <p class="${ styles.subTitle}">No HTML found!!!</p>
            </div>
          </div>
        </div>
      </div>
      `;
  }

  /**
   * Load HTML from Link
   * @param htmlLink Link
   * @returns Response based on HTML Link
   */
  protected loadHTML(htmlLink: string): Promise<string> {
    return this.context.spHttpClient.get(htmlLink, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.text();
    });
  }

  /**
   * Add _spPageContextInfo to the Page
   */
  protected async addSpContextInfo() {
    if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }
  }

  /**
   * Add Request Digest
   */
  protected async addRequestDigest() {
    if (this.properties.enableRequestDigest) {
      this.getDigest();
    }
  }

  /**
   * Get Request Digest to the page
   * @returns Loads Digest and adds it to Hidden Field
   */
  protected getDigest(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const digestCache: IDigestCache = this.context.serviceScope.consume(DigestCache.serviceKey);
      digestCache.fetchDigest(this.context.pageContext.web.serverRelativeUrl).then((digest: string): void => {
        if (document.getElementById("__REQUESTDIGEST") === null) {
          const requestDigestElement = document.createElement("input");
          const head = document.getElementsByTagName("head")[0];
          requestDigestElement.type = "hidden";
          requestDigestElement.value = digest;
          requestDigestElement.id = "__REQUESTDIGEST";
          head.insertBefore(requestDigestElement, head.lastChild);
        }
        else {
          const requestDigestElement = document.getElementById("__REQUESTDIGEST");
          requestDigestElement.nodeValue = digest;
        }
        resolve();
      });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'addHtmlDirectly') {
      if (newValue) {
        let newHtmLink: URL;
        this.properties.htmlUrl = newHtmLink;
      }
      else {
        this.properties.addHtmlScript = "";
      }
    }
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
                // PropertyPaneTextField('description', {
                //   label: strings.DescriptionFieldLabel
                // }),
                PropertyPaneTextField("htmlUrl", {
                  label: strings.HtmlUrlFieldLabel,
                  disabled: this.properties.addHtmlDirectly,
                  value: this.properties.htmlUrl == undefined ? "" : this.properties.htmlUrl.toString()
                }),
                PropertyPaneLink("htmlUrl", {
                  href: this.properties.htmlUrl == undefined ? "#" : this.properties.htmlUrl.toString(),
                  text: strings.HTMLUrlLinkLabel,
                  target: "_blank",
                  disabled: this.properties.htmlUrl == undefined ? true : this.properties.htmlUrl.toString() == ""
                }),
                PropertyPaneToggle("addHtmlDirectly", {
                  label: strings.AddHtmlDirectlyFieldLabel,
                  checked: this.properties.addHtmlDirectly,
                  onText: strings.EnabledText,
                  offText: strings.DisabledText
                }),
                PropertyPaneTextField("addHtmlScript", {
                  label: strings.AddHtmlScriptFieldLabel,
                  multiline: true,
                  disabled: !this.properties.addHtmlDirectly,
                  value: this.properties.addHtmlScript
                }),
                PropertyPaneToggle('spPageContextInfo', {
                  label: strings.SpPageContextInfoFieldLabel,
                  checked: this.properties.spPageContextInfo,
                  onText: strings.EnabledText,
                  offText: strings.DisabledText
                }),
                PropertyPaneToggle('enableRequestDigest', {
                  label: strings.EnableRequestDigestFieldLabel,
                  checked: this.properties.enableRequestDigest,
                  onText: strings.EnabledText,
                  offText: strings.DisabledText
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
