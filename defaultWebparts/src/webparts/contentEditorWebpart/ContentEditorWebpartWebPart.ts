import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneLink
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPComponentLoader } from '@microsoft/sp-loader';
import {
  SPHttpClient,
  SPHttpClientResponse,
  IDigestCache,
  DigestCache
} from '@microsoft/sp-http';

import styles from './ContentEditorWebpartWebPart.module.scss';
import * as strings from 'ContentEditorWebpartWebPartStrings';
import { SPHttpClientConfiguration } from '@microsoft/sp-http';

export interface IContentEditorWebpartWebPartProps {
  // description: string;
  spPageContextInfo: boolean;
  htmlUrl: URL;
  addHtmlDirectly: boolean;
  addHtmlScript: string;
  enableSODFunctions: boolean;
  enableRequestDigest: boolean;
}

export default class ContentEditorWebpartWebPart extends BaseClientSideWebPart<IContentEditorWebpartWebPartProps> {

  public render(): void {
    this.addSpContextInfo();
    this.addSODFunctions();
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
    else if (this.properties.htmlUrl != undefined) {
      let htmlLink: string = this.properties.htmlUrl.toString();
      if (htmlLink != "") {
        this.loadHTML(htmlLink).then((response) => {
          if (response.trim() != '') {
            let responseHTML = this.convertStringToHTML(response);
            this.addAllScripts(responseHTML).then((response) => {
              this.domElement.innerHTML = responseHTML.outerHTML;
            });

          }
          else
            this.defaultHTML();
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
    }
    else {
      this.defaultHTML();
    }
  }
  private nodeName(elem, name) {
    return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
  }

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

  protected convertStringToHTML(htmlString: string): HTMLDivElement {
    const div = document.createElement('div');
    div.innerHTML = htmlString;
    return div;
  }

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
  protected loadHTML(htmlLink: string): Promise<string> {
    return this.context.spHttpClient.get(htmlLink, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      return response.text();
    });
  }
  protected async addSpContextInfo() {
    if (this.properties.spPageContextInfo && !window["_spPageContextInfo"]) {
      window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
    }
  }

  protected async addSODFunctions() {
    if (this.properties.enableSODFunctions && !window["SP"]) {
      const head = document.getElementsByTagName("head")[0];
      const allScriptReferences: string[] = [
        this.context.pageContext.web.absoluteUrl + "/ScriptResource.axd?d=DAcecIMKyRVIe2katSv_eSqkfzwWu66cDhiDAZgTIPkwiDG0s7JyEY89zijyQrsgxv2WAm7hCRFej7EoXjMgpZY0NNc4kkPd4rfYU7kyBoGmBwrLZ3NUz4ig94J6fTJBOcv6Tarf8boKZ3nF8-wibRqESkQkuCs3N7yh2UuqR3aMu0hj55t48S0XnLXPTzFq0&t=72fc8ae3",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/1033/initstrings.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/init.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/1033/strings.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/clienttemplates.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/theming.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/ie55up.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206//online/scripts/sposuitenav.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/blank.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/1033/sp.res.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.runtime.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.init.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.ui.dialog.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/core.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.core.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/ms.rte.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.ui.rte.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/1033/sp.jsgrid.res.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.taxonomy.js",
        // this.context.pageContext.web.absoluteUrl + "/_layouts/15/ScriptResx.ashx?culture=en%2Dus&amp;name=ScriptResources&amp;rev=pMUQ%2Fe2tBQON96NLsfaMtA%3D%3D",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/scriptforwebtaggingui.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.ui.taxonomy.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.ui.reputation.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/sp.ui.listsearchboxbootstrap.js",
        "https://static.sharepointonline.com/bld/_layouts/15/16.0.7618.1206/calloutusagecontrolscript.js"
      ];
      for (let index = 0; index < allScriptReferences.length; index++) {
        const scriptSrc = allScriptReferences[index];
        setTimeout(() => {
          const scriptElement = document.createElement("script");
          scriptElement.type = "text/javascript";
          scriptElement.src = scriptSrc;
          head.insertBefore(scriptElement, head.lastChild);
        }, 50);
      }
      console.info("Display mode:" + this.displayMode);
      if (this.displayMode !== DisplayMode.Edit) {
        const addCoreV4CSS = document.createElement("link");
        addCoreV4CSS.href = this.context.pageContext.web.absoluteUrl + "/_layouts/15/1033/styles/Themable/corev15.css?rev=mX2UOCi99%2FD8gyljp67ezg%3D%3DTAG120";
        addCoreV4CSS.rel = "stylesheet";
        head.insertBefore(addCoreV4CSS, head.lastChild);
      }
      setTimeout(() => {
        let SP = <any>window["SP"];
        const addSODScripts = document.createElement("script");
        addSODScripts.type = "text/javascript";
        addSODScripts.innerText = `
        SP.SOD.registerSod("require.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002frequire.js");
        SP.SOD.registerSod("menu.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fmenu.js");
        SP.SOD.registerSod("mQuery.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fmquery.js");
        SP.SOD.registerSod("callout.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fcallout.js");
        SP.SOD.registerSodDep("callout.js", "mQuery.js");
        SP.SOD.registerSod("sharedhovercard.strings.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002f1033\u002fsharedhovercard.strings.js");
        SP.SOD.registerSod("sharedhovercard.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsharedhovercard.js");
        SP.SOD.registerSodDep("sharedhovercard.js", "sharedhovercard.strings.js");
        SP.SOD.registerSod("sharing.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsharing.js");
        SP.SOD.registerSodDep("sharing.js", "mQuery.js");
        SP.SOD.registerSod("suitelinks.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsuitelinks.js");
        SP.SOD.registerSod("clientrenderer.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fclientrenderer.js");
        SP.SOD.registerSod("srch.resources.resx", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002f1033\u002fsrch.resources.js");
        SP.SOD.registerSod("search.clientcontrols.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsearch.clientcontrols.js");
        SP.SOD.registerSodDep("search.clientcontrols.js", "clientrenderer.js");
        SP.SOD.registerSodDep("search.clientcontrols.js", "srch.resources.resx");
        SP.SOD.registerSod("sp.search.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.search.js");
        SP.SOD.registerSod("ajaxtoolkit.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fajaxtoolkit.js");
        SP.SOD.registerSodDep("ajaxtoolkit.js", "search.clientcontrols.js");
        SP.SOD.registerSod("userprofile", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.userprofiles.js");
        SP.SOD.registerSod("followingcommon.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002ffollowingcommon.js");
        SP.SOD.registerSodDep("followingcommon.js", "userprofile");
        SP.SOD.registerSodDep("followingcommon.js", "mQuery.js");
        SP.SOD.registerSod("profilebrowserscriptres.resx", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002f1033\u002fprofilebrowserscriptres.js");
        SP.SOD.registerSod("sp.ui.mysitecommon.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.ui.mysitecommon.js");
        SP.SOD.registerSodDep("sp.ui.mysitecommon.js", "userprofile");
        SP.SOD.registerSodDep("sp.ui.mysitecommon.js", "profilebrowserscriptres.resx");
        SP.SOD.registerSod("inplview", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002finplview.js");
        SP.SOD.registerSod("datepicker.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fdatepicker.js");
        SP.SOD.registerSod("jsgrid.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fjsgrid.js");
        SP.SOD.registerSod("sp.datetimeutil.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.datetimeutil.js");
        SP.SOD.registerSod("jsgrid.gantt.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fjsgrid.gantt.js");
        SP.SOD.registerSodDep("jsgrid.gantt.js", "jsgrid.js");
        SP.SOD.registerSodDep("jsgrid.gantt.js", "sp.datetimeutil.js");
        SP.SOD.registerSod("spgantt.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fspgantt.js");
        SP.SOD.registerSodDep("spgantt.js", "jsgrid.js");
        SP.SOD.registerSodDep("spgantt.js", "jsgrid.gantt.js");
        SP.SOD.registerSod("clientforms.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fclientforms.js");
        SP.SOD.registerSod("autofill.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fautofill.js");
        SP.SOD.registerSod("clientpeoplepicker.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fclientpeoplepicker.js");
        SP.SOD.registerSodDep("clientpeoplepicker.js", "autofill.js");
        SP.SOD.registerSod("sp.ui.combobox.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.ui.combobox.js");
        SP.SOD.registerSod("jsapiextensibilitymanager.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fjsapiextensibilitymanager.js");
        SP.SOD.registerSod("ganttsharedapi.generated.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fganttsharedapi.generated.js");
        SP.SOD.registerSod("ganttapishim.generated.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fganttapishim.generated.js");
        SP.SOD.registerSod("createsharedfolderdialog.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fcreatesharedfolderdialog.js");
        SP.SOD.registerSodDep("createsharedfolderdialog.js", "clientpeoplepicker.js");
        SP.SOD.registerSodDep("createsharedfolderdialog.js", "clientforms.js");
        SP.SOD.registerSod("reputation.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002freputation.js");
        SP.SOD.registerSod("sp.ui.listsearchbox.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.ui.listsearchbox.js");
        SP.SOD.registerSodDep("sp.ui.listsearchbox.js", "search.clientcontrols.js");
        SP.SOD.registerSodDep("sp.ui.listsearchbox.js", "profilebrowserscriptres.resx");
        SP.SOD.registerSod("sp.search.apps.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.search.apps.js");
        SP.SOD.registerSod("offline.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002foffline.js");
        SP.SOD.registerSod("dragdrop.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fdragdrop.js");
        SP.SOD.registerSod("online/scripts/suiteextensions.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fonline\u002fscripts\u002fsuiteextensions.js");
        SP.SOD.registerSod("filePreview.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002ffilepreview.js");
        SP.SOD.registerSod("movecopy.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fmovecopy.js");
        SP.SOD.registerSod("dlp.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fdlp.js");
        SP.SOD.registerSod("shell/shell15.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002f1033\u002fshell\u002fshell15.js");
        SP.SOD.registerSod("cui.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fcui.js");
        SP.SOD.registerSod("ribbon", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.ribbon.js");
        SP.SOD.registerSodDep("ribbon", "cui.js");
        SP.SOD.registerSodDep("ribbon", "inplview");
        SP.SOD.registerSod("WPAdderClass", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fwpadder.js");
        SP.SOD.registerSod("quicklaunch.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fquicklaunch.js");
        SP.SOD.registerSodDep("quicklaunch.js", "dragdrop.js");
        SP.SOD.registerSod("sp.ui.pub.ribbon.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.ui.pub.ribbon.js");
        SP.SOD.registerSod("sp.publishing.resources.resx", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002f1033\u002fsp.publishing.resources.js");
        SP.SOD.registerSod("sp.documentmanagement.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.documentmanagement.js");
        SP.SOD.registerSod("assetpickers.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fassetpickers.js");
        SP.SOD.registerSod("sp.ui.rte.publishing.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.ui.rte.publishing.js");
        SP.SOD.registerSod("spellcheckentirepage.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fspellcheckentirepage.js");
        SP.SOD.registerSod("sp.ui.spellcheck.js", "https:\u002f\u002fstatic.sharepointonline.com\u002fbld\u002f_layouts\u002f15\u002f16.0.7618.1206\u002fsp.ui.spellcheck.js");
        SP.SOD.registerSodDep("spgantt.js", "SP.UI.Rte.js");
        `;
        head.insertBefore(addSODScripts, head.lastChild);
      }, 2000);
    }
  }

  protected async addRequestDigest() {
    if (this.properties.enableRequestDigest) {
      this.getDigest();
    }
  }

  protected getDigest(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      const digestCache: IDigestCache = this.context.serviceScope.consume(DigestCache.serviceKey);
      digestCache.fetchDigest(this.context.pageContext.web.serverRelativeUrl).then((digest: string): void => {
        // use the digest here
        const requestDigestElement = document.createElement("input");
        const head = document.getElementsByTagName("head")[0];
        requestDigestElement.type = "hidden";
        requestDigestElement.value = digest;
        requestDigestElement.id = "__REQUESTDIGEST";
        head.insertBefore(requestDigestElement, head.lastChild);
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
                PropertyPaneToggle("enableSODFunctions", {
                  label: strings.EnableSODFunctionsFieldLabel,
                  checked: this.properties.enableSODFunctions,
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
