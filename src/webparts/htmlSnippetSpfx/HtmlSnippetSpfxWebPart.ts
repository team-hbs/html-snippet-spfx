import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import pnp, { List, Item, ListEnsureResult, ItemAddResult, FieldAddResult, Web } from "sp-pnp-js"
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'HtmlSnippetSpfxWebPartStrings';
import HtmlSnippetSpfx from './components/HtmlSnippetSpfx';
import { IHtmlSnippetSpfxProps } from './components/IHtmlSnippetSpfxProps';

export interface IHtmlSnippetSpfxWebPartProps {
  description: string;
  listName: string;
  item: string;
  rawHtml: string;
}


export default class HtmlSnippetSpfxWebPart extends BaseClientSideWebPart<IHtmlSnippetSpfxWebPartProps> {

  private itemsDropDown: PropertyPaneAsyncDropdown;

  private loadDocuments(): Promise<IDropdownOption[]> {
    return this.getDocuments();
  }

  private getDocuments() : Promise<IDropdownOption[]> { 
    return new Promise<IDropdownOption[]>((resolve: (items: IDropdownOption[]) => void, reject: (err: string) => void): void => {
           let web = new Web(this.context.pageContext.web.absoluteUrl);
           web.lists.getByTitle("Scriptsx").items.select(
            "ID","FileRef").get().then((items) => {
                let options:IDropdownOption[] = [];
                for (let _i = 0;_i < items.length; _i++)
                {
                  let path = items[_i]["FileRef"];
                  let parts = path.split("/");
                   let fileName = parts[parts.length - 1];
                  options.push({
                      key: path,
                      text: fileName
                    });
                }
                console.log(options);
                resolve(options);
            }).catch(e => { 
              console.error(e); 
              console.log("LIST DOES NOT EXIST");
            });;
     });
  }

  private onListChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    // reset selected item
    this.properties.item = undefined;
    // store new value in web part properties
    update(this.properties, 'item', (): any => { return this.properties.item; });
    // refresh web part
    this.render();
    // reset selected values in item dropdown
    this.itemsDropDown.properties.selectedKey = this.properties.item;
    // allow to load items
    this.itemsDropDown.properties.disabled = false;
    // load items and re-render items dropdown
    this.getDocuments().then((items) => {
      //TODO
          this.itemsDropDown.render();
    });
  }

  private onDocumentChange(propertyPath: string, newValue: any): void {
    this.render();
  }

  private loadItems(): Promise<IDropdownOption[]> {
    if (!this.properties.listName) {
      // resolve to empty options since no list has been selected
      return Promise.resolve();
    }
    const wp: HtmlSnippetSpfxWebPart = this;
    return this.getDocuments();
  }

  private onListItemChange(propertyPath: string, newValue: any): void {
    const oldValue: any = get(this.properties, propertyPath);
    // store new value in web part properties
    update(this.properties, propertyPath, (): any => { return newValue; });
    // refresh web part
    this.render();
  }

  public render(): void {
    let web = new Web(this.context.pageContext.web.absoluteUrl);
    web.getFileByServerRelativeUrl(this.properties.item).getText().then((text: string) => {
        this.properties.rawHtml = text;
        const element: React.ReactElement<IHtmlSnippetSpfxProps > = React.createElement(
          HtmlSnippetSpfx,
          {
            description: this.properties.description,
            listName: "Scripts",
            item: this.properties.item,
            context: this.context,
            rawHtml: this.properties.rawHtml
          }
        );
        this.domElement.innerHTML = this.properties.rawHtml;
        this.executeScript(this.domElement);
        ReactDom.render(element, this.domElement);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // reference to item dropdown needed later after selecting a list

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
                new PropertyPaneAsyncDropdown('item', {
                  label: "Files",
                  loadOptions: this.loadDocuments.bind(this),
                  onPropertyChange: this.onListItemChange.bind(this),
                  selectedKey: this.properties.item
                })
                
              ]
            }
          ]
        }
      ]
    };
  }


  private nodeName(elem, name) {
    return elem.nodeName && elem.nodeName.toUpperCase() === name.toUpperCase();
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

    headTag.insertBefore(scriptTag, headTag.firstChild);
    headTag.removeChild(scriptTag);
}



private async executeScript(element: HTMLElement) {
  if (this.context.pageContext && !window["_spPageContextInfo"]) {
    window["_spPageContextInfo"] = this.context.pageContext.legacyPageContext;
}

(<any>window).ScriptGlobal = {};

      // main section of function
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

}
