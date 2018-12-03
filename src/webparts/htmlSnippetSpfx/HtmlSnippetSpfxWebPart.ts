import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { update, get } from '@microsoft/sp-lodash-subset';
import pnp, { List, Item, ListEnsureResult, ItemAddResult, FieldAddResult, Web } from "sp-pnp-js"

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
}

export default class HtmlSnippetSpfxWebPart extends BaseClientSideWebPart<IHtmlSnippetSpfxWebPartProps> {

  private itemsDropDown: PropertyPaneAsyncDropdown;

  private loadLists(): Promise<IDropdownOption[]> {
     return this.getLists();
  }

  private getLists() : Promise<IDropdownOption[]> { 
    console.log("getLists()");
    return new Promise<IDropdownOption[]>((resolve: (items: IDropdownOption[]) => void, reject: (err: string) => void): void => {
           let web = new Web(this.context.pageContext.web.absoluteUrl);
           web.lists.get().then((lists:List[]) => {
               let options:IDropdownOption[] = [];
               for (let _i = 0;_i < lists.length; _i++)
               {
                 //TODO: filter by document library
                 options.push({
                     key: lists[_i]["Title"],
                     text: lists[_i]["Title"]
                   });
               }
               console.log(options);
               resolve(options);
           });
     });
    }

  private getDocuments() : Promise<IDropdownOption[]> { 
    return new Promise<IDropdownOption[]>((resolve: (items: IDropdownOption[]) => void, reject: (err: string) => void): void => {
           let web = new Web(this.context.pageContext.web.absoluteUrl);
           web.lists.getByTitle(this.properties.listName).items.select(
            "ID","FileRef").get().then((items) => {
                let options:IDropdownOption[] = [];
                for (let _i = 0;_i < items.length; _i++)
                {
                  let path = items[_i]["FileRef"];
                  let parts = path.split("/");
                   let fileName = parts[parts.length - 1];
                  options.push({
                      key: items[_i]["ID"],
                      text: fileName
                    });
                }
                console.log(options);
                resolve(options);
            });
     });
    }

  private getDocumentsOld() : Promise<Item[]> { 
    return new Promise<Item[]>((resolve: (items: Item[]) => void, reject: (err: string) => void): void => {
           let web = new Web(this.context.pageContext.web.absoluteUrl);
           web.lists.getById(this.properties.listName).items.select(
             "ID",
            "Title"
            ).get().then((items) => {
                resolve(items);
            });
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
    const element: React.ReactElement<IHtmlSnippetSpfxProps > = React.createElement(
      HtmlSnippetSpfx,
      {
        description: this.properties.description,
        listName: this.properties.listName,
        item: this.properties.item
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }



  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // reference to item dropdown needed later after selecting a list
    this.itemsDropDown = new PropertyPaneAsyncDropdown('item', {
      label: strings.ItemFieldLabel,
      loadOptions: this.loadItems.bind(this),
      onPropertyChange: this.onListItemChange.bind(this),
      selectedKey: this.properties.item,
      // should be disabled if no list has been selected
      disabled: !this.properties.listName
    });
 
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
                new PropertyPaneAsyncDropdown('listName', {
                  label: strings.ListFieldLabel,
                  loadOptions: this.loadLists.bind(this),
                  onPropertyChange: this.onListChange.bind(this),
                  selectedKey: this.properties.listName
                }),
                this.itemsDropDown
              ]
            }
          ]
        }
      ]
    };
  }


  

}
