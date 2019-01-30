import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'ToDoListWebPartStrings';
import ToDoList from './components/ToDoList';
import { IToDoListProps } from './components/IToDoListProps';
import { IToDoListItem } from './components/IToDoListItem';
import SPService from '../../api/SPService';
import ICommonWebPartProps from '../ICommonWebPartProps';
import { elementContains } from 'office-ui-fabric-react';

export interface IToDoListWebPartProps{
  listTitle: string,
}

export default class ToDoListWebPart extends BaseClientSideWebPart<IToDoListWebPartProps> {
  private service:SPService;
  private webUrl:string;
  private element:ToDoList;
  private items:IToDoListItem[];

  protected async onInit(): Promise<void> {    
    this.service = new SPService(this.context);        
    await this.service.init();    

    this.webUrl = this.context.pageContext.web.absoluteUrl;
    this.items = await this.service.getItems(this.properties.listTitle);
    
    this._onCreateToDoItem = this._onCreateToDoItem.bind(this);
    this._onDeleteToDoItem = this._onDeleteToDoItem.bind(this);
  }

  public render(): void {
    let props:IToDoListProps = {
      toDoList: this.items,
      onAddItem: this._onCreateToDoItem,
      onDeleteItem: this._onDeleteToDoItem,
    };
    
    const elem = React.createElement(
      ToDoList,
      {
        ...props, ref: (elem) =>{ this.element = elem; }
      }
    );

    ReactDom.render(elem, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _onCreateToDoItem(toDo:string): Promise<void>{
    if(toDo.trim() != ''){
      const it = await this.service.saveItem(this.properties.listTitle, {Title: toDo});
      if(it != null){
        const item:IToDoListItem = {Id: it.Id, Title: it.Title};
        this.element.addItem(item);
      }
    }
  }

  private async _onDeleteToDoItem(id:number): Promise<void>{
    if(await this.service.deleteItem(this.properties.listTitle, id)){
      this.element.deleteItem(id);
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneTitle
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listTitle', {
                  label: strings.ListTitleFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
