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

export interface IToDoListWebPartProps extends ICommonWebPartProps{}

export default class ToDoListWebPart extends BaseClientSideWebPart<IToDoListWebPartProps> {
  private service:SPService;
  private element:ToDoList;
  private items;

  protected async onInit(): Promise<void> {    
    this.service = new SPService(this.context);
        
    await this.service.init();    

    this.properties.webUrl = this.context.pageContext.web.absoluteUrl;
    this.properties.userEmail = this.context.pageContext.user.email;
    this.properties.userId = await this.service.getUserId(this.properties.userEmail);
    this.items = await this.service.getItems('LST_ToDoList');
    
    this._onCreateToDoItem = this._onCreateToDoItem.bind(this);
    this._onDeleteToDoItem = this._onDeleteToDoItem.bind(this);
    this.setRef = this.setRef.bind(this);
  }

  public render(): void {
    let props:IToDoListProps = {
      toDoList: this.items,
      onAddItem: this._onCreateToDoItem,
      onDeleteItem: this._onDeleteToDoItem,
    };
    
    let elem = React.createElement(
      ToDoList,
      {
        ...props, ref: this.setRef
      }
    );

    ReactDom.render(elem, this.domElement);
  }

  private setRef(elm:any){
    this.element = elm;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  private async _onCreateToDoItem(toDo:string): Promise<void>{
    if(toDo.trim() != ''){
      let it = await this.service.saveItem("LST_ToDoList",{Title: toDo});
      if(it != null){
        let item:IToDoListItem = {Id: it.Id, Title: it.Title};
        this.element.addItem(item);
      }
    }
  }

  private async _onDeleteToDoItem(id:number): Promise<void>{
    if(await this.service.deleteItem("LST_ToDoList", id)){
      this.element.deleteItem(id);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
