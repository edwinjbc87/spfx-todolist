import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
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
  listTitle: string;
}

export default class ToDoListWebPart extends BaseClientSideWebPart<IToDoListWebPartProps> {
  private service:SPService;
  private webUrl:string;
  private element:ToDoList;
  private todoProps:IToDoListProps;
  private isValidToDoList:boolean;

  protected async onInit(): Promise<void> {   
        
    this._onCreateToDoItem = this._onCreateToDoItem.bind(this);
    this._onDeleteToDoItem = this._onDeleteToDoItem.bind(this);
    this._setToDoList = this._setToDoList.bind(this);
    

    this.todoProps = {
      items: []
    };

    if (!DEBUG || Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) { 
      this.service = new SPService(this.context);        
      await this.service.init();    
  
      this.webUrl = this.context.pageContext.web.absoluteUrl;

      this.todoProps.onAddItem = this._onCreateToDoItem;
      this.todoProps.onDeleteItem = this._onDeleteToDoItem;
    }

    await this._setToDoList(this.properties.listTitle);
  }

  public render(): void {
    const element = React.createElement(
      ToDoList,
      {
        ...this.todoProps, ref: (elm:ToDoList) =>{ this.element = elm; }
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

  private async _setToDoList(listTitle:string):Promise<void>{
    if(listTitle && listTitle.trim() != ''){
      try{
        let list = await this.service.getList(listTitle);
        this.isValidToDoList = (list != null);
        if(this.isValidToDoList){
          this.todoProps.items = await this.service.getItems(listTitle);
        } else {
          this.todoProps.items = [];
        }        
      }
      catch(ex){
        this.isValidToDoList = false;
        this.todoProps.items = [];
      }      
    } else {
      this.isValidToDoList = false;
      this.todoProps.items = [];
    }
    if(this.element) this.element.setItems(this.todoProps.items);
  }

  private async _onCreateToDoItem(toDo:string): Promise<void>{    
    if(toDo.trim() != ''){
      let it = null;
      if(this.isValidToDoList){
        it = await this.service.saveItem(this.properties.listTitle, {Title: toDo});
      } else {
        it = {Id: (new Date()).getTime(), Title: toDo};
      }
      if(it != null){
        const item:IToDoListItem = {Id: it.Id, Title: it.Title};
        this.element.addItem(item);
      }
    }    
  }

  private async _onDeleteToDoItem(id:number): Promise<void>{
    if(!this.isValidToDoList || await this.service.deleteItem(this.properties.listTitle, id)){
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
                  label: strings.ListTitleFieldLabel,
                  validateOnFocusOut: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
  
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {    
    if (propertyPath === 'listTitle') {
      this._setToDoList(newValue);
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}
