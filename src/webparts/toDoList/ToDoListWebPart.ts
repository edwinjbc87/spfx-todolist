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
import { IToDoListProps, IToDoListItem } from './components/IToDoListProps';
import SPService from '../../api/SPService';
import ICommonWebPartProps from '../ICommonWebPartProps';

export interface IToDoListWebPartProps extends ICommonWebPartProps{
  description: string;
  items: IToDoListItem[];
}

export interface IToDoListWebPartState{
  items: IToDoListItem[];
}

export default class ToDoListWebPart extends BaseClientSideWebPart<IToDoListWebPartProps> {
  private service:SPService;

  protected async onInit(): Promise<void> {    
    this.service = new SPService(this.context);
    await this.service.init();    

    this.properties.webUrl = this.context.pageContext.web.absoluteUrl;
    this.properties.userEmail = this.context.pageContext.user.email;
    this.properties.userId = await this.service.getUserId(this.properties.userEmail);
    this.properties.items = await this.service.getItems('LST_ToDoList');
    
  }

  public render(): void {
    const element: React.ReactElement<IToDoListProps > = React.createElement(
      ToDoList,
      {
        description: this.properties.description,
        toDoList: this.properties.items,
        service: this.service
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
