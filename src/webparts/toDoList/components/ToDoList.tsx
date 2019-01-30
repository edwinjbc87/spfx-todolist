import * as React from 'react';
import styles from './ToDoList.module.scss';
import { IToDoListProps, IToDoListItem } from './IToDoListProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { 
  TextField, 
  ProgressIndicator,
  PrimaryButton, 
  IconButton, 
  List, 
  Label, 
  FocusZone, 
  FocusZoneDirection, 
  getRTLSafeKeyCode, 
  KeyCodes 
} from 'office-ui-fabric-react';

export interface IToDoListState {
  toDo?: string;
  toDoList: IToDoListItem[];
}

export default class ToDoList extends React.Component<IToDoListProps, IToDoListState> {
  private inputToDo:any;

  constructor(props: IToDoListProps) {
    super(props);

    this.state = {
      toDo: '',
      toDoList: this.props.toDoList
    };

    
    this._onRenderCell = this._onRenderCell.bind(this);
    this._deleteToDoItem = this._deleteToDoItem.bind(this);
    this._onChange = this._onChange.bind(this);
    this._getErrorMessage = this._getErrorMessage.bind(this);
    this._createToDoItem = this._createToDoItem.bind(this);

  }

  public render(): React.ReactElement<IToDoListProps> {


    return (
      <div className={ styles.toDoList }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <FocusZone
                direction={ FocusZoneDirection.vertical }
                isInnerZoneKeystroke={ (ev: React.KeyboardEvent<HTMLElement>) => ev.which === getRTLSafeKeyCode(KeyCodes.right) }
                >
                <h1 className={ styles.title }>To Do List</h1>
                <div className={ styles.toDoListForm }>
                  <TextField 
                    className={ styles.toDoTextField } 
                    value={this.state.toDo} 
                    ref={(input) => this.inputToDo = input}                    
                    onGetErrorMessage={this._getErrorMessage}
                    validateOnLoad={false}></TextField>                  
                  <PrimaryButton className={ styles.toDoButton } iconProps={{ iconName: 'Add' }} onClick={(event:any)=>{
                    if(this.inputToDo.value.trim() != ''){
                      this._createToDoItem(this.inputToDo.value);
                    }
                  }}>Add</PrimaryButton>
                </div>
                <List items={this.state.toDoList} onRenderCell={this._onRenderCell}></List>
              </FocusZone>
            </div>
            
          </div>
        </div>
      </div>
    );
  }
  

  private _onChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
    this.setState({
      toDo: newValue
    });
  }

  private _getErrorMessage(value:string):string{
    return value?'':"This field is required";
  }
  
  private async _createToDoItem(toDo:string): Promise<any>{
    await this.props.service.saveItem("LST_ToDoList",{Title: toDo});
    let items:IToDoListItem[] = await this.props.service.getItems("LST_ToDoList");
    
    this.setState({toDoList: items, toDo: ''});
  }

  private async _deleteToDoItem(id:number): Promise<any>{
    await this.props.service.deleteItem("LST_ToDoList", id);
    let items:IToDoListItem[] = await this.props.service.getItems("LST_ToDoList");
    
    this.setState({toDoList: items, toDo: ''});   
  }

  private _onRenderCell = (item: IToDoListItem, index: number): JSX.Element =>{
    return (
      <div className={ styles.toDoListItem } data-is-focusable={true}>
        <Label data-id={item.Id}>{item.Title}</Label>
        <IconButton  iconProps={{ iconName: 'Delete' }} data-id={item.Id} onClick={(evt:any)=>{this._deleteToDoItem(item.Id)}} />
      </div>
    );
  }
}