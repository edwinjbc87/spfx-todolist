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

    this.addItem = this.addItem.bind(this);
    this.deleteItem = this.deleteItem.bind(this);
    this._onRenderCell = this._onRenderCell.bind(this);
    this._onChange = this._onChange.bind(this);
    this._getErrorMessage = this._getErrorMessage.bind(this);
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
                    this.props.onAddItem(this.inputToDo.value);                    
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

  public addItem(item: IToDoListItem):void{
    let toDoList = JSON.parse(JSON.stringify(this.state.toDoList));
    
    toDoList.push(item);
    
    this.setState({toDo: '', toDoList: toDoList});
  }
  
  public deleteItem(id: number):void{    
    let idx = -1;
    let toDoList = JSON.parse(JSON.stringify(this.state.toDoList));

    for(let i=0; i < toDoList.length; i++){
      if(toDoList[i].Id == id){
        idx = i; break;
      }
    }
    
    if(idx>=0) toDoList.splice(idx, 1);

    this.setState({toDoList: toDoList});
  }

  private _onChange(event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
    this.setState({
      toDo: newValue
    });
  }

  private _getErrorMessage(value:string):string{
    return value?'':"This field is required";
  }
  
  private _onRenderCell = (item: IToDoListItem, index: number): JSX.Element =>{
    return (
      <div className={ styles.toDoListItem } data-is-focusable={true}>
        <Label data-id={item.Id}>{item.Title}</Label>
        <IconButton  iconProps={{ iconName: 'Delete' }} data-id={item.Id} onClick={(evt:any)=>{this.props.onDeleteItem(item.Id)}} />
      </div>
    );
  }
}