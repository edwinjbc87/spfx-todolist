import * as React from 'react';
import styles from './ToDoList.module.scss';
import { IToDoListProps } from './IToDoListProps';
import { IToDoListItem } from './IToDoListItem';
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
                  <PrimaryButton className={ styles.toDoButton } iconProps={{ iconName: 'Add' }} onClick={(e)=>this.handleAddItem(this.inputToDo.value)}>Add</PrimaryButton>
                </div>
                <List items={this.state.toDoList} onRenderCell={this._onRenderCell}></List>
              </FocusZone>
            </div>
            
          </div>
        </div>
      </div>
    );
  }

  private handleAddItem(toDo:string){
    if(this.props.onAddItem){
      this.props.onAddItem(toDo);
    } else {
      if(toDo.trim() != ''){
        let item = {Id: (new Date()).getTime(), Title: toDo};
        this.addItem(item);
      }      
    }
  }

  private handleDeleteItem(id:number){        
    if(this.props.onDeleteItem){
      this.props.onDeleteItem(id);
    } else {
      this.deleteItem(id);
    }
  }

  public addItem(item: IToDoListItem):void{
    this.setState({toDo: '', toDoList: [...this.state.toDoList,item]});
  }
  
  public deleteItem(id: number):void{
    this.setState({toDoList: this.state.toDoList.filter((_, i) => _.Id !== id)});
  }

  public getItems():IToDoListItem[]{
    return this.state.toDoList;
  }

  public setItems(items:IToDoListItem[]):void{
    this.setState({toDoList: items});
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
        <IconButton  iconProps={{ iconName: 'Delete' }} data={{Id: item.Id}} onClick={(e)=>this.handleDeleteItem(item.Id)} />
      </div>
    );
  }
}