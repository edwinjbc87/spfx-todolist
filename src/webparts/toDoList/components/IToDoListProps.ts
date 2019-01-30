export interface IToDoListProps {
  toDoList: IToDoListItem[] ,
  onDeleteItem?(id: number): Promise<void> ,
  onAddItem?(toDo: string): Promise<void>
}

export interface IToDoListItem {
  Id: number;
  Title: string;
}