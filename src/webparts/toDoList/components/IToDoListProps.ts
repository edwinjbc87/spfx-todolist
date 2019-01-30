import { IToDoListItem } from './IToDoListItem';

export interface IToDoListProps {
  toDoList: IToDoListItem[];
  onDeleteItem?(id: number): Promise<void>;
  onAddItem?(toDo: string): Promise<void>;
}