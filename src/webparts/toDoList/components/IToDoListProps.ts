import { IToDoListItem } from './IToDoListItem';

export interface IToDoListProps {
  items: IToDoListItem[];
  onDeleteItem?(id: number): Promise<void>;
  onAddItem?(toDo: string): Promise<void>;
}