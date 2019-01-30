import SPService from "../../../api/SPService";

export interface IToDoListProps {
  description: string;
  toDo?: string;
  toDoList: IToDoListItem[];
  service:SPService;
}

export interface IToDoListItem {
  Id: number;
  Title: string;
}