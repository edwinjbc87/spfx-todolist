declare interface IToDoListWebPartStrings {
  PropertyPaneTitle: string;
  BasicGroupName: string;
  ListTitleFieldLabel: string;
  Title: string;
  AddItemButtonText: string;
  AddItemPlaceholder: string;
}

declare module 'ToDoListWebPartStrings' {
  const strings: IToDoListWebPartStrings;
  export = strings;
}
