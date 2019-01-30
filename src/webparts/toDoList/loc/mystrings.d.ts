declare interface IToDoListWebPartStrings {
  PropertyPaneTitle: string;
  BasicGroupName: string;
  ListTitleFieldLabel: string;
  Title: string;
  AddItemButtonText: string;
  ToDoPlaceholder: string;
}

declare module 'ToDoListWebPartStrings' {
  const strings: IToDoListWebPartStrings;
  export = strings;
}
