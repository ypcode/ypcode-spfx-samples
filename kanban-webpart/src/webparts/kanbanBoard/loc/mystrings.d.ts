declare interface IKanbanBoardStrings {
  HeaderDescription: string;
  TasksConfigurationGroup: string;
  SourceTasksList: string;
  StatusFieldInternalName: string;
  PleaseConfigureWebPartMessage: string;
}

declare module 'kanbanBoardStrings' {
  const strings: IKanbanBoardStrings;
  export = strings;
}
