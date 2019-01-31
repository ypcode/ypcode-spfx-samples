declare interface IWebPartWithPanelWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'WebPartWithPanelWebPartStrings' {
  const strings: IWebPartWithPanelWebPartStrings;
  export = strings;
}
