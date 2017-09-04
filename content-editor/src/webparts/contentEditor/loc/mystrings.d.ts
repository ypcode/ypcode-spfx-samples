declare interface IContentEditorWebPartStrings {
  SourceLinkLabel: string;
  SourceFormatLabel: string;
  SourceTypeLabel: string;
  SourceConfigGroup: string;
  PropertyPaneHeader: string;
  ContentSourceTypeOption: string;
  LinkSourceTypeOption: string;
  displaySettingsGroupLabel: string;
  showCaptionSwitchLabel: string;

}

declare module 'ContentEditorWebPartStrings' {
  const strings: IContentEditorWebPartStrings;
  export = strings;
}
