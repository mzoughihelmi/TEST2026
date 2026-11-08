declare interface IMesFavorisWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  ListTitleFieldLabel: string;
  ListTitleFieldDescription: string;
  MaxItemsFieldLabel: string;
  VoirPlusUrlFieldLabel: string;
  VoirPlusUrlFieldDescription: string;
  VoirPlusLabelFieldLabel: string;
}

declare module 'MesFavorisWebPartStrings' {
  const strings: IMesFavorisWebPartStrings;
  export = strings;
}
