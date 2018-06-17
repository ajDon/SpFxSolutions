declare interface IMyTestWebpartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'MyTestWebpartWebPartStrings' {
  const strings: IMyTestWebpartWebPartStrings;
  export = strings;
}
