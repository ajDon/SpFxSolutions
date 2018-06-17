declare interface IContentEditorWebpartWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  HtmlUrlFieldLabel: string;
  HTMLUrlLinkLabel: string;
  AddHtmlDirectlyFieldLabel: string;
  EnabledText: string;
  DisabledText: string;
  AddHtmlScriptFieldLabel: string;
  SpPageContextInfoFieldLabel: string;
  EnableSODFunctionsFieldLabel: string;
  EnableRequestDigestFieldLabel: string;
}

declare module 'ContentEditorWebpartWebPartStrings' {
  const strings: IContentEditorWebpartWebPartStrings;
  export = strings;
}
