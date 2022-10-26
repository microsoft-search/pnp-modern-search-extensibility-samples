declare interface IMyCompanyLibraryLibraryStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
    WordModifier: {
      GroupName:string;
      PrefixLabel:string;
      PrefixDescription:string;
      PrefixPlaceholder:string;
      SuffixLabel:string;
      SuffixDescription:string;
      SuffixPlaceholder:string;
    }
}

declare module 'MyCompanyLibraryLibraryStrings' {
  const strings: IMyCompanyLibraryLibraryStrings;
  export = strings;
}
