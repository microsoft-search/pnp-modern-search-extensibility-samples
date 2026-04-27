declare interface IMyCompanyLibraryLibraryStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  CustomQueryModifier: {
      GroupName: string;
      PrefixLabel: string;
      PrefixDescription: string;
      PrefixPlaceholder: string;
      SuffixLabel: string;
      SuffixDescription: string;
      SuffixPlaceholder: string;
    };
  Layouts: {
    CustomSimpleList: {
      ShowFileIconLabel: string;
      ShowItemThumbnailLabel: string;
      OpenLinkInNewTab: string;
    };
    People: {
      ProfilePageURL: string;
    };
  };
}

declare module 'MyCompanyLibraryLibraryStrings' {
  const strings: IMyCompanyLibraryLibraryStrings;
  export = strings;
}
