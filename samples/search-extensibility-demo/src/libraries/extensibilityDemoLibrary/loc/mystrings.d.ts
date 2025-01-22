declare interface IExtensibilityDemoLibraryStrings {
  Layouts: {
    CustomSimpleList: {
      ShowFileIconLabel: string;
      ShowItemThumbnailLabel: string;
      OpenLinkInNewTab: string;
    },
    People:{
      ProfilePageURL: string
    }
  }
}

declare module 'ExtensibilityDemoLibraryStrings' {
  const strings: IExtensibilityDemoLibraryStrings;
  export = strings;
}
