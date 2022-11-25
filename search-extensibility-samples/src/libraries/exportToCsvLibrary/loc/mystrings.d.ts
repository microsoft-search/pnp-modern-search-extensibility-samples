declare interface IExportToCsvLibraryLibraryStrings {
  ExportButtonText: string,
  ExportInfoText: string,
  ExportDialogHelpText: string,
  ExportBrowserNotSupportedText: string
  ExportCurrentPageLabel: string
  ExportAllLabel: string
  ExportDialogOKButtonText: string
}

declare module 'ExportToCsvLibraryStrings' {
  const strings: IExportToCsvLibraryLibraryStrings;
  export = strings;
}
