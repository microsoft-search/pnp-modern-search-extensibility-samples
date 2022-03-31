import * as React from "react";
import * as ReactDOM from "react-dom";
import { DefaultButton, Dialog, DialogFooter, DialogType, ITheme, PrimaryButton } from 'office-ui-fabric-react';
import { BaseWebComponent, IDataContext, IDataSource } from '@pnp/modern-search-extensibility';
import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { ExportHelper } from "../helpers/ExportHelper";
import * as strings from 'ExportToCsvLibraryStrings'

export interface IExportComponentProps {
  /**
   * The columns to export
   */
  columns: string[];
  /**
  * The template context
  */
  context: IDataContext | any;
  /**
   * The data source
   */
  dataSource?: IDataSource;
}

enum ExportType {
  CurrentPage,
  All
}

export interface IExportComponentState {
  hideInfoDialog: boolean;
  isExporting: boolean;
  exportType: ExportType;
}

export class ExportComponent extends React.Component<IExportComponentProps, IExportComponentState> {
  private readonly maxhits = 10000;
  private readonly pagesize = 500;
  private readonly extension = ".csv";
  private readonly isExportSupported: boolean = false;

  constructor(props: IExportComponentProps) {
    super(props);

    this.isExportSupported = document && ("download" in document.createElement("a"));

    this.state = {
      hideInfoDialog: true,
      isExporting: false,
      exportType: ExportType.CurrentPage,
    };

    this.toggleInfoDialog = this.toggleInfoDialog.bind(this);
    this.exportTrigger = this.exportTrigger.bind(this);
  }

  private toggleInfoDialog(): void {
    this.setState(p => ({ hideInfoDialog: !p.hideInfoDialog }));
  }

  private async exportTrigger(exportAll?: boolean): Promise<void> {
    const { context, dataSource, columns } = this.props;
    const { exportType } = this.state;
    const fileName = "csvExport_" + (new Date().toLocaleDateString() + "_" + new Date().toLocaleTimeString()).replace(/[^\d_-]/g, "").trim();
    this.setState({ isExporting: true, hideInfoDialog: true });
    try {
      let items: any[] = [];
      let errorOccured = false;
      let errorColumnValue = false;
      if (dataSource && (exportAll === true || exportType == ExportType.All)) {
        console.time("fetching");
        try {
          let currentPageNumber = 0;
          let fetchMore = true;
          let itemsFetched = 0;
          while (items.length < this.maxhits && fetchMore) {
            const pagesToProcess = [++currentPageNumber, ++currentPageNumber, ++currentPageNumber, ++currentPageNumber];
            const itemResults = await Promise.all(pagesToProcess.map(async page => {
              const data = await dataSource.getData({ ...context, itemsCountPerPage: this.pagesize, pageNumber: page });
              itemsFetched += data.items?.length || 0;
              return data.items || [];
            }));

            itemResults.forEach(i => {
              fetchMore = fetchMore && i.length == this.pagesize;
              if (i.length > 0) { items = items.concat(i); }
            });
            console.log(`Processed '${pagesToProcess.join(", ")}', items total fetched ${items.length}`);
          }
        }
        catch (error) {
          errorOccured = true;
          console.log(`Error occurred while fetching result for csv export. ${error}`);
        }
        finally {
          console.timeEnd("fetching");
        }
      }
      else {
        items = context?.data?.items;
      }

      if (items) {
        console.time("mapvalues");
        const itemKeys = Object.keys(items && items.length && items[0] || {});
        const existingColumns = columns.filter(c => itemKeys.indexOf(c) >= 0);
        var result = items.map((item) => existingColumns.map(column => item[column]));
        console.timeEnd("mapvalues");

        console.time("exporttocsv");
        const emptyRows = result.filter(r => r.every(c => !c)).length;
        ExportHelper.exportToCsv(fileName + this.extension, result, existingColumns);
        console.log(`Processed '${fileName + this.extension}', items total exported ${result.length}, has error column value: ${errorColumnValue}, empty rows: ${emptyRows}`);
        console.timeEnd("exporttocsv");
      }
    }
    finally {
      this.setState({ isExporting: false });
    }
  }

  public render() {
    const { context, dataSource } = this.props;
    if (!context) return null;
    const { isExporting, hideInfoDialog } = this.state;
    const { totalItemsCount } = context.data;
    const disableExport = !this.isExportSupported || isExporting || !totalItemsCount;
    return <>
      <DefaultButton text={strings.ExportButtonText} split onClick={() => this.exportTrigger()}
        theme={context.theme as ITheme}
        primaryDisabled={disableExport}
        iconProps={{ iconName: "Save" }}
        menuProps={{
          items: [
            {
              key: 'exportAll',
              text: strings.ExportAllLabel?.replace("{maxhits}", this.maxhits.toString()),
              iconProps: { iconName: 'SaveAll' },
              onClick: () => { this.exportTrigger(true); },
              disabled: disableExport || dataSource == undefined
            },
            {
              key: 'information',
              text: strings.ExportInfoText,
              iconProps: { iconName: 'Info' },
              onClick: this.toggleInfoDialog
            }
          ]
        }} />
      {!hideInfoDialog && <Dialog
        hidden={hideInfoDialog}
        onDismiss={this.toggleInfoDialog}
        dialogContentProps={{
          type: DialogType.normal,
          title: strings.ExportInfoText,
          showCloseButton: false,
          subText: strings.ExportDialogHelpText?.replace("{maxhits}", this.maxhits.toString())
        }}
        modalProps={{ isBlocking: true }}
        theme={context.theme as ITheme}>
        {!this.isExportSupported && strings.ExportBrowserNotSupportedText}
        <DialogFooter>
          <PrimaryButton onClick={this.toggleInfoDialog} text={strings.ExportDialogOKButtonText} />
        </DialogFooter>
      </Dialog>}
    </>;
  }
}

export class ExportWebComponent extends BaseWebComponent {

  public constructor() {
    super();
  }

  public async connectedCallback() {
    let props = this.resolveAttributes();

    const columns = (props.columns || "").toString().split(',').map(s => s.trim()).filter(s => s);
    if (!columns || !props.context) return;

    let serviceScope: ServiceScope = this._serviceScope; // Default is the root shared service scope regardless the current Web Part
    let dataSourceServiceKey: ServiceKey<any>;

    if (props.instanceId || props.context.instanceId) {
      const instanceId = props.instanceId || props.context.instanceId;
      // Get the service scope and keys associated to the current Web Part displaying the component
      serviceScope = this._webPartServiceScopes.get(instanceId) ? this._webPartServiceScopes.get(instanceId) : serviceScope;
      dataSourceServiceKey = this._webPartServiceKeys.get(props.instanceId) ? this._webPartServiceKeys.get(props.instanceId).TheresNoDataSourceServiceKeyHereOnlyTemplateService : dataSourceServiceKey;
    }

    const dataSource = dataSourceServiceKey ? serviceScope.consume<IDataSource>(dataSourceServiceKey) : undefined;

    const exportComponent = <ExportComponent columns={columns} context={props.context} dataSource={dataSource} />;
    ReactDOM.render(exportComponent, this);
  }
}