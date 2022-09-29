import { BaseDataSource, BuiltinTemplateSlots, IDataContext, IDataSourceData, ITemplateSlot } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { PropertyPaneLabel } from "@microsoft/sp-property-pane";

export interface ICustomDataSourceProperties {
  possibleResults: string;
}

export class CustomDataSource extends BaseDataSource<ICustomDataSourceProperties> {
    private _itemsCount: number;

    public async getData(dataContext?: IDataContext): Promise<IDataSourceData> {
      const possibleResults: string[] = this.properties.possibleResults ? this.properties.possibleResults.split(/\r?\n/) : [];
      const possibleResultItems: any = possibleResults.map((r, i) => {return {key: i, Title: r};});
      const results: string[] = possibleResultItems.filter(r => r.Title.toLocaleUpperCase() == dataContext.inputQueryText.toLocaleUpperCase());
      let data: IDataSourceData = {
        items: results,
        filters: []
      };
      this._itemsCount = results.length;
      return data;
    }

  public getItemCount(): number {
    return this._itemsCount;
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    return [
      {
          groupName: "Search settings",
          groupFields: [
            PropertyPaneTextField('dataSourceProperties.possibleResults', {
              label: 'Search results',
              multiline: true,
              rows: 5
            }),
            PropertyPaneLabel("", {
              text: "List of possible search results, one per line"
            })
          ]
      }
    ];
  }

  public getTemplateSlots(): ITemplateSlot[] {
    return [
        {
            slotName: BuiltinTemplateSlots.Title,
            slotField: 'Title'
        }
    ];
  }
}
