import { BaseDataSource, BuiltinTemplateSlots, IDataContext, IDataSourceData, ITemplateSlot, PagingBehavior } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { PropertyPaneLabel } from "@microsoft/sp-property-pane";
import { PropertyPaneSlider } from "@microsoft/sp-property-pane";
import { ServiceScope } from "@microsoft/sp-core-library";
import { HttpClient, HttpClientConfiguration } from "@microsoft/sp-http";

export interface ICustomDataSourceProperties {
  qualityEffect: number;
  popularityEffect: number;
  maintenanceEffect: number;
}

export class CustomDataSource extends BaseDataSource<ICustomDataSourceProperties> {
    private _itemsCount: number;
    private _httpClient: HttpClient;

    public constructor(serviceScope: ServiceScope) {
      super(serviceScope);

      serviceScope.whenFinished(() => {
          this._httpClient = serviceScope.consume<HttpClient>(HttpClient.serviceKey);
      });
    }

    public async getData(dataContext?: IDataContext): Promise<IDataSourceData> {
      const rowLimit = dataContext.itemsCountPerPage ? (dataContext.itemsCountPerPage > 250 ? 250 : dataContext.itemsCountPerPage) : 50;
      let startRow = 0;
      if (dataContext.pageNumber > 1) {
          startRow = (dataContext.pageNumber - 1) * rowLimit;
      }
      const response = await this._httpClient.get(`https://registry.npmjs.org/-/v1/search?text=${dataContext.inputQueryText}&size=${rowLimit}&from=${startRow}&quality=${this.properties.qualityEffect ?? 0}&popularity=${this.properties.popularityEffect ?? 0}&maintenance=${this.properties.maintenanceEffect ?? 0}`,HttpClient.configurations.v1);
      const results = await response.json();
      let data: IDataSourceData = {
        items: results.objects,
        filters: []
      };
      this._itemsCount = results.total;
      return data;
    }

  public getItemCount(): number {
    return this._itemsCount;
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    return [
      {
          groupName: "Npm Source",
          groupFields: [
            PropertyPaneSlider('dataSourceProperties.qualityEffect', {
              label: 'Quality effect',
              max: 1,
              min: 0,
              step: 0.01
            }),
            PropertyPaneLabel("", {
              text: "How much of an effect should quality have on search results"
            }),
            PropertyPaneSlider('dataSourceProperties.popularityEffect', {
              label: 'Popularity effect',
              max: 1,
              min: 0,
              step: 0.01
            }),
            PropertyPaneLabel("", {
              text: "How much of an effect should popularity have on search results"
            }),
            PropertyPaneSlider('dataSourceProperties.maintenanceEffect', {
              label: 'Maintenance effect',
              max: 1,
              min: 0,
              step: 0.01
            }),
            PropertyPaneLabel("", {
              text: "How much of an effect should maintenance have on search results"
            })
          ]
      }
    ];
  }

  public getPagingBehavior(): PagingBehavior {
    return PagingBehavior.Dynamic;
  }

  public getTemplateSlots(): ITemplateSlot[] {
    return [
        {
            slotName: BuiltinTemplateSlots.Title,
            slotField: 'package.name'
        },
        {
          slotName: BuiltinTemplateSlots.PreviewUrl,
          slotField: 'package.links.npm'
        },
        {
          slotName: BuiltinTemplateSlots.Author,
          slotField: 'package.author.name'
        },
        {
          slotName: BuiltinTemplateSlots.Summary,
          slotField: 'package.description'
        },
        {
          slotName: BuiltinTemplateSlots.Date,
          slotField: 'package.date'
        },
    ];
  }
}
