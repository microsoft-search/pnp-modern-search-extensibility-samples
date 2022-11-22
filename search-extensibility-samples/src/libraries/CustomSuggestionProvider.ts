import { BaseSuggestionProvider, ISuggestion } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { MSGraphClientFactory } from "@microsoft/sp-http";
import { IMicrosoftSearchResponse, IMicrosoftSearchResultSet } from "../models/IMicrosoftSearchResponse";
import { IMicrosoftSearchQuery } from "../models/IMicrosoftSearchRequest";

const PARKER_ICON_URL = 'https://raw.githubusercontent.com/pnp/media/master/parker/pnp/300w/parker.png';
const PNP_ICON_URL = 'https://raw.githubusercontent.com/pnp/media/master/pnp-logos-generics/png/teal/300w/pnp-samples-teal-300.png';

export interface ICustomSuggestionProviderProperties {
  myProperty: string;
}

export class CustomSuggestionProvider extends BaseSuggestionProvider<ICustomSuggestionProviderProperties> {

    private _zeroTermSuggestions: ISuggestion[] = [];
  
    public async onInit(): Promise<void> {
    }
  
    public get isZeroTermSuggestionsEnabled(): boolean {
        return true;
    }
  
    public async getSuggestions(queryText: string): Promise<ISuggestion[]> {

        const productsSuggestions = await (await this.getProductsSuggestions(queryText)).map(item => {
          return {
            displayText: item.name,
            groupName: item.productCategory,
            description: item.productNumber,
            iconSrc: `https://sonbaedev.sharepoint.com/sites/espc22/ProductImages/${item.productId}.jpg`
          } as ISuggestion;

        });

        return productsSuggestions;
    }
  
    public async getZeroTermSuggestions(): Promise<ISuggestion[]> {
        return [];
    }

    private async getProductsSuggestions(queryText: string): Promise<{[key: string]: any}[]>  {

      let items: {[key: string]: any}[] = [];

      const msGraphClientFactory = this.context.serviceScope.consume<MSGraphClientFactory>(MSGraphClientFactory.serviceKey);
      const msGraphClient = await msGraphClientFactory.getClient('3');
      const request = await msGraphClient.api("https://graph.microsoft.com/v1.0/search/query");

      const searchQuery: IMicrosoftSearchQuery = {
        requests: [
          {
            entityTypes: ["externalItem"],
            contentSources: [
              "/external/connections/advworksproducts"
            ],
            query: {
              queryString: queryText
            },
            fields: ["name","productId","productNumber","productCategory"]
          }
        ]
      };

      const jsonResponse: IMicrosoftSearchResponse = await request.headers(
        { 'SdkVersion': 'pnpmodernsearch/' + this.context.manifest.version }).post(searchQuery);

      if (jsonResponse.value && Array.isArray(jsonResponse.value)) {

          jsonResponse.value.forEach((value: IMicrosoftSearchResultSet) => {

              // Map results
              value.hitsContainers.forEach(hitContainer => {

                  if (hitContainer.hits) {

                      const hits = hitContainer.hits.map(hit => {

                        // 'externalItem' will contain resource.properties but 'listItem' will be resource.fields
                        const propertiesFieldName = hit.resource.properties ? 'properties' : (hit.resource.properties ? 'fields' : null)

                          if (propertiesFieldName) {

                              // Flatten 'fields' to be usable with the Search Fitler WP as refiners
                              Object.keys(hit.resource[propertiesFieldName]).forEach(field => {
                                  hit[field] = hit.resource[propertiesFieldName][field];
                              });
                          }

                          return hit;
                      });

                      items = items.concat(hits);
                  }
              });
          });
        }

        return items;
    }

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

       return [
         {
           groupName: 'Custom Search Suggestions',
           groupFields: [
             PropertyPaneTextField('providerProperties.myProperty', {
              label: 'My property'
             })
           ]
         }
       ];
    }
}