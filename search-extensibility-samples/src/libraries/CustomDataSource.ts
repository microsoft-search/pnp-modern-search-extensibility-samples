import { IPropertyPaneField, IPropertyPaneGroup, PropertyPaneDropdown,PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { ServiceScope } from '@microsoft/sp-core-library';
import { BaseDataSource, ITokenService, FilterBehavior, PagingBehavior, IDataContext, ITemplateSlot } from '@pnp/modern-search-extensibility';
import { AnonymousRestService } from './anonymousRestService/AnonymousRestService';

export interface ICustomDataSourceProperties {

    /**
     * url template
     */
    urlTemplate: string;

    /**
     * Method to use
     */
    method: 'GET' | 'POST';

    /**
     * BodyTemplate
     */

    bodyTemplate: string | null;

    /**
    * The search query template
    */
    queryTemplate: string;
}

export class CustomDataSource extends BaseDataSource<ICustomDataSourceProperties> {

    private _anonRestService: AnonymousRestService;
    private _tokenService: ITokenService;
    /**
     * The data source items count
     */
    private _itemsCount = 0;

    public constructor(serviceScope: ServiceScope) {
        super(serviceScope);

        serviceScope.whenFinished(() => {
            
            this._tokenService = serviceScope.consume<ITokenService>(this.serviceKeys.TokenService);
            this._anonRestService = serviceScope.consume<AnonymousRestService>(AnonymousRestService.ServiceKey);
        });
    }

    public async onInit(): Promise<void> {   
    }

    public getItemCount(): number {
        return this._itemsCount;
    }

    public getFilterBehavior(): FilterBehavior {
        return FilterBehavior.Dynamic;
    }

    public getPagingBehavior(): PagingBehavior {
        return PagingBehavior.Dynamic;
    }

    public async getData(dataContext: IDataContext): Promise<{items:any[]}> {

        let results = {
            items: []
        };

        results = await this.search(dataContext);

        return results;
    }

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

        const requestFields: IPropertyPaneField<any>[] = [];
       
        requestFields.push(
            PropertyPaneTextField('dataSourceProperties.urlTemplate', {
                value: this.properties.urlTemplate,
                label: 'Url Template',
                placeholder: 'e.g. http://abc/{TOKEN}?q={QueryToken}',
                multiline: false,
                description: "Enter url template",            
            }),
            PropertyPaneDropdown('dataSourceProperties.method',
                {
                    label: 'Method',
                    options: [
                        {
                            key: 'GET',
                            text: 'GET'
                        },
                        {
                            key: 'POST',
                            text: 'POST'
                        },
                    ]

                })
        );

        if (this.properties.method === 'POST') {
            requestFields.push(
                 PropertyPaneTextField('dataSourceProperties.bodyTemplate', {
                    value: this.properties.bodyTemplate,
                    label: 'Body Template',
                    placeholder: '{ xx:yyy}',
                    multiline: true,
                    description: "Enter body template"
                }));
        }

       
        const groupFields: IPropertyPaneField<any>[] = [
            ...requestFields
        ];

        return [
            {
                groupName: "Rest Request",
                groupFields: groupFields
            }
        ];
    }

    public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any) {

    }

    public onCustomPropertyUpdate(propertyPath: string, newValue: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        //no-code
    }

    public getTemplateSlots(): ITemplateSlot[] {
        return [            
        ];
    }

    public getSortableFields(): string[] {
        return  [];
    }


    /**
     * Retrieves data from Microsoft Graph API
     * @param searchRequest the Microsoft Search search request
     */
    private async search(dataContext: IDataContext): Promise<any> {

        const response= {
            items: []
        };

        const url= await this._tokenService.resolveTokens(this.properties.urlTemplate);        
        const body= await this._tokenService.resolveTokens(this.properties.bodyTemplate);
        const method= this.properties.method;

        const jsonResponse = await this._anonRestService.requestData(url,method,body);

        console.log(jsonResponse);

        if (jsonResponse && Array.isArray(jsonResponse)) {

            jsonResponse.forEach(item => response.items.push(item));            
        }

        this._itemsCount = response.items?.length ?? 0;

        return response;
    }
}