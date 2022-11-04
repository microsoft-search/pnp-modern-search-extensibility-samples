import { IPropertyPaneField, IPropertyPaneGroup, PropertyPaneDropdown, PropertyPaneTextField } from "@microsoft/sp-property-pane";
import { ServiceScope } from '@microsoft/sp-core-library';
import { BaseDataSource, ITokenService, FilterBehavior, PagingBehavior, IDataContext, ITemplateSlot, BuiltinTemplateSlots } from '@pnp/modern-search-extensibility';
import { AnonymousRestService } from './AnonymousRestService';

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
    * The root key for the returned data
    */
    rootKey?: string;
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

    }

    public async onInit(): Promise<void> {

        this.serviceScope.whenFinished(() => {

            this._tokenService = this.serviceScope.consume<ITokenService>(this.serviceKeys.TokenService);
            this._anonRestService = this.serviceScope.consume<AnonymousRestService>(AnonymousRestService.ServiceKey);
        });

        this.initProperties();
    }

    private initProperties() {
        this.properties.urlTemplate = this.properties.urlTemplate || '';
        this.properties.method = this.properties.method || 'GET';
        this.properties.bodyTemplate = this.properties.bodyTemplate || '';
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

    public async getData(dataContext: IDataContext): Promise<{ items: any[] }> {

        return await this.search();
    }

    public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

        const requestFields: IPropertyPaneField<any>[] = [];

        requestFields.push(
            PropertyPaneTextField('dataSourceProperties.urlTemplate', {
                value: this.properties.urlTemplate,
                label: 'Url Template',
                placeholder: 'e.g. https://registry.npmjs.org/-/v1/search?text={inputQueryText}&size=10',
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

                }),
            PropertyPaneTextField('dataSourceProperties.rootKey', {
                value: this.properties.rootKey,
                label: 'JSON Return Root Key',
                placeholder: 'e.g. objects',
                multiline: false,
                description: "Enter key of the root array containing the search result items",
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
        //no-code
    }

    public onCustomPropertyUpdate(propertyPath: string, newValue: any, changeCallback?: (targetProperty?: string, newValue?: any) => void): void {
        //no-code
    }

    public getTemplateSlots(): ITemplateSlot[] {
        return [
            {
                slotName: BuiltinTemplateSlots.Title,
                slotField: 'name'
            },
            {
                slotName: BuiltinTemplateSlots.Path,
                slotField: 'webUrl'
            },
            {
                slotName: BuiltinTemplateSlots.Id,
                slotField: 'key'
            }
        ];
    }

    public getSortableFields(): string[] {
        return [];
    }


    /**
     * Retrieves data from any rest endpoint     
     */
    private async search(): Promise<{ items: any[] }> {

        const response = {
            items: []
        };

        const url = await this._tokenService.resolveTokens(this.properties.urlTemplate);
        const body = await this._tokenService.resolveTokens(this.properties.bodyTemplate);
        const method = this.properties.method;

        const jsonResponse = await this._anonRestService.requestData(url, method, body);

        if (jsonResponse) {

            const items = this.properties.rootKey && jsonResponse.hasOwnProperty(this.properties.rootKey) ? jsonResponse[this.properties.rootKey] : jsonResponse;
            if (items && Array.isArray(items)) {
                items.forEach(item => response.items.push(item));
            }
        }

        this._itemsCount = response.items?.length ?? 0;

        return response;
    }
}

