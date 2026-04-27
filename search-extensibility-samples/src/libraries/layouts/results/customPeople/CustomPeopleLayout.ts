import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from "MyCompanyLibraryLibraryStrings";

/**
 * Custom People Layout properties
 */
export interface ICustomPeopleLayoutProperties {
    profilePageURL?: string;
}

export class CustomPeopleLayout extends BaseLayout<ICustomPeopleLayoutProperties> {

    public async onInit(): Promise<void> {
        await this.loadMsGraphToolkit();
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {
        return [
            PropertyPaneTextField('layoutProperties.profilePageURL', {
                label: strings.Layouts.People.ProfilePageURL
            }),
        ];
    }

    /**
     * Loads the Microsoft Graph Toolkit library dynamically
     */
    private async loadMsGraphToolkit(): Promise<void> {

        const { Providers } = await import(
            /* webpackChunkName: 'microsoft-graph-toolkit' */
            '@microsoft/mgt-react/dist/es6'
        );

        const { SharePointProvider } = await import(
            /* webpackChunkName: 'microsoft-graph-toolkit' */
            '@microsoft/mgt-sharepoint-provider/dist/es6'
        );

        if (!Providers.globalProvider) {
            Providers.globalProvider = new SharePointProvider(this.context);
        }
    }
}
