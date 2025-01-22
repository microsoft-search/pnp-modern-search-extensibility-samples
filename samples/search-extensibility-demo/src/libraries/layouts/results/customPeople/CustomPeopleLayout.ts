import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as strings from "ExtensibilityDemoLibraryStrings";

/**
 * Custom Layout properties
 */
export interface ICustomPeopleLayoutProperties {
    profilePageURL?: string;
}

export class CustomPeoplelayout extends BaseLayout<ICustomPeopleLayoutProperties> {

    /**
  * Dynamically loaded components for property pane
  */

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
    private async loadMsGraphToolkit() {

        // Load Microsoft Graph Toolkit dynamically
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
