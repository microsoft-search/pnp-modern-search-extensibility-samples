import { BaseLayout } from "@pnp/modern-search-extensibility";
import { IPropertyPaneField, PropertyPaneToggle } from '@microsoft/sp-property-pane';
import * as strings from "MyCompanyLibraryLibraryStrings";

/**
 * Custom Simple List Layout properties
 */
export interface ICustomSimpleListLayoutProperties {
    /**
     * Show or hide the file icon
     */
    showFileIcon: boolean;

    /**
     * Show or hide the item thumbnail
     */
    showItemThumbnail: boolean;

    /**
     * Open links in a new tab
     */
    openLinkInNewTab: boolean;
}

export class CustomSimpleListLayout extends BaseLayout<ICustomSimpleListLayoutProperties> {

    public async onInit(): Promise<void> {

        this.properties.showFileIcon = this.properties.showFileIcon !== null && this.properties.showFileIcon !== undefined ? this.properties.showFileIcon : true;
        this.properties.showItemThumbnail = this.properties.showItemThumbnail !== null && this.properties.showItemThumbnail !== undefined ? this.properties.showItemThumbnail : true;
        this.properties.openLinkInNewTab = this.properties.openLinkInNewTab !== null && this.properties.openLinkInNewTab !== undefined ? this.properties.openLinkInNewTab : true;
    }

    public getPropertyPaneFieldsConfiguration(availableFields: string[]): IPropertyPaneField<any>[] {

        return [
            PropertyPaneToggle('layoutProperties.showFileIcon', {
                label: strings.Layouts.CustomSimpleList.ShowFileIconLabel
            }),
            PropertyPaneToggle('layoutProperties.showItemThumbnail', {
                label: strings.Layouts.CustomSimpleList.ShowItemThumbnailLabel
            }),
            PropertyPaneToggle('layoutProperties.openLinkInNewTab', {
                label: strings.Layouts.CustomSimpleList.OpenLinkInNewTab
            })
        ];
    }
}
