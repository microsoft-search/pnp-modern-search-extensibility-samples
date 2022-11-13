import { BaseQueryModifier } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';

export interface ICustomQueryModifierProperties {
  prefix: string;
  suffix: string;
}

export class CustomQueryModifier extends BaseQueryModifier<ICustomQueryModifierProperties> {

  private _regex: RegExp;

  public async onInit(): Promise<void> {
    this._regex = new RegExp('\\b(?!OR|NEAR|ONEAR|WORDS|XRANK\\b)\\w+(?!(\\s)*[:,=,<,>]|\\.\\.)\\b', 'gm');
  }

  public async modifyQuery(queryText: string): Promise<string> {

    this._regex.lastIndex = 0;

    const alteredQueryText = queryText?.replace(this._regex, (match => {
      return `${this.properties.prefix}${match}${this.properties.suffix}`;
    }));

    return alteredQueryText;
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

    return [
      {
        groupName: myLibraryStrings.CustomQueryModifier.GroupName,
        groupFields: [
          PropertyPaneTextField('queryModifierProperties.prefix', {
            label: myLibraryStrings.CustomQueryModifier.PrefixLabel,
            description: myLibraryStrings.CustomQueryModifier.PrefixDescription,
            placeholder: myLibraryStrings.CustomQueryModifier.PrefixPlaceholder,
          }),
          PropertyPaneTextField('queryModifierProperties.suffix', {
            label: myLibraryStrings.CustomQueryModifier.SuffixLabel,
            description: myLibraryStrings.CustomQueryModifier.SuffixDescription,
            placeholder: myLibraryStrings.CustomQueryModifier.SuffixPlaceholder,
          })
        ],
      },
    ];
  }
}