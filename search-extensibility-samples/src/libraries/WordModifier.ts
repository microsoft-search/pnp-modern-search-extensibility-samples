import { BaseQueryModifier, IDataContext } from "@pnp/modern-search-extensibility";
import { IPropertyPaneGroup, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import * as myLibraryStrings from 'MyCompanyLibraryLibraryStrings';

export interface IWordModifierProperties {
  prefix: string;
  suffix: string;
}

export class WordModifier extends BaseQueryModifier<IWordModifierProperties> {

  private _regex: RegExp;

  public async onInit(): Promise<void> {
    this._regex = new RegExp('\\b(?!OR|NEAR|ONEAR|WORDS|XRANK\\b)\\w+(?!(\\s)*[:,=,<,>]|\\.\\.)\\b', 'gm');
  }

  public async modifyQuery(queryText: string, dataContext: IDataContext, resolveTokens: (string: string) => Promise<string>): Promise<string> {

    this._regex.lastIndex = 0;

    const alteredQueryText = queryText?.replace(this._regex, (match => {
      return  `${this.properties.prefix}${match}${this.properties.suffix}`;
    }));

    return  alteredQueryText; 
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {

    return [
      {
        groupName: myLibraryStrings.WordModifier.GroupName,
        groupFields: [
          PropertyPaneTextField('queryModifierProperties.prefix', {
            label: myLibraryStrings.WordModifier.PrefixLabel,
            description: myLibraryStrings.WordModifier.PrefixDescription,
            placeholder: myLibraryStrings.WordModifier.PrefixPlaceholder,
          }),
          PropertyPaneTextField('queryModifierProperties.suffix', {
            label: myLibraryStrings.WordModifier.SuffixLabel,
            description: myLibraryStrings.WordModifier.SuffixDescription,
            placeholder: myLibraryStrings.WordModifier.SuffixPlaceholder,
          })
        ],
      },
    ];
  }
}