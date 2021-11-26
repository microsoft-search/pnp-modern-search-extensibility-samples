import { ServiceKey } from "@microsoft/sp-core-library";
import {  IExtensibilityLibrary, 
          IComponentDefinition, 
          ISuggestionProviderDefinition, 
          ISuggestionProvider,
          ILayoutDefinition, 
          LayoutType, 
          ILayout 
} from "@pnp/modern-search-extensibility";
import { MyCustomComponentWebComponent } from "../CustomComponent";
import { Customlayout } from "../CustomLayout";
import { CustomSuggestionProvider } from "../CustomSuggestionProvider";

export class MyCompanyLibraryLibrary implements IExtensibilityLibrary {

  public getCustomLayouts(): ILayoutDefinition[] {
    return [
      {
        name: 'PnP Custom layout',
        iconName: 'Color',
        key: 'CustomLayout',
        type: LayoutType.Results,
        templateContent: require('../custom-layout.html'),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomLayout', Customlayout),
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'my-custom-component',
        componentClass: MyCustomComponentWebComponent
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [
        {
          name: 'Custom Suggestions Provider',
          key: 'CustomSuggestionsProvider',
          description: 'A demo custom suggestions provider from the extensibility library',
          serviceKey: ServiceKey.create<ISuggestionProvider>('MyCompany:CustomSuggestionsProvider', CustomSuggestionProvider)
      }
    ];
  }

  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {

    // Register custom Handlebars helpers
    // Usage {{myHelper 'value'}}
    namespace.registerHelper('myHelper', (value: string) => {
      return new namespace.SafeString(value.toUpperCase());
    });
  }
}
