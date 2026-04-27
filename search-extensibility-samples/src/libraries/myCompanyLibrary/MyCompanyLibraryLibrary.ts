import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import {  IExtensibilityLibrary, 
          IComponentDefinition, 
          ISuggestionProviderDefinition, 
          ISuggestionProvider,
          ILayoutDefinition, 
          LayoutType, 
          ILayout,
          LayoutRenderType,
          IQueryModifierDefinition,
          IQueryModifier,
          IDataSourceDefinition,
          IDataSource
} from "@pnp/modern-search-extensibility";
import * as Handlebars from "handlebars";
import { MyCustomComponentWebComponent } from "../CustomComponent";
import { CustomLayout } from "../CustomLayout";
import { CustomSuggestionProvider } from "../CustomSuggestionProvider";
import { CustomQueryModifier } from "../CustomQueryModifier";
import { CustomDataSource } from "../CustomDataSource";
import { CustomPersonaWebComponent } from "../components/CustomPersonaComponent";
import { CustomPersonCardWebComponent } from "../components/CustomPersonCardComponent";
import { CustomSimpleListLayout } from "../layouts/results/customSimpleList/CustomSimpleListLayout";
import { CustomPeopleLayout } from "../layouts/results/customPeople/CustomPeopleLayout";

export class MyCompanyLibraryLibrary implements IExtensibilityLibrary {
  

  public static readonly serviceKey: ServiceKey<MyCompanyLibraryLibrary> =
  ServiceKey.create<MyCompanyLibraryLibrary>('SPFx:MyCustomLibraryComponent', MyCompanyLibraryLibrary);

  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _currentWebUrl: string;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._currentWebUrl = this._pageContext.web.absoluteUrl;
    });
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    /* eslint-disable @typescript-eslint/no-var-requires */
    
    return [
      {
        name: 'PnP Custom layout (Handlebars)',
        iconName: 'Color',
        key: 'CustomLayoutHandlebars',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        templateContent: require('../custom-layout.html').default.toString(),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomLayoutHandlebars', CustomLayout),
      },
      {
        name: 'PnP Custom layout (Adaptive Cards)',
        iconName: 'Color',
        key: 'CustomLayoutAdaptive',
        type: LayoutType.Results,
        renderType: LayoutRenderType.AdaptiveCards,
        templateContent: JSON.stringify(require('../custom-layout.json'), null, "\t"),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomLayoutAdaptive', CustomLayout),
      },
      {
        name: 'Custom Simple List',
        iconName: 'List',
        key: 'CustomSimpleListLayoutHandlebars',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        templateContent: require('../layouts/results/customSimpleList/custom-simple-list.html').default.toString(),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomSimpleListLayoutHandlebars', CustomSimpleListLayout),
      },
      {
        name: 'Custom People',
        iconName: 'People',
        key: 'CustomPeopleLayoutHandlebars',
        type: LayoutType.Results,
        renderType: LayoutRenderType.Handlebars,
        templateContent: require('../layouts/results/customPeople/custom-people.html').default.toString(),
        serviceKey: ServiceKey.create<ILayout>('PnP:CustomPeopleLayoutHandlebars', CustomPeopleLayout),
      }
    ];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'my-custom-component',
        componentClass: MyCustomComponentWebComponent
      },
      {
        componentName: 'custom-persona',
        componentClass: CustomPersonaWebComponent
      },
      {
        componentName: 'custom-person-card',
        componentClass: CustomPersonCardWebComponent
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

    // Trim text to a max length with ellipsis
    // Usage: {{trim description}}
    namespace.registerHelper('trim', (description?: string) => {
      if (description && description.length > 180) {
        return description.substring(0, 180) + '...';
      }
      return description;
    });

    // Check if text was trimmed and return expanded content
    // Usage: {{{wasTrimmed description}}}
    namespace.registerHelper('wasTrimmed', (description?: string) => {
      if (description && description.length > 180) {
        return new namespace.SafeString(
          '<div class="item-main-content"><div class="metadata">' +
          '<span class="metadata-label">Project Summary:&nbsp;</span>' +
          '<span class="metadata-result">' + namespace.Utils.escapeExpression(description) + '</span>' +
          '</div></div>'
        );
      }
      return null;
    });
  }

  public invokeCardAction(action: any): void {
    
    // Process the action based on type
    if (action.type == "Action.OpenUrl") {
      window.open(action.url, "_blank");
    } else if (action.type == "Action.Submit") {
      // Process the action based on title
      switch (action.title) {

        case 'Click on item':

           // Invoke the currentUser endpoing
           this._spHttpClient.get(
            `${this._currentWebUrl}/_api/web/currentUser`,
            SPHttpClient.configurations.v1, 
            null).then((response: SPHttpClientResponse) => {
              response.json().then((json) => {
                console.log(JSON.stringify(json));
              });
            });

          break;

        case 'Global click':
          alert(JSON.stringify(action.data));
          break;
        default:
          console.log('Action not supported!');
          break;
      }
    }
  }

  public getCustomQueryModifiers(): IQueryModifierDefinition[]
  {
    return [
      {
        name: 'Word Modifier',
        key: 'WordModifier',
        description: 'A demo query modifier from the extensibility library',
        serviceKey: ServiceKey.create<IQueryModifier>('MyCompany:CustomQueryModifier', CustomQueryModifier)

      }
    ];
  
    }
  public getCustomDataSources(): IDataSourceDefinition[] {
    return [
      {
          name: 'NPM Search',
          iconName: 'Database',
          key: 'CustomDataSource',
          serviceKey: ServiceKey.create<IDataSource>('MyCompany:CustomDataSource', CustomDataSource)
      }
    ];
  }

  public name(): string {
    return 'MyCustomLibraryComponent';
  }
}
