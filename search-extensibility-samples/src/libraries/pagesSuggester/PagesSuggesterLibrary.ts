import { ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import {
  IExtensibilityLibrary,
  ISuggestionProviderDefinition,
  ISuggestionProvider,
  ILayoutDefinition,
  IComponentDefinition,
  IDataSourceDefinition
} from '@aequos/extensibility';
import * as Handlebars from 'handlebars';
import { PageSuggestionProvider } from './PageSuggestionProvider';

export class PagesSuggesterLibrary implements IExtensibilityLibrary {

  public static readonly serviceKey: ServiceKey<PagesSuggesterLibrary> =
    ServiceKey.create<PagesSuggesterLibrary>('SharePointPagesSuggester:ExtensibilityLibrary', PagesSuggesterLibrary);

  private static readonly _suggestionProviderServiceKey: ServiceKey<ISuggestionProvider> =
    ServiceKey.create<ISuggestionProvider>('SharePointPagesSuggester:PageSuggestions', PageSuggestionProvider);

  private _serviceScope: ServiceScope;

  public constructor(serviceScope: ServiceScope) {
    this._serviceScope = serviceScope;
    console.log('[SharePointPagesSuggesterLibrary] Library constructor called');
  }

  public name(): string {
    console.log('[SharePointPagesSuggesterLibrary] name() method called');
    return 'SharePointPagesSuggesterLibrary';
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    console.log('[SharePointPagesSuggesterLibrary] getCustomSuggestionProviders() called');
    const providers = [
      {
        name: PageSuggestionProvider.ProviderName,
        key: 'SharePointPagesSuggester',
        description: PageSuggestionProvider.ProviderDescription,
        serviceKey: PagesSuggesterLibrary._suggestionProviderServiceKey
      }
    ];
    console.log('[SharePointPagesSuggesterLibrary] Returning providers:', providers);
    return providers;
  }

  public getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [];
  }

  public getCustomDataSources(): IDataSourceDefinition[] {
    return [];
  }

  public registerHandlebarsCustomizations(handlebarsNamespace: typeof Handlebars): void {
    // No custom Handlebars helpers
  }
}

export default PagesSuggesterLibrary;
