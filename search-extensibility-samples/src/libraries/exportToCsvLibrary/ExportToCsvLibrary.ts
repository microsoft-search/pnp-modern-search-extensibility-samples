import {
  IExtensibilityLibrary,
  IComponentDefinition,
  ISuggestionProviderDefinition,
  ILayoutDefinition
} from "@pnp/modern-search-extensibility";
import { ExportWebComponent } from "./components/ExportComponent";

export class ExportToCsvLibrary implements IExtensibilityLibrary {

  public getCustomLayouts(): ILayoutDefinition[] {
    return [];
  }

  public getCustomWebComponents(): IComponentDefinition<any>[] {
    return [
      {
        componentName: 'pnp-export',
        componentClass: ExportWebComponent
      }
    ];
  }

  public getCustomSuggestionProviders(): ISuggestionProviderDefinition[] {
    return [];
  }

  public registerHandlebarsCustomizations(namespace: typeof Handlebars) {
  }
}
