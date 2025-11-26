import { BaseSuggestionProvider } from '@pnp/modern-search-extensibility';
import { IPropertyPaneGroup } from '@microsoft/sp-property-pane';
import { SPFI, spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/search";
import { ISuggestion } from '@pnp/modern-search-extensibility';

export interface IPageSuggestionProviderProperties {
  maxSuggestions?: number;
  minQueryLength?: number;
}

export class PageSuggestionProvider extends BaseSuggestionProvider<IPageSuggestionProviderProperties> {

  public static readonly ProviderName: string = 'Site Page Suggestions';
  public static readonly ProviderDescription: string = 'Suggests popular pages from Site Page titles and descriptions';

  private _sp: SPFI;
  private readonly MAX_DESCRIPTION_LENGTH = 100;

  public onInit(): void {
    this._sp = spfi().using(SPFx(this.context));
    console.log('[PageSuggestionProvider] Provider initialized');
  }

  private sanitizeQuery(input: string): string {
    if (!input) return '';
    return input.replace(/[():"*<>]/g, ' ').trim();
  }

  private truncateDescription(text: string | undefined, maxLength: number = this.MAX_DESCRIPTION_LENGTH): string {
    if (!text) return '';
    if (text.length <= maxLength) return text;
    return text.substring(0, maxLength).trim() + '...';
  }

  public async getSuggestions(queryText: string): Promise<ISuggestion[]> {
    console.log('[PageSuggestionProvider] getSuggestions called with:', queryText);
    const minLength = this.properties?.minQueryLength || 3;

    if (!queryText || queryText.trim().length < minLength) {
      console.log('[PageSuggestionProvider] Query too short, returning empty');
      return [];
    }

    const sanitizedQuery = this.sanitizeQuery(queryText);
    if (!sanitizedQuery) {
      console.log('[PageSuggestionProvider] Query sanitized to empty, returning empty');
      return [];
    }

    console.log('[PageSuggestionProvider] Sanitized query:', sanitizedQuery);

    try {
      const kqlQuery = `(Title:${sanitizedQuery}* OR Description:${sanitizedQuery}*) AND ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118* AND FileExtension:aspx`;

      const searchResults = await this._sp.search({
        Querytext: kqlQuery,
        RowLimit: 10,
        SelectProperties: [
          "Title",
          "Description",
          "Path",
          "ViewsLastMonths1",
          "ViewsRecent",
          "ViewsLifeTime",
          "LastModifiedTime"
        ],
        TrimDuplicates: true,
        SortList: [
          { Property: "ViewsLastMonths1", Direction: 1 }
        ]
      });

      if (!searchResults?.PrimarySearchResults || searchResults.PrimarySearchResults.length === 0) {
        console.log('[PageSuggestionProvider] No search results found');
        return [];
      }

      console.log('[PageSuggestionProvider] Found', searchResults.PrimarySearchResults.length, 'results');

      const sortedResults = this.sortByPopularity(searchResults.PrimarySearchResults as any[]);
      const maxResults = this.properties?.maxSuggestions || 5;
      const topResults = sortedResults.slice(0, maxResults);

      console.log('[PageSuggestionProvider] Returning', topResults.length, 'suggestions');

      return topResults.map((result: any) => {
        const viewCount = this.getViewCount(result);
        const viewText = viewCount > 0 ? ` (${viewCount} views)` : '';
        const description = this.truncateDescription(result.Description);

        return {
          displayText: result.Title || 'Untitled Page',
          groupName: "Popular Pages",
          description: description ? `${description}${viewText}` : viewText.trim(),
          onSuggestionSelected: () => {
            window.location.href = result.Path;
          }
        };
      });
    } catch (error) {
      console.warn('PageSuggestionProvider: Error fetching suggestions', error);
      return this.getBasicSuggestions(sanitizedQuery);
    }
  }

  public async getZeroTermSuggestions(): Promise<ISuggestion[]> {
    try {
      const kqlQuery = `ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118* AND FileExtension:aspx`;

      const searchResults = await this._sp.search({
        Querytext: kqlQuery,
        RowLimit: 10,
        SelectProperties: [
          "Title",
          "Description",
          "Path",
          "ViewsLastMonths1",
          "ViewsRecent",
          "ViewsLifeTime",
          "LastModifiedTime"
        ],
        TrimDuplicates: true,
        SortList: [
          { Property: "ViewsLastMonths1", Direction: 1 }
        ]
      });

      if (!searchResults?.PrimarySearchResults || searchResults.PrimarySearchResults.length === 0) {
        return [];
      }

      const sortedResults = this.sortByPopularity(searchResults.PrimarySearchResults as any[]);
      const maxResults = this.properties?.maxSuggestions || 5;
      const topResults = sortedResults.slice(0, maxResults);

      const hasViewData = topResults.some((r: any) => this.getViewCount(r) > 0);

      if (!hasViewData) {
        return [];
      }

      return topResults.map((result: any) => {
        const viewCount = this.getViewCount(result);
        const viewText = viewCount > 0 ? `${viewCount} views this month` : 'Recently viewed';

        return {
          displayText: result.Title || 'Untitled Page',
          groupName: "Trending Pages",
          description: viewText,
          onSuggestionSelected: () => {
            window.location.href = result.Path;
          }
        };
      });
    } catch (error) {
      console.warn('PageSuggestionProvider: Error fetching trending pages', error);
      return [];
    }
  }

  private sortByPopularity(results: any[]): any[] {
    if (!results || results.length === 0) {
      return [];
    }

    return [...results].sort((a, b) => {
      const viewsA = this.getViewCount(a);
      const viewsB = this.getViewCount(b);
      return viewsB - viewsA;
    });
  }

  private getViewCount(result: any): number {
    if (result.ViewsLastMonths1 && result.ViewsLastMonths1 > 0) {
      return result.ViewsLastMonths1;
    }
    if (result.ViewsRecent && result.ViewsRecent > 0) {
      return result.ViewsRecent;
    }
    if (result.ViewsLifeTime && result.ViewsLifeTime > 0) {
      return result.ViewsLifeTime;
    }
    return 0;
  }

  private async getBasicSuggestions(sanitizedQuery: string): Promise<ISuggestion[]> {
    try {
      const kqlQuery = `(Title:${sanitizedQuery}* OR Description:${sanitizedQuery}*) AND ContentTypeId:0x0101009D1CB255DA76424F860D91F20E6C4118* AND FileExtension:aspx`;

      const searchResults = await this._sp.search({
        Querytext: kqlQuery,
        RowLimit: this.properties?.maxSuggestions || 5,
        SelectProperties: ["Title", "Description", "Path"],
        TrimDuplicates: true
      });

      if (!searchResults?.PrimarySearchResults || searchResults.PrimarySearchResults.length === 0) {
        return [];
      }

      return searchResults.PrimarySearchResults.map((result: any) => {
        const description = this.truncateDescription(result.Description);

        return {
          displayText: result.Title || 'Untitled Page',
          groupName: "Recommended Pages",
          description: description,
          onSuggestionSelected: () => {
            window.location.href = result.Path;
          }
        };
      });
    } catch (error) {
      console.error('PageSuggestionProvider: All search attempts failed', error);
      return [];
    }
  }

  public getPropertyPaneGroupsConfiguration(): IPropertyPaneGroup[] {
    return [];
  }

  public onPropertyUpdate(propertyPath: string, oldValue: any, newValue: any): void {
    // No custom properties to handle
  }
}
