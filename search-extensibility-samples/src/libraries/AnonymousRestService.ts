import { Log, ServiceKey, ServiceScope } from '@microsoft/sp-core-library';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
const SearchService_ServiceKey = 'pnpModernSearchExtensibilitySamples:AnonymousRestService';

export class AnonymousRestService {
  public static ServiceKey: ServiceKey<AnonymousRestService> = ServiceKey.create(SearchService_ServiceKey, AnonymousRestService);

  /**
   * The current service scope
   */
  private serviceScope: ServiceScope;

  /**
   * The HttpClient instance
   */
  private httpClient: HttpClient;

  constructor(serviceScope: ServiceScope) {
    this.serviceScope = serviceScope;
    serviceScope.whenFinished(async () => {
      this.httpClient = serviceScope.consume<HttpClient>(HttpClient.serviceKey);
    });
  }

  public async requestData(url:string, method:'GET'|'POST', body?:string): Promise<any> {

    let response = null;

    switch (method) {
      case 'GET':
        {
          response = await this.httpClient.get(url, HttpClient.configurations.v1, {
            method: 'GET',
            headers: new Headers({
              Accept: 'application/json',
            }),

            mode: 'cors',
          });
        }
        break;

      case 'POST':
        {
          response = await this.httpClient.post(url, HttpClient.configurations.v1, {
            method: 'POST',
            headers: new Headers({
              Accept: 'application/json',

            }),
            mode: 'cors',
            body:body  
          });
        }
        break;
    }

    return this.processSearchResult(response);
  }

  /**
   * Processes the generic api client response
   * @param response
   * @returns
   */
  protected processSearchResult(response: HttpClientResponse): Promise<any> {
    return response.text().then((_responseText) => {
      if (response.status === 200) {
        return _responseText === '' ? null : JSON.parse(_responseText);
      }

      const error = new Error(_responseText);
      Log.error('[processSearchResult]', error, this.serviceScope);
      return error;
    });
  }
}
