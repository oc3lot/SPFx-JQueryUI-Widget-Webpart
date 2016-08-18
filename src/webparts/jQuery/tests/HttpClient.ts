// Setup mock Http client
import { ISPList } from '../JQueryWebPart';

export default class HttpClient {

    private static _items: ISPList[] = [{ Title: 'Mock List', Id: '1' }];

    public static get(restUrl: string, options?: any): Promise<ISPList[]> {
      return new Promise<ISPList[]>((resolve) => {
            resolve(HttpClient._items);
        });
    }
}