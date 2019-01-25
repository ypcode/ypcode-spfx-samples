import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";
import { SPHttpClient } from "@microsoft/sp-http";

/**
 * The List service contract
 */
export interface IListService {
    configure(webUrl: string, listId: string);
    getListData(): Promise<IListData>;
}

export interface IListData {
    Title: string;
    ItemsCount: number;
}

const SERVICE_KEY_TOKEN = "ListService";

/**
 * List service default implementation
 */
export class ListService implements IListService {

    private _listId: string;
    private _webUrl: string;

    constructor(private serviceScope: ServiceScope) {}

    /**
     * Set the configuration of the service
     * @param webUrl The URL of the SharePoint web
     * @param listId THe ID of the list to work on
     */
    public configure(webUrl: string, listId: string) {
        this._webUrl = webUrl;
        this._listId = listId;
    }    

    /**
     * Gets basic information about the configured SharePoint list
     */
    public getListData(): Promise<IListData> {
        const apiUrl = `${this._webUrl}/_api/web/lists(guid'${this._listId}')?$select=Title,ItemCount`;
        const client = this.serviceScope.consume(SPHttpClient.serviceKey);
        return client.get(apiUrl, SPHttpClient.configurations.v1)
        .then(r => r.json())
        .then(r => ({
            Title: r.Title,
            ItemsCount: r.ItemCount
        } as IListData));
    }

    public static serviceKey: ServiceKey<IListService> = ServiceKey.create(SERVICE_KEY_TOKEN, ListService);
}
