import { IList } from "../models/IList";
import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";
import { ComponentContextServiceKey } from "./ComponentContextService";
import { SPHttpClient } from "@microsoft/sp-http";
import { PageContextServiceKey } from "./PageContextService";

export interface IListService {
    getListByTitle(listTitle: string): Promise<IList>;
}

export class ListService implements IListService {

    constructor(private serviceScope: ServiceScope) {
    }

    public getListByTitle(listTitle: string): Promise<IList> {

        return new Promise<IList>((resolve, reject) => {
            // Ensure the service scope is completely configured before we can consume any service
            this.serviceScope.whenFinished(() => {
                const pageContext = this.serviceScope.consume(PageContextServiceKey);
                const spHttpClient = this.serviceScope.consume(SPHttpClient.serviceKey);
                const url = `${pageContext.webAbsoluteUrl}/_api/web/lists/getbytitle('${escape(listTitle)}')`;
                spHttpClient.get(url, SPHttpClient.configurations.v1)
                    .then(r => r.json())
                    .then(l => {
                        resolve({
                            id: l.Id,
                            title: l.Title,
                            itemsCount: l.ItemCount
                        } as IList);
                    });
            });
        });
    }
}

export const ListServiceKey = ServiceKey.create<IListService>("ypcode::ListService", ListService);