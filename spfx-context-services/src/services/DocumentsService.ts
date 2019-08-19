import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { ListServiceKey } from "./ListsService";
import { ComponentContextServiceKey } from "./ComponentContextService";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentsService {
    getDocumentsCount(): Promise<number>;
}

export class DocumentsService implements IDocumentsService {
    constructor(private serviceScope: ServiceScope) { }

    public getDocumentsCount(): Promise<number> {
        return new Promise<number>((resolve, reject) => {

            // Ensure the service scope is completely configured before we can consume any service
            this.serviceScope.whenFinished(() => {
                const listService = this.serviceScope.consume(ListServiceKey);
                const componentContextService = this.serviceScope.consume(ComponentContextServiceKey);
                const docLibName = componentContextService.properties.documentLibraryName;
                listService.getListByTitle(docLibName).then(list => {
                    resolve(list.itemsCount);
                });
            });
        });
    }


}

export const DocumentsServiceKey = ServiceKey.create<IDocumentsService>("ypcode::DocumentsService", DocumentsService);