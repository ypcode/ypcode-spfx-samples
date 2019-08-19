import { ServiceScope } from "@microsoft/sp-core-library";
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ComponentContextServiceKey, ComponentContextService } from "../services/ComponentContextService";
import { DocumentsServiceKey, DocumentsService } from "../services/DocumentsService";
import { PageContextServiceKey } from "../services/PageContextService";
import { ISpContextServiceLabWebPartProps } from "../webparts/spContextServiceLab/SpContextServiceLabWebPart";

export const configure = (componentContext: BaseComponentContext, properties: ISpContextServiceLabWebPartProps): Promise<ServiceScope> => {
    const rootScope = componentContext.serviceScope;

    return new Promise((resolve, reject) => {
        try {
            // The default implementation of all services (built-in AND custom) are available at root scope
            // We should be extremely cautious of altering a root-scoped service 'state' from a specific component instance
            // This might be not that important in the context of an app-part page        
            // All services directly usable from root scope should not have any component-specific dependencies
            const pageContextService = rootScope.consume(PageContextServiceKey);
            pageContextService.configure(componentContext);

            const scopedService = rootScope.startNewChild();
            // TODO Here create and initialize the component scoped custom service instances
            // TODO Initialize and configure scoped services based on component configuration

            // The component-scoped context should be created here to ensure it will remain tied to the proper instance
            const componentContextService = scopedService.createAndProvide(ComponentContextServiceKey, ComponentContextService);
            componentContextService.configure(componentContext, properties);

            // Create and provide new instances of services that uses component specific context (configuration, instance id, ...)
            // (e.g. In this example the Documents service relies of the component configuration)
            scopedService.createAndProvide(DocumentsServiceKey, DocumentsService);

            // Finish the child scope initalization
            scopedService.finish();

            resolve(scopedService);

        } catch (error) {
            reject(error);
        }
    });
};