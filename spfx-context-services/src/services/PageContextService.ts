import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";

export interface IPageContextService {
    configure(spfxComponentContext: BaseComponentContext): void;
    webAbsoluteUrl: string;
    // TODO Expose any needed page scoped property
    // NOTE Avoid simply exposing the whole page context
    // Exposing only the needed information in that service allows to have better control
    // and better understanding of what's really page specific or not
    // It also mitigates risk of unexpected behavior in OTB API
}

export class PageContextService implements IPageContextService {

    private _webAbsoluteUrl: string;
    private _configured: boolean = false;

    constructor(private serviceScope: ServiceScope) {

    }

    public configure(spfxComponentContext: BaseComponentContext): void {
       // Note this service should be reconfigured at each call of the configure() method
       // Because the SPA context of SharePoint might involve changes in page context

        if (!spfxComponentContext) {
            throw new Error("The SPFx component context is not specified.");
        }


        const webAbsoluteUrlFromContext = spfxComponentContext.pageContext && spfxComponentContext.pageContext.web && spfxComponentContext.pageContext.web.absoluteUrl;
        if (webAbsoluteUrlFromContext && webAbsoluteUrlFromContext != this._webAbsoluteUrl) {
            this._webAbsoluteUrl = webAbsoluteUrlFromContext;
        }

        this._configured = (this._webAbsoluteUrl && true) || false;
    }

    public get webAbsoluteUrl(): string {
        if (!this._configured) {
            throw new Error("The Page Context Service has not been properly configured.");
        }

        return this._webAbsoluteUrl;
    }
}

export const PageContextServiceKey = ServiceKey.create<IPageContextService>("ypcode::PageContextService", PageContextService);