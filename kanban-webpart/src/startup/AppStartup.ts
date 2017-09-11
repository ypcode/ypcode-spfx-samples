
import {
    IWebPartContext,
} from '@microsoft/sp-webpart-base';
import {
    Environment,
    EnvironmentType,
    ServiceScope
} from '@microsoft/sp-core-library';

// sp-pnp-js for SPFx context configuration
import pnp from "sp-pnp-js";

// Services
import {
    ConfigurationServiceKey, DataServiceKey,
    IConfigurationService, MockDataService
} from "../services";


export class AppStartup {
    private static configured: boolean = false;
    private static serviceScope: ServiceScope = null;

    public static configure(ctx: IWebPartContext, props: any): Promise<ServiceScope> {

        switch (Environment.type) {
            case EnvironmentType.SharePoint:
            case EnvironmentType.ClassicSharePoint:
                return AppStartup.configureForSharePointContext(ctx, props);
            // case EnvironmentType.Local:
            // case EnvironmentType.Test:
            default:
                return AppStartup.configureForLocalOrTestContext(ctx, props);

        }
    }

    public static getServiceScope(): ServiceScope {
        if (AppStartup.configured)
            throw new Error("The application is not properly configured");

        return AppStartup.serviceScope;
    }

    private static configureForSharePointContext(ctx: IWebPartContext, props: any): Promise<ServiceScope> {
        return new Promise<any>((resolve, reject) => {
            ctx.serviceScope.whenFinished(() => {
                // Configure PnP Js for working seamlessly with SPFx
                pnp.setup({
                    spfxContext: ctx
                });

                let config: IConfigurationService = ctx.serviceScope.consume(ConfigurationServiceKey);

                // Initialize the config with WebPart Properties
                config.statusFieldInternalName = props["statusFieldName"];
                config.tasksListId = props["tasksListId"];

                AppStartup.serviceScope = ctx.serviceScope;
                AppStartup.configured = true;

                resolve(ctx.serviceScope);
            });
        });
    }

    private static configureForLocalOrTestContext(ctx: IWebPartContext, props: any): Promise<ServiceScope> {
        return new Promise<any>((resolve, reject) => {
            // Here create a dedicated service scope for test or local context
            const childScope: ServiceScope = ctx.serviceScope.startNewChild();
            // Register the services that will override default implementation
            childScope.createAndProvide(DataServiceKey, MockDataService);
            // Must call the finish() method to make sure the child scope is ready to be used
            childScope.finish();

            childScope.whenFinished(() => {
                // If other services must be used, it must done HERE

                AppStartup.serviceScope = childScope;
                AppStartup.configured = true;
                resolve(childScope);
            });
        });
    }
}