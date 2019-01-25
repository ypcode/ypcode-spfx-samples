
import { BaseComponentContext } from "@microsoft/sp-component-base";
import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";

interface IServiceConfiguration {
    serviceKey: any;
    config: (serviceInstance: any) => void;
}

export class ComponentServices {
    constructor(private serviceScope: ServiceScope) {

    }

    public registerScopedServiceInstance<TService>(serviceKey: ServiceKey<TService>, instance: TService): TService {
        this.serviceScope.provide(serviceKey, instance);
        return instance;
    }

    public registerScopedService<TService>(serviceKey: ServiceKey<TService>, serviceType: new (serviceScope: ServiceScope) => any): TService {
        return this.serviceScope.createAndProvide(serviceKey, serviceType);
    }

    private _serviceConfigurations: { [id: string] : IServiceConfiguration} = {};
    public configureService<TService>(serviceKey: ServiceKey<TService>, config: (serviceInstance: TService) => void) {
        this._serviceConfigurations[serviceKey.id] = {
            serviceKey,
            config
        };
    }

    public static init<TProps>(componentContext: BaseComponentContext,
        properties: TProps,
        configureServices: (startup: ComponentServices, ctx: BaseComponentContext, props: TProps) => Promise<void>): Promise<ServiceScope> {

        console.log('ComponentContext: ', componentContext);
        console.log('Properties: ', properties);

        if (!configureServices) {
            return Promise.resolve();
        }

        return new Promise((resolve, reject) => {
            try {
                const childScope = componentContext.serviceScope.startNewChild();
                const startup = new ComponentServices(childScope);
                configureServices(startup, componentContext, properties)
                    .then(() => {
                        console.log('Services are configured.');
                        childScope.finish();
                        childScope.whenFinished(() => {
                            ComponentServices.serviceScope = childScope;

                            for (let k in startup._serviceConfigurations) {
                                const configHandle = startup._serviceConfigurations[k];
                                const serviceInstance = childScope.consume(configHandle.serviceKey);
                                configHandle.config(serviceInstance);
                            }

                            resolve(childScope);
                        });
                    }).catch(err => reject(err));
            } catch (error) {
                reject(error);
            }
        });
    }

    public static serviceScope: ServiceScope;
}