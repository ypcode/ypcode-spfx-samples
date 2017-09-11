import { ServiceKey} from "@microsoft/sp-core-library";

/**
 * The Configuration Service public interface
 */
export interface IConfigurationService {
   tasksListId: string;
   statusFieldInternalName: string;
}

/**
 * The default implementation of the Configuration service class
 * It is a simple class with 2 public properties that represent the settings in the property pane
 */
export default class ConfigurationService implements IConfigurationService {
    public tasksListId: string;
    public statusFieldInternalName: string;
}

export const ConfigurationServiceKey = ServiceKey.create<IConfigurationService>("kanban:config", ConfigurationService);