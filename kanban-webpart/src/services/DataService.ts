
import { ServiceScope, ServiceKey } from "@microsoft/sp-core-library";
import { firstOrDefault } from "../helpers/CollectionHelper";
import { IFieldInfo, IListInfo, ITask } from "../models";
import { IConfigurationService, ConfigurationServiceKey } from "./ConfigurationService";

import pnp from "sp-pnp-js";

export const ChoiceFieldType = "Choice";

export interface IDataService {
    /**
     *  Get the statuses (the available choices) from the specified choice field
     */
    getStatuses(): Promise<string[]>;

    /**
     * Get the available lists in the current web
     */
    getAvailableLists(refresh?: boolean): Promise<IListInfo[]>;

    /**
     * Get the available choice fields for the specified list
     */
    getAvailableChoiceFields(): Promise<IFieldInfo[]>;

    /**
     * Get the available choice fields from lists aleady loaded
     */
    getAvailableChoiceFieldsFromLoadedLists(): IFieldInfo[];

    /**
     * Get all tasks
     */
    getAllTasks(): Promise<ITask[]>;

    /**
     * Update the status to newStatus for the specified task
     */
    updateTaskStatus(taskId: number, newStatus: string): Promise<void>;
}


export default class SharePointDataService implements IDataService {

    private config: IConfigurationService = null;
    private cachedAvailableLists: IListInfo[] = null;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            // Configure the required dependencies
            this.config = serviceScope.consume(ConfigurationServiceKey);
        });

    }

    public getStatuses(): Promise<string[]> {
        return pnp.sp.web.lists.getById(this.config.tasksListId).fields
            .getByInternalNameOrTitle(this.config.statusFieldInternalName)
            .get()
            .then((fieldInfo: IFieldInfo) => fieldInfo.Choices || []);
    }

    public getAvailableLists(refresh: boolean = false): Promise<IListInfo[]> {
        if (!refresh && this.cachedAvailableLists)
            return new Promise<IListInfo[]>(resolve => resolve(this.cachedAvailableLists));

        return pnp.sp.web.lists
            .expand("Fields")
            .select("Id", "Title", "Fields/Title", "Fields/InternalName", "Fields/TypeAsString")
            .get()
            .then(lists => {
                this.cachedAvailableLists = lists;
                return lists;
            });


           
    }

    public getAvailableChoiceFields(): Promise<IFieldInfo[]> {
        return this.getAvailableLists(false).then(lists => {
            let list = firstOrDefault(lists, l => l.Id == this.config.tasksListId);
            if (!list)
                return [];

            return list.Fields.filter(f => f.TypeAsString == ChoiceFieldType);
        });
    }

    public getAvailableChoiceFieldsFromLoadedLists() {
        if (!this.cachedAvailableLists)
            return [];

        let list = firstOrDefault(this.cachedAvailableLists, l => l.Id == this.config.tasksListId);
        if (!list)
            return [];

        return list.Fields.filter(f => f.TypeAsString == ChoiceFieldType);
    }


    public updateTaskStatus(taskId: number, newStatus: string): Promise<any> {
        // Set the value for the configured "status" field
        let fieldsToUpdate = {};
        fieldsToUpdate[this.config.statusFieldInternalName] = newStatus;

        // Update the property on the list item
        return pnp.sp.web.lists.getById(this.config.tasksListId).items.getById(taskId).update(fieldsToUpdate);
    }

    public getAllTasks(): Promise<ITask[]> {
        return pnp.sp.web.lists.getById(this.config.tasksListId).items
            .select("Id", "Title", this.config.statusFieldInternalName)
            .get()
            .then((results: ITask[]) => results && results.map(t => ({
                Id: t.Id,
                Title: t.Title,
                Status: t[this.config.statusFieldInternalName]
            })));
    }
}

export const DataServiceKey = ServiceKey.create<IDataService>("kanban:data-service", SharePointDataService);
