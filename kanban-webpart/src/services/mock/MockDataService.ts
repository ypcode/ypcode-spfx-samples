import { ServiceScope } from "@microsoft/sp-core-library";

import { firstOrDefault } from "../../helpers/CollectionHelper";
import {
    IDataService,
    IConfigurationService,
    ConfigurationServiceKey,
    ChoiceFieldType
} from "../";
import { IFieldInfo, IListInfo, ITask } from "../../models";

const mockListsInfo: IListInfo[] = [
    {
        Id: "##1",
        Title: "Tasks List 1",
        Fields: [
            {
                Title: "Title",
                InternalName: "Title",
                TypeAsString: "Text"
            },
            {
                Title: "Task Status",
                InternalName: "Status",
                TypeAsString: ChoiceFieldType,
                Choices: ["Open", "On going", "Done", "Canceled"]
            },
            {
                Title: "Priority",
                InternalName: "Priority",
                TypeAsString: ChoiceFieldType,
                Choices: ["Low", "Medium", "High"]
            }
        ]
    },
    {
        Id: "##2",
        Title: "Tasks List 2",
        Fields: [
            {
                Title: "Title",
                InternalName: "Title",
                TypeAsString: "Text"
            },
            {
                Title: "Status",
                InternalName: "Status",
                TypeAsString: ChoiceFieldType,
                Choices: ["New", "In Progress", "Completed"]
            }
        ]
    }
];

const mockTasks = {
    "##1": [
        {
            Id: 1,
            Title: "Task 1 from list 1",
            Status: "Open",
            Priority: "Low"
        },
        {
            Id: 2,
            Title: "Task 2 from list 1",
            Status: "On going",
            Priority: "Medium"
        },
        {
            Id: 3,
            Title: "Task 3 from list 1",
            Status: "On going",
            Priority: "Medium"
        },
        {
            Id: 4,
            Title: "Task 4 from list 1",
            Status: "Done",
            Priority: "High"
        },
        {
            Id: 5,
            Title: "Task 5 from list 1",
            Status: "Canceled",
            Priority: "Low"
        },
    ],
    "##2": [
        {
            Id: 1,
            Title: "Task 1 from list 2",
            Status: "New"
        },
        {
            Id: 2,
            Title: "Task 2 from list 2",
            Status: "New"
        },
        {
            Id: 3,
            Title: "Task 3 from list 2",
            Status: "In Progress"
        },
        {
            Id: 4,
            Title: "Task 4 from list 2",
            Status: "Canceled"
        },
    ]
};




export class MockDataService implements IDataService {
    private config: IConfigurationService = null;

    constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            // Configure the required dependencies
            this.config = serviceScope.consume(ConfigurationServiceKey);
        });

    }


    public getStatuses(): Promise<string[]> {
        return new Promise<string[]>((resolve) => {
            let list = firstOrDefault(mockListsInfo, ld => ld.Id == this.config.tasksListId);
            if (!list) {
                resolve([]);
                return;
            }

            let choiceField = firstOrDefault(list.Fields, f => f.InternalName == this.config.statusFieldInternalName);
            resolve((choiceField && choiceField.Choices) || []);
        });

    }

    public getAvailableLists(): Promise<IListInfo[]> {
        return new Promise<IListInfo[]>((resolve) => {
            resolve(mockListsInfo);
        });
    }

    public getAvailableChoiceFields(): Promise<IFieldInfo[]> {
        return this.getAvailableLists().then(lists => {
            let list = firstOrDefault(lists, l => l.Id == this.config.tasksListId);
            if (!list)
                return [];

            return list.Fields.filter(f => f.TypeAsString == ChoiceFieldType);
        });
    }

    public getAvailableChoiceFieldsFromLoadedLists() {
        let list = firstOrDefault(mockListsInfo, l => l.Id == this.config.tasksListId);
        if (!list)
            return [];

        return list.Fields.filter(f => f.TypeAsString == ChoiceFieldType);
    }

    public updateTaskStatus(taskId: number, newStatus: string): Promise<any> {
        return new Promise<IListInfo[]>((resolve) => {
            let tasks: ITask[] = mockTasks[this.config.tasksListId];

            // For each task of the list with the Id Task ID (Always only one!), update the status
            tasks.filter(t => t.Id == taskId)
                .forEach(t => t.Status = newStatus);


            resolve();
        });
    }

    public getAllTasks(): Promise<ITask[]> {
        return new Promise<ITask[]>((resolve) => {
            let tasks: ITask[] = mockTasks[this.config.tasksListId]
                .map(t => ({
                    Id: t.Id,
                    Title: t.Title,
                    Status: t[this.config.statusFieldInternalName]
                }));
            resolve(tasks);
        });
    }
}