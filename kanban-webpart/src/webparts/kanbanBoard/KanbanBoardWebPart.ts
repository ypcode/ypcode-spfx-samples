import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
} from '@microsoft/sp-webpart-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import styles from './KanbanBoard.module.scss';
import * as strings from 'kanbanBoardStrings';
import { IKanbanBoardWebPartProps } from './IKanbanBoardWebPartProps';
import { AppStartup } from "../../startup";

// jQuery and jQuery UI Drag&Drop extensions
import * as $ from "jquery";
require("jquery-ui/ui/widgets/draggable");
require("jquery-ui/ui/widgets/droppable");

// Models
import { ITask, IListInfo, IFieldInfo } from "../../models/";

// Services
import {
  IConfigurationService, ConfigurationServiceKey,
  IDataService, DataServiceKey
} from "../../services";

// Constants
const LAYOUT_MAX_COLUMNS = 12;

export default class KanbanBoardWebPart extends BaseClientSideWebPart<IKanbanBoardWebPartProps> {

  private statuses: string[] = [];
  private tasks: ITask[] = [];
  private availableLists: IListInfo[] = [];
  private availableChoiceFields: IFieldInfo[] = [];
  private dataService: IDataService = null;
  private config: IConfigurationService = null;

  constructor() {
    super();
    SPComponentLoader.loadCss('https://code.jquery.com/ui/1.12.1/themes/base/jquery-ui.min.css');
  }

  public onInit(): Promise<any> {
    return super.onInit()
      // Set the global configuration of the application
      // This is where we will define the proper services according to the context (Local, Test, Prod,...) 
      // or according to specific settings
      .then(_ => AppStartup.configure(this.context, this.properties))
      // When configuration is done, we get the instances of the services we want to use
      .then(serviceScope => {
        this.dataService = serviceScope.consume(DataServiceKey);
        this.config = serviceScope.consume(ConfigurationServiceKey);
      })
      // Then, we are able to use the methods of the services
      // Load the available lists in the current site
      .then(() => this.dataService.getAvailableLists())
      .then((lists: IListInfo[]) => this.availableLists = lists);
  }



  public render(): void {

    // Only if properly configured
    if (this.properties.statusFieldName && this.properties.tasksListId) {
      // Load the data
      this.dataService.getStatuses()
        .then((statuses: string[]) => this.statuses = statuses)
        .then(() => this.dataService.getAllTasks())
        .then((tasks: ITask[]) => this.tasks = tasks)
        // And then render
        .then(() => {
          this.domElement.innerHTML = this.renderBoard();
          this.enableDragAndDrop();
        })
        .catch(error => {
          console.log(error);
          console.log("An error occured while loading data...");
        });
    } else {
      this.domElement.innerHTML = `<div class="ms-MessageBar ms-MessageBar--warning">${strings.PleaseConfigureWebPartMessage}</div>`;
    }
  }

  private _getColumnSizeClassName(columnsCount: number): string {
    if (columnsCount < 1) {
      console.log("Invalid number of columns");
      return "";
    }

    if (columnsCount > LAYOUT_MAX_COLUMNS) {
      console.log("Too many columns for responsive UI");
      return "";
    }

    let columnSize = Math.floor(LAYOUT_MAX_COLUMNS / columnsCount);

    return "ms-u-sm" + (columnSize || 1);
  }

  /**
   * Generates and inject the HTML of the Kanban Board
   */
  public renderBoard(): string {
    let columnSizeClass = this._getColumnSizeClassName(this.statuses.length);

    // The begininning of the WebPart
    let html = `
      <div class="${styles.kanbanBoard}">
        <div class="${styles.container}">
          <div class="ms-Grid-row ${styles.row}">`;

    // For each status
    this.statuses.forEach(status => {
      // Append a new Office UI Fabric column with the appropriate width to the row
      html += `<div class="${styles.kanbanColumn} ms-Grid-col ${columnSizeClass}" data-status="${status}">
                <h3 class="ms-fontColor-themePrimary">${status}</h3>`;
      // Get all the tasks in the current status
      let currentGroupTasks = this.tasks.filter(t => t.Status == status);
      // Add a new tile for each task in the current column
      currentGroupTasks.forEach(task => {
        html += `<div class="${styles.task}" data-taskid="${task.Id}">
          <div class="ms-fontSize-xl">${task.Title}</div></div>`;
      });
      // Close the column element
      html += `</div>`;
    });

    // Ends the WebPart HTML
    html += `</div></div></div>`;

    return html;
  }

  private enableDragAndDrop() {
    // Make the kanbanColumn elements droppable areas
    let webpart = this;
    $(this.domElement).find(`.${styles.kanbanColumn}`).droppable({
      tolerance: "intersect",
      accept: `.${styles.task}`,
      activeClass: "ui-state-default",
      hoverClass: "ui-state-hover",
      drop: (event, ui) => {
        // Here the code to execute whenever an element is dropped into a column
        let taskItem = $(ui.draggable);
        let source = taskItem.parent();
        let previousStatus = source.data("status");
        let taskId = taskItem.data("taskid");
        let target = $(event.target);
        let newStatus = target.data("status");
        taskItem.appendTo(target);

        // If the status has changed, apply the changes
        if (previousStatus != newStatus) {
          webpart.dataService.updateTaskStatus(taskId, newStatus);
        }
      }
    });

    // Make the task items draggable
    $(this.domElement).find(`.${styles.task}`).draggable({
      classes: {
        "ui-draggable-dragging": styles.dragging
      },
      opacity: 0.7,
      helper: "clone",
      cursor: "move",
      revert: "invalid"
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('2.0');
  }

  public onPropertyPaneFieldChanged(propertyName: string, oldValue: string, newValue: string) {
    this.config.statusFieldInternalName = this.properties.statusFieldName;
    this.config.tasksListId = this.properties.tasksListId;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.HeaderDescription
          },
          groups: [
            {
              groupName: strings.TasksConfigurationGroup,
              groupFields: [
                PropertyPaneDropdown('tasksListId', {
                  label: strings.SourceTasksList,
                  options: this.availableLists.map(l => ({
                    key: l.Id,
                    text: l.Title
                  }))
                }),
                PropertyPaneDropdown('statusFieldName', {
                  label: strings.StatusFieldInternalName,
                  options: this.dataService.getAvailableChoiceFieldsFromLoadedLists().map(f => ({
                    key: f.InternalName,
                    text: f.Title
                  }))
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
