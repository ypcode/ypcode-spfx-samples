import { Version, ServiceScope } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './ItemsCounterWebPart.module.scss';
import * as strings from 'ItemsCounterWebPartStrings';
import { ListService, IListData, IListService } from '../../services/ListService';
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from '@pnp/spfx-property-controls/lib/PropertyFieldListPicker';
import { ComponentServices } from '../../startup/ComponentServices';

export interface IItemsCounterWebPartProps {
  description: string;
  listId: string;
}

export default class ItemsCounterWebPart extends BaseClientSideWebPart<IItemsCounterWebPartProps> {

  private listData: IListData = null;
  private isLoading: boolean = true;
  private error: string = null;

  private listDataService: IListService;

  public onInit(): Promise<void> {
    return ComponentServices.init(this.context,
      this.properties, (startup, ctx, props) => {

        // // Register a new scoped instance of the service
        startup.registerScopedService(ListService.serviceKey, ListService);
        // Configure the service instance with the component specific properties
        startup.configureService(ListService.serviceKey, service => {
          service.configure(ctx.pageContext.web.absoluteUrl, props.listId);
        });
      
        // Must return a resolved promise 
        // (useless here but needed in case on async needs in the config process)
        return Promise.resolve();
      }).then(serviceScope => {

        // Consume the list service
        // Instead of keeping a service reference assigned here,
        // one can use ComponentServices.serviceScope.consume(ListService.serviceKey);
        this.listDataService = serviceScope.consume(ListService.serviceKey);
        this._refresh();
      }).catch(error => {
        console.log('Error on init: ', error);
      });
  }

  public onPropertyPaneFieldChanged() {
    this.listDataService.configure(this.context.pageContext.web.absoluteUrl, this.properties.listId);
    this._refresh();
  }

  private _refresh() {
    this.isLoading = true;
    this.render();

    console.log('Loading data...');
    this.listDataService.getListData().then(list => {
      console.log('Data fetched: ', list);
      this.listData = list;
      this.isLoading = false;
      this.render();
    }).catch(error => {
      this.error = error;
      this.render();
    });
  }

  public render(): void {

    if (this.isLoading) {
      this.domElement.innerHTML = `<div>Loading...</div>`;
      return;
    }

    if (this.error) {
      this.domElement.innerHTML = `<div class='error'>${this.error}</div>`;
      return;
    }

    if (!this.listData) {
      this.domElement.innerHTML = `<div>No data...</div>`;
      return;
    }
    this.domElement.innerHTML = `
      <div class="${ styles.itemsCounter}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
             <h1>${this.listData.ItemsCount} ${this.listData.Title}</h1>
             <button class="btnRefresh">Refresh</button>
            </div>
          </div>
        </div>
      </div>`;

    this.domElement.querySelector("button.btnRefresh").addEventListener("click", () => {
      this._refresh();
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyFieldListPicker('listId', {
                  label: 'Select a list',
                  selectedList: this.properties.listId,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  multiSelect: false,
                  key: 'listPickerFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
