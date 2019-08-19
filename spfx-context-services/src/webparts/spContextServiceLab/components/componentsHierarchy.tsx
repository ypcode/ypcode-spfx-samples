import * as React from "react";
import { ServiceScope } from "@microsoft/sp-core-library";
import { ListServiceKey } from "../../../services/ListsService";
import { IList } from "../../../models/IList";
import { DocumentsServiceKey } from "../../../services/DocumentsService";
import { ComponentContextServiceKey } from "../../../services/ComponentContextService";
import styles from "./component.module.scss";

export interface IComponentProps {
    serviceScope: ServiceScope;
}


export class FourthLevelSubComponent extends React.Component<IComponentProps, any> {

    constructor(props: IComponentProps) {
        super(props);

        this.state = {
            listInfo: null as IList,
            documentsCount: null,
        };
    }

    private get instanceId(): string {
        return this.props.serviceScope.consume(ComponentContextServiceKey).instanceId;
    }

    public componentWillMount() {

        const listService = this.props.serviceScope.consume(ListServiceKey);
        const componentContext = this.props.serviceScope.consume(ComponentContextServiceKey);

        listService.getListByTitle(componentContext.properties.documentLibraryName)
            .then(docLib => {
                this.setState({
                    listInfo: docLib
                });
            }).catch(error => {
                console.log("Error: ", error);
            });
    }

    private _getDocumentsCount() {
        const documentsService = this.props.serviceScope.consume(DocumentsServiceKey);
        documentsService.getDocumentsCount()
            .then(itemsCount => {
                this.setState({
                    documentsCount: itemsCount
                });
            });
    }

    public render() {
        if (!this.state.listInfo) {
            return <div>Loading...</div>;
        }

        return <div className={styles.component}>
            <div className={styles.title}>
                WebPart: <br/>
                <span className={styles.instanceIdFromService}>
                    {this.instanceId}
                </span>
            </div>
            <br />
            <br />
            <h3>Loaded list</h3>
            ID: {this.state.listInfo.id}<br />Title: {this.state.listInfo.title}
            <br /><br />
            <button onClick={() => this._getDocumentsCount()}>Get count of documents</button>
            <br />
            {(this.state.documentsCount || this.state.documentsCount == 0) && <div>Count: {this.state.documentsCount}</div>}
        </div>;
    }
}

export class ThirdLevelSubComponent extends React.Component<IComponentProps, any> {
    public render() {
        return <FourthLevelSubComponent serviceScope={this.props.serviceScope} />;
    }
}


export class SecondLevelSubComponent extends React.Component<IComponentProps, any> {
    public render() {
        return <ThirdLevelSubComponent serviceScope={this.props.serviceScope} />;
    }

}


export class FirstLevelSubComponent extends React.Component<IComponentProps, any> {
    public render() {
        return <SecondLevelSubComponent serviceScope={this.props.serviceScope} />;
    }
}