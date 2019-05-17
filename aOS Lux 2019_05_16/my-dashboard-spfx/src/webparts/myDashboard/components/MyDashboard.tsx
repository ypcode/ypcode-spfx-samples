import * as React from 'react';
import * as moment from "moment";
import { List } from "office-ui-fabric-react/lib/List";
import { Icon } from "office-ui-fabric-react/lib/Icon";
import styles from './MyDashboard.module.scss';
import { MSGraphClientFactory, MSGraphClient } from '@microsoft/sp-http';
import * as microsoftTeams from '@microsoft/teams-js';
import { IDocument } from '../../../models/IDocument';
import { IEmail } from '../../../models/IEmail';
import { IEvent } from '../../../models/IEvent';
import { DiscussThis } from './DiscussThis';

export interface IMyDashboardState {
  myUpcomingEvents: IEvent[];
  eventsError: string;
  myRecentEmails: IEmail[];
  emailsError: string;
  myRecentDocuments: IDocument[];
  documentsError: string;
  debug: string;
}

export interface IMyDashboardProps {
  title: string;
  subTitle: string;
  graphClientFactory: MSGraphClientFactory;
  teamsContext: microsoftTeams.Context;
  showDebug: boolean;
}

export default class MyDashboard extends React.Component<IMyDashboardProps, IMyDashboardState> {

  constructor(props: IMyDashboardProps) {
    super(props);

    this.state = {
      myRecentDocuments: [],
      myRecentEmails: [],
      myUpcomingEvents: [],
      documentsError: null,
      eventsError: null,
      emailsError: null,
      debug: null
    };
  }

  public componentWillMount() {
    this._loadDataFromMicrosoftGraph();
  }

  private _debug(message: string, error?: any) : void {
    console.log(message, error);
    let debugMessage = `${message} : ${error && JSON.stringify(error)}`;
    if (this.props.showDebug) {
      this.setState({
        debug: debugMessage
      });
    }
  }

  public render(): React.ReactElement<IMyDashboardProps> {
    let { myRecentDocuments, myRecentEmails, myUpcomingEvents, documentsError, eventsError, emailsError } = this.state;
    return (
      <div className={styles.myDashboard}>
        <div className={styles.container}>
          {this.state.debug && <div className={styles.debug}>{this.state.debug}</div>}
          <div>
            <h1>{this.props.title}</h1>
            <h2>{this.props.subTitle}</h2>
          </div>
          <div className={styles.row}>
            <div className={styles.column}>
              <h1>My recent e-mails</h1>
              <List items={myRecentEmails} onRenderCell={this._onRenderEmail.bind(this)} />
              {emailsError && `ERROR: ${emailsError}`}
            </div>
            <div className={styles.column}>
              <h1>My recent documents</h1>
              <List items={myRecentDocuments} onRenderCell={this._onRenderDocument.bind(this)} />
              {documentsError && `ERROR: ${documentsError}`}
            </div>
            <div className={styles.column}>
              <h1>Upcoming events</h1>
              <List items={myUpcomingEvents} onRenderCell={this._onRenderEvent.bind(this)} />
              {eventsError && `ERROR: ${eventsError}`}
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _onRenderEmail(item: IEmail, index: number | undefined): JSX.Element {
    return (
      <div className={`${styles.emailItem} ${item.isRead ? '' : styles.unread}`} data-is-focusable={true}>
        <div>
          <div className={styles.subject}>
            <Icon iconName={item.isRead ? 'Read' : 'Mail'} className={styles.icon} />
            {item.subject}
          </div>
          <div>{item.bodyPreview}</div>
          <DiscussThis
            graphClientFactory={this.props.graphClientFactory}
            teamsContext={this.props.teamsContext}
            subject={item.subject}
            messagePattern={`Let's discuss this e-mail: <br/><h2>${item.subject}</h2><br/><div>${item.bodyPreview}</div>`} />
        </div>
      </div>
    );
  }

  private _onRenderDocument(item: IDocument, index: number | undefined): JSX.Element {
    return (<div className={styles.documentItem}>
      <Icon iconName="Document" className={styles.icon} />
      <a href={item.webUrl}>{item.name}</a>
      <DiscussThis
            graphClientFactory={this.props.graphClientFactory}
            teamsContext={this.props.teamsContext}
            subject={item.name}
            messagePattern={`Let's discuss this document: <br/><a href="${item.webUrl}">${item.name}</a>`} />
 
    </div>);
  }

  private _onRenderEvent(item: IEvent, index: number | undefined): JSX.Element {
    return (<div className={styles.eventItem}>

      <h3 className={styles.subject}> <Icon iconName="Event" className={styles.icon} /> {item.subject}</h3>
      <h4>{item.location} - {moment(item.when).format('LLL')}</h4>
      <DiscussThis
            graphClientFactory={this.props.graphClientFactory}
            teamsContext={this.props.teamsContext}
            subject={item.subject}
            messagePattern={`Let's discuss this event: <br/><h2>${item.subject}</h2><br/><div>${item.when}</div>`} />
  
    </div>);
  }



  private _loadDataFromMicrosoftGraph() {
    if (!this.props.graphClientFactory) {
      this._debug('No specified Graph client factory...');
      return;
    }

    this._debug("Loading recent documents...");
    // Load recent documents
    this.props.graphClientFactory.getClient()
      .then((client: MSGraphClient) => client.api('drive/recent').version('v1.0').top(3).get())
      .then(res => {
        let recentDocuments = res.value.map(item => ({
          name: item.name,
          lastModifiedDateTime: item.lastModifiedDateTime,
          webUrl: item.webUrl
        }) as IDocument);

        this._debug("Loaded recent documents");
        this.setState({
          myRecentDocuments: recentDocuments
        });
      }).catch(err => {
        this._debug("Error while loading recent documents...", err);
        this.setState({
          documentsError: err.message
        });
      });

    // Load upcoming events
    this._debug("Loading upcoming events...");
    this.props.graphClientFactory.getClient()
      .then((client: MSGraphClient) => {

        const today = new Date();

        return client.api('me/events').version('v1.0').filter(`start/dateTime ge '${today.toISOString()}'`).top(3).get();

      }).then(res => {

        let upcomingEvents = res.value.map(item => ({
          subject: item.subject,
          when: item.start.dateTime,
          location: item.location.displayName
        }) as IEvent);
        this._debug("Loaded upcoming events");
        this.setState({
          myUpcomingEvents: upcomingEvents
        });
      }).catch(err => {
        this._debug("Error while loading upcoming events...", err);
        this.setState({
          eventsError: err.message
        });
      });

    // Load recent e-mails
    this._debug("Loading recent e-mails...");
    this.props.graphClientFactory.getClient()
      .then((client: MSGraphClient) => client.api('me/messages').version('v1.0').top(3).get())
      .then(res => {

        let recentEmails = res.value.map(item => ({
          subject: item.subject,
          bodyPreview: item.bodyPreview,
          receivedOn: item.receivedOn,
          isRead: item.isRead
        }) as IEmail);
        this._debug("Loaded recent e-mails");
        this.setState({
          myRecentEmails: recentEmails
        });
      }).catch(err => {
        this._debug("Error while loading recent e-mails", err);
        this.setState({
          emailsError: err.message
        });
      });
  }
}
