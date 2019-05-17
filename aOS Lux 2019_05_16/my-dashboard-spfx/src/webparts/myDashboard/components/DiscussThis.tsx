import * as React from 'react';
import { IconButton } from "office-ui-fabric-react/lib/Button";
import { MSGraphClientFactory, MSGraphClient } from '@microsoft/sp-http';
import * as microsoftTeams from '@microsoft/teams-js';

export interface IDiscussThisState {
    statusMessage: string;
}

export interface IDiscussThisProps {
    graphClientFactory: MSGraphClientFactory;
    teamsContext: microsoftTeams.Context;
    subject: string;
    messagePattern: string;
}

export class DiscussThis extends React.Component<IDiscussThisProps, IDiscussThisState> {

    constructor(props: IDiscussThisProps) {
        super(props);

        this.state = {
            statusMessage: ''
        };
    }

    public render(): React.ReactElement<IDiscussThisProps> {
        if (!this.props.teamsContext) {
            return <div />;
        }

        return <IconButton onClick={this._onDiscussButtonClick.bind(this)} text="Discuss this" iconProps={{ iconName: 'Chat' }} />;
    }

    private _onDiscussButtonClick() {
        this.sendChannelMessage(this.props.subject, this.props.messagePattern);
    }

    private sendChannelMessage(subject: string, message: string) {
        this.props.graphClientFactory.getClient()
            .then((client: MSGraphClient) => client.api(`teams/${this.props.teamsContext.groupId}/channels/${this.props.teamsContext.channelId}/messages`).version('beta').post({
                'subject': subject,
                'body': {
                    'contentType': 'html',
                    'content': message
                }
            }
            ))
            .then(res => {

                console.log("Channel message posted");
                this.setState({
                    statusMessage: 'Discussion started !'
                });
            }).catch(err => {
                console.log("Cannot start conversation", err);
                this.setState({
                    statusMessage: 'Cannot start Discussion'
                });
            });
    }

}