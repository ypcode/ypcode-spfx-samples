import * as React from "react";
import { TeachingBubble, autobind, Button, Link, IButtonProps } from "office-ui-fabric-react";

import { ITutorialItem } from "../model/ITutorialItem";
import { ITutorial } from "../model/ITutorial";
import playTutorial, { shouldPlayTutorial } from "../core/TutorialCore";

export interface ITutorialItemDisplayState {

}

export interface ITutorialItemDisplayProps {
    targetElement: HTMLElement;
    tutorialItem: ITutorialItem;
    next?: (item: ITutorialItem) => void;
}



export class TutorialItemDisplay extends React.Component<ITutorialItemDisplayProps, ITutorialItemDisplayState> {
    constructor(props: ITutorialItemDisplayProps) {
        super(props);
        this.state = {};
    }

    public render(): React.ReactElement<ITutorialItemDisplayProps> {
        console.log(`${this.props.tutorialItem.key} is rendered by React!`);
        const actionButton: IButtonProps = this.props.tutorialItem.nextTrigger == "action" ? {
            text: this.props.tutorialItem.nextActionText || "Next",
            onClick: () => {
                this._onDismiss();
            }
        }
            : null;
        const moreButton: IButtonProps = this.props.tutorialItem.learnMore ? {
            text: this.props.tutorialItem.learnMore.text,
        }
            : null;
        return <div>
            <TeachingBubble
                targetElement={this.props.targetElement}
                headline={this.props.tutorialItem.caption}
                primaryButtonProps={actionButton}
                secondaryButtonProps={moreButton}
                onDismiss={this._onDismiss.bind(this)}
            >
                {this.props.tutorialItem.content}
            </TeachingBubble>
        </div>;
    }

    private _onDismiss() {
        console.log(`${this.props.tutorialItem.key} is dismissed!`);
        if (this.props.next) {
            this.props.next(this.props.tutorialItem);
        }
    }
}

export interface ITutorialState {
    shouldPlayTutorial: boolean;
}

export interface ITutorialProps {
    showReplayTutorial?: boolean;
    replayTutorialLabel?: string;
    replayTutorialLabelProps?: any;
    tutorial: ITutorial;
}

export class Tutorial extends React.Component<ITutorialProps, ITutorialState> {

    constructor(props: ITutorialProps) {
        super(props);
    }

    public componentDidMount() {

        if (shouldPlayTutorial(this.props.tutorial)) {
            this.setState({ shouldPlayTutorial: true });
        }
    }

    public componentDidUpdate() {
        if (this.state.shouldPlayTutorial) {
            playTutorial(this.props.tutorial);
            return;
        }
    }

    public render(): React.ReactElement<ITutorialProps> {
        const content: any[] = [];

        if (this.props.showReplayTutorial) {
            const replayLabel = this.props.replayTutorialLabel || "Play tutorial again";
            content.push(<Link onClick={() => this._requestPlay()} {...this.props.replayTutorialLabelProps}>{replayLabel}</Link>);
        }

        return <div>
            {content}
        </div>;
    }

    private _requestPlay() {
        this.setState({
            shouldPlayTutorial: true
        });
    }
}