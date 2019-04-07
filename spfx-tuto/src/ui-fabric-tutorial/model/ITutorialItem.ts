
export interface ILink {
    url: string;
    text: string;
}
export interface ITutorialItem {
    key: string;
    selector?: string;
    caption: string;
    content: string|JSX.Element;
    delay?: number;
    learnMore?: ILink;
    nextActionText?: string;
    nextTrigger: "action"|"delay";
}