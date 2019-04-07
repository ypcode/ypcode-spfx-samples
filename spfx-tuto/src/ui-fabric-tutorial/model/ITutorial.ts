import { ITutorialItem } from "./ITutorialItem";

export interface ITutorial {
    id: string;
    items: ITutorialItem[];
}