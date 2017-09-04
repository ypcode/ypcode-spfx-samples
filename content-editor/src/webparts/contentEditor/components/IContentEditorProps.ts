
import { IContentService } from "../services/ContentService";
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IContentEditorProps {
  contentService: IContentService;
  displayMode: DisplayMode;
  showCaption: boolean;
}
