import * as React from 'react';
import { TextField } from "office-ui-fabric-react";
import styles from './ContentEditor.module.scss';
import { IContentEditorProps } from './IContentEditorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { DisplayMode } from '@microsoft/sp-core-library';

import { SourceFormat, SourceType, IContentService } from "../services/ContentService";

export interface IContentEditorState {
  content: string;
}

export default class ContentEditor extends React.Component<IContentEditorProps, IContentEditorState> {

  private contentService: IContentService;

  constructor(props: IContentEditorProps) {
    super(props);

    this.state = {
      content: null
    };

    this.contentService = props.contentService;
  }

  public componentWillMount() {

    this.contentService.getContent().then(content => {
      this.setState({
        content: content
      });
    });
  }

  public componentWillReceiveProps(oldProps, nextProps) {
    this.contentService.getContent().then(content => {
      this.setState({
        content: content
      });
    });
  }

  public render(): React.ReactElement<IContentEditorProps> {

    let { content } = this.state;
    let { contentService, displayMode, showCaption } = this.props;
    let config = contentService.configuration();

    return (
      <div className={styles.contentEditor}>
        <div className={styles.container}>
          <div className="ms-Grid-row">
            <div className="ms-Grid-col ms-u-md12">
              {
                displayMode == DisplayMode.Edit
                  ? contentService.isProperlyConfigured()
                    ? config.sourceType == SourceType.Content
                      ? <TextField
                        multiline={true}
                        value={content}
                        rows={15}
                        onChanged={v => this._updateContent(v)} />
                      : <div>
                        {showCaption && <div className={styles.caption} >
                          {config.sourceLink || ""}
                        </div>}
                        <div className={styles.content} dangerouslySetInnerHTML={this._unsafeHtml(content)} >
                        </div>
                      </div>
                    : <div>
                      <a>Please configure the Content source</a>
                    </div>
                  : <div>
                    {showCaption && config.sourceType == SourceType.Link && <div className={styles.caption} >
                      {config.sourceLink || ""}
                    </div>}
                    <div className={styles.content} dangerouslySetInnerHTML={this._unsafeHtml(content)} >
                    </div>
                  </div>
              }
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _unsafeHtml(html: string) {
    return { __html: (html || "") };
  }

  private _updateContent(content: string): void {
    let { contentService } = this.props;
    console.log("CONTENT HAS CHANGED TO " + content);
    contentService.setContent(content);
  }
}
