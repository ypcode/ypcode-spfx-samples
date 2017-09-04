

import { markdown } from "markdown";
import pnp from "sp-pnp-js";

export enum SourceFormat {
    Auto,
    Html,
    Markdown
}

export enum SourceType {
    Content,
    Link
}

export interface IContentServiceConfiguration {
    sourceType: SourceType;
    sourceFormat: SourceFormat;
    sourceContent: string;
    sourceLink: string;
}

export interface IContentService {

    configuration(): IContentServiceConfiguration;

    configure(propertyBag: IContentServiceConfiguration);

    isProperlyConfigured(): boolean;

    getContent(): Promise<string>;

    setContent(content: string): void;

    setContentLink(link: string): void;
}

export class ContentService implements IContentService {

    private propertyBag: IContentServiceConfiguration;

    public configuration(): IContentServiceConfiguration {
        return {
            sourceContent: this.propertyBag.sourceContent,
            sourceLink: this.propertyBag.sourceLink,
            sourceFormat: this.propertyBag.sourceFormat,
            sourceType: this.propertyBag.sourceType
        };
    }


    public configure(propertyBag: IContentServiceConfiguration) {
        this.propertyBag = propertyBag;
    }

    public isProperlyConfigured(): boolean {
        let config = this.configuration();

        return (config.sourceFormat != null
            && config.sourceType != null
            && (config.sourceContent || config.sourceLink))
            && true;
    }

    public getContent(): Promise<string> {

        let config = this.configuration();
        switch (config.sourceType) {
            case SourceType.Link:
                return pnp.sp.web.getFileByServerRelativeUrl(config.sourceLink)
                    .getText()
                    .then(content => ContentService.formatContent(content, this._getSourceFormat()));
            case SourceType.Content:
            default:
                return Promise.resolve(ContentService.formatContent(config.sourceContent, this._getSourceFormat()));
        }
    }

    public setContent(content: string): void {
        if (this.propertyBag) {
            this.propertyBag.sourceContent = content;
        }
    }

    public setContentLink(link: string): void {
        if (this.propertyBag) {
            this.propertyBag.sourceLink = link;
        }
    }

    private _getSourceFormat() : SourceFormat {
        let config = this.configuration();

        if (config.sourceFormat != SourceFormat.Auto)
            return config.sourceFormat;

        if (config.sourceLink.indexOf(".html") == config.sourceLink.length-6)
            return SourceFormat.Html;

        // if (config.sourceLink.indexOf(".md") == config.sourceLink.length-4)
        //     return SourceFormat.Markdown;

        return SourceFormat.Markdown;
    }

    private static formatContent(content: string, format: SourceFormat): string {
        if (!content)
            return "";

        switch (format) {
            case SourceFormat.Markdown:
                return markdown.toHTML(content);
            case SourceFormat.Html:
            default:
                return content;
        }
    }
}