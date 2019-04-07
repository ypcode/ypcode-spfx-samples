import { ITutorial } from "../../ui-fabric-tutorial/index";

export default {
    id: 'tuto_sample',
    items: [
        {
            key: 'newItem',
            caption: 'New Item',
            content: 'This action allows to add a new item',
            delay: 3000,
            nextTrigger: 'delay'
        },
        {
            key: 'upload',
            caption: 'Upload',
            content: 'This action allows to upload a new document',
            delay: 3000,
            nextTrigger: 'delay'
        },
        {
            key: 'share',
            caption: 'Share',
            content: 'Share the document with a contact',
            delay: 3000,
            nextTrigger: 'delay'
        },
        {
            key: 'download',
            caption: 'Download',
            content: 'This action allows to download a document',
            delay: 3000,
            nextTrigger: 'delay'
        },
        {
            key: 'dateModifiedColumn',
            caption: 'Modified date',
            content: 'This columns displays the last modification date',
            delay: 3000,
            nextTrigger: 'delay'
        },
        {
            key: 'modifiedByColumn',
            caption: 'Modified by',
            content: 'This columns displays the last person who modified the document',
            delay: 3000,
            nextTrigger: 'delay'
        }

    ]
} as ITutorial;