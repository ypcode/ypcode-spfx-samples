## spfx-panel

# MPA (Minimum Path to Awesome)

- Update the config/serve.json to set the URL of a page on your tenant

```json
  "panelCommandSet": {
      "pageUrl": "https://.sharepoint.com/sites/mySite/SitePages/myPage.aspx", // This should be updated
      "customActions": {
        "506494b1-0cb4-4f4a-a5e5-80c2ea92d191": {
          "location": "ClientSideExtension.ListViewCommandSet.CommandBar",
          "properties": {
          }
        }
      }
    }
```

- type command gulp serve --config panelCommandSet
