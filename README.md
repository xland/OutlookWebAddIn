# OutlookWebAddIn

- npm install -g yo generator-office
- yo office
- Choose a project type: Office Add-in Task Pane project
- Choose a script type: TypeScript
- What do you want to name your add-in? OutlookWebAddin
- Which Office client application would you like to support? Outlook
- webpack.config.js: entry 
- webpack.config.js: HtmlWebpackPlugin
- package.json: manifest.xml
- mainfest.xml: MessageReadCommandSurface： msgReadGroup
- mainfest.xml: AppointmentOrganizerCommandSurface
- mainfest.xml: ShortStrings，LongStrings
- mainfest.xml: ItemRead
- mainfest.xml: RuleCollection


# ToDo

- 特殊邮件，第一个按钮禁用状态
- 插件名，提供商
- 多语言
- 必须用外网