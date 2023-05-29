declare interface IGetSpListItemsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'GetSpListItemsWebPartStrings' {
  const strings: IGetSpListItemsWebPartStrings;
  export = strings;
}
