declare interface IIssueTrackerWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'IssueTrackerWebPartStrings' {
  const strings: IIssueTrackerWebPartStrings;
  export = strings;
}
