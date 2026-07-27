declare interface IPlannerGanttWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PlanFieldLabel: string;
  DisplayGroupName: string;
  ViewModeFieldLabel: string;
  ViewModeDay: string;
  ViewModeWeek: string;
  ViewModeMonth: string;
  ShowCompletedFieldLabel: string;
  ShowBucketsAsPhasesFieldLabel: string;
  SortTasksByStartDateFieldLabel: string;
  ColorBarsByStatusFieldLabel: string;
  ShowCurrentDateLineFieldLabel: string;
  ColumnsGroupName: string;
  ShowStartDateColumnFieldLabel: string;
  ShowEndDateColumnFieldLabel: string;
  ShowAssigneeColumnFieldLabel: string;
  ColumnNameHeader: string;
  ColumnStartHeader: string;
  ColumnEndHeader: string;
  ColumnAssigneeHeader: string;
  BucketsGroupName: string;
  LoadingBucketsLabel: string;
  ErrorLoadingBucketsLabel: string;
  NoBucketsLabel: string;
  LoadingPlansLabel: string;
  ErrorLoadingPlansLabel: string;
  NoPlanSelectedLabel: string;
  LoadingTasksLabel: string;
  ErrorLoadingTasksLabel: string;
  NoTasksFoundLabel: string;
}

declare module 'PlannerGanttWebPartStrings' {
  const strings: IPlannerGanttWebPartStrings;
  export = strings;
}
