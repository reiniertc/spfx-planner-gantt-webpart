declare interface IPlannerGanttWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  PlanFieldLabel: string;
  ViewModeFieldLabel: string;
  ViewModeDay: string;
  ViewModeWeek: string;
  ViewModeMonth: string;
  ShowCompletedFieldLabel: string;
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
