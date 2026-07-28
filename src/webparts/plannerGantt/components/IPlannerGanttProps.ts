import { PlannerService } from '../services/PlannerService';
import { IBucketFilter } from '../models/IPlannerModels';

export interface IPlannerGanttProps {
  planId: string;
  planTitle: string;
  customTitle: string;
  showTitle: boolean;
  viewMode: string;
  showCompletedTasks: boolean;
  showBucketsAsPhases: boolean;
  colorBarsByStatus: boolean;
  sortTasksByStartDate: boolean;
  sortBucketsByStartDate: boolean;
  showTaskNameOnBar: boolean;
  showCurrentDateLine: boolean;
  showStartDateColumn: boolean;
  showEndDateColumn: boolean;
  showAssigneeColumn: boolean;
  defaultZoomLevel: number;
  showZoomControl: boolean;
  showPrintButton: boolean;
  showAssigneeFilter: boolean;
  showCompletedFilterControl: boolean;
  scrollToToday: boolean;
  scrollToTodayMarginUnits: number;
  bucketFilter?: IBucketFilter;
  isDarkTheme: boolean;
  themePrimary: string;
  themeSecondary: string;
  themeGrey: string;
  plannerService: PlannerService;
}
