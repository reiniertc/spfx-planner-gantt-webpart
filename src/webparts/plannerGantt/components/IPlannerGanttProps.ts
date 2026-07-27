import { PlannerService } from '../services/PlannerService';
import { IBucketFilter } from '../models/IPlannerModels';

export interface IPlannerGanttProps {
  planId: string;
  planTitle: string;
  viewMode: string;
  showCompletedTasks: boolean;
  showBucketsAsPhases: boolean;
  colorBarsByStatus: boolean;
  sortTasksByStartDate: boolean;
  showTaskNameOnBar: boolean;
  showCurrentDateLine: boolean;
  showStartDateColumn: boolean;
  showEndDateColumn: boolean;
  showAssigneeColumn: boolean;
  bucketFilter?: IBucketFilter;
  isDarkTheme: boolean;
  themePrimary: string;
  themeSecondary: string;
  themeGrey: string;
  plannerService: PlannerService;
}
