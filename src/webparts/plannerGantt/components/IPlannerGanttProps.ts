import { PlannerService } from '../services/PlannerService';
import { IBucketFilter } from '../models/IPlannerModels';

export interface IPlannerGanttProps {
  planId: string;
  planTitle: string;
  viewMode: string;
  showCompletedTasks: boolean;
  showBucketsAsPhases: boolean;
  showProgressOnBar: boolean;
  showCurrentDateLine: boolean;
  showStartDateColumn: boolean;
  showEndDateColumn: boolean;
  showAssigneeColumn: boolean;
  bucketFilter?: IBucketFilter;
  isDarkTheme: boolean;
  plannerService: PlannerService;
}
