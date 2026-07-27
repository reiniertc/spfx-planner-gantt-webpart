import { PlannerService } from '../services/PlannerService';
import { IBucketFilter } from '../models/IPlannerModels';

export interface IPlannerGanttProps {
  planId: string;
  planTitle: string;
  viewMode: string;
  showCompletedTasks: boolean;
  showBucketsAsPhases: boolean;
  bucketFilter?: IBucketFilter;
  isDarkTheme: boolean;
  plannerService: PlannerService;
}
