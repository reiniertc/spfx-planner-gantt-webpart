import { PlannerService } from '../services/PlannerService';

export interface IPlannerGanttProps {
  planId: string;
  planTitle: string;
  viewMode: string;
  showCompletedTasks: boolean;
  isDarkTheme: boolean;
  plannerService: PlannerService;
}
