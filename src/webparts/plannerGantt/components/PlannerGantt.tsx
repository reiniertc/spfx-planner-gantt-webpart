import * as React from 'react';
import { Gantt, Task as GanttTask, ViewMode } from 'gantt-task-react';
import 'gantt-task-react/dist/index.css';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import styles from './PlannerGantt.module.scss';
import type { IPlannerGanttProps } from './IPlannerGanttProps';
import { IGanttRow } from '../models/IPlannerModels';
import * as strings from 'PlannerGanttWebPartStrings';

const COLUMN_WIDTH_BY_VIEW_MODE: Record<string, number> = {
  Day: 65,
  Week: 165,
  Month: 300
};

function toGanttTasks(rows: IGanttRow[]): GanttTask[] {
  return rows.map(row => ({
    id: row.id,
    name: row.name,
    start: row.start,
    end: row.end,
    progress: row.progress,
    type: row.type,
    project: row.project,
    isDisabled: true // Planner is the system of record; the chart is read-only.
  }));
}

const PlannerGantt: React.FC<IPlannerGanttProps> = (props: IPlannerGanttProps) => {
  const { planId, viewMode, showCompletedTasks, plannerService } = props;

  const [rows, setRows] = React.useState<IGanttRow[] | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  React.useEffect(() => {
    if (!planId) {
      setRows(undefined);
      setError(undefined);
      return;
    }

    let isCancelled: boolean = false;
    setIsLoading(true);
    setError(undefined);

    plannerService
      .getGanttRows(planId, showCompletedTasks)
      .then(result => {
        if (!isCancelled) {
          setRows(result);
          setIsLoading(false);
        }
      })
      .catch(() => {
        if (!isCancelled) {
          setError(strings.ErrorLoadingTasksLabel);
          setIsLoading(false);
        }
      });

    return () => {
      isCancelled = true;
    };
  }, [planId, showCompletedTasks, plannerService]);

  if (!planId) {
    return (
      <div className={styles.plannerGantt}>
        <MessageBar messageBarType={MessageBarType.info}>{strings.NoPlanSelectedLabel}</MessageBar>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className={styles.plannerGantt}>
        <Spinner size={SpinnerSize.large} label={strings.LoadingTasksLabel} />
      </div>
    );
  }

  if (error) {
    return (
      <div className={styles.plannerGantt}>
        <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>
      </div>
    );
  }

  if (!rows || rows.length === 0) {
    return (
      <div className={styles.plannerGantt}>
        <MessageBar messageBarType={MessageBarType.warning}>{strings.NoTasksFoundLabel}</MessageBar>
      </div>
    );
  }

  const resolvedViewMode: ViewMode = ViewMode[viewMode as keyof typeof ViewMode] || ViewMode.Week;

  return (
    <div className={styles.plannerGantt}>
      <h2 className={styles.planTitle}>{props.planTitle}</h2>
      <Gantt
        tasks={toGanttTasks(rows)}
        viewMode={resolvedViewMode}
        columnWidth={COLUMN_WIDTH_BY_VIEW_MODE[viewMode] || COLUMN_WIDTH_BY_VIEW_MODE.Week}
        listCellWidth="220px"
        rowHeight={42}
        ganttHeight={520}
        locale={navigator.language}
      />
    </div>
  );
};

export default PlannerGantt;
