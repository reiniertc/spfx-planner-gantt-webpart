import * as React from 'react';
import { Gantt, ViewMode } from 'gantt-task-react';
import 'gantt-task-react/dist/index.css';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';

import styles from './PlannerGantt.module.scss';
import type { IPlannerGanttProps } from './IPlannerGanttProps';
import { IGanttRow } from '../models/IPlannerModels';
import { IGanttChartTask, createTaskListHeader, createTaskListTable } from './GanttTaskList';
import * as strings from 'PlannerGanttWebPartStrings';

const COLUMN_WIDTH_BY_VIEW_MODE: Record<string, number> = {
  Day: 65,
  Week: 165,
  Month: 300
};

const TODAY_LINE_COLOR: string = 'rgba(232, 17, 35, 0.15)';
const TODAY_LINE_COLOR_HIDDEN: string = 'transparent';

function toGanttTasks(rows: IGanttRow[], showProgressOnBar: boolean): IGanttChartTask[] {
  return rows.map(row => ({
    id: row.id,
    name: showProgressOnBar && row.type !== 'milestone' ? `${row.name} (${row.progress}%)` : row.name,
    start: row.start,
    end: row.end,
    progress: row.progress,
    type: row.type,
    project: row.project,
    assigneeNames: row.assigneeNames,
    isDisabled: true // Planner is the system of record; the chart is read-only.
  }));
}

const PlannerGantt: React.FC<IPlannerGanttProps> = (props: IPlannerGanttProps) => {
  const {
    planId, viewMode, showCompletedTasks, showBucketsAsPhases, showProgressOnBar,
    showCurrentDateLine, showStartDateColumn, showEndDateColumn, showAssigneeColumn,
    bucketFilter, plannerService
  } = props;
  // Stringified so an equivalent-but-new object reference (the web part
  // shallow-copies bucketFilter on every render) doesn't trigger a refetch.
  const bucketFilterKey: string = JSON.stringify(bucketFilter || {});

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
      .getGanttRows(planId, {
        includeCompleted: showCompletedTasks,
        showBucketsAsPhases,
        bucketFilter
      })
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
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [planId, showCompletedTasks, showBucketsAsPhases, bucketFilterKey, plannerService]);

  const columns = React.useMemo(() => ({
    showStartDate: showStartDateColumn,
    showEndDate: showEndDateColumn,
    showAssignee: showAssigneeColumn
  }), [showStartDateColumn, showEndDateColumn, showAssigneeColumn]);

  const TaskListHeader = React.useMemo(() => createTaskListHeader(columns), [columns]);
  const TaskListTable = React.useMemo(() => createTaskListTable(columns), [columns]);

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
        tasks={toGanttTasks(rows, showProgressOnBar)}
        viewMode={resolvedViewMode}
        columnWidth={COLUMN_WIDTH_BY_VIEW_MODE[viewMode] || COLUMN_WIDTH_BY_VIEW_MODE.Week}
        listCellWidth="220px"
        rowHeight={42}
        ganttHeight={520}
        locale={navigator.language}
        todayColor={showCurrentDateLine ? TODAY_LINE_COLOR : TODAY_LINE_COLOR_HIDDEN}
        TaskListHeader={TaskListHeader}
        TaskListTable={TaskListTable}
      />
    </div>
  );
};

export default PlannerGantt;
