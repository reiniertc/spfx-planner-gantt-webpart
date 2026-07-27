import * as React from 'react';
import { Task as GanttTask } from 'gantt-task-react';
import * as strings from 'PlannerGanttWebPartStrings';
import styles from './PlannerGantt.module.scss';

/** Extra data gantt-task-react's own Task type has no slot for. */
export interface IGanttChartTask extends GanttTask {
  assigneeNames?: string;
}

export interface IColumnVisibility {
  showStartDate: boolean;
  showEndDate: boolean;
  showAssignee: boolean;
}

const NAME_COLUMN_WIDTH: string = '220px';
const DATE_COLUMN_WIDTH: string = '100px';
const ASSIGNEE_COLUMN_WIDTH: string = '160px';

interface ITaskListHeaderProps {
  headerHeight: number;
  fontFamily: string;
  fontSize: string;
}

interface ITaskListTableProps {
  rowHeight: number;
  fontFamily: string;
  fontSize: string;
  locale: string;
  tasks: GanttTask[];
  selectedTaskId: string;
  setSelectedTask: (taskId: string) => void;
  onExpanderClick: (task: GanttTask) => void;
}

function expanderSymbol(task: GanttTask): string {
  if (task.hideChildren === false) {
    return '▼';
  }
  if (task.hideChildren === true) {
    return '▶';
  }
  return '';
}

export function createTaskListHeader(columns: IColumnVisibility): React.FC<ITaskListHeaderProps> {
  return function TaskListHeader(props: ITaskListHeaderProps): React.ReactElement {
    return (
      <div className={styles.ganttTableHeader} style={{ height: props.headerHeight - 2, fontFamily: props.fontFamily, fontSize: props.fontSize }}>
        <div className={styles.ganttTableHeaderCell} style={{ width: NAME_COLUMN_WIDTH }}>{strings.ColumnNameHeader}</div>
        {columns.showStartDate && (
          <div className={styles.ganttTableHeaderCell} style={{ width: DATE_COLUMN_WIDTH }}>{strings.ColumnStartHeader}</div>
        )}
        {columns.showEndDate && (
          <div className={styles.ganttTableHeaderCell} style={{ width: DATE_COLUMN_WIDTH }}>{strings.ColumnEndHeader}</div>
        )}
        {columns.showAssignee && (
          <div className={styles.ganttTableHeaderCell} style={{ width: ASSIGNEE_COLUMN_WIDTH }}>{strings.ColumnAssigneeHeader}</div>
        )}
      </div>
    );
  };
}

export function createTaskListTable(columns: IColumnVisibility): React.FC<ITaskListTableProps> {
  return function TaskListTable(props: ITaskListTableProps): React.ReactElement {
    return (
      <div style={{ fontFamily: props.fontFamily, fontSize: props.fontSize }}>
        {props.tasks.map(task => {
          const chartTask: IGanttChartTask = task as IGanttChartTask;
          const isPhase: boolean = task.type === 'project';
          const isPhaseChild: boolean = !!task.project;
          return (
            <div className={styles.ganttTableRow} style={{ height: props.rowHeight }} key={`${task.id}-row`}>
              <div className={styles.ganttTableCell} style={{ width: NAME_COLUMN_WIDTH }} title={task.name}>
                <span
                  className={expanderSymbol(task) ? styles.ganttTableExpander : styles.ganttTableExpanderEmpty}
                  onClick={() => props.onExpanderClick(task)}
                >
                  {expanderSymbol(task)}
                </span>
                <span
                  className={isPhase ? styles.ganttTablePhaseName : undefined}
                  style={isPhaseChild ? { paddingLeft: '1.2em' } : undefined}
                >
                  {task.name}
                </span>
              </div>
              {columns.showStartDate && (
                <div className={styles.ganttTableCell} style={{ width: DATE_COLUMN_WIDTH }}>
                  {task.start.toLocaleDateString(props.locale)}
                </div>
              )}
              {columns.showEndDate && (
                <div className={styles.ganttTableCell} style={{ width: DATE_COLUMN_WIDTH }}>
                  {task.end.toLocaleDateString(props.locale)}
                </div>
              )}
              {columns.showAssignee && (
                <div className={styles.ganttTableCell} style={{ width: ASSIGNEE_COLUMN_WIDTH }} title={chartTask.assigneeNames || ''}>
                  {chartTask.assigneeNames || '–'}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  };
}
