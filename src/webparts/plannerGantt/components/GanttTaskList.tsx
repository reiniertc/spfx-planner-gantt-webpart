import * as React from 'react';
import { Task as GanttTask } from 'gantt-task-react';
import * as strings from 'PlannerGanttWebPartStrings';
import styles from './PlannerGantt.module.scss';
import { PlannerService } from '../services/PlannerService';
import { TaskInfoButton, ITaskInfoOptions } from './TaskInfoPopover';

/** Extra data gantt-task-react's own Task type has no slot for. */
export interface IGanttChartTask extends GanttTask {
  assignees?: string[];
  labels?: string[];
  hasDescription?: boolean;
}

export interface IColumnVisibility {
  showStartDate: boolean;
  showEndDate: boolean;
  showAssignee: boolean;
  showLabel: boolean;
}

export interface IColumnWidths {
  name: number;
  start: number;
  end: number;
  assignee: number;
  label: number;
}

export const DEFAULT_COLUMN_WIDTHS: IColumnWidths = { name: 220, start: 100, end: 100, assignee: 160, label: 140 };

const MIN_COLUMN_WIDTH: number = 60;

export interface IColumnWidthsContextValue {
  widths: IColumnWidths;
  setColumnWidth: (column: keyof IColumnWidths, width: number) => void;
}

/**
 * Shared between the header and body components gantt-task-react renders
 * for us: they're separate component instances with no direct link to each
 * other, so a plain prop can't keep their column widths in sync. Context
 * works because both are still rendered inside our own React tree (Gantt
 * doesn't portal them elsewhere).
 */
export const ColumnWidthsContext: React.Context<IColumnWidthsContextValue> = React.createContext<IColumnWidthsContextValue>({
  widths: DEFAULT_COLUMN_WIDTHS,
  setColumnWidth: () => { /* no provider mounted */ }
});

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

/**
 * Opens the actual task, on its own plan board, in a new tab. Not valid for
 * phase/project rows - those are synthetic bucket summaries, not real
 * Planner tasks. The task id alone isn't enough for a working deep link in
 * the current ("new") Planner - it needs the plan id too.
 */
export function openPlannerTask(planId: string, taskId: string): void {
  const url: string = `https://planner.cloud.microsoft/webui/plan/${encodeURIComponent(planId)}/view/board/task/${encodeURIComponent(taskId)}`;
  window.open(url, '_blank', 'noopener,noreferrer');
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

const ResizeHandle: React.FC<{ column: keyof IColumnWidths }> = ({ column }) => {
  const { widths, setColumnWidth } = React.useContext(ColumnWidthsContext);
  const dragStateRef = React.useRef<{ startX: number; startWidth: number } | undefined>(undefined);

  const onMouseMove = React.useCallback((event: MouseEvent) => {
    if (!dragStateRef.current) {
      return;
    }
    const delta: number = event.clientX - dragStateRef.current.startX;
    setColumnWidth(column, Math.max(MIN_COLUMN_WIDTH, dragStateRef.current.startWidth + delta));
  }, [column, setColumnWidth]);

  const onMouseUp = React.useCallback(() => {
    dragStateRef.current = undefined;
    window.removeEventListener('mousemove', onMouseMove);
    window.removeEventListener('mouseup', onMouseUp);
  }, [onMouseMove]);

  const onMouseDown = (event: React.MouseEvent): void => {
    event.preventDefault();
    dragStateRef.current = { startX: event.clientX, startWidth: widths[column] };
    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);
  };

  return <span className={styles.ganttTableResizeHandle} onMouseDown={onMouseDown} />;
};

export function createTaskListHeader(columns: IColumnVisibility): React.FC<ITaskListHeaderProps> {
  return function TaskListHeader(props: ITaskListHeaderProps): React.ReactElement {
    const { widths } = React.useContext(ColumnWidthsContext);
    return (
      <div className={styles.ganttTableHeader} style={{ height: props.headerHeight - 2, fontFamily: props.fontFamily, fontSize: props.fontSize }}>
        <div className={styles.ganttTableHeaderCell} style={{ width: widths.name }}>
          {strings.ColumnNameHeader}
          <ResizeHandle column="name" />
        </div>
        {columns.showStartDate && (
          <div className={styles.ganttTableHeaderCell} style={{ width: widths.start }}>
            {strings.ColumnStartHeader}
            <ResizeHandle column="start" />
          </div>
        )}
        {columns.showEndDate && (
          <div className={styles.ganttTableHeaderCell} style={{ width: widths.end }}>
            {strings.ColumnEndHeader}
            <ResizeHandle column="end" />
          </div>
        )}
        {columns.showAssignee && (
          <div className={styles.ganttTableHeaderCell} style={{ width: widths.assignee }}>
            {strings.ColumnAssigneeHeader}
            <ResizeHandle column="assignee" />
          </div>
        )}
        {columns.showLabel && (
          <div className={styles.ganttTableHeaderCell} style={{ width: widths.label }}>
            {strings.ColumnLabelHeader}
            <ResizeHandle column="label" />
          </div>
        )}
      </div>
    );
  };
}

export function createTaskListTable(
  columns: IColumnVisibility,
  planId: string,
  taskInfoOptions: ITaskInfoOptions,
  plannerService: PlannerService
): React.FC<ITaskListTableProps> {
  return function TaskListTable(props: ITaskListTableProps): React.ReactElement {
    const { widths } = React.useContext(ColumnWidthsContext);

    return (
      <div style={{ fontFamily: props.fontFamily, fontSize: props.fontSize }}>
        {props.tasks.map(task => {
          const chartTask: IGanttChartTask = task as IGanttChartTask;
          const isPhase: boolean = task.type === 'project';
          const isPhaseChild: boolean = !!task.project;
          const assignees: string[] = chartTask.assignees || [];
          const assigneeText: string = assignees.join(', ');
          const labelText: string = (chartTask.labels || []).join(', ');
          return (
            <div className={styles.ganttTableRow} style={{ height: props.rowHeight }} key={`${task.id}-row`}>
              <div className={styles.ganttTableCell} style={{ width: widths.name }} title={task.name}>
                <span
                  className={expanderSymbol(task) ? styles.ganttTableExpander : styles.ganttTableExpanderEmpty}
                  onClick={() => props.onExpanderClick(task)}
                >
                  {expanderSymbol(task)}
                </span>
                {isPhase ? (
                  <span className={styles.ganttTablePhaseName}>{task.name}</span>
                ) : (
                  <span
                    className={styles.ganttTableTaskLink}
                    style={isPhaseChild ? { paddingLeft: '1.2em' } : undefined}
                    onClick={() => openPlannerTask(planId, task.id)}
                    title={strings.OpenInPlannerLabel}
                  >
                    {task.name}
                  </span>
                )}
                {!isPhase && taskInfoOptions.show && (
                  <TaskInfoButton
                    data={{
                      taskId: task.id,
                      start: task.start,
                      end: task.end,
                      progress: task.progress,
                      assignees,
                      hasDescription: !!chartTask.hasDescription
                    }}
                    options={taskInfoOptions}
                    plannerService={plannerService}
                  />
                )}
              </div>
              {columns.showStartDate && (
                <div className={styles.ganttTableCell} style={{ width: widths.start }}>
                  {task.start.toLocaleDateString(props.locale)}
                </div>
              )}
              {columns.showEndDate && (
                <div className={styles.ganttTableCell} style={{ width: widths.end }}>
                  {task.end.toLocaleDateString(props.locale)}
                </div>
              )}
              {columns.showAssignee && (
                <div className={styles.ganttTableCell} style={{ width: widths.assignee }} title={assigneeText}>
                  {assigneeText || '–'}
                </div>
              )}
              {columns.showLabel && (
                <div className={styles.ganttTableCell} style={{ width: widths.label }} title={labelText}>
                  {labelText || '–'}
                </div>
              )}
            </div>
          );
        })}
      </div>
    );
  };
}
