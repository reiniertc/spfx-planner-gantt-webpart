import * as React from 'react';
import { Gantt, ViewMode } from 'gantt-task-react';
import 'gantt-task-react/dist/index.css';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Slider } from '@fluentui/react/lib/Slider';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';

import styles from './PlannerGantt.module.scss';
import type { IPlannerGanttProps } from './IPlannerGanttProps';
import { IGanttRow } from '../models/IPlannerModels';
import {
  IGanttChartTask, IColumnWidths, DEFAULT_COLUMN_WIDTHS, ColumnWidthsContext,
  createTaskListHeader, createTaskListTable
} from './GanttTaskList';
import * as strings from 'PlannerGanttWebPartStrings';

const COLUMN_WIDTH_BY_VIEW_MODE: Record<string, number> = {
  Hour: 60,
  'Quarter Day': 80,
  'Half Day': 100,
  Day: 65,
  Week: 165,
  Month: 300
};

const ZOOM_LEVELS: Array<{ key: keyof typeof ViewMode; label: string }> = [
  { key: 'Hour', label: strings.ViewModeHour },
  { key: 'QuarterDay', label: strings.ViewModeQuarterDay },
  { key: 'HalfDay', label: strings.ViewModeHalfDay },
  { key: 'Day', label: strings.ViewModeDay },
  { key: 'Week', label: strings.ViewModeWeek },
  { key: 'Month', label: strings.ViewModeMonth }
];

function indexOfZoomLevel(key: string): number {
  for (let i: number = 0; i < ZOOM_LEVELS.length; i++) {
    if (ZOOM_LEVELS[i].key === key) {
      return i;
    }
  }
  return -1;
}

const DEFAULT_ZOOM_INDEX: number = indexOfZoomLevel('Week');
const UNASSIGNED_KEY: string = '__unassigned__';

const TODAY_LINE_COLOR: string = 'rgba(232, 17, 35, 0.15)';
const TODAY_LINE_COLOR_HIDDEN: string = 'transparent';

interface IThemeColors {
  primary: string;
  secondary: string;
  grey: string;
}

/**
 * Planner only has three real progress states (0 / 50 / 100 - not started,
 * in progress, done), so a numeric percentage on the bar is more noise than
 * signal. Colour-coding the whole bar by status reads better in practice.
 */
function colorForStatus(progress: number, theme: IThemeColors): string {
  if (progress >= 100) {
    return theme.grey;
  }
  if (progress <= 0) {
    return theme.secondary;
  }
  return theme.primary;
}

// gantt-task-react always renders the task name on/next to the bar with no
// prop to suppress it; when the "show task name on bar" toggle is off, the
// label is hidden with CSS instead (see .hideBarLabels), so `name` here is
// always the real title (also used by the tooltip and left-hand task list).
function toGanttTasks(rows: IGanttRow[], colorBarsByStatus: boolean, theme: IThemeColors): IGanttChartTask[] {
  return rows.map(row => {
    const statusColor: string | undefined = colorBarsByStatus ? colorForStatus(row.progress, theme) : undefined;

    return {
      id: row.id,
      name: row.name,
      start: row.start,
      end: row.end,
      progress: row.progress,
      type: row.type,
      project: row.project,
      assignees: row.assignees,
      isDisabled: true, // Planner is the system of record; the chart is read-only.
      styles: statusColor
        ? {
          backgroundColor: statusColor,
          backgroundSelectedColor: statusColor,
          progressColor: statusColor,
          progressSelectedColor: statusColor
        }
        : undefined
    };
  });
}

function findZoomIndex(viewModeKey: string): number {
  const index: number = indexOfZoomLevel(viewModeKey);
  return index === -1 ? DEFAULT_ZOOM_INDEX : index;
}

const PlannerGantt: React.FC<IPlannerGanttProps> = (props: IPlannerGanttProps) => {
  const {
    planId, viewMode, showCompletedTasks, showBucketsAsPhases, colorBarsByStatus, sortTasksByStartDate,
    showTaskNameOnBar, showCurrentDateLine, showStartDateColumn, showEndDateColumn, showAssigneeColumn,
    bucketFilter, plannerService, themePrimary, themeSecondary, themeGrey
  } = props;
  // Stringified so an equivalent-but-new object reference (the web part
  // shallow-copies bucketFilter on every render) doesn't trigger a refetch.
  const bucketFilterKey: string = JSON.stringify(bucketFilter || {});

  const [rows, setRows] = React.useState<IGanttRow[] | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  // Live, viewer-side controls - not persisted to the web part's properties.
  const [zoomIndex, setZoomIndex] = React.useState<number>(() => findZoomIndex(viewMode));
  const [selectedAssignee, setSelectedAssignee] = React.useState<string>('');
  const [columnWidths, setColumnWidths] = React.useState<IColumnWidths>(DEFAULT_COLUMN_WIDTHS);

  React.useEffect(() => {
    setZoomIndex(findZoomIndex(viewMode));
    setSelectedAssignee('');
  }, [planId, viewMode]);

  const setColumnWidth = React.useCallback((column: keyof IColumnWidths, width: number) => {
    setColumnWidths(previous => ({ ...previous, [column]: width }));
  }, []);

  const columnWidthsContextValue = React.useMemo(() => ({ widths: columnWidths, setColumnWidth }), [columnWidths, setColumnWidth]);

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
        sortByStartDate: sortTasksByStartDate,
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
  }, [planId, showCompletedTasks, showBucketsAsPhases, sortTasksByStartDate, bucketFilterKey, plannerService]);

  const columns = React.useMemo(() => ({
    showStartDate: showStartDateColumn,
    showEndDate: showEndDateColumn,
    showAssignee: showAssigneeColumn
  }), [showStartDateColumn, showEndDateColumn, showAssigneeColumn]);

  const TaskListHeader = React.useMemo(() => createTaskListHeader(columns), [columns]);
  const TaskListTable = React.useMemo(() => createTaskListTable(columns), [columns]);

  const assigneeOptions: IDropdownOption[] = React.useMemo(() => {
    const options: IDropdownOption[] = [{ key: '', text: strings.AllAssigneesOption }];
    if (!rows) {
      return options;
    }

    let hasUnassigned: boolean = false;
    const namesSeen: Record<string, boolean> = {};
    rows.forEach(row => {
      if (row.type === 'project') {
        return;
      }
      if (row.assignees.length === 0) {
        hasUnassigned = true;
      }
      row.assignees.forEach(name => {
        namesSeen[name] = true;
      });
    });

    if (hasUnassigned) {
      options.push({ key: UNASSIGNED_KEY, text: strings.UnassignedOption });
    }
    Object.keys(namesSeen).sort((a, b) => a.localeCompare(b)).forEach(name => {
      options.push({ key: name, text: name });
    });

    return options;
  }, [rows]);

  const visibleRows: IGanttRow[] | undefined = React.useMemo(() => {
    if (!rows || !selectedAssignee) {
      return rows;
    }

    const isMatch = (row: IGanttRow): boolean =>
      selectedAssignee === UNASSIGNED_KEY ? row.assignees.length === 0 : row.assignees.indexOf(selectedAssignee) !== -1;

    const bucketsWithVisibleTasks: Record<string, boolean> = {};
    rows.forEach(row => {
      if (row.type !== 'project' && row.project && isMatch(row)) {
        bucketsWithVisibleTasks[row.project] = true;
      }
    });

    return rows.filter(row => (row.type === 'project' ? !!bucketsWithVisibleTasks[row.id] : isMatch(row)));
  }, [rows, selectedAssignee]);

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

  const resolvedViewMode: ViewMode = ViewMode[ZOOM_LEVELS[zoomIndex].key];
  const rootClassName: string = showTaskNameOnBar ? styles.plannerGantt : `${styles.plannerGantt} ${styles.hideBarLabels}`;

  return (
    <div className={rootClassName}>
      <h2 className={styles.planTitle}>{props.planTitle}</h2>
      <div className={styles.toolbar}>
        <div className={styles.toolbarItem}>
          <Slider
            label={strings.ZoomSliderLabel}
            min={0}
            max={ZOOM_LEVELS.length - 1}
            step={1}
            value={zoomIndex}
            showValue={true}
            valueFormat={(value: number) => ZOOM_LEVELS[value].label}
            onChange={(value: number) => setZoomIndex(value)}
          />
        </div>
        <div className={styles.toolbarItem}>
          <Dropdown
            label={strings.AssigneeFilterLabel}
            selectedKey={selectedAssignee}
            options={assigneeOptions}
            onChange={(_event, option) => setSelectedAssignee(option ? String(option.key) : '')}
          />
        </div>
      </div>
      {visibleRows && visibleRows.length > 0 ? (
        <ColumnWidthsContext.Provider value={columnWidthsContextValue}>
          <Gantt
            tasks={toGanttTasks(visibleRows, colorBarsByStatus, { primary: themePrimary, secondary: themeSecondary, grey: themeGrey })}
            viewMode={resolvedViewMode}
            columnWidth={COLUMN_WIDTH_BY_VIEW_MODE[resolvedViewMode] || COLUMN_WIDTH_BY_VIEW_MODE.Week}
            listCellWidth="220px"
            rowHeight={42}
            ganttHeight={520}
            locale={navigator.language}
            todayColor={showCurrentDateLine ? TODAY_LINE_COLOR : TODAY_LINE_COLOR_HIDDEN}
            TaskListHeader={TaskListHeader}
            TaskListTable={TaskListTable}
          />
        </ColumnWidthsContext.Provider>
      ) : (
        <MessageBar messageBarType={MessageBarType.warning}>{strings.NoTasksMatchFilterLabel}</MessageBar>
      )}
    </div>
  );
};

export default PlannerGantt;
