import * as React from 'react';
import { Gantt, ViewMode, Task as GanttTask } from 'gantt-task-react';
import 'gantt-task-react/dist/index.css';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Slider } from '@fluentui/react/lib/Slider';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Toggle } from '@fluentui/react/lib/Toggle';

import styles from './PlannerGantt.module.scss';
import type { IPlannerGanttProps } from './IPlannerGanttProps';
import { IGanttRow } from '../models/IPlannerModels';
import {
  IGanttChartTask, IColumnWidths, DEFAULT_COLUMN_WIDTHS, ColumnWidthsContext,
  createTaskListHeader, createTaskListTable, openPlannerTask
} from './GanttTaskList';
import * as strings from 'PlannerGanttWebPartStrings';

// How many of the chosen time unit (Day/Week/Month) should be visible at
// once - the slider adjusts this, and columnWidth is derived from it and
// the chart's actual rendered width, so "zooming out" really does bring
// more weeks/days/months into view instead of switching time units.
export const MIN_VISIBLE_UNITS: number = 2;
export const MAX_VISIBLE_UNITS: number = 16;
const MIN_COLUMN_WIDTH_PX: number = 28;
const FALLBACK_CONTAINER_WIDTH: number = 900;

const UNASSIGNED_KEY: string = '__unassigned__';

function clampVisibleUnits(value: number): number {
  return Math.min(MAX_VISIBLE_UNITS, Math.max(MIN_VISIBLE_UNITS, value));
}

/** Subtracts whole zoom periods (day/week/month) from a date, calendar-aware. */
function subtractUnits(date: Date, viewMode: string, units: number): Date {
  const result: Date = new Date(date.getTime());
  if (viewMode === 'Month') {
    result.setMonth(result.getMonth() - units);
  } else if (viewMode === 'Week') {
    result.setDate(result.getDate() - units * 7);
  } else {
    result.setDate(result.getDate() - units);
  }
  return result;
}

const TODAY_LINE_COLOR: string = 'rgba(232, 17, 35, 0.15)';
const TODAY_LINE_COLOR_HIDDEN: string = 'transparent';
const MS_PER_DAY: number = 24 * 60 * 60 * 1000;

interface IThemeColors {
  primary: string;
  secondary: string;
  grey: string;
}

interface ITooltipContentProps {
  task: GanttTask;
  fontSize: string;
  fontFamily: string;
}

// Replaces gantt-task-react's default tooltip, which always includes a
// "Progress: NN %" line - status is already conveyed by the bar color
// (see colorForStatus), so the percentage would just be noise here.
const PlannerTooltipContent: React.FC<ITooltipContentProps> = ({ task, fontSize, fontFamily }) => {
  const durationDays: number = Math.round((task.end.getTime() - task.start.getTime()) / MS_PER_DAY);
  return (
    <div className={styles.ganttTooltip} style={{ fontSize, fontFamily }}>
      <b style={{ fontSize: `calc(${fontSize} + 4px)` }}>{task.name}</b>
      <div className={styles.ganttTooltipLine}>{task.start.toLocaleDateString()} – {task.end.toLocaleDateString()}</div>
      {durationDays > 0 && <div className={styles.ganttTooltipLine}>{durationDays} day(s)</div>}
    </div>
  );
};

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

/**
 * Opens the chart in its own popup window (title + chart only, no toolbar,
 * no SharePoint page chrome) and triggers the browser's print dialog there,
 * so "Save as PDF" produces a clean export instead of the whole page.
 */
function handlePrintExport(chartElement: HTMLElement | undefined, title: string): void {
  if (!chartElement) {
    return;
  }

  const printWindow: Window | null = window.open('', '_blank');
  if (!printWindow) {
    return; // popup blocked - the browser will usually show its own notice
  }

  printWindow.document.title = title || document.title;
  printWindow.document.head.innerHTML = document.head.innerHTML;
  printWindow.document.body.style.margin = '0';
  printWindow.document.body.appendChild(chartElement.cloneNode(true));

  // Linked stylesheets copied above load asynchronously; give them a moment
  // before printing so the popup isn't printed unstyled.
  setTimeout(() => {
    printWindow.focus();
    printWindow.print();
  }, 500);
}

const PlannerGantt: React.FC<IPlannerGanttProps> = (props: IPlannerGanttProps) => {
  const {
    planId, viewMode, showCompletedTasks, showBucketsAsPhases, colorBarsByStatus, sortTasksByStartDate,
    sortBucketsByStartDate, showTaskNameOnBar, showCurrentDateLine, showStartDateColumn, showEndDateColumn,
    showAssigneeColumn, bucketFilter, plannerService, themePrimary, themeSecondary, themeGrey,
    defaultZoomLevel, showZoomControl, showPrintButton, showAssigneeFilter, showCompletedFilterControl,
    customTitle, showTitle, scrollToToday, scrollToTodayMarginUnits
  } = props;
  // Stringified so an equivalent-but-new object reference (the web part
  // shallow-copies bucketFilter on every render) doesn't trigger a refetch.
  const bucketFilterKey: string = JSON.stringify(bucketFilter || {});
  const clampedDefaultZoom: number = clampVisibleUnits(defaultZoomLevel);

  const [rows, setRows] = React.useState<IGanttRow[] | undefined>(undefined);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [error, setError] = React.useState<string | undefined>(undefined);

  // Live, viewer-side controls - not persisted to the web part's properties.
  const [visibleUnits, setVisibleUnits] = React.useState<number>(clampedDefaultZoom);
  const [selectedAssignee, setSelectedAssignee] = React.useState<string>('');
  const [liveShowCompleted, setLiveShowCompleted] = React.useState<boolean>(showCompletedTasks);
  const [columnWidths, setColumnWidths] = React.useState<IColumnWidths>(DEFAULT_COLUMN_WIDTHS);
  const [containerWidth, setContainerWidth] = React.useState<number>(FALLBACK_CONTAINER_WIDTH);
  const [chartWrapperEl, setChartWrapperEl] = React.useState<HTMLDivElement | null>(null);

  React.useEffect(() => {
    setVisibleUnits(clampedDefaultZoom);
    setSelectedAssignee('');
    setLiveShowCompleted(showCompletedTasks);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [planId, viewMode, clampedDefaultZoom, showCompletedTasks]);

  // Re-measures whenever the wrapper mounts/unmounts (e.g. switching between
  // the loading/empty/chart states) and whenever the page resizes it.
  React.useEffect(() => {
    if (!chartWrapperEl || typeof ResizeObserver === 'undefined') {
      return undefined;
    }
    const observer: ResizeObserver = new ResizeObserver(entries => {
      const width: number = entries[0].contentRect.width;
      if (width > 0) {
        setContainerWidth(width);
      }
    });
    observer.observe(chartWrapperEl);
    return () => observer.disconnect();
  }, [chartWrapperEl]);

  const setColumnWidth = React.useCallback((column: keyof IColumnWidths, width: number) => {
    setColumnWidths(previous => ({ ...previous, [column]: width }));
  }, []);

  const columnWidthsContextValue = React.useMemo(() => ({ widths: columnWidths, setColumnWidth }), [columnWidths, setColumnWidth]);

  // Computed once per plan/view-mode/setting change (not on every render),
  // so it positions the initial scroll without fighting the viewer's own
  // scrolling afterwards.
  const viewDate: Date | undefined = React.useMemo(() => {
    if (!scrollToToday) {
      return undefined;
    }
    return subtractUnits(new Date(), viewMode, scrollToTodayMarginUnits);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [planId, viewMode, scrollToToday, scrollToTodayMarginUnits]);

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
        includeCompleted: liveShowCompleted,
        showBucketsAsPhases,
        sortByStartDate: sortTasksByStartDate,
        sortBucketsByStartDate,
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
  }, [planId, liveShowCompleted, showBucketsAsPhases, sortTasksByStartDate, sortBucketsByStartDate, bucketFilterKey, plannerService]);

  const columns = React.useMemo(() => ({
    showStartDate: showStartDateColumn,
    showEndDate: showEndDateColumn,
    showAssignee: showAssigneeColumn
  }), [showStartDateColumn, showEndDateColumn, showAssigneeColumn]);

  const TaskListHeader = React.useMemo(() => createTaskListHeader(columns), [columns]);
  const TaskListTable = React.useMemo(() => createTaskListTable(columns, planId), [columns, planId]);

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

  const resolvedViewMode: ViewMode = ViewMode[viewMode as keyof typeof ViewMode] || ViewMode.Week;
  const columnWidth: number = Math.max(MIN_COLUMN_WIDTH_PX, Math.floor(containerWidth / visibleUnits));
  const zoomUnitLabel: string = ({
    Day: strings.ZoomUnitDay,
    Week: strings.ZoomUnitWeek,
    Month: strings.ZoomUnitMonth
  } as Record<string, string>)[viewMode] || strings.ZoomUnitWeek;
  const rootClassName: string = showTaskNameOnBar ? styles.plannerGantt : `${styles.plannerGantt} ${styles.hideBarLabels}`;
  const hasAnyToolbarControl: boolean = showZoomControl || showAssigneeFilter || showPrintButton || showCompletedFilterControl;

  const handleTaskSelect = (task: IGanttChartTask, isSelected: boolean): void => {
    if (isSelected && task.type !== 'project') {
      openPlannerTask(planId, task.id);
    }
  };

  return (
    <div className={rootClassName}>
      {hasAnyToolbarControl && (
        <div className={styles.toolbar}>
          {showCompletedFilterControl && (
            <div className={styles.toolbarItem}>
              <Toggle
                label={strings.CompletedTasksToggleLabel}
                onText={strings.CompletedTasksToggleOnText}
                offText={strings.CompletedTasksToggleOffText}
                checked={liveShowCompleted}
                onChange={(_event, checked) => setLiveShowCompleted(!!checked)}
              />
            </div>
          )}
          {showZoomControl && (
            <div className={styles.toolbarItem}>
              <Slider
                label={strings.ZoomSliderLabel}
                min={MIN_VISIBLE_UNITS}
                max={MAX_VISIBLE_UNITS}
                step={1}
                value={visibleUnits}
                showValue={true}
                valueFormat={(value: number) => `${value} ${zoomUnitLabel}`}
                onChange={(value: number) => setVisibleUnits(value)}
              />
            </div>
          )}
          {showAssigneeFilter && (
            <div className={styles.toolbarItem}>
              <Dropdown
                label={strings.AssigneeFilterLabel}
                selectedKey={selectedAssignee}
                options={assigneeOptions}
                onChange={(_event, option) => setSelectedAssignee(option ? String(option.key) : '')}
              />
            </div>
          )}
          {showPrintButton && (
            <div className={styles.toolbarItem}>
              <DefaultButton
                text={strings.PrintButtonLabel}
                iconProps={{ iconName: 'Print' }}
                onClick={() => handlePrintExport(chartWrapperEl || undefined, customTitle || props.planTitle)}
              />
            </div>
          )}
        </div>
      )}
      <div ref={setChartWrapperEl}>
        {showTitle && <h2 className={styles.planTitle}>{customTitle || props.planTitle}</h2>}
        {visibleRows && visibleRows.length > 0 ? (
          <ColumnWidthsContext.Provider value={columnWidthsContextValue}>
            <Gantt
              tasks={toGanttTasks(visibleRows, colorBarsByStatus, { primary: themePrimary, secondary: themeSecondary, grey: themeGrey })}
              viewMode={resolvedViewMode}
              viewDate={viewDate}
              columnWidth={columnWidth}
              listCellWidth="220px"
              rowHeight={42}
              ganttHeight={520}
              locale={navigator.language}
              todayColor={showCurrentDateLine ? TODAY_LINE_COLOR : TODAY_LINE_COLOR_HIDDEN}
              TaskListHeader={TaskListHeader}
              TaskListTable={TaskListTable}
              TooltipContent={PlannerTooltipContent}
              onSelect={handleTaskSelect}
            />
          </ColumnWidthsContext.Provider>
        ) : (
          <MessageBar messageBarType={MessageBarType.warning}>{strings.NoTasksMatchFilterLabel}</MessageBar>
        )}
      </div>
    </div>
  );
};

export default PlannerGantt;
