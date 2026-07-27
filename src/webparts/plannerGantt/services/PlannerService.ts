import { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  IPlannerPlanOption,
  IPlannerBucket,
  IPlannerTask,
  IGanttRow,
  IBucketFilter
} from '../models/IPlannerModels';

const MS_PER_DAY: number = 24 * 60 * 60 * 1000;

export interface IGanttRowOptions {
  includeCompleted: boolean;
  showBucketsAsPhases: boolean;
  /** Bucket id -> whether to include it. A bucket missing from the map is included by default. */
  bucketFilter?: IBucketFilter;
}

/**
 * Wraps the Microsoft Graph calls needed to turn a Planner plan into rows
 * a Gantt chart can render. Delegated auth via MSGraphClientV3, so results
 * are always scoped to what the signed-in user can already see in Planner.
 */
export class PlannerService {
  constructor(private graphClientFactory: () => Promise<MSGraphClientV3>) {}

  /**
   * Graph has no "list every plan I can see" endpoint, so plans are
   * discovered by walking the Microsoft 365 groups the user belongs to.
   */
  public async getAvailablePlans(): Promise<IPlannerPlanOption[]> {
    const client: MSGraphClientV3 = await this.graphClientFactory();

    const groupsResponse = await client
      .api('/me/memberOf/microsoft.graph.group')
      .select('id,displayName,groupTypes')
      .top(999)
      .get();

    const unifiedGroups: Array<{ id: string; displayName: string }> = (groupsResponse.value || [])
      .filter((group: { groupTypes?: string[] }) => (group.groupTypes || []).indexOf('Unified') !== -1);

    const plansByGroup: IPlannerPlanOption[][] = await Promise.all(
      unifiedGroups.map(async group => {
        try {
          const plansResponse = await client.api(`/groups/${group.id}/planner/plans`).select('id,title').get();
          return (plansResponse.value || []).map((plan: { id: string; title: string }) => ({
            planId: plan.id,
            planTitle: plan.title,
            groupId: group.id,
            groupName: group.displayName
          }));
        } catch {
          // The user is a group member but has no Planner access (or the
          // group simply has no plans) - skip it rather than fail the whole list.
          return [];
        }
      })
    );

    return plansByGroup
      .reduce((all, plans) => all.concat(plans), [])
      .sort((a, b) => a.planTitle.localeCompare(b.planTitle));
  }

  public async getBuckets(planId: string): Promise<IPlannerBucket[]> {
    const client: MSGraphClientV3 = await this.graphClientFactory();
    const bucketsResponse = await client.api(`/planner/plans/${planId}/buckets`).select('id,name,orderHint').get();

    return (bucketsResponse.value || [])
      .sort((a: IPlannerBucket, b: IPlannerBucket) => a.orderHint.localeCompare(b.orderHint));
  }

  public async getGanttRows(planId: string, options: IGanttRowOptions): Promise<IGanttRow[]> {
    const client: MSGraphClientV3 = await this.graphClientFactory();

    const [buckets, tasksResponse] = await Promise.all([
      this.getBuckets(planId),
      client
        .api(`/planner/plans/${planId}/tasks`)
        .select('id,title,bucketId,percentComplete,priority,startDateTime,dueDateTime,createdDateTime,completedDateTime,assignments')
        .get()
    ]);

    const bucketNameById: Record<string, string> = {};
    buckets.forEach(bucket => {
      bucketNameById[bucket.id] = bucket.name;
    });

    const isBucketVisible = (bucketId: string): boolean =>
      !options.bucketFilter || options.bucketFilter[bucketId] !== false;

    const tasks: IPlannerTask[] = (tasksResponse.value || [])
      .filter((task: { completedDateTime?: string }) => options.includeCompleted || !task.completedDateTime)
      .filter((task: { bucketId: string }) => isBucketVisible(task.bucketId))
      .map((task: {
        id: string; title: string; bucketId: string; percentComplete: number; priority: number;
        startDateTime?: string; dueDateTime?: string; createdDateTime: string; completedDateTime?: string;
        assignments?: Record<string, unknown>;
      }) => ({
        id: task.id,
        title: task.title || '(Untitled task)',
        bucketId: task.bucketId,
        percentComplete: task.percentComplete,
        priority: task.priority,
        startDateTime: task.startDateTime,
        dueDateTime: task.dueDateTime,
        createdDateTime: task.createdDateTime,
        completedDateTime: task.completedDateTime,
        assigneeCount: Object.keys(task.assignments || {}).length
      }));

    const rows: IGanttRow[] = [];

    buckets
      .filter(bucket => tasks.some(task => task.bucketId === bucket.id))
      .forEach(bucket => {
        const bucketTasks: IPlannerTask[] = tasks.filter(task => task.bucketId === bucket.id);
        const taskRows: IGanttRow[] = bucketTasks.map(task => this.toGanttRow(task, bucket.id, options.showBucketsAsPhases));

        if (options.showBucketsAsPhases) {
          rows.push(this.toBucketProjectRow(bucket.id, bucket.name, taskRows));
        }
        rows.push(...taskRows);
      });

    // Tasks whose bucket no longer resolves (rare, but Graph doesn't guarantee referential integrity on read).
    const orphanTasks: IPlannerTask[] = tasks.filter(task => !bucketNameById[task.bucketId]);
    if (orphanTasks.length > 0) {
      const orphanRows: IGanttRow[] = orphanTasks.map(task => this.toGanttRow(task, 'unbucketed', options.showBucketsAsPhases));
      if (options.showBucketsAsPhases) {
        rows.push(this.toBucketProjectRow('unbucketed', 'Other tasks', orphanRows));
      }
      rows.push(...orphanRows);
    }

    return rows;
  }

  private toGanttRow(task: IPlannerTask, bucketId: string, showBucketsAsPhases: boolean): IGanttRow {
    const created: Date = new Date(task.createdDateTime);
    const start: Date = task.startDateTime ? new Date(task.startDateTime) : created;
    const project: string | undefined = showBucketsAsPhases ? bucketId : undefined;

    // A Planner task with no due date has nothing to draw a bar between,
    // so it's rendered as a milestone at its start date instead.
    if (!task.dueDateTime) {
      return {
        id: task.id,
        name: task.title,
        start,
        end: start,
        progress: task.percentComplete || 0,
        type: 'milestone',
        project,
        bucketName: ''
      };
    }

    let end: Date = new Date(task.dueDateTime);
    if (end.getTime() <= start.getTime()) {
      end = new Date(start.getTime() + MS_PER_DAY);
    }

    return {
      id: task.id,
      name: task.title,
      start,
      end,
      progress: task.percentComplete || 0,
      type: 'task',
      project,
      bucketName: ''
    };
  }

  private toBucketProjectRow(bucketId: string, bucketName: string, taskRows: IGanttRow[]): IGanttRow {
    const starts: number[] = taskRows.map(row => row.start.getTime());
    const ends: number[] = taskRows.map(row => row.end.getTime());

    return {
      id: bucketId,
      name: bucketName,
      start: new Date(Math.min(...starts)),
      end: new Date(Math.max(...ends)),
      progress: Math.round(taskRows.reduce((sum, row) => sum + row.progress, 0) / taskRows.length),
      type: 'project',
      bucketName
    };
  }
}
