import { MSGraphClientV3 } from '@microsoft/sp-http';
import {
  IPlannerPlanOption,
  IPlannerBucket,
  IPlannerTask,
  IGanttRow,
  IBucketFilter,
  IPlannerTaskComment
} from '../models/IPlannerModels';

const MS_PER_DAY: number = 24 * 60 * 60 * 1000;

export interface IGanttRowOptions {
  includeCompleted: boolean;
  showBucketsAsPhases: boolean;
  /** Bucket id -> whether to include it. A bucket missing from the map is included by default. */
  bucketFilter?: IBucketFilter;
  /** When true, tasks within each bucket are ordered by start date instead of their Planner (manual) order. */
  sortByStartDate: boolean;
  /** When true, buckets are ordered by the earliest start date among their tasks instead of their Planner board order. */
  sortBucketsByStartDate: boolean;
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
   * Since Planner plans in private/shared Teams channels live in their own
   * container (not the team's group-level plan list), each group's private
   * and shared channels are also checked for their own channel-scoped plans.
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
        // Channel discovery is the newer, less-proven half of this call - a
        // single group's channel plans hanging or misbehaving must never be
        // able to hold up (or fail) the group's own, already-working plans.
        const [groupPlans, channelPlans] = await Promise.all([
          this.getGroupPlans(client, group),
          this.withTimeout(this.getChannelPlans(client, group), [], 15000)
        ]);
        return groupPlans.concat(channelPlans);
      })
    );

    const allPlans: IPlannerPlanOption[] = plansByGroup.reduce((all, plans) => all.concat(plans), []);
    const allPlanIds: string[] = allPlans.map(plan => plan.planId);
    const distinctPlans: IPlannerPlanOption[] = allPlans.filter((plan, index) => allPlanIds.indexOf(plan.planId) === index);

    return distinctPlans.sort((a, b) => a.planTitle.localeCompare(b.planTitle));
  }

  /** Bounds how long a single promise can hold up the rest of the plan list, falling back instead of hanging forever. */
  private withTimeout<T>(promise: Promise<T>, fallback: T, ms: number): Promise<T> {
    return new Promise<T>(resolve => {
      const timer: number = window.setTimeout(() => resolve(fallback), ms);
      promise
        .then(value => {
          window.clearTimeout(timer);
          resolve(value);
        })
        .catch(() => {
          window.clearTimeout(timer);
          resolve(fallback);
        });
    });
  }

  private async getGroupPlans(client: MSGraphClientV3, group: { id: string; displayName: string }): Promise<IPlannerPlanOption[]> {
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
  }

  /**
   * Private and shared channels each get their own plan container, separate
   * from their team's group-level one - not a Teams member, or a group with
   * no Teams provisioned on it, means this 404s and is skipped the same way.
   */
  private async getChannelPlans(client: MSGraphClientV3, group: { id: string; displayName: string }): Promise<IPlannerPlanOption[]> {
    try {
      // The $filter form of this call (membershipType eq 'private' or ... eq
      // 'shared') 404s outright rather than filtering, at least on some
      // tenants - fetch all channels and filter client-side instead.
      const channelsResponse = await client
        .api(`/teams/${group.id}/channels`)
        .select('id,displayName,membershipType')
        .get();
      const channels: Array<{ id: string; displayName: string; membershipType?: string }> = (channelsResponse.value || [])
        .filter((channel: { membershipType?: string }) => channel.membershipType === 'private' || channel.membershipType === 'shared');

      const plansByChannel: IPlannerPlanOption[][] = await Promise.all(
        channels.map(async channel => {
          try {
            const plansResponse = await client
              .api(`/teams/${group.id}/channels/${channel.id}/planner/plans`)
              .version('beta')
              .select('id,title')
              .get();
            return (plansResponse.value || []).map((plan: { id: string; title: string }) => ({
              planId: plan.id,
              planTitle: plan.title,
              groupId: group.id,
              groupName: group.displayName,
              channelName: channel.displayName
            }));
          } catch {
            return [];
          }
        })
      );

      return plansByChannel.reduce((all, plans) => all.concat(plans), []);
    } catch {
      return [];
    }
  }

  public async getBuckets(planId: string): Promise<IPlannerBucket[]> {
    const client: MSGraphClientV3 = await this.graphClientFactory();
    const bucketsResponse = await client.api(`/planner/plans/${planId}/buckets`).select('id,name,orderHint').get();

    return (bucketsResponse.value || [])
      .sort((a: IPlannerBucket, b: IPlannerBucket) => a.orderHint.localeCompare(b.orderHint));
  }

  public async getGanttRows(planId: string, options: IGanttRowOptions): Promise<IGanttRow[]> {
    const client: MSGraphClientV3 = await this.graphClientFactory();

    const [buckets, tasksResponse, categoryNameByKey] = await Promise.all([
      this.getBuckets(planId),
      client
        .api(`/planner/plans/${planId}/tasks`)
        .select('id,title,bucketId,orderHint,percentComplete,priority,startDateTime,dueDateTime,createdDateTime,completedDateTime,assignments,appliedCategories,hasDescription')
        .get(),
      this.getCategoryNames(client, planId)
    ]);

    const bucketNameById: Record<string, string> = {};
    buckets.forEach(bucket => {
      bucketNameById[bucket.id] = bucket.name;
    });

    const isBucketVisible = (bucketId: string): boolean =>
      !options.bucketFilter || options.bucketFilter[bucketId] !== false;

    const toLabels = (appliedCategories?: Record<string, boolean>): string[] =>
      Object.keys(appliedCategories || {})
        .filter(key => appliedCategories && appliedCategories[key] && categoryNameByKey[key])
        .map(key => categoryNameByKey[key]);

    const tasks: IPlannerTask[] = (tasksResponse.value || [])
      .filter((task: { completedDateTime?: string }) => options.includeCompleted || !task.completedDateTime)
      .filter((task: { bucketId: string }) => isBucketVisible(task.bucketId))
      .map((task: {
        id: string; title: string; bucketId: string; orderHint: string; percentComplete: number; priority: number;
        startDateTime?: string; dueDateTime?: string; createdDateTime: string; completedDateTime?: string;
        assignments?: Record<string, unknown>; appliedCategories?: Record<string, boolean>;
        hasDescription: boolean;
      }) => ({
        id: task.id,
        title: task.title || '(Untitled task)',
        bucketId: task.bucketId,
        orderHint: task.orderHint,
        percentComplete: task.percentComplete,
        priority: task.priority,
        startDateTime: task.startDateTime,
        dueDateTime: task.dueDateTime,
        createdDateTime: task.createdDateTime,
        completedDateTime: task.completedDateTime,
        assigneeIds: Object.keys(task.assignments || {}),
        labels: toLabels(task.appliedCategories),
        hasDescription: !!task.hasDescription
      }));

    const assigneeNameById: Record<string, string> = await this.resolveAssigneeNames(client, tasks);

    const rows: IGanttRow[] = [];

    const bucketMinStartById: Record<string, number> = {};
    tasks.forEach(task => {
      const start: number = this.getEffectiveStart(task).getTime();
      if (bucketMinStartById[task.bucketId] === undefined || start < bucketMinStartById[task.bucketId]) {
        bucketMinStartById[task.bucketId] = start;
      }
    });

    const orderedBuckets: IPlannerBucket[] = buckets.filter(bucket => tasks.some(task => task.bucketId === bucket.id));
    if (options.sortBucketsByStartDate) {
      // gantt-task-react renders rows in exactly the order they're supplied
      // (verified empirically against the pinned 0.3.8 build - no internal
      // reversal or resort by date), so ascending here puts the
      // earliest-starting bucket first, i.e. at the top.
      orderedBuckets.sort((a, b) => bucketMinStartById[a.id] - bucketMinStartById[b.id]);
    }

    orderedBuckets.forEach(bucket => {
      const bucketTasks: IPlannerTask[] = tasks
        .filter(task => task.bucketId === bucket.id)
        .sort((a, b) => this.compareTasks(a, b, options.sortByStartDate));
      const taskRows: IGanttRow[] = bucketTasks.map(task =>
        this.toGanttRow(task, bucket.id, options.showBucketsAsPhases, assigneeNameById));

      if (options.showBucketsAsPhases) {
        rows.push(this.toBucketProjectRow(bucket.id, bucket.name, taskRows));
      }
      rows.push(...taskRows);
    });

    // Tasks whose bucket no longer resolves (rare, but Graph doesn't guarantee referential integrity on read).
    const orphanTasks: IPlannerTask[] = tasks
      .filter(task => !bucketNameById[task.bucketId])
      .sort((a, b) => this.compareTasks(a, b, options.sortByStartDate));
    if (orphanTasks.length > 0) {
      const orphanRows: IGanttRow[] = orphanTasks.map(task =>
        this.toGanttRow(task, 'unbucketed', options.showBucketsAsPhases, assigneeNameById));
      if (options.showBucketsAsPhases) {
        rows.push(this.toBucketProjectRow('unbucketed', 'Other tasks', orphanRows));
      }
      rows.push(...orphanRows);
    }

    return rows;
  }

  /**
   * The plan's up-to-25 category slots each have an optional display name
   * (empty/null slots are unused labels), fetched once per plan load rather
   * than per task.
   */
  private async getCategoryNames(client: MSGraphClientV3, planId: string): Promise<Record<string, string>> {
    try {
      const response = await client.api(`/planner/plans/${planId}/details`).select('categoryDescriptions').get();
      const descriptions: Record<string, string | null> = response.categoryDescriptions || {};
      const nameByKey: Record<string, string> = {};
      Object.keys(descriptions).forEach(key => {
        if (descriptions[key]) {
          nameByKey[key] = descriptions[key] as string;
        }
      });
      return nameByKey;
    } catch {
      return {};
    }
  }

  /** Fetches a task's notes on demand - not part of the main task list call since most tasks are never opened. */
  public async getTaskNotes(taskId: string): Promise<string> {
    const client: MSGraphClientV3 = await this.graphClientFactory();
    const response = await client.api(`/planner/tasks/${taskId}/details`).select('description').get();
    return response.description || '';
  }

  /**
   * The modern Planner task chat ("Taakchat" in the current UI) is served by
   * a dedicated, currently beta-only Graph endpoint - the older approach of
   * following a task's conversationThreadId into the owning group's Outlook
   * conversation is Planner's previous/legacy comment system and often
   * doesn't reflect what today's UI actually shows. Resolved on demand, only
   * when the info popup's comments section is both enabled and opened.
   */
  public async getTaskComments(taskId: string): Promise<IPlannerTaskComment[]> {
    try {
      const client: MSGraphClientV3 = await this.graphClientFactory();

      const messagesResponse = await client
        .api(`/planner/tasks/${taskId}/messages`)
        .version('beta')
        .select('content,createdDateTime,createdBy')
        .get();

      const messages: Array<{
        content?: string;
        createdDateTime: string;
        createdBy?: { user?: { id?: string } };
      }> = messagesResponse.value || [];

      const userIds: string[] = messages
        .map(message => message.createdBy && message.createdBy.user && message.createdBy.user.id)
        .filter((id): id is string => !!id)
        .filter((id, index, all) => all.indexOf(id) === index);
      const nameById: Record<string, string> = await this.resolveUserNames(client, userIds);

      const comments: IPlannerTaskComment[] = messages.map(message => {
        const userId: string | undefined = message.createdBy && message.createdBy.user && message.createdBy.user.id;
        return {
          from: (userId && nameById[userId]) || '',
          createdDateTime: message.createdDateTime,
          bodyText: this.stripHtml(message.content || '')
        };
      });

      comments.sort((a, b) => new Date(b.createdDateTime).getTime() - new Date(a.createdDateTime).getTime());
      return comments;
    } catch {
      // Best-effort: missing permissions, beta endpoint hiccups, etc. shouldn't break the popup.
      return [];
    }
  }

  private stripHtml(html: string): string {
    const parsed: Document = new DOMParser().parseFromString(html, 'text/html');
    return parsed.body.textContent || '';
  }

  /**
   * Planner only gives us assignee user ids on the task; resolving them to
   * display names is a single batched Graph call for every distinct id
   * across the whole plan, rather than one call per task.
   */
  private async resolveAssigneeNames(client: MSGraphClientV3, tasks: IPlannerTask[]): Promise<Record<string, string>> {
    const allIds: string[] = tasks.reduce<string[]>((ids, task) => ids.concat(task.assigneeIds), []);
    const distinctIds: string[] = allIds.filter((id, index) => allIds.indexOf(id) === index);
    return this.resolveUserNames(client, distinctIds);
  }

  /** Batch-resolves AAD user ids to display names in a single Graph call. */
  private async resolveUserNames(client: MSGraphClientV3, userIds: string[]): Promise<Record<string, string>> {
    if (userIds.length === 0) {
      return {};
    }

    try {
      const response = await client
        .api('/directoryObjects/getByIds')
        .post({ ids: userIds, types: ['user'] });

      const nameById: Record<string, string> = {};
      (response.value || []).forEach((user: { id: string; displayName?: string }) => {
        nameById[user.id] = user.displayName || user.id;
      });
      return nameById;
    } catch {
      // Best-effort: if resolution fails (e.g. missing User.ReadBasic.All), fall back to blank names.
      return {};
    }
  }

  private compareTasks(a: IPlannerTask, b: IPlannerTask, sortByStartDate: boolean): number {
    if (sortByStartDate) {
      return this.getEffectiveStart(a).getTime() - this.getEffectiveStart(b).getTime();
    }
    return a.orderHint.localeCompare(b.orderHint);
  }

  private getEffectiveStart(task: IPlannerTask): Date {
    return new Date(task.startDateTime || task.createdDateTime);
  }

  private toGanttRow(
    task: IPlannerTask,
    bucketId: string,
    showBucketsAsPhases: boolean,
    assigneeNameById: Record<string, string>
  ): IGanttRow {
    const start: Date = this.getEffectiveStart(task);
    const project: string | undefined = showBucketsAsPhases ? bucketId : undefined;
    const assignees: string[] = task.assigneeIds
      .map(id => assigneeNameById[id])
      .filter(name => !!name);

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
        bucketName: '',
        assignees,
        labels: task.labels,
        hasDescription: task.hasDescription
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
      bucketName: '',
      assignees,
      labels: task.labels,
      hasDescription: task.hasDescription
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
      bucketName,
      assignees: [],
      labels: [],
      hasDescription: false
    };
  }
}
