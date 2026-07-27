/**
 * A Planner plan, together with the Microsoft 365 group it belongs to.
 * Microsoft Graph has no single endpoint that lists every plan a user can
 * see, so plans are discovered by walking the groups the user is a member
 * of and asking each group for its plans.
 */
export interface IPlannerPlanOption {
  planId: string;
  planTitle: string;
  groupId: string;
  groupName: string;
}

export interface IPlannerBucket {
  id: string;
  name: string;
  orderHint: string;
}

/** Bucket id -> whether it should be shown. A bucket missing from the map is shown by default. */
export type IBucketFilter = Record<string, boolean>;

export interface IPlannerTask {
  id: string;
  title: string;
  bucketId: string;
  percentComplete: number;
  priority: number;
  startDateTime?: string;
  dueDateTime?: string;
  createdDateTime: string;
  completedDateTime?: string;
  assigneeCount: number;
}

/**
 * A task shaped for rendering with gantt-task-react, after resolving the
 * date fallbacks Planner tasks don't guarantee (start/due are both optional
 * in the Planner UI).
 */
export interface IGanttRow {
  id: string;
  name: string;
  start: Date;
  end: Date;
  progress: number;
  type: 'task' | 'project' | 'milestone';
  project?: string;
  bucketName: string;
}
