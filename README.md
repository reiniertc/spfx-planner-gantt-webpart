# spfx-planner-gantt-webpart

A SharePoint Framework (SPFx) web part that reads a Microsoft Planner plan
through Microsoft Graph and renders its tasks as an interactive Gantt chart.
The site owner picks the plan once in the web part's property pane; every
visitor then sees that plan's tasks, grouped by bucket, as a Gantt chart
(no re-configuration per visitor).

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant) with Microsoft Planner enabled
- [Microsoft Graph](https://learn.microsoft.com/graph/use-the-api) (delegated permissions)

## Features

- **Plan picker in the property pane.** The dropdown is populated live from
  Microsoft Graph: it walks the Microsoft 365 groups the signed-in user
  belongs to and lists every Planner plan owned by those groups (Graph has
  no single "list all my plans" endpoint, so this is the standard way to
  discover them).
- **Gantt chart rendering** via [`gantt-task-react`](https://github.com/MaTeMaTuK/gantt-task-react).
- **Buckets as phases (toggle).** When on (default), each Planner bucket is
  shown as a bold phase bar spanning from the earliest start date to the
  latest end date of its tasks, labeled with the bucket's name; its tasks
  are indented underneath it in the task list. Turn it off to see a flat
  task list without the phase bars.
- **Bucket filter.** A checkbox per bucket in the property pane controls
  which buckets are shown at all, so a bucket like "Backlog" can be hidden
  from the chart entirely (its tasks are excluded, not just the phase bar).
- **Sort tasks by start date (toggle).** Off by default, which keeps each
  bucket's tasks in Planner's own (manual/drag-and-drop) order. On sorts
  each bucket's tasks by start date instead.
- Tasks with a due date but no start date fall back to their creation date;
  tasks with no due date at all are rendered as milestones instead of bars.
- Zoom level (Day / Week / Month) and a "show completed tasks" toggle,
  configurable per web part instance from the property pane.
- **Color bars by status.** Planner only has three real progress states
  (not started / in progress / done - 0%, 50%, 100%), so instead of a
  numeric percentage the whole bar is colored by status: the site's
  secondary theme color for not-started, primary theme color for in
  progress, and grey for done. Can be turned off to use the chart's
  default single-color bars.
- **Today indicator.** Today's column is highlighted; can be turned off.
- **Optional columns** in the task list next to the chart: start date, due
  date, and who a task is assigned to (assignee names are resolved from
  Planner's assignee user ids via a single batched Graph call per load).
- Read-only by design: Planner remains the system of record. The chart does
  not write back to Planner.

## Prerequisites

This web part calls Microsoft Graph on behalf of the signed-in user, so the
following Graph **delegated** permissions must be approved once per tenant:

| Permission        | Why it's needed                                              |
| ------------------ | ------------------------------------------------------------- |
| `Group.Read.All`   | Enumerate the Microsoft 365 groups the user belongs to, to find their Planner plans |
| `Tasks.Read`       | Read plan buckets and tasks from Planner                     |
| `User.ReadBasic.All` | Resolve assignee user ids to display names for the "Assigned to" column |

These are declared in [`config/package-solution.json`](./config/package-solution.json)
under `webApiPermissionRequests`. After deploying the `.sppkg`:

1. Go to the **SharePoint Admin Center** → **Advanced** → **API access**.
2. Approve the three pending Microsoft Graph requests for this solution.
3. Users can only see plans/tasks they already have access to in Planner —
   this web part does not elevate anyone's permissions.

## Build and deploy

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

This produces `sharepoint/solution/spfx-planner-gantt-webpart.sppkg`. Upload
it to your tenant's [App Catalog](https://learn.microsoft.com/sharepoint/use-app-catalog),
approve the API permissions as described above, then add the **Planner
Gantt** web part to any modern page.

For local development against your own tenant:

```bash
npm install
gulp serve
```

## Using the web part

1. Add the **Planner Gantt** web part to a page and edit it.
2. In the property pane, pick a **Plan** from the dropdown (it lists every
   plan from every Microsoft 365 group you belong to).
3. Optionally adjust, under **Display**: **Zoom level**, **Show completed
   tasks**, **Show buckets as phases**, **Sort tasks by start date**,
   **Color bars by status**, and **Highlight today's date**.
4. Under **Columns**, toggle the **Start date**, **Due date** and **Assigned
   to** columns next to the chart.
5. Under **Buckets to show**, uncheck a bucket (e.g. Backlog) to hide it and
   its tasks from the chart entirely.
6. Save the page — all visitors now see that plan's Gantt chart, scoped to
   whatever Planner tasks they're individually allowed to see.

## Solution architecture

- `src/webparts/plannerGantt/PlannerGanttWebPart.ts` — web part shell and
  property pane (async plan picker via `MSGraphClientV3`).
- `src/webparts/plannerGantt/services/PlannerService.ts` — all Graph calls
  and the mapping from Planner tasks/buckets to Gantt rows.
- `src/webparts/plannerGantt/components/PlannerGantt.tsx` — React component
  that fetches the selected plan's tasks and renders `gantt-task-react`.
- `src/webparts/plannerGantt/components/GanttTaskList.tsx` — custom task
  list header/table so the start/due/assignee columns can be toggled
  (`gantt-task-react`'s built-in table only ever shows name/from/to).
- `src/webparts/plannerGantt/models/IPlannerModels.ts` — shared TypeScript
  interfaces for plans, buckets, tasks and Gantt rows.

## Known limitations

- Planner tasks have no notion of dependencies between tasks, so the chart
  never draws dependency arrows.
- A plan with many buckets/tasks means many Graph round trips on first load
  (one per group to discover plans); this is a Graph API constraint, not a
  caching layer this web part currently implements.
- **Zoom only goes up to Month.** [`gantt-task-react`](https://github.com/MaTeMaTuK/gantt-task-react)
  (the charting library used here) doesn't have a Year view mode — its
  coarsest zoom level is Month. A React-17-compatible version with Year
  support doesn't exist upstream; adding one would mean switching to a
  different charting library. Open an issue/PR if that's a hard
  requirement for your use case.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Planner Graph API overview](https://learn.microsoft.com/graph/api/resources/planner-overview)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
