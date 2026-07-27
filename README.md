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
- **Show task name on bars (toggle).** Turn off to remove the name label
  from the bars themselves (the task list on the left still always shows
  the full name).
- **Optional columns** in the task list next to the chart: start date, due
  date, and who a task is assigned to (assignee names are resolved from
  Planner's assignee user ids via a single batched Graph call per load).
  Every visible column (including Name) can be resized by dragging its
  right edge.
- **Live toolbar above the chart** (viewer-side, not saved to the page):
  - A **zoom slider** that controls how many of the chosen time unit
    (Day/Week/Month, set via the property pane) fit in view at once - drag
    it to zoom out from e.g. 3 weeks visible to 10+, or zoom in for more
    detail. This reflows the column width to fit the chart's actual
    rendered width, it doesn't just switch time units.
  - An **"Assigned to" filter** that hides tasks (and any phase left with
    nothing under it) that don't match the selected person, including an
    "Unassigned" option.
  - A **Print / export to PDF** button. It opens the chart (title + Gantt,
    exactly as currently zoomed/filtered) in its own window with none of
    the page's other chrome, and triggers the browser's print dialog there
    - pick "Save as PDF" as the destination for an export. Regular Ctrl+P
    on the page itself also hides the toolbar via print CSS, but will still
    include whatever else is on the page since that's outside this web
    part's control.
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
   **Color bars by status**, **Show task name on bars**, and **Highlight
   today's date**.
4. Under **Columns**, toggle the **Start date**, **Due date** and **Assigned
   to** columns next to the chart.
5. Under **Buckets to show**, uncheck a bucket (e.g. Backlog) to hide it and
   its tasks from the chart entirely.
6. Save the page. Visitors then see that plan's Gantt chart, scoped to
   whatever Planner tasks they're individually allowed to see, and can each
   independently use the **zoom slider**, **Assigned to filter** and
   **Print / export to PDF** button above the chart, and **drag column
   borders** to resize them, without affecting what anyone else sees or
   changing the saved page.

## Solution architecture

- `src/webparts/plannerGantt/PlannerGanttWebPart.ts` — web part shell and
  property pane (async plan picker via `MSGraphClientV3`).
- `src/webparts/plannerGantt/services/PlannerService.ts` — all Graph calls
  and the mapping from Planner tasks/buckets to Gantt rows.
- `src/webparts/plannerGantt/components/PlannerGantt.tsx` — React component
  that fetches the selected plan's tasks and renders `gantt-task-react`.
- `src/webparts/plannerGantt/components/GanttTaskList.tsx` — custom task
  list header/table so the start/due/assignee columns can be toggled and
  resized (`gantt-task-react`'s built-in table only ever shows name/from/to
  at a single fixed width). Column widths live in a small React Context
  since the header and body are separate component instances gantt-task-react
  renders independently.
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
- **The on-bar task name can only be shown or hidden, not aligned.**
  `gantt-task-react` centers the label inside the bar if it fits, otherwise
  pushes it just outside to the right — that x-position is computed
  internally and isn't exposed as a prop, so left/center/right alignment
  of that specific label isn't achievable without forking the library.
  Same underlying cause as the Year-zoom gap above: it would take
  switching (or patching) the charting library to unlock.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Planner Graph API overview](https://learn.microsoft.com/graph/api/resources/planner-overview)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
