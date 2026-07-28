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
- **Title (toggle).** The heading above the chart automatically mirrors the
  SharePoint page's own title (it's what the print/export window's title
  uses too, and falls back to the plan's name if the page has none yet),
  so there's no separate title field to keep in sync - just rename the
  page. Can be hidden entirely.
- **Open scrolled to today (toggle).** By default the chart otherwise opens
  scrolled to the very start of the plan's date range, which for a plan
  spanning several years means landing on its first weeks rather than
  anywhere near the present. When on, it opens scrolled to today's date
  instead, with a configurable margin (0-4 zoom periods) of context before
  it.
- **Gantt chart rendering** via [`gantt-task-react`](https://github.com/MaTeMaTuK/gantt-task-react).
- **Buckets as phases (toggle).** When on (default), each Planner bucket is
  shown as a bold phase bar spanning from the earliest start date to the
  latest end date of its tasks, labeled with the bucket's name; its tasks
  are indented underneath it in the task list. Turn it off to see a flat
  task list without the phase bars.
- **Bucket filter.** A checkbox per bucket in the property pane controls
  which buckets are shown at all, so a bucket like "Backlog" can be hidden
  from the chart entirely (its tasks are excluded, not just the phase bar).
- **Bucket order (toggle).** "Planner board order" (default) keeps buckets
  in the same left-to-right order as the Planner board itself; "By earliest
  start date" instead orders buckets by the earliest start date among their
  own tasks.
- **Sort tasks by start date (toggle).** Off by default, which keeps each
  bucket's tasks in Planner's own (manual/drag-and-drop) order. On sorts
  each bucket's tasks by start date instead.
- Tasks with a due date but no start date fall back to their creation date;
  tasks with no due date at all are rendered as milestones instead of bars.
- Zoom level (Day / Week / Month) and a "show completed tasks" toggle,
  configurable per web part instance from the property pane.
- **Color bars by status, no percentage shown anywhere.** Planner only has
  three real progress states (not started / in progress / done - 0%, 50%,
  100%), so instead of a numeric percentage the whole bar is colored by
  status: the site's secondary theme color for not-started, primary theme
  color for in progress, and grey for done. Can be turned off to use the
  chart's default single-color bars. The hover tooltip is also replaced
  (name + dates + duration only) since gantt-task-react's default tooltip
  always includes a "Progress: NN %" line.
- **Today indicator.** Today's column is highlighted; can be turned off.
- **Show task name on bars (toggle).** Turn off to remove the name label
  from the bars themselves (the task list on the left still always shows
  the full name).
- **Optional columns** in the task list next to the chart: start date, due
  date, who a task is assigned to, and its Planner label(s) (assignee names
  are resolved from Planner's assignee user ids via a single batched Graph
  call per load; labels come from the plan's own category names). Every
  visible column (including Name) can be resized by dragging its right edge.
- **Task info popup.** An "i" icon next to each task's name (toggleable)
  opens a popup with more detail, and each section in it can be turned on
  or off independently: the task's description (always shown first when
  enabled), start date, due date, assigned to, status, and its Planner
  comments/conversation (scrollable if there are many). The description and
  comments are only fetched from Graph the first time a given task's popup
  is opened, not for every task up front.
- **Live toolbar above the chart** (viewer-side, not saved to the page).
  Each control can be turned on or off from the property pane's **Viewer
  toolbar** group (the toolbar disappears entirely if all four are off):
  - A **completed-tasks toggle** ("All tasks" / "Hide completed") letting
    each visitor override the **Show completed tasks by default** property
    for their own view.
  - A **zoom slider** that controls how many of the chosen time unit
    (Day/Week/Month, set via the property pane) fit in view at once - drag
    it to zoom out from e.g. 3 weeks visible to 10+, or zoom in for more
    detail. This reflows the column width to fit the chart's actual
    rendered width, it doesn't just switch time units. Its starting
    position is set by the **Default zoom level** property (2-16).
  - An **"Assigned to" filter** that hides tasks (and any phase left with
    nothing under it) that don't match the selected person, including an
    "Unassigned" option.
  - A **label filter**, working the same way, for the plan's Planner
    categories/labels, including a "No label" option. Both filters apply
    together when both are in use.
  - A **Print / export to PDF** button. It opens the chart (title + Gantt,
    exactly as currently zoomed/filtered) in its own window with none of
    the page's other chrome, and triggers the browser's print dialog there
    - pick "Save as PDF" as the destination for an export. Regular Ctrl+P
    on the page itself also hides the toolbar via print CSS, but will still
    include whatever else is on the page since that's outside this web
    part's control.
- **Click a task to open it in Planner.** Both the bar itself and its name
  in the task list open that exact task, on its own plan board, at
  `planner.cloud.microsoft` in a new tab. Phase/bucket rows aren't
  clickable - they're a synthetic summary, not a real Planner task.
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
   plan from every Microsoft 365 group you belong to). The heading above
   the chart shows the page's own title; turn it off with **Show title**.
3. Optionally adjust, under **Display**: **Zoom level**, **Show completed
   tasks**, **Show buckets as phases**, **Bucket order** (board order vs.
   by earliest start date), **Sort tasks by start date**, **Color bars by
   status**, **Show task name on bars**, **Highlight today's date**, **Open
   scrolled to today** and its margin.
4. Under **Columns**, toggle the **Start date**, **Due date** and **Assigned
   to** columns next to the chart.
5. Under **Buckets to show**, uncheck a bucket (e.g. Backlog) to hide it and
   its tasks from the chart entirely.
6. Under **Viewer toolbar**, set the **Default zoom level** and turn the
   **completed-tasks toggle**, **zoom slider**, **Assigned to filter** and
   **print button** on or off for visitors.
7. Save the page. Visitors then see that plan's Gantt chart, scoped to
   whatever Planner tasks they're individually allowed to see, and can each
   independently use whichever toolbar controls are enabled and **drag
   column borders** to resize them, without affecting what anyone else sees
   or changing the saved page. Clicking a task (bar or name) opens it in
   Planner in a new tab.

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
- **Task info popup comments use a beta Graph endpoint.** The current
  Planner task chat ("Taakchat") is only exposed via
  `/planner/tasks/{id}/messages?$select=...` on the **beta** Graph API - the
  older, v1.0 `conversationThreadId`/group-conversation route is Planner's
  previous comment system and often doesn't reflect what today's UI
  actually shows. Beta endpoints can change without notice; if Microsoft
  ships a v1.0 equivalent, switch to it.

## Copyright

Built for Andersom BV. © 2026 Andersom BV.

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Use Microsoft Graph in your solution](https://docs.microsoft.com/sharepoint/dev/spfx/web-parts/get-started/using-microsoft-graph-apis)
- [Planner Graph API overview](https://learn.microsoft.com/graph/api/resources/planner-overview)
- [Publish SharePoint Framework applications to the Marketplace](https://docs.microsoft.com/sharepoint/dev/spfx/publish-to-marketplace-overview)
