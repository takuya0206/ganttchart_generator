



# Gantt Chart Generator
This is an add-on for Google apps script. You can automatically create a gantt chart, which is suitable to manage a small or middle project. English & Japanese are available.



![f:id:takuya0206:20171223222331g:plain](https://cdn-ak.f.st-hatena.com/images/fotolife/t/takuya0206/20171223/20171223222331.gif)


## Installation

[Gantt Chart Generator - Google Sheets add-on](https://chrome.google.com/webstore/detail/gantt-chart-generator/bnaicalmdphddkedcgchnfbjohmhdgni?utm_source=permalink)

Access the above URL, log in your Google account, and click "＋ Free."


## What You Can Do

* Break down tasks into five levels
* Automatically paint a chart based on a date you enter
* Automatically calculate workloads related to a parent task
* Automatically place bars based on progress you enter
* Automatically calculate progress related to a parent task based on workload weight
* Edit holidays as you want (*the default is Japanese holidays)


## Specification

### Add-on Menu

Item         | Action
---------- | -------------------------
Create Gantt Chart | Create a schedule sheet and a holiday sheet
Show Sidebar   | Show a sidebar

### scheudle sheet

Item           | Input  | Action
------------ | --- | ----------------------------------------
Work Breakdown Structure    | String | Assign task ID
Planned Start & Planned Finish | Date  | Paint a chart in blu
Planned Finish         | Date  | Set a milestone in orange
Actual Start & Actual Finish | Date  | Paint a chart in green（*be hidden in the default）
Worklaod(plan)        | Number  | Calculate parent's workload (sum childrens' workload)
Progress           | Number  | Place bars in a chart & calculate parent's progress if tasks have workload (plan)


### holiday sheet

Item | Input | Actuon
-- | -- | -------------------
A column | Date | Make a holiday line in pink

### Sidebar

Item           | Actio
------------ | --------------------------
Change Start Date and Chart Width       | Change start date<br />Change chart width in week
Recalculate Workload (plan) & Progress | Calculate all parents' worklaod and progress
Repaint Gantt Chart | Repaint all of the Chart
Show Color Indication      | Indicate progress like blue means "completed" or "not start," yellow means "in progress" and red means "delayed"
Show Parents' Charts      | Automatically show total duration of parents' charts
Initalize Gantt Chart          | Initalize schedule sheet and holiday sheet

## Recommended Usage

* Place a project name in the top hierarchy and make all of the tasks its children. That is to say, you can automatically calculate progress of your project.
* Break down tasks in detail like Parent tasks, child tasks and grandchild tasks.
* Place planned start and planned finish in tasks which you have to watch daily or weekly.
* Use "show parents' charts" when you check progress from a broad viewpoint.
* Place workload in man-day.
* Utilize "workday function." You can refer to A column in holiday sheet and workload (plan).
* Watch progress by using progress bars or actual start and actual finish.
* If you use progress bars, you may want to active "Color Indication."
* If you update numbers in progress column by using functions, charts won't be automatically updated. You need to use "Repaint Gantt Chart."

## Restriction

* Do not change the sheet name (schedule sheet and holiday sheet).
* Do not insert rows before the item row.
* Do not edit or delete the hidden second row.
* Do not insert columns between the start and the finish date column and after the progress column.
* Do not edit date in a gantt chart

## License
GNU General Public License (GPL)

## FAQ

 - The painting function does not work in my gantt chart..

This add-on is sometimes updated without notification. In most of those cases, you can activate the new version by using "Change Start Date," If your gantt chart still does not work, you may have to create a new spreadsheet and use this add-on from scratch. Then, you transfer the date from the previous one by copy & paste.

 - I got a system error when using sidebar...

There is possibility that a problem related to authority has occurred. Please try to delete browser cache. If you are using Google docs with more than two IDs at the same time, please use only one ID or utilize something like a secret browser which means there is no cache.

 - Date on my gantt chart is not correct...

There is a possibility that the timezone on your laptop differs from the timezone on your spreadsheet. Please check the setting and ensure the same timezone is used.
[*How to check the timezone on your spreadsheet](https://support.google.com/docs/answer/58515?co=GENIE.Platform%3DDesktop&hl=en)  

After version 30, as a workaround, you can manually change the first date in your gantt chart (*the following dates automatically change). Then, by using "Repaint Gantt Chart" and editing the holiday sheet, your gantt chart will be fixed.

 - Date is not correct after a specific month

Daylight saving time may cause this kind of problem. As a workaround, you can solve this by adding balance values like "2019/04/01 1:00:00" as [this picture shown](https://github.com/takuya0206/ganttchart_generator/issues/4#issuecomment-465597251).


 - How do I extend the width of my gantt chart?

You can change the width on the sidebar. However, it is not recommended that you extend the width too much because processing speed in spread sheets gets slow as the number of columns increases.


 - How do I change the color of charts?

The color of charts can not be changed as you like. In the current specification,  Show Color Indication is the function related to change of color, which shows blue means "completed" or "not yet start," yellow means "in progress," and red means "delayed".


 - Parents' progress always show zero...

Parents' progress are calculated based on the weighted average of their child's workload, which means that parents’s progress always shows 0% unless you enter "planned workload." Note: "Actual workload" does not have to do with any program. This is just space to make a note which allows you to look back your project after it finished.


## Privacy Policy
We treat your privacy with respect and it is secured and will never be sold, shared or rented to third parties.

### Information We Collect
In operating our add-on, we may collect and process the following data about you:

* Details of your visits to our website and the resources that you access, including, but not limited to, traffic data, location data, weblogs and other communication data
* Information that you provide by filling in forms on our website, such as when you registered for information or make a purchase
* Information provided to us when you communicate with us for any reason.
