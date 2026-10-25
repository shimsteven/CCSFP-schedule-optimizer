# Schedule Optimizer

Finds the best combination of class sections for your timetable. It searches every
combination (branch-and-bound) and ranks them by what you care about: fewest conflicts,
fewest days at school, and your preferred sections per year level. Results can be
exported to Excel or CSV.

Two interchangeable front-ends share the same engine:

| Version | File | Look |
|---|---|---|
| Modern | `src/schedule_optimizer_ctk.py` | customtkinter, dark/light theme |
| Classic | `src/schedule_optimizer_gui.py` | plain tkinter |

## Download

Grab `ScheduleOptimizer-Modern.exe` or `ScheduleOptimizer-Classic.exe` from the
[Releases](../../releases) page. No Python needed.

> The exes are unsigned, so Windows SmartScreen may warn on first run
> ("More info" → "Run anyway"). Some antivirus tools also flag PyInstaller one-file builds.

## Run from source

Python 3.10+.

```
pip install -r requirements.txt
python src/schedule_optimizer_ctk.py     # modern
python src/schedule_optimizer_gui.py     # classic
```

`openpyxl` is only needed for Excel export; `customtkinter` only for the modern version.

## How to use

**Contents:**
[The idea](#the-idea) ·
[Quick start](#quick-start-with-the-sample-data) ·
[Step 1: enter schedules](#step-1--enter-your-schedules) ·
[Step 2: set options](#step-2--set-your-options) ·
[Step 3: results](#step-3--find-and-read-your-results) ·
[Saving your work](#saving-and-reopening-your-work) ·
[Recipes](#recipes) ·
[Troubleshooting](#troubleshooting)

### The idea

The app answers one question: *"Which sections should I enrol in so my week works out best?"*
It is built for **irregular students** who take different subjects from different sections
(for example a Year 2 lecture from section 2-B but its lab from 2-D, plus a Year 3 subject from 3-E).

A few terms used throughout:

| Term | Meaning | Example |
|---|---|---|
| **Section** | A group of students with a fixed weekly timetable. Its name starts with its year level. | `2-B`, `3-D` |
| **Meeting** | One weekly class: a section, day, subject, type, start and end time. | `2-B`, Mon, DSA, LEC, 5:30PM-7:30PM |
| **Type** | `LEC` (lecture), `LAB` (laboratory) or `ALL` (one combined class). | |
| **Subject part** | A subject + type, e.g. "DSA Lec", "DSA Lab" or an `ALL` subject like "QA". | |
| **Year level** | The leading number of a section name (`3-D` is Year 3). Sections that don't start with a number are grouped as "Other sections". | |

You tell the app every meeting of every section you could choose from. It then picks **exactly one
section for each subject part** (a subject's Lec and Lab may come from different sections), tries
**every possible combination**, and ranks them by what you care about most.

### Quick start with the sample data

1. Start the app. The sample data loads automatically (you can reload it any time with **Load sample data**).
2. Look through the three screens: **1 · Schedules** (the meetings), **2 · Options** (your priorities), **3 · Results**.
3. Press **▶ Find best schedules** (the bar at the bottom of the window).
4. The Results screen opens. Click different rows on the left to see each schedule's weekly timetable.

With the sample data and its default settings, the app searches 262,144 combinations of 13 subject parts
and the top schedule has **0.0 hours of conflict, 5 days at school**, and sections `2-B`/`2-A`/`2-D` for Year 2
and `3-D`/`3-E` for Year 3.

### Step 1 · Enter your schedules

Open **1 · Schedules**. There are four ways to get meetings in.

**A. One at a time.** Fill in the form and press **Add** (or press <kbd>Enter</kbd> in a text box):

| Field | What to enter |
|---|---|
| Section | The section name, e.g. `3-D` |
| Day | Mon to Sat |
| Subject | The subject name, e.g. `Data Structures`. Reuse the exact same spelling for the same subject in every section. |
| Type | `LEC`, `LAB` or `ALL` |
| Start / End | `7:00AM`, `2 PM`, `2:30pm`, `14:00` and `1430` all work. The end must be after the start. |

After adding, the Subject and time boxes clear so you can type the next meeting quickly; Section, Day and Type stay put.

**B. Paste rows** (fastest, e.g. straight from Excel). Click **Paste rows…** and enter one meeting per line as
`Day, Subject, Type, Start, End` (full day names like `Monday` are fine). Commas, tabs or `|` all work as separators. Choose the section in the box at the top
and it is applied to every line that doesn't have its own. If a line has **6** values, the first is read as the section:

```
Mon, DS and A, LEC, 3:30PM, 5:30PM
Mon, AdvNet, LAB, 6:00PM, 9:00PM
3-E, Thu, QA, ALL, 2:00PM, 3:30PM
```

If any line has a problem, nothing is added and the app lists the lines to fix. Exact duplicates are skipped.

**C. Import CSV.** Click **Import CSV…** and choose a file with a header row and these columns
(any order, case doesn't matter; `part` is accepted instead of `type`):

```
section,day,subject,type,start,end
3-D,Mon,DS and A,LEC,3:30PM,5:30PM
3-D,Mon,AdvNet,LAB,6:00PM,9:00PM
```

Unreadable rows are skipped, and the app tells you which file lines they were. **Export CSV…** writes your
current meetings in this same format, so you can edit them in a spreadsheet and import them back.

**D. Load sample data** replaces everything with the built-in example (it asks first if you already have data).

**Editing and removing.**
- Click a row in the table to load it into the form, change what you need, then press **Save changes**.
- Select one or more rows (hold <kbd>Ctrl</kbd>/<kbd>Shift</kbd>) and press **Delete** (button or key) to remove them.
- **Clear form** empties the form; **Clear all** removes every meeting *and* also clears your section rankings, unticked subjects and priority order (hard limits are left as they are).
- Use **Show:** above the table to look at a single section. The counter shows how many meetings and sections you have.

**Rules and tips**
- Section and subject names can't contain `[` or `]`.
- Days are Mon, Tue, Wed, Thu, Fri, Sat. Saturday only appears in the timetable if a chosen class is on Saturday.
- Meetings may be entered in any order; the table sorts them for you.
- **Enter every section that offers a subject you might take**, from every year level you could pull classes from. The app can only choose among sections you've entered.
- Keep subject names identical across sections. `DSA` and `D.S.A.` are treated as two different subjects.

### Step 2 · Set your options

Open **2 · Options**. It has up to four parts.

#### What matters most? (priority list)

Schedules are compared using this list **from the top down**: the first item decides, and each later item only
breaks ties between schedules that are *exactly equal* on everything above it. The list can contain:

- **Fewest conflict hours**: total hours where two of your classes overlap (classes that merely touch, like 10:00 end and 10:00 start, don't count).
- **Fewest days at school**: how many different weekdays you have to attend.
- **Year N section preference**: how well the schedule matches your ranked sections for that year (see below). One line appears for each year level where you've entered sections.

Select a line and use **▲ / ▼** to move it. Or use a quick preset (it keeps your preference lines in their current order):

| Preset | Resulting order |
|---|---|
| Fewest conflicts → fewest days → preferences | conflicts, days, then preferences |
| Fewest days first | days, conflicts, then preferences |
| Section preferences first | preferences, then conflicts, then days |

A preference line says "(nothing ranked)" until you rank at least one section for that year; it has no effect until then.

#### Hard limits and output

| Setting | What it does |
|---|---|
| **Max days at school** | Throws out every schedule that needs more than this many days. "No limit" turns it off. |
| **Max conflict hours** | Throws out every schedule with more conflict than this. Blank = no limit, `0` = no overlaps allowed, `1.5` = up to 90 minutes. |
| **Schedules to keep** | How many of the best schedules to return (default 15, maximum 500). |

Hard limits are different from priorities: a priority only *orders* schedules, while a limit *removes* them.
If your limits are impossible to meet, the app says so and you can relax them.

#### Subjects to schedule

Every subject part found in your meetings is listed with a checkbox and the number of sections that offer it.
Click a row to tick or untick it (**Select all** / **Select none** are below the list).
**Untick the subjects you are not taking**: only ticked subjects are scheduled.

#### Section preferences by year level

For each year level there is a ranking box with **Not ranked** and **Ranked (top = best)** lists. Put the sections you'd
like most at the top; leave sections you don't care about unranked.
- **Classic:** double-click a section to rank or un-rank it, or use the **Rank →**, **← Remove**, **▲** and **▼** buttons.
- **Modern:** click a section in the "not ranked" area to add it, then use **▲ / ▼ / ✕** beside it to reorder or remove.

**How points work.** If a year has *N* ranked sections, every subject part placed in your #1 section earns *N* points, in your
#2 section *N − 1*, and so on; unranked sections earn 0. More points is better. For example, with the ranking `3-D`, `3-E` for
Year 3, a subject part placed in `3-D` earns 2 points and in `3-E` earns 1. Results show your score as `got/best possible`,
e.g. `Year 3 preference: 3/7`.

Rankings only affect the year they belong to, and only if that year's preference line is in your priority list
(it is added automatically). Rankings are forgotten for sections you later delete.

### Step 3 · Find and read your results

Press **▶ Find best schedules**. A progress bar and status line show the number of combinations, elapsed time and percent done.
You can press **Stop** at any time; you'll still get the best schedules found so far, but they may not be the very best.
When the search finishes, the **3 · Results** screen opens by itself.

**The results list** (left) has one row per schedule, best first (#1 is the best):

| Column | Meaning |
|---|---|
| # | Rank |
| Conflict hr | Hours of overlap in that schedule |
| Days / Days used | How many days at school, and which (Mo Tu We Th Fr) |
| Yr N pref | Your preference score for that year, as `points/best possible`. There is one such column for each year level in which you ranked at least one section. |

**Click a row** to see it on the right:
- **The weekly timetable.** Each subject has its own colour. Classes that overlap are drawn **red** and side by side.
- **The detail box** underneath: the summary line, the **section chosen for every subject part** (with "your choice #n" if it's one you ranked), and a day-by-day list of times in which conflicts are tagged `<-- CONFLICT`.

**Exporting**
- **Export selected to Excel…**: the schedule you've highlighted, as a one-sheet weekly grid ready to print (landscape, 7 AM to 9 PM, wider if you have earlier or later classes). Conflicts are marked in red. The rank, summary and chosen sections are written below the grid.
- **Export all to Excel…**: one sheet per schedule (`#1`, `#2`, …).
- **Export CSV…**: one row per schedule with its rank, conflict hours, days, preference scores and the section chosen for every subject part. Handy for comparing in a spreadsheet.

Excel export needs `openpyxl`, which is listed in `requirements.txt`. If you run from source without it, only Excel export is disabled; everything else still works.

### Saving and reopening your work

Everything you've entered can be saved as a small `.json` project file:
- **Classic:** *File → Save project…*, *Open project…*, *New (empty)*.
- **Modern:** the **PROJECT** buttons in the sidebar (*Save project…*, *Open project…*, *New project*).

A project stores your meetings, section rankings, priority order, unticked subjects, hard limits and "schedules to keep".
Results are not saved; just press **Find best schedules** again after reopening. The **How the ranking works** entry
(Help menu / sidebar) shows a short summary of the rules in this guide. The Modern version also has a **Dark mode** switch.

### Recipes

| I want… | Do this |
|---|---|
| The most compact schedule with no overlaps | Preset **Fewest days first**, set **Max conflict hours** to `0`. |
| To be in my favourite sections, even if some classes overlap | Rank your sections, then use the preset **Section preferences first**. Check the red conflicts on the timetable before deciding. |
| No overlaps, then my favourite sections | Set **Max conflict hours** to `0` and use the preset **Section preferences first**. |
| At most 4 days at school | Set **Max days at school** to `4`; keep whichever priority order you like. |
| Year 3 preferences to beat Year 2 preferences | In the priority list, move **Year 3 section preference** above **Year 2** (the list is ordered with the highest year first by default). |
| To compare many options | Raise **Schedules to keep**, then **Export CSV…** or **Export all to Excel…**. |
| To drop a subject I'm not taking | Untick it under **Subjects to schedule**. |
| To try "what if I take only these subjects?" | **Select none**, then tick just the subjects you want. |

A worked example with the sample data: with the default priorities the best schedule has 0 conflict hours over 5 days.
Switch to **Section preferences first** and run again: the best schedule now gets every top-ranked section
(Year 3: 7/7, Year 2: 6/6) but at the cost of **4.0 hours of conflict**. That shows how the priority order, not just
your rankings, decides which schedule comes out on top.

### Troubleshooting

| Problem | Fix |
|---|---|
| "Can't read the time…" | Use a form like `7:00AM`, `2 PM`, `2:30pm` or `14:00`. Minutes must be 0-59. |
| "The end time must be after the start time." | Check AM/PM, or use 24-hour times such as `13:30`. |
| "Nothing to schedule" | You have no meetings, or every subject is unticked on the Options screen. |
| "No schedules found / no schedule satisfies your hard limits" | Raise **Max days** or **Max conflict hours**, or clear them. |
| A subject I'm taking is missing from the results | Make sure it is ticked under **Subjects to schedule**, and that it was entered with consistent spelling. |
| Two subjects that should be the same are listed separately | Their names differ (spelling, spacing or capitals). Edit the rows so they match exactly. |
| A preference line says "(nothing ranked)" | Rank at least one section for that year on the Options screen. |
| The search takes a long time | The number of combinations multiplies with every subject part. Untick subjects you don't need, or set **Max days** and **Max conflict hours**, which let the search skip bad branches early. You can also press **Stop** and use the best found so far. |
| "Excel export needs openpyxl" | You're running from source without it. Run `pip install openpyxl`. |
| Import CSV skipped rows | Check the header (`section,day,subject,type,start,end`), day names, `LEC`/`LAB`/`ALL`, and that times are valid. |
| Windows SmartScreen warning on the `.exe` | See the note under [Download](#download). |

## Releasing

Releases are built by GitHub Actions (`.github/workflows/release.yml`).

```
git tag v1.0.0
git push origin v1.0.0
```

The workflow builds both exes on Windows and attaches them to a new GitHub release.
To test the build without releasing, run the workflow manually from the Actions tab;
the exes are then available as run artifacts.

## Build locally

```
pip install pyinstaller
pyinstaller --onefile --windowed --name ScheduleOptimizer-Modern --collect-all customtkinter src/schedule_optimizer_ctk.py
pyinstaller --onefile --windowed --name ScheduleOptimizer-Classic src/schedule_optimizer_gui.py
```

Output goes to `dist/`.
