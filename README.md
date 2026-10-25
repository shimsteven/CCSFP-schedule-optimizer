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
