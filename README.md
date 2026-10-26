# Student Schedule Audit Dashboard

An interactive **R Shiny** application for detecting schedule anomalies, checking
master-schedule consistency, reviewing teacher-period enrollment, and analyzing
service placements by core subject.

**Status:** Used successfully by the project author in R Shiny; the packaged version has not been independently rerun here. The included demonstration CSVs use user-supplied sample records with newly substituted, fictional course codes and only the fields needed by the application.

## Dashboard preview

The author has run the application successfully with sample records. New screenshots
should be captured using the simplified demonstration CSVs in this repository.
Earlier screenshots have been deliberately excluded because they displayed course
codes resembling those used by a school district.

## Why this project exists

School schedule files, master schedules, and student service placement records
often require reconciliation before operational decisions can be made. The
dashboard presents exception reports and aggregates across these inputs.

## Features

- Duplicate-course, period-overlap, missing-period, and excess-period reports
- Master schedule teacher and period consistency checks
- Teacher-period enrollment and service placement summaries
- Core subject classification with manual overrides
- Interactive filtering, CSV exports, and a multi-sheet Excel report

## Technology

R (with base pipe `|>`), Shiny, dplyr, DT, openxlsx.

## Run locally

1. Install R 4.1 or later and RStudio (optional).
2. Install the dependencies:

   ```r
   install.packages(c("shiny", "DT", "dplyr", "openxlsx"))
   ```

3. From the repository folder, run:

   ```r
   shiny::runApp(".")
   ```

4. Upload the three CSV files from `data/synthetic/` in this order:
   - `sample_student_schedules.csv`
   - `sample_master_schedule.csv`
   - `sample_dl_core_level.csv`
5. Review the automatically selected field mappings. Set expected periods
   to `01(Reg),02(Reg),03(Reg),04(Reg),05(Reg),06(Reg),07(Reg),08(Reg)`,
   active IEP value to `Active`, and service codes to `SC,CTT`.
6. Select **Run dashboard analysis** and inspect exception tabs.
   The fictional data intentionally contains missing periods, a duplicate course,
   an overlap, and a master-schedule exception.

## Repository layout

- `Student_Schedule_Audit_Dashboard.R`: main R Shiny application (original logic, updated displayed title)
- `app.R`: lightweight launch entry point for `shiny::runApp(".")`
- `data/synthetic/`: wholly fictional demonstration CSVs
- `docs/review-notes.md`: limitations and publication checklist

## Limitations and work in progress

- The author confirmed that the application runs locally, but this packaged revision
  and included fictional sample files have not been independently regression-tested.
- The uploaded `app.R` does not visibly implement the requested exclusion for
  course codes ending in `N`; reconcile that business rule before release.
- Classification relies partly on course-description keywords and can need overrides.
- The application can display and export row-level names, identifiers, and service
  information. It is intended for controlled local use with authorized inputs.
- Assess eligibility to share school-derived work and source code before
  making this repository public.

## Privacy

**Never commit actual school records, identifiable student data, staff exports,
or confidential compliance files.** The three files in `data/synthetic/` are
newly constructed fictional examples; they are not de-identified copies of
the uploaded samples.

## Author's contribution

Requirements definition, data validation rules, service-placement logic,
operational reporting workflow, and dashboard design. Specify development
collaborators and coding assistance accurately as appropriate before publication.

## License

No license is included pending review of code ownership and publication rights.

## Simplified demonstration schema

All nonessential columns have been removed. The application accepts these names
through its column-mapping controls and should auto-detect them:

| Sample file | Required columns |
|---|---|
| Student schedules | Student ID, First Name, Last Name, IEP Status, Course Code, Course Description, Teacher, Period |
| Master schedule | Course Code, Course Description, Period, Teacher |
| DL core level | Student ID, English, Math, Science, Social Science |

Each fictional `DEMO-###` code maps consistently across the student and master
schedule samples. The demonstration inputs preserve their original test-case
row counts and relationships.

**Note:** Records and example names are from the author's sample CSVs. They
must be confirmed fictional before public release. No original school export
columns beyond those required for analysis are included.
