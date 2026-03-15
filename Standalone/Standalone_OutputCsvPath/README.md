# Standalone_OutputCsvPath

Default output folder for `Standalone` scripts.

- Default result file pattern: `Results_SA-<P|M|IR><WWNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.csv`
- Required transcript file pattern: `Transcript_SA-<P|M|IR><WWNN>-<Action>-<Target>_<yyyyMMdd-HHmmss>.log`
- Override per run with `-OutputCsvPath`
- Standalone scripts should write both result CSV and transcript log here by default
