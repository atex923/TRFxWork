# BUILD_NOTES

## Nuitka exe build memo

- Do not rebuild exe after every source edit. Rebuild only when explicitly requested.
- App icon is stored at `assets/icon_result.ico` and should be committed to Git.
- Database files must not be committed or packaged into Git: `TRFxWork_db`, `TR_FxWork.db`, `*.db`, `*.sqlite`, `*.sqlite3`.

## Onefile command

Run from the repository root, replacing the version names when the project version changes:

```bat
python -m nuitka --onefile --enable-plugin=tk-inter --windows-console-mode=disable --windows-icon-from-ico=assets\icon_result.ico --output-dir=outputs\TRFxWork_V0_1_8_onefile --output-filename=TRFxWork_V0_1_8.exe TR_FxWork_V0_1_8.py
```

The onefile exe stores `TRFxWork_db` next to the exe at runtime. Keep `outputs/` out of Git.
