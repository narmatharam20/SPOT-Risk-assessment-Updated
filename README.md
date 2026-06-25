# SPOTT Multi-Project E&I Tracker

This Streamlit app stores multiple SPOTT project Excel files, saves changes into a SQLite database, and lets users reload projects anytime.

## Key features
- Upload multiple Excel files for different projects.
- Store each project separately.
- Save latest edits for documents, procurement and FAT milestones.
- Reload any saved project.
- Portfolio dashboard across all projects.
- Export selected project to Excel.
- Download/restore complete database backup JSON.

## Run locally
```powershell
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Important persistence note
The app uses a local SQLite database file (`spott_project_tracker.db`). This is suitable for local/server use. On Streamlit Community Cloud, local storage may reset after redeploys/restarts, so use the Backup feature regularly or connect a permanent database/SharePoint for production use.
