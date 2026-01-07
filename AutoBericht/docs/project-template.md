# Project Template Structure

Use this layout when zipping a project for engineers. The user selects the project root folder once in the app.

```
Projekt-<Kunde>-<YYYY>/
  AutoBericht/              # app bundle + seeds + docs (do not rename)
  Photos/                   # raw photos (root = unsorted)
  Inputs/                   # self-assessment + customer docs
  Outputs/                  # Word/PPT/PDF outputs
  project_sidecar.json      # created by the app
  library_user_<initials>_de-CH.json  # optional starter library
```

Notes
- If you want materialized photo folders later, they can live under `Photos/` as subfolders
  while the root stays as the “unsorted” bucket.
- The app will prefer `AutoBericht/data/seed` for seed data if no sidecar exists.
