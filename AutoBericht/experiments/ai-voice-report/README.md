# AI Voice Report (Experiment)

Standalone offline-first voice workflow for drafting findings/recommendations per row ID in `project_sidecar.json`.

Deployment is browser-only in managed Microsoft Edge: no native helper,
localhost service, extension, cloud inference, or machine installation. The
selected target stack is Whisper Large-v3 Turbo ONNX for transcription followed
sequentially by a text-only LFM2.5 ONNX model for reviewed report drafting; see
the model decision in `SPEC.md` section 7.

- Main spec: `SPEC.md`
- Build checklist: `TASKS.md`
- Runtime entry: `index.html` + `app.js`

This experiment is intentionally isolated from the main AutoBericht runtime for now.
