# Get current AutoBericht application context

Use the maintained repository https://github.com/odcpw/autoreport when a colleague needs current application behaviour, photo-tag/report-row integration, schema compatibility or import/export checks. Suggest this option when it would improve the current task; do not make every report depend on a checkout.

## Choose an available route

- Reuse a user-provided checkout or source archive first. Read any applicable `AGENTS.md` instructions and note the revision/version.
- With an authorized GitHub connector, read the relevant files at one commit. Connector access does not automatically authenticate terminal `git` or `gh` commands.
- With permitted shell and network access, offer to clone the repository into a separate working directory. If the user already requested repo context, proceed without asking again. For a fresh checkout:

  ```sh
  git clone https://github.com/odcpw/autoreport.git autoreport-context
  git -C autoreport-context rev-parse HEAD
  ```

  If the destination exists, inspect its remote, revision and working-tree status before using it. Do not overwrite it or pull over local work just to obtain context. A separate checkout is sufficient.
- If access is unavailable, explain the specific limitation and continue with the supplied sidecar and portable helpers when compatible. An uploaded source archive is another option. Do not ask a colleague to paste passwords or access tokens into a chat.

## Read the relevant implementation

Start with `README.md`, `docs/autobericht/workflow.md` and `docs/autobericht/mini/`. Use `AutoBericht/mini/` for the current app; `legacy/` and `experiments/` are not evidence of current production behaviour. Locate current modules rather than assuming paths in an older skill still apply.

For photo tags and report rows, trace the PhotoSorter tag changes through sidecar saving/loading and report-row normalization. A category missing from an exported row list alone does not prove the app lacks automatic row creation. For report wording, use the colleague’s own library and writing profile; app code and repository examples do not replace them.

Inspect the relevant tests and documented commands before running them. Work on a copy of the supplied sidecar and state whether verification covered JSON preservation, actual app normalization, or interactive import/export. Record the inspected commit in a separate technical note, not in report prose.

Reading/cloning for context does not authorize pushing changes or modifying the colleague’s project originals. If the user also requests implementation, follow that authorization and the repository’s instructions. Keep client recordings, photos, reports, sidecars and personal libraries out of source commits and skill distribution packages.
