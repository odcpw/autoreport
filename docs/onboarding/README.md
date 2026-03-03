# Onboarding Guides

This folder contains onboarding material in three languages:

- `AutoBericht-Schnellstart-DE.docx`
- `AutoBericht-Guide-Rapide-FR.docx`
- `AutoBericht-Guida-Rapida-IT.docx`

Source files:

- `quickstart_de.md`
- `quickstart_fr.md`
- `quickstart_it.md`
- `screenshots/annotated/*.png`

To regenerate DOCX files:

```bash
cd docs/onboarding
pandoc quickstart_de.md -o AutoBericht-Schnellstart-DE.docx --resource-path=.
pandoc quickstart_fr.md -o AutoBericht-Guide-Rapide-FR.docx --resource-path=.
pandoc quickstart_it.md -o AutoBericht-Guida-Rapida-IT.docx --resource-path=.
```
