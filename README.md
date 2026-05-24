# Word Flowchart Canvas / 论文图画布

Codex skill for generating thesis-style Microsoft Word flowcharts as editable Word Drawing Canvas content.

This skill is useful when a paper or thesis requires flowcharts drawn with:

- Word Drawing Canvas
- Rounded rectangle process nodes
- Diamond decision nodes
- Connector arrows
- `是/否` labels placed beside branch arrows
- Elbow-style return connectors that avoid crossing shapes

## Skill

The skill folder is:

```text
word-flowchart-canvas/
```

Install it by copying `word-flowchart-canvas` into your Codex skills directory:

```text
C:\Users\<you>\.codex\skills\word-flowchart-canvas
```

## Included Template

The skill includes an approved Word template:

```text
word-flowchart-canvas/assets/wardrobe_flowchart_canvas_spaced_logic_labels.docx
```

It is a real Word drawing canvas with editable shapes and connectors, not an inserted image.

## Script

On Windows with Microsoft Word installed, generate the sample flowchart with:

```powershell
powershell -ExecutionPolicy Bypass -File ".\word-flowchart-canvas\scripts\create-wardrobe-flowchart.ps1" -OutPath ".\output.docx"
```

The script uses Microsoft Word COM automation.
