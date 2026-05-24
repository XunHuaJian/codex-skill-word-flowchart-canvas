---
name: word-flowchart-canvas
description: Create thesis-style flowcharts in Microsoft Word using a real Drawing Canvas, Word shapes, and connector arrows. Use when the user asks for a Word-editable paper/thesis/论文 flowchart, especially when they require 插入画布, 形状, 连接箭头, decision diamonds with 是/否 labels beside arrows, spaced right-side branch boxes, and elbow connectors that avoid crossing shapes.
---

# Word Flowchart Canvas

Use this skill to create Word-editable academic flowcharts made from:

- Word Drawing Canvas
- Rounded rectangles
- Diamond decision nodes
- Connector arrows
- Separate "是/否" text labels placed beside the corresponding arrows

Do not insert a raster image when the user asks for Word drawing-canvas output.

## Workflow

1. Create a new `.docx` rather than modifying the user's source document unless explicitly asked.
2. Use `scripts/create-wardrobe-flowchart.ps1` when the user wants the STM32 wardrobe-care flowchart layout from this skill.
3. Keep the layout principles:
   - Leave horizontal space between each diamond and the right-side rounded rectangle.
   - Put "是" beside the right-branch arrow, not inside the line.
   - Put "否" beside the downward arrow, not inside the line.
   - Use elbow-style return lines for right-side branches so lines do not cross boxes or diamonds.
   - Keep every element editable as Word shapes inside the drawing canvas.
4. If the script is unavailable or Word COM cannot run, use `assets/wardrobe_flowchart_canvas_spaced_logic_labels.docx` as the ready-made template.

## Script

Run from PowerShell:

```powershell
powershell -ExecutionPolicy Bypass -File "<skill-dir>\scripts\create-wardrobe-flowchart.ps1" -OutPath "<output.docx>"
```

The script requires Microsoft Word to be installed and callable through COM automation.

## Template

`assets/wardrobe_flowchart_canvas_spaced_logic_labels.docx` is the approved layout template. It contains the flowchart as Word Drawing Canvas + shapes + connector arrows, not as an inserted image.
