# ka-pook-look

PowerPoint redesign workspace for **Present Asset**.

## Output

- `Present Asset - Redesigned.pptx` — redesigned 19-slide deck.
- Content is preserved; changes are visual/layout only.
- `redesign_ppt.py` — reproducible Python script used to restyle the original deck.

## Run

```bash
pip install python-pptx
python redesign_ppt.py "Present Asset.pptx" -o "Present Asset - Redesigned.pptx"
```

The script validates the original text/table-cell content before saving and fails if any text content changes.
