# Integrate MCP with Copilot

<img src="https://octodex.github.com/images/Professortocat_v2.png" align="right" height="200px" />

Hey mflashjuan!

Mona here. I'm done preparing your exercise. Hope you enjoy! 💚

Remember, it's self-paced so feel free to take a break! ☕️

[![](https://img.shields.io/badge/Go%20to%20Exercise-%E2%86%92-1f883d?style=for-the-badge&logo=github&labelColor=197935)](https://github.com/mflashjuan/skills-integrate-mcp-with-copilot/issues/1)
## Data ingestion pipeline

This repository now includes a data ingestion pipeline for activity records stored in Excel and DOCX files.

- `scripts/build_activity_data.py` builds `data/activities.json`
- `src/data_ingestion.py` parses Excel worksheets and DOCX tables into structured activity data
- `src/data_loader.py` loads activities from JSON into the FastAPI app

### Usage

```bash
python3 -m pip install -r requirements.txt
python3 scripts/build_activity_data.py --input data/input --output data/activities.json
```

The pipeline supports:

- Excel files with activity metadata and participant rows
- DOCX files containing tables of activity participants
---

&copy; 2025 GitHub &bull; [Code of Conduct](https://www.contributor-covenant.org/version/2/1/code_of_conduct/code_of_conduct.md) &bull; [MIT License](https://gh.io/mit)

