# Survey Validation Studio — Final App

## Included workspaces
- Dashboard with modern responsive Streamlit styling
- Brand Validation
- Text Response Validation (classification, suspicious short/repeated patterns, near duplicates)
- Separate Questionnaire Contradiction Checker

## Contradiction Checker workflow
1. Upload questionnaire as DOCX, selectable-text PDF, TXT or Markdown.
2. Upload respondent dataset as XLSX, XLS or CSV.
3. Select the unique respondent ID column every time. Blank or duplicate IDs stop execution.
4. AI proposes check names, descriptions, variables, logic and severity.
5. Review/edit every check, map it to exact dataset columns, approve or reject.
6. Run approved checks across respondent batches.
7. Download an Excel workbook with one sheet per check containing only flagged respondents, plus Check Audit.

## Providers
Gemini, Groq, OpenRouter and local Ollama. API keys are entered in the sidebar. Ollama requires its local server to be running on the machine hosting Streamlit.

## Windows setup
1. Extract ZIP.
2. Open PowerShell in `survey_validation_streamlit`.
3. `python -m venv .venv`
4. `.venv\Scripts\Activate.ps1`
5. `pip install -r requirements.txt`
6. `streamlit run app.py`

## Notes
- Provider catalogues and free quotas change. The app uses provider model-list endpoints where available; check access/quotas in your provider account.
- AI-generated checks must be reviewed and approved. Test against known examples before using in production.
- Scanned PDFs require OCR before upload.
- The selected ID column must be populated and unique.
- On Streamlit Cloud, Ollama localhost refers to the cloud host, not your personal PC.
- API keys are held in Streamlit session state and are not written into reports. For shared deployments, configure secrets and authentication.
