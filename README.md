# BIXEPS Automatic Analyser

A reusable Streamlit application for analysing completed pre/post BIXEPS Excel workbooks.

## Main features

- Automatically detects the BIXEPS assessment sheet
- Classifies each assessment area as Improved, No Change, Declined or Missing Data
- Generates overall outcomes: Improved, Maintained, No Change, Declined or Insufficient Data
- Shows average SPPB, pain and health-score changes
- Analyses frailty-status changes
- Groups results into physical function, pain, daily living and wellbeing domains
- Flags seniors requiring follow-up
- Ranks the top 10 improving seniors
- Generates an analysed Excel workbook and a management PDF report
- Preserves all original workbook sheets

## Run locally

Open Command Prompt in this folder and run:

```bash
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

## Deploy using Streamlit Community Cloud

1. Upload `app.py`, `analyser.py`, `requirements.txt` and this README to a GitHub repository.
2. Open Streamlit Community Cloud and select **Create app**.
3. Choose the repository and branch `main`.
4. Set the main file path to `app.py`.
5. Deploy the app.

## Data privacy

The application analyses the workbook in memory. It does not contain a database or permanently save uploaded files. Streamlit hosting and organisational data-governance requirements should still be reviewed before uploading identifiable senior data.

## Data privacy and retention

- Uploaded workbooks are read from memory (`BytesIO`) only.
- The application does not write uploaded files, generated reports, senior names, or assessment results to local storage or a database.
- No Streamlit caching is used for uploaded data or analysis results.
- Use **Clear uploaded data** after downloading the reports. This resets the upload widget and releases session references.
- Closing the browser tab ends normal user access to the session. On Streamlit Community Cloud, the hosting provider still operates the underlying infrastructure; for confidential personal data, use an organisation-approved private deployment with access controls rather than a publicly reachable Community Cloud app.
- Do not add logging statements that print dataframe contents, names, or uploaded bytes.
