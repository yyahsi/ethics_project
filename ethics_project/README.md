# AI Ethics Rating App

Run the Streamlit app with:

```bat
streamlit run app.py
```

The app reads `data/ratings.csv` for the incident-rating page. Ratings are 0-10.

## Israeli vs Russian Crimes workbook

The **Israeli vs Russian Crimes** page lets you upload six files:

- ChatGPT Israel and ChatGPT Russia
- Claude Israel and Claude Russia
- Gemini Israel and Gemini Russia

The app generates an Excel workbook with three comparison sheets. Each sheet has this layout:

- News Title → Israel only
- Ethicality → Israel / Russia
- Visibility → Israel / Russia
- Accountability → Israel / Russia

The Streamlit page displays the full 200-row comparison table for every model, not just a preview.

Each comparison sheet outputs 200 rows. If an input file has fewer rows, the missing cells stay blank.

You can also generate the workbook from the command line using the sample files in `data/comparison_inputs`:

```bat
python generate_comparison_tables.py --inputs data/comparison_inputs --output israeli_vs_russian_crimes.xlsx
```

## Final Comparison averages

The **Final Comparison** page displays one combined average-score table for the three AI models:

- ChatGPT
- Claude
- Gemini

The table mirrors the Israeli vs Russian Crimes structure, but each row is an AI model average:

- Ethicality → Israel / Russia
- Visibility → Israel / Russia
- Accountability → Israel / Russia

It uses the files uploaded on the **Israeli vs Russian Crimes** page when available. Otherwise, it calculates from the bundled files in `data/comparison_inputs`. Blank or invalid cells are ignored.


## Incident ratings averages in Final Comparison

The **Final Comparison** page now also includes a second table called **Incident Ratings Average Scores**. It averages the incident-rating columns from `data/ratings.csv`:

- My Ground Truth
- ChatGPT
- Claude
- Gemini

Blank or invalid cells are ignored. The generated Excel workbook includes this as the second table on the **Final Comparison** sheet when `data/ratings.csv` is available.
