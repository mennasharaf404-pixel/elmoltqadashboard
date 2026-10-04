# ELMOLTQA Developments — Department Dashboards (Streamlit)

Initial UI/UX scaffold for a multi-department analytics dashboard, branded
with the ELMOLTQA Developments logo and color palette (navy / gold / teal).

## Structure
```
dashboard/
├── app.py                       # Entry point — page config + navigation
├── theme.py                     # Shared branding, CSS, and color palette
├── assets/
│   └── logo.jpg                 # Company logo
├── views/
│   ├── home.py                  # Overview page
│   ├── operations.py
│   ├── customer_service.py
│   ├── financial.py
│   ├── marketing.py
│   └── hr.py
└── requirements.txt
```

Navigation is defined explicitly in `app.py` via `st.navigation()` / `st.Page()`,
with each page's title and icon set in Python code rather than in the filename.
(An earlier version used emoji in the `pages/` filenames for Streamlit's
auto-discovery — that's fragile on Windows, since zip/extract tools can
mis-decode the emoji bytes and turn the sidebar labels into garbled text.
This structure avoids that entirely.)

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Streamlit auto-detects the `pages/` folder and builds the left-hand
navigation menu from the emoji-prefixed filenames — that's how department
switching works, alongside the dropdown selector on the home page.

## What's real vs. placeholder
- **Branding, layout, navigation, KPI cards, chart types, filters** — final,
  ready to use.
- **The numbers themselves** — sample/random data generated in each page's
  `load_data()` function, clearly marked with `@st.cache_data`. Replace the
  body of that function with a query to your actual source (SQL database,
  data warehouse, API, or a CSV/Excel load) and every chart/KPI updates
  automatically since they all reference the same DataFrame columns.

## Next steps to wire up real data
1. Pick a connection method per department (e.g. `sqlalchemy` for a
   warehouse, `pandas.read_sql`, or an internal API client).
2. Replace each `load_data()` body, keeping the same output column names
   (or update the chart calls to match your schema).
3. Add authentication if this will be deployed outside a trusted network
   (e.g. `streamlit-authenticator`) before sharing externally.
4. Deploy via Streamlit Community Cloud, an internal server, or a Docker
   container behind your company's SSO.
