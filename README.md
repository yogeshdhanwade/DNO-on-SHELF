# Smart DNO Automation

Streamlit tool: upload portfolio -> build checkup (SKU / ASIN / Master ID / MPN) -> generate the DNO upload file from your formatted template.

## Files

- `streamlit_app.py` - UI (main file)
- `core.py` - portfolio parsing, checkup, DNO template filling
- `requirements.txt` - dependencies

## Run locally

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Host on Streamlit Community Cloud

share.streamlit.io -> Create app -> this repo, branch `main`, main file `streamlit_app.py`.

The DNO template is kept in server memory and must be re-uploaded (Step 0) after an app restart. Do not commit `.xlsx` files to the repo.
