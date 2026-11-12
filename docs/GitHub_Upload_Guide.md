# Publishing this project on GitHub (browser only, no Git install needed)

## 1. Create the repository
1. Sign in at https://github.com.
2. Click the **+** icon (top right) > **New repository**.
3. Repository name: `dc-asset-lifecycle-dashboard` (or similar).
4. Description: `Simulated data centre IT asset lifecycle, spare parts, RMA and budget-forecast workbook (Excel + Python)`.
5. Choose **Public** so recruiters can see it.
6. Leave "Add a README file" **unticked** (this project already has one). Click **Create repository**.

## 2. Upload the files

GitHub's browser upload places files in the folder you are currently viewing, and dragged folders are usually flattened. So upload the root files first, then create each folder and upload its files from inside it.

**Root files.** On the empty repository page click **uploading an existing file**, add `README.md`, `requirements.txt` and `.gitignore`, and click **Commit changes**.

**Each folder (`data`, `docs`, `output`, `scripts`):**
1. On the repository's main page click **Add file > Create new file**.
2. In the file-name box type the folder name, a slash and `README.md`, for example `data/README.md`. Typing the slash turns `data` into a folder.
3. Paste the contents of that folder's `README.md` from the unzipped package and click **Commit changes**.
4. Click the folder name (e.g. `data`) in the path at the top of the page so you are inside the folder.
5. Click **Add file > Upload files**, select all the remaining files from the matching local folder, and click **Commit changes**.
6. Click the repository name to return to the main page and repeat for the next folder.

## 3. Check the result
- The repository front page should show the README with the project summary.
- Click `docs/Management_Report.md` and `docs/Formula_and_PivotTable_Guide.md`; GitHub renders them as formatted pages.
- Click `output/DC_Asset_Lifecycle_Financial_Dashboard.xlsx` > **Download raw file** (or **View raw**) to confirm the workbook downloads. GitHub does not preview Excel files.

## 4. Polish (optional, 5 minutes)
- On the repository page click the gear icon next to **About** and add topics: `excel`, `asset-management`, `data-center`, `inventory-management`, `financial-forecasting`, `python`, `openpyxl`.
- Add the repository link to your CV and LinkedIn under Projects, described as a *simulated case study*.
- After you build the three PivotTables in Excel, save the workbook and upload it again (**Add file > Upload files**, drop the new file into `output/`, commit). GitHub replaces the old version and keeps the history.

## 5. Reproducing the files (for reviewers)
Anyone can rebuild everything from the repository:

```bash
pip install -r requirements.txt
python scripts/generate_data.py
python scripts/build_workbook.py
```

The data are seeded, so the rebuilt workbook contains exactly the same records.
