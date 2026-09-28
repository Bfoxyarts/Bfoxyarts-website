# Bfoxyarts

Official website for artist Bernadette Fox. Bfoxyarts showcases original artwork inspired by wildlife, birds, nature, personal experiences, and everyday life.

## Website sections

- Home
- Gallery
- Exhibitions
- About the Artist
- Contact

## Exhibition automation

`Bfoxyarts_Exhibitions_Content_Template.xlsx` is the source of truth for the Exhibitions page.

When that workbook is changed and committed to the `main` branch, GitHub Actions runs `scripts/generate_exhibitions.py`, rebuilds `exhibitions.html`, and commits the generated page to the repository. Cloudflare can then deploy the new commit through the site's existing GitHub connection.

### Update an exhibition

1. Download `Bfoxyarts_Exhibitions_Content_Template.xlsx` from the repository.
2. Edit the `Exhibitions` worksheet in Microsoft Excel.
3. Keep the workbook filename exactly `Bfoxyarts_Exhibitions_Content_Template.xlsx`.
4. Enter real Excel dates in `Start Date` and `End Date`.
5. Set `Publish?` to `Yes` for every row that should appear on the website.
6. Set `Featured?` to `Yes` only when the show should display the Featured badge.
7. Put exhibition image files in the repository's `images` folder and enter exact filenames, including capitalization, in the workbook.
8. Upload the revised workbook to the repository root and commit the change.
9. Open the repository's Actions tab and confirm that `Build exhibitions page` completed successfully.

Do not manually edit `exhibitions.html`. The next successful automation run will replace manual changes to that file.

## Automatic exhibition status

The generated page assigns status using each exhibition's dates:

- Before `Start Date`: Upcoming
- From `Start Date` through `End Date`: Current
- After `End Date`: Past

The browser recalculates status whenever the page is opened, so shows continue moving between sections even when the workbook has not been edited that day. The Upcoming, Current, and Past counts are also calculated automatically.

The calculated `Status`, `Sort Order`, `Display Date`, `URL Slug`, and `Generated HTML Snippet` workbook columns are not used as the website's source for status or layout. The generator uses the primary content fields and dates.

## Repository automation files

- `Bfoxyarts_Exhibitions_Content_Template.xlsx`: exhibition content source
- `scripts/generate_exhibitions.py`: page generator
- `.github/workflows/build-exhibitions.yml`: GitHub Actions workflow
- `requirements.txt`: Python dependency used by the workflow
- `exhibitions.html`: generated website page
- `images/`: website and exhibition images

## Troubleshooting

### The workbook changed but the website did not

1. Confirm that the workbook is in the repository root and its filename has not changed.
2. Confirm that the workbook change was committed to `main`.
3. Check the Actions tab for the `Build exhibitions page` run.
4. If the workflow failed to push, open **Settings > Actions > General > Workflow permissions**, select **Read and write permissions**, and save.
5. Confirm that the new `exhibitions.html` commit appears in the repository.
6. Confirm that the image filenames in the workbook exactly match files in `images`, including uppercase and lowercase letters.

### Run the generator manually in GitHub

Open **Actions > Build exhibitions page > Run workflow**, choose `main`, and run the workflow.

## Local generation

With Python installed:

```bash
pip install -r requirements.txt
python scripts/generate_exhibitions.py
```

## Public contact

- Email: hello@bfoxyarts.com
- Instagram: @bfoxyarts
