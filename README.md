# Python Data Cleaner

I built this because cleaning Excel files by hand is slow, and you can never tell afterwards what was changed. You run one command and get back a clean workbook plus a log of every fix.

[Watch the 38-second demo video](docs/demo.mp4)

![Results](docs/step4_results.png)

## The problem

Most exports I've seen have the same issues. There are duplicate rows and extra spaces. IDs get saved as numbers like 17850.0. Blanks are typed as "N/A" or "-", and dates sit in the file as text. Fixing that by hand in Excel takes ages and is easy to get wrong.

## What it does

It reads CSV, TSV, TXT, XLSX and XLS files, or a whole folder of them, and every sheet in each workbook. It works out the delimiter and encoding by itself. It matches column names across files, trims spaces and turns blank-looking values into real blanks. It drops fully empty rows and columns. Then it converts number and date columns, and it can handle things like "$1,200", "15%" and "(300)". Last, it removes exact duplicates, or duplicates on the key columns you pick.

The output workbook has five sheets: Cleaned, Removed Duplicates, Missing Values, Column Types and Change Log. You also get a short summary in plain English.

## What happened when I tested it

I ran it on the full, unmodified UCI Online Retail file, which has 541,909 rows and 8 columns.

| | |
|---|---|
| Rows in / rows out | 541,909 / 536,641 |
| Duplicate rows removed | 5,268 |
| Cells with extra spaces fixed | 128,367 |
| Missing CustomerID | 135,037 rows (25.2%) |
| Missing Description | 1,505 rows (0.3%) |

InvoiceNo, StockCode and CustomerID stayed as text. Quantity and UnitPrice became numbers, and InvoiceDate became a date. Running it on this real file also turned up two bugs in my first version, and I fixed both.

## Try it

```bash
pip install -r requirements.txt
python scripts/download_data.py        # downloads the test file into data/
python data_clean.py "data/Online Retail.xlsx" -o cleaned_online_retail.xlsx
```

The useful options are `--key COL ...`, `--dayfirst`, `--no-merge` and `--fill-numeric {none,zero,median}`. Run `python data_clean.py -h` to see them all.

## Screenshots

| | |
|---|---|
| ![Input](docs/step1_input.png) | ![Run](docs/step2_run.png) |
| ![Output](docs/step3_output.png) | ![Results](docs/step4_results.png) |

## Data and licence

The test data is Chen, D. (2015), *Online Retail*, UCI Machine Learning Repository, https://doi.org/10.24432/C5BW33, licensed CC BY 4.0. I didn't commit it to this repo. The download script fetches it.

The code is MIT licensed. See [LICENSE](LICENSE).

Made by Shanmukh Narayana Chaganti. I'm an engineering student and I'm starting out as a freelancer, so if you have a messy file, feel free to get in touch.
