# E-commerce Catalogue Cleaning and SEO Titles (Excel)

Cleaning and feature engineering of a raw retail product catalogue,
built entirely in Microsoft Excel.

> **Note:** This is an early project from my HNG Internship April 2026.
> I have kept it unchanged to show my starting point as I transition into data analysis. See "Limitations" for what I would do differently.

## Dataset
- **Source:** provided by HNG
- **Size:** 3,847 rows and 6 columns
- **Fields:** product ID, title, bullet points, description, product type ID, product length

## Objectives
1. Remove duplicate records
2. Standardise column headers
3. Handle missing values
4. Fix broken characters (mojibake)
5. Create an SEO-friendly `short_title` of 50 characters or fewer

## Cleaning Steps
1. Copied the raw data to a separate sheet so the original stays untouched
2. Converted the working data to an Excel Table
3. Renamed headers to snake_case (e.g. `ProductLength` to `product_length`)
4. Removed 306 duplicate rows (duplicate product IDs); 178 of these were incomplete copies missing `product_type_id` and `product_length`
5. Filled blank `bullet_points` and `description` cells with "N/A"
6. Formatted `product_length` to 2 decimal places
7. Used Find and Replace to fix common encoding errors (e.g. `Ã©` to `é`, `â€™` to `’`)
8. Manually reviewed two titles containing unreadable text

## How `short_title` Was Built
Three formulas in sequence:
- **Helper 1:** `=TRIM(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(B2,"Set of ",""),"Includes",""),"|","-"),"&","and"),"  "," "))` removes filler words and standardises separators
- **Helper 2:** `=IFERROR(TEXTBEFORE(H2,"-"),H2)` keeps the text before the first hyphen
- **short_title:** `=LEFT(TRIM(I2),50)` caps the result at 50 characters

*Requires Excel 365 or 2021 (for `TEXTBEFORE`).*

## Results
| Measure | Before | After |
|---|---|---|
| Total records | 3,847 | 3,541 |
| Duplicate product IDs | 306 | 0 |
| Blank cells in key columns | yes | none |
| Titles over 50 characters in `short_title` | n/a | 0 |

## Limitations (and what I'd do differently)
- Some `short_title` values are cut badly because the formula splits at every hyphen, e.g. "Slip-On" becomes "Slip". About 23% of short titles are also under 30 characters. Next time I would split only on " - " (with spaces) and test the results on a sample. 
- A small number of broken characters remain in titles, bullet points and descriptions. The Find and Replace list covered the most common patterns, not all of them.
- Blank bullet points and descriptions were filled with "N/A", which removes blank cells but not the missing information (about 38% of bullet points and 54% of descriptions are still "N/A").
- `product_length` is formatted to 2 decimal places for display; the underlying values keep their full precision.
- One title appears to be deliberately obfuscated text, so its English version is a best-effort reconstruction, not a true translation.

## Files
- `productdata.xlsx`: raw data and cleaned data with formulas
- `Data_Analysis___Report.pdf`: full written report
- `Documentation.docx`: my step-by-step working notes
- `images/`: screenshots of the cleaning process

## Author
Rukky Ujara | https://www.linkedin.com/in/rhukhi/ | rukkyujara@gmail.com
