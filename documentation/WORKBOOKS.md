# Simulation Workbook Notes

Updated: 2026-05-19 14:40:05 AEST

This package stages the two Excel simulator/result workbooks identified by David Levinson for `paper-2003-08`. The original files are preserved unchanged in `code/simulation_workbooks/original_xls/`, and LibreOffice-generated `.xlsx` sidecars are in `code/simulation_workbooks/xlsx_converted/`. Nonempty sheets were exported as CSV under `data/simulation_workbook_csv/`; formula-preserving sheet CSVs with the suffix `__formulas.csv` and `metadata/FORMULA_INVENTORY.csv` expose the spreadsheet logic.

## Workbook Scope

The source folder also contains manuscript drafts, correspondence-like Word files, GIF figure exports, PDFs, and earlier TRB-era workbook variants. Those are not copied because the archive boundary is the published paper plus the two identified simulator/result workbooks.

## Conversion Notes

LibreOffice headless conversion was used for `.xls` to `.xlsx`. The original `14nov.xls` contains a VBA project marker; macro source was not separately extracted in this environment, so the original workbook is the authoritative preservation copy for any embedded VBA. Spreadsheet formulas are preserved in the `.xlsx` workbooks and are also indexed in `metadata/FORMULA_INVENTORY.csv`.

## Sheet Profile

- `14nov.xls` / `toll`: 1011 rows x 45 columns; 43121 nonempty cells; 42012 formula cells; values CSV `data/simulation_workbook_csv/14nov/toll.csv`.
- `14nov.xls` / `comp`: 63 rows x 11 columns; 549 nonempty cells; 27 formula cells; values CSV `data/simulation_workbook_csv/14nov/comp.csv`.
- `14nov.xls` / `results`: 43 rows x 11 columns; 410 nonempty cells; 0 formula cells; values CSV `data/simulation_workbook_csv/14nov/results.csv`.
- `14nov_results.xls` / `case2 (2)`: 43 rows x 15 columns; 533 nonempty cells; 120 formula cells; values CSV `data/simulation_workbook_csv/14nov_results/case2_2.csv`.
- `14nov_results.xls` / `delay`: 1 rows x 1 columns; 0 nonempty cells; 0 formula cells; values CSV `not exported (empty)`.
- `14nov_results.xls` / `Sys`: 1 rows x 1 columns; 0 nonempty cells; 0 formula cells; values CSV `not exported (empty)`.
- `14nov_results.xls` / `CS`: 1 rows x 1 columns; 0 nonempty cells; 0 formula cells; values CSV `not exported (empty)`.
- `14nov_results.xls` / `profit`: 1 rows x 1 columns; 0 nonempty cells; 0 formula cells; values CSV `not exported (empty)`.
- `14nov_results.xls` / `case 1`: 45 rows x 15 columns; 536 nonempty cells; 121 formula cells; values CSV `data/simulation_workbook_csv/14nov_results/case_1.csv`.
- `14nov_results.xls` / `case 2`: 43 rows x 15 columns; 535 nonempty cells; 120 formula cells; values CSV `data/simulation_workbook_csv/14nov_results/case_2.csv`.
- `14nov_results.xls` / `case 3`: 43 rows x 15 columns; 535 nonempty cells; 120 formula cells; values CSV `data/simulation_workbook_csv/14nov_results/case_3.csv`.

## First Rows Preview

### 14nov.xls / toll

```text
free way conditions |  |  |  |  |  |  |  |  |  |  | %vehicles
 | max | current |  | freeway flow |  | ramp meter rate |  | by pass rate |  |  | car
speed | 65 | 25 |  | Q |  | q1 |  | q2 |  |  | truck
density | 120 | 75 |  |  |  |  |  | q1+q2 | 0.1 | vps | 
```

### 14nov.xls / comp

```text
 | tolled |  |  |  |  |  |  |  |  | 
 | trucks | tolled units | utility | delay  | truckes-d | CS |  | utility | delay | truckers-d
 | 94 | 1 | -587.609382815455 | 147177.208779751 | 16907.7655833028 | 3.60948032968099 |  | -591.05944864617 | 147854.679305261 | 17585.236108813
```

### 14nov.xls / results

```text
toll | trucks | tolled units | utility | delay  | truckes-d | CS |  | utility | delay | truckers-d
0 | 102.1 | 68.22 | -617.028127188763 | 175004.561804841 | 3930.14592416423 | 189.39298839949 |  | -772.101794596521 | 196840.808880682 | 19265.1373837189
0.5 | 98.96 | 56.2 | -732.811891402339 | 200387.012951145 | 5620.93832165436 | 174.15158624575 |  | -869.324412981119 | 221957.670650958 | 22019.1786960146
1 | 100.68 | 48.3 | -746.183760045588 | 198685.781011913 | 7293.95657236872 | 163.669641605782 |  | -870.250231329336 | 218025.120589989 | 22269.5969817051
```

### 14nov_results.xls / case2 (2)

```text
toll | trucks | tolled units | utility | delay  | truckes-d | CS |  | utility | delay | truckers-d | 
0 | 102.1 | 68.22 | -617.028127188763 | 175004.561804841 | 3930.14592416423 | 189.39298839949 |  | -772.101794596521 | 196840.808880682 | 19265.1373837189 | 
0.5 | 98.96 | 56.2 | -732.811891402339 | 200387.012951145 | 5620.93832165436 | 174.15158624575 |  | -869.324412981119 | 221957.670650958 | 22019.1786960146 | 
1 | 100.68 | 48.3 | -746.183760045588 | 198685.781011913 | 7293.95657236872 | 163.669641605782 |  | -870.250231329336 | 218025.120589989 | 22269.5969817051 | 
```

### 14nov_results.xls / case 1

```text
scenario | 1 |  |  |  |  |  |  |  |  |  | 
toll | trucks | tolled units | utility | delay  | truckes-d | CS |  | utility | delay | truckers-d | 
0 | 99.68 | 61.38 | -403.50509467392 | 114028.116865132 | 3201.99087772032 | 138.485216686099 |  | -492.198565092364 | 126182.10664987 | 12492.6180418952 | 
0.5 | 102.52 | 52.18 | -420.398745844722 | 111481.161663632 | 4168.81079829912 | 117.1206320445 |  | -482.907791569639 | 121948.635865296 | 12406.7875601253 | 
```

### 14nov_results.xls / case 2

```text
scenario | 2 |  |  |  |  |  |  |  |  |  | 
toll | trucks | tolled units | utility | delay  | truckes-d | CS |  | utility | delay | truckers-d | 
0 | 98.06 | 65.48 | -610.757499165035 | 172845.93839929 | 3587.49436039269 | 180.090859514768 |  | -756.719072351804 | 194172.136665921 | 19103.0425458542 | 
0.5 | 100.66 | 57.54 | -700.45905736686 | 190561.663616775 | 5724.86328235063 | 180.812700063459 |  | -843.102410471968 | 212569.37833642 | 22427.4817165424 | 
```

### 14nov_results.xls / case 3

```text
Scenario  | 3 |  |  |  |  |  |  |  |  |  | 
toll | trucks | tolled units | utility | delay  | truckes-d | CS |  | utility | delay | truckers-d | 
0 | 101.6 | 69.46 | -1022.57195660235 | 289686.620413408 | 4060.05543566698 | 294.489507631997 |  | -1291.07591833262 | 326239.615501563 | 33044.1884010427 | 
0.5 | 99.38 | 59.1 | -1055.36304661266 | 291369.841075087 | 5312.41239896273 | 247.956021916226 |  | -1274.42622099185 | 325179.114801598 | 32265.2001534013 | 
```
