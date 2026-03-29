# Primebuild Payroll Journals Automation

Streamlit app that converts KeyPay raw journal exports into GL Journal Download files,
replicating the VBA macro in `XXX_XX_2023MMDD_Payroll_Jnl_Workings_v14.xlsm`.

## How it works

1. Upload one or more `*_JNL_Raw.xlsx` files from KeyPay
2. Enter the Payment Date (DD/MM/YYYY)
3. Click **Generate Journal Files**
4. Download the ZIP containing one `.xlsx` per input file

## File naming convention (inputs)

`STATE_FREQ_YYYYMMDD_JNL_Raw.xlsx`

Examples:
- `NSW_FN_20260318_JNL_Raw.xlsx`
- `VIC_WCOMP_FN_20260318_JNL_Raw.xlsx`
- `ROL_FN_20260318_JNL_Raw.xlsx`

## Output naming

`STATE FREQ YYYYMMDD JNL.xlsx` using the **payment date** for the date stamp.  
WComp files output as: `WC VIC YYYYMMDD JNL.xlsx`

## State → Costing Work ID mapping

| State | CWI |
|-------|-----|
| NSW   | 10  |
| QLD   | 40  |
| VIC   | 20  |
| ROL   | 11  |
| SVS   | 85  |
| CON   | 50  |

## Deployment

1. Push to GitHub
2. Connect repo to Streamlit Cloud
3. Set main file to `app.py`
