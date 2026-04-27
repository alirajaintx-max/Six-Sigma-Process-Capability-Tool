# Six Sigma Process Capability Tool

A Python quality engineering tool that calculates Cp and Cpk, plots control charts, and exports a Six Sigma quality report to Excel.

## What it does

* Calculates process capability indices Cp, Cpk, CPU, and CPL
* Estimates sigma level and defect rate in PPM (parts per million)
* Builds X-bar and R control charts with UCL and LCL limits
* Automatically flags out-of-control points beyond 3 sigma limits
* Plots a process distribution histogram with specification limits overlaid
* Exports all results to a formatted Excel quality report

## Results (Shaft Diameter Process)

* Cp: 1.247 — Marginal
* Cpk: 1.138 — Off-Centre
* Sigma Level: 3.41σ
* Estimated Defects: 832 PPM
* Out-of-Control Points Detected: 2

## How to run it

1. Install dependencies: `pip install numpy pandas matplotlib scipy openpyxl`
2. Run the script: `python sixsigma_tool.py`

## Technologies

Python · NumPy · Pandas · Matplotlib · SciPy · openpyxl
