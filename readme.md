# FootCalc

This project reads Excel files containing multiple measurements and evaluates it.
The files contains timeseries data of multiple measurements of an experiment.

The experiment has a pressure sensitive matrix and subjects persorm actions on this sensor mat.
The Input of this project is a temporal map of the center of mass on the mat.

Each measurement has a preamble that is irrelevant for this project and is ignored. The dataset looks like this:

| Frame | ms  | X (mm) | Y (mm) | Force (N) |
|-------|-----|--------|--------|-----------|
| 0     | 0   | -3.14  | 23.53  | 1.36      |
| ...   | ... | ...    | ...    | ...       |

## Dependencies

* pandas
* openpyxl
* matplotlib