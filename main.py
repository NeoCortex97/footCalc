import pathlib
from typing import List

import pandas as pd
from openpyxl import load_workbook

path = pathlib.Path('./data')


class MeasurementRange:
    def __init__(self, start: int, end: int):
        self.start = start
        self.end = end
        self.df: pd.DataFrame = None

    def __repr__(self):
        return f'({self.start} ; {self.end})'


class Dataset:
    def __init__(self):
        self.filename: pathlib.Path = pathlib.Path()
        self.sheet_name: str = ''
        self.measurement_ranges: List[MeasurementRange] = []

    def __repr__(self):
        return (str(self.filename)
                + ':'
                + self.sheet_name
                + ', '.join([str(f) for f in self.measurement_ranges]))


datasets: List[Dataset] = []
for file in [f for f in path.iterdir() if f.is_file() and f.name.endswith('.xlsx')]:
    wb = load_workbook(file, False, False, False, False, False)
    datasets.append(Dataset())
    datasets[-1].filename = file
    datasets[-1].sheet_name = wb.sheetnames[0]
    ws = wb[datasets[-1].sheet_name]
    r = MeasurementRange(0, 0)
    active = False
    for cell in ws['A']:
        if cell.value == 'Frame':
            r.start = cell.row
            active = True
        if active and cell.value is None:
            r.end = cell.row
            datasets[-1].measurement_ranges.append(r)

            r = MeasurementRange(0, 0)
            active = False


for dataset in datasets:
    for r in dataset.measurement_ranges:
        df = pd.read_excel(str(dataset.filename),
                           sheet_name=dataset.sheet_name,
                           header=r.start - 1,
                           index_col=0,
                           usecols='A:E',
                           nrows=r.end - r.start)
        print(df)
        r.df = df







