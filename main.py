import pathlib
from typing import List

import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt

path = pathlib.Path('./data')


# Some Simple wrapper classes to verbalize structure

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
        self.data_frame: pd.DataFrame = None
        self.max_rows: int = 0

    def __repr__(self):
        return (str(self.filename)
                + ':'
                + self.sheet_name
                + ', '.join([str(f) for f in self.measurement_ranges]))


datasets: List[Dataset] = [] # This contains all files

# Iterate over all files and scan datasets.
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

    datasets[-1].max_rows = max([r.end - r.start for r in datasets[-1].measurement_ranges])


# Load data-frames for every range
for dataset in datasets:
    for r in dataset.measurement_ranges:
        df = pd.read_excel(str(dataset.filename),
                           sheet_name=dataset.sheet_name,
                           header=r.start - 1,
                           index_col=0,
                           usecols='A:E',
                           nrows=r.end - r.start)
        r.df = df

    # Merge dataframes
    dataset.data_frame = pd.DataFrame()
    dataset.data_frame.index.name = 'Frame'  # Set name of index column
    dataset.data_frame['t'] = sorted([f.df for f in dataset.measurement_ranges], key=len)[-1]['ms']
    for index, r in enumerate(dataset.measurement_ranges):
        dataset.data_frame[f'x{index}'] = r.df['X (mm)']
        dataset.data_frame[f'y{index}'] = r.df['Y (mm)']
        dataset.data_frame[f'f{index}'] = r.df[f'Force (N)']

    # Calculate averages for x, y and force
    dataset.data_frame['avg_x'] = dataset.data_frame.apply(
        lambda row: sum([row[f'x{i}'] for i in range(len(dataset.measurement_ranges))])/len(dataset.measurement_ranges),
        axis=1)
    dataset.data_frame['avg_y'] = dataset.data_frame.apply(
        lambda row: sum([row[f'y{i}'] for i in range(len(dataset.measurement_ranges))]) / len(
            dataset.measurement_ranges),
        axis=1)
    dataset.data_frame['avg_f'] = dataset.data_frame.apply(
        lambda row: sum([row[f'f{i}'] for i in range(len(dataset.measurement_ranges))]) / len(
            dataset.measurement_ranges),
        axis=1)

    # Kill all rows that contain empty values.
    dataset.data_frame.dropna(axis=0, inplace=True)

    print(dataset)
    print(dataset.data_frame)
    dataset.data_frame.plot.scatter(x='avg_x', y='avg_y', c='avg_f')
    plt.show()
