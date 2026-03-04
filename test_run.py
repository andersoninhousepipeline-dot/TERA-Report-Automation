import pandas as pd
from tera_template import TERAReportGenerator

df = pd.read_excel('TERA evaluation data_06-02-26.xls')
df = df.dropna(how='all')

# Pick patient 4 which is Mrs. Rajathi (Pre-receptive)
row_idx = 4
data_row = df.loc[row_idx].to_dict()

generator = TERAReportGenerator(data_row, '.')
out_file = generator.generate()
print(f"Generated test file: {out_file}")
