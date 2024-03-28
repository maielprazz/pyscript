import pandas as pd

# Sample DataFrame
df = pd.DataFrame({
    'datetime_column': pd.date_range('2022-01-01', periods=5),
    'other_column': [1, 2, 3, 4, 5]
})

print (df)
# Convert to records with datetimes as dates
records = df.to_records(index=False)
records = [(pd.Timestamp(r[0]).to_pydatetime().date(), r[1], r[2],r[3]) for r in records]

print(records)