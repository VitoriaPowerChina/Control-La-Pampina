import openpyxl
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

FILE_PATH = r"D:\PowerChina\extracted\Anexo 02 Progreso Fisico.xlsm"

# Read only the fixed columns (cols 1-31) using pandas with header row 2
df = pd.read_excel(FILE_PATH, sheet_name=0, header=2, engine='openpyxl', usecols=range(0, 31))
print(f"Fixed columns shape: {df.shape}")
print(f"Columns: {list(df.columns)}")

print("\n=== SAMPLE DATA (first 30 rows, fixed columns) ===")
pd.set_option('display.max_columns', 35)
pd.set_option('display.max_colwidth', 60)
pd.set_option('display.width', 300)
print(df.head(30).to_string())

print("\n=== UNIQUE SUBCONTRATISTA VALUES ===")
sub = df['Subcontratista'].dropna()
for v in sorted(sub.unique()):
    print(f"  '{v}': {(df['Subcontratista']==v).sum()} rows")

print("\n=== UNIQUE EDT LEVEL STRUCTURE (first 50 unique EDT values) ===")
edt = df['EDT'].dropna().unique()[:50]
for v in edt:
    print(f"  {v}")

print("\n=== NIVEL DE ESQUEMA distribution ===")
if 'Nivel de esquema' in df.columns:
    print(df['Nivel de esquema'].value_counts().sort_index().to_string())

print("\n=== UNIDAD unique values ===")
if 'Unidad' in df.columns:
    unidades = df['Unidad'].dropna().unique()
    for u in sorted([str(x) for x in unidades]):
        c = (df['Unidad'].astype(str)==u).sum()
        print(f"  '{u}': {c}")

print("\n=== DATE RANGES ===")
for col in ['Comienzo PLAN', 'Comienzo REAL', 'Fin PLAN', 'Fin REAL', 'Fin proyectado']:
    if col in df.columns:
        s = pd.to_datetime(df[col], errors='coerce').dropna()
        if len(s) > 0:
            print(f"  {col}: min={s.min().date()}, max={s.max().date()}, count={len(s)}")

print("\n=== NUMERIC STATS (fixed columns) ===")
num_df = df.select_dtypes(include='number')
for col in num_df.columns:
    s = num_df[col].dropna()
    if len(s) > 0:
        print(f"  '{col}': min={s.min():.4f}, max={s.max():.4f}, mean={s.mean():.4f}, non-null={len(s)}/{len(df)}")

print("\n=== MISSING DATA SUMMARY (fixed columns) ===")
total = len(df)
print(f"  Total rows: {total}")
miss = df.isnull().sum()
for col, cnt in miss.items():
    pct = 100*cnt/total
    if cnt > 0:
        print(f"  '{col}': {cnt} missing ({pct:.1f}%)")

print("\n=== COMPLETELY EMPTY ROWS ===")
# Check across all fixed columns
empty = df.isnull().all(axis=1).sum()
print(f"  Fully empty rows: {empty}")

print("\n=== TASK NAME SAMPLE (first 50 unique task names at level 1-3) ===")
if 'Task Name' in df.columns and 'Nivel de esquema' in df.columns:
    top = df[df['Nivel de esquema'].isin([1.0, 2.0, 3.0])][['Nivel de esquema','EDT','Task Name']].dropna(subset=['Task Name'])
    print(top.head(50).to_string())

print("\n=== DAILY QUANTITY COLUMNS - first date column value range ===")
# Read first date column (col 31)
df_dates = pd.read_excel(FILE_PATH, sheet_name=0, header=2, engine='openpyxl', usecols=range(31, 40))
print(f"  First 9 date columns: {list(df_dates.columns)}")
for col in df_dates.columns:
    s = df_dates[col].dropna()
    if len(s) > 0:
        print(f"  '{col}': min={s.min():.4f}, max={s.max():.4f}, non-null={len(s)}, dtype={s.dtype}")

print("\nDONE")
