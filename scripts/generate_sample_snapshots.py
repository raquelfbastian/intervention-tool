import os
import pandas as pd

os.makedirs('historical_dumps', exist_ok=True)

rows = [
    {'Enterpriseid': 'eid001', 'Personnel No': '1001', 'SkillName': 'Python', 'proficiency': 1, 'career_level': 10, 'Business Group': 'Tech_Song'},
    {'Enterpriseid': 'eid002', 'Personnel No': '1002', 'SkillName': 'Python', 'proficiency': None, 'career_level': 11, 'Business Group': 'Tech_Adobe Platform'},
    {'Enterpriseid': 'eid003', 'Personnel No': '1003', 'SkillName': 'SQL', 'proficiency': 2, 'career_level': 9, 'Business Group': 'Tech_Song'},
]

# Create three snapshots with different filename date formats

df = pd.DataFrame(rows)

fn1 = 'historical_dumps/Dump_20_7_2026.xlsx'
fn2 = 'historical_dumps/snapshot_2026-07-15.xlsx'
fn3 = 'historical_dumps/snapshot_20260710.xlsx'

# Write slightly different proficiencies across snapshots

df1 = df.copy()
df1.loc[0, 'proficiency'] = 1

df2 = df.copy()
df2.loc[0, 'proficiency'] = 2

df3 = df.copy()
df3.loc[0, 'proficiency'] = 3

with pd.ExcelWriter(fn1) as w:
    df1.to_excel(w, index=False)

with pd.ExcelWriter(fn2) as w:
    df2.to_excel(w, index=False)

with pd.ExcelWriter(fn3) as w:
    df3.to_excel(w, index=False)

print('Sample snapshots written:', fn1, fn2, fn3)
