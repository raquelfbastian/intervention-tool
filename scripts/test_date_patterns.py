import re
import pandas as pd


def extract_snapshot_date(filename):
    patterns = [
        r"\d{4}-\d{2}-\d{2}",
        r"\d{4}_\d{1,2}_\d{1,2}",
        r"\d{1,2}_\d{1,2}_\d{4}",
        r"\d{8}",
        r"\d{2}-\d{1,2}-\d{4}",
    ]

    for pat in patterns:
        match = re.search(pat, filename)
        if match:
            token = match.group(0)
            token_norm = re.sub(r"[_]", "-", token)
            for fmt in ["%Y-%m-%d", "%d-%m-%Y", "%Y-%m-%d"]:
                try:
                    dt = pd.to_datetime(token_norm, format=fmt, errors="coerce")
                    if not pd.isna(dt):
                        return dt
                except Exception:
                    pass
            dt = pd.to_datetime(token_norm, errors="coerce", dayfirst=False)
            if not pd.isna(dt):
                return dt

    match = re.search(r"(\d{1,2})[._-](\d{1,2})[._-](\d{4})", filename)
    if match:
        token = "-".join(match.groups()[::-1])
        dt = pd.to_datetime(token, errors="coerce", dayfirst=False)
        if not pd.isna(dt):
            return dt

    match = re.search(r"\d{6,8}", filename)
    if match:
        token = match.group(0)
        try:
            dt = pd.to_datetime(token, format="%Y%m%d", errors="coerce")
            if not pd.isna(dt):
                return dt
        except Exception:
            pass
        try:
            dt = pd.to_datetime(token, format="%d%m%Y", errors="coerce")
            if not pd.isna(dt):
                return dt
        except Exception:
            pass

    return None


if __name__ == '__main__':
    samples = [
        'Dump_06_07_2026.xlsx',
        'Dump_15_7_2026.xlsx',
        'Dump_20_7_2026.xlsx',
        'snapshot_2026-07-15.xlsx',
        'snapshot_20260710.xlsx',
    ]

    for s in samples:
        dt = extract_snapshot_date(s)
        print(s, '->', dt)
