import pandas as pd
import numpy as np

# ---------- Normalize ----------
def normalize_data(input_file):
    df = pd.read_excel(input_file)
    time = df.iloc[:, 0]
    data = df.iloc[:, 1:]

    numeric = data.select_dtypes(include='number')
    norm = numeric.apply(lambda col: col / col.min(), axis=0)

    norm_df = pd.concat([time, norm], axis=1)
    norm_df.columns = df.columns
    return norm_df


# ---------- AUC + Amplitude ----------
def compute_interval_metrics(time, data, start, end, delta_t):
    mask = (time >= start) & (time <= end + delta_t)
    t_seg = time[mask]
    y_seg = data[mask].values

    auc_records, trapezoids = [], {c: [] for c in data.columns}

    for i in range(len(t_seg) - 1):
        t1, t2 = t_seg[i], t_seg[i + 1]
        row = {'Time Start': t1, 'Time End': t2}
        for j, col in enumerate(data.columns):
            y1, y2 = y_seg[i, j], y_seg[i + 1, j]
            area = (y1 + y2) / 2 * (t2 - t1)
            row[col] = area
            trapezoids[col].append(area)
        auc_records.append(row)

    auc_df = pd.DataFrame(auc_records)
    auc_sums = {col: sum(trapezoids[col]) for col in data.columns}

    amp_vals = (data[mask].max() - data[mask].min()).to_dict()

    mins = (end - start) / 60
    summary = {
        'Total_AUC': np.mean(list(auc_sums.values())),
        'Average_AUC': np.mean(list(auc_sums.values())),
        'Interval_Duration_min': mins,
        'Average_AUC_per_min': np.mean(list(auc_sums.values())) / mins,
        'Average_Amplitude': np.mean(list(amp_vals.values())),
        'Avg_Amplitude_per_min': np.mean(list(amp_vals.values())) / mins
    }

    return auc_df, pd.Series(auc_sums), pd.Series(amp_vals), summary


# ---------- All Intervals ----------
def analyse_intervals(df, cut_points):
    time = df.iloc[:, 0].values
    data = df.iloc[:, 1:]
    dt = time[1] - time[0]

    cuts = sorted(cut_points)
    intervals, s = [], 0
    for c in cuts:
        intervals.append((s, c))
        s = c + dt
    intervals.append((s, time[-1]))

    auc_tables, auc_summaries, amp_summaries, meta_rows = {}, [], [], []

    for idx, (a, b) in enumerate(intervals, 1):
        print(f"📐 Processing interval: {a}–{b}")
        auc_df, auc_sum, amp_vals, summary = compute_interval_metrics(time, data, a, b, dt)

        tag = f"{int(a)}-{int(b)}"
        auc_tables[tag] = auc_df
        auc_sum.name = tag
        amp_vals.name = tag

        auc_summaries.append(auc_sum)
        amp_summaries.append(amp_vals)
        meta_rows.append(pd.Series(summary, name=tag))

    auc_summary_df = pd.concat(auc_summaries, axis=1).T
    amp_summary_df = pd.concat(amp_summaries, axis=1).T
    meta_df = pd.DataFrame(meta_rows)

    return auc_tables, auc_summary_df, amp_summary_df, meta_df


# ---------- Main ----------
if __name__ == "__main__":
    IN_FILE = "tudca 10 uM p1s1.xlsx"
    OUT_FILE = "tudca p1s1 wamp.xlsx"
    CUTS = [2830]

    original_df = pd.read_excel(IN_FILE)
    normalized_df = normalize_data(IN_FILE)

    auc_tables, auc_sum_df, amp_df, meta_df = analyse_intervals(normalized_df, CUTS)

    with pd.ExcelWriter(OUT_FILE, engine="xlsxwriter") as writer:
        sh = 'Processed_Data'
        wb = writer.book
        ws = wb.add_worksheet(sh)
        writer.sheets[sh] = ws
        r = 0

        # Original
        ws.write(r, 0, "Original Data")
        r += 1
        original_df.to_excel(excel_writer=writer, sheet_name=sh, startrow=r, index=False)
        r += len(original_df) + 4

        # Normalized
        ws.write(r, 0, "Normalized Data")
        r += 1
        normalized_df.to_excel(excel_writer=writer, sheet_name=sh, startrow=r, index=False)
        r += len(normalized_df) + 4

        for i, (tag, auc_df) in enumerate(auc_tables.items(), 1):
            ws.write(r, 0, f"AUC Data - Interval {i} ({tag})")
            r += 1
            auc_df.drop(columns=['Time End'], errors='ignore') \
                  .rename(columns={'Time Start': 'Time'}) \
                  .to_excel(excel_writer=writer, sheet_name=sh, startrow=r, index=False)
            r += len(auc_df) + 2

            ws.write(r, 0, f"AUC Sums - Interval {i}")
            r += 1
            auc_sum_df.loc[tag:tag].to_excel(excel_writer=writer, sheet_name=sh, startrow=r, index=False)
            r += 2

            ws.write(r, 0, f"AUC Average - Interval {i}")
            ws.write(r + 1, 0, meta_df.loc[tag, 'Average_AUC'])
            r += 3

            ws.write(r, 0, f"AUC per Minute - Interval {i}")
            ws.write(r + 1, 0, meta_df.loc[tag, 'Average_AUC_per_min'])
            r += 3

            # Amplitude section
            ws.write(r, 0, f"Amplitude - Interval {i}")
            r += 1
            amp_row = amp_df.loc[tag:tag]
            amp_row.to_excel(excel_writer=writer, sheet_name=sh, startrow=r, index=False)
            r += 2

            ws.write(r, 0, f"Average Amplitude - Interval {i}")
            ws.write(r + 1, 0, meta_df.loc[tag, 'Average_Amplitude'])
            r += 3

            ws.write(r, 0, f"Avg Amplitude per Minute - Interval {i}")
            ws.write(r + 1, 0, meta_df.loc[tag, 'Avg_Amplitude_per_min'])
            r += 3

    print(f"✅ Results saved to {OUT_FILE}")
