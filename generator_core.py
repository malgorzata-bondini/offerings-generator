# generator_core.py
import datetime as dt
import re, warnings
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

need_cols = [
    "Name (Child Service Offering lvl 1)", "Parent Offering",
    "Service Offerings | Depend On (Application Service)", "Service Commitments",
    "Delivery Manager", "Subscribed by Location", "Phase", "Status",
    "Life Cycle Stage", "Life Cycle Status", "Support group", "Managed by Group",
    "Subscribed by Company"
]

skip_hs_for = {"MD", "UA", "PL"}
discard_lc = {"retired", "retiring", "end of life", "end of support"}

def run_generator(
    *,
    keywords, new_apps, days, hours,
    delivery_manager, global_prod,
    rsp_duration, rsl_duration,
    sr_or_im, require_corp, delivering_tag,
    support_group, managed_by_group, aliases_on,
    src_dir: Path, out_dir: Path
):
    schedule_suffix = " ".join(f"{d} {h}" for d, h in zip(days, hours))
    aliases_value = "" if aliases_on else ""
    sheets = {}
    seen = set()

    def row_keywords_ok(row):
        p = str(row["Parent Offering"]).lower()
        n = str(row["Name (Child Service Offering lvl 1)"]).lower()
        hit_p = {k for k in keywords if re.search(rf"\b{re.escape(k)}\b", p)}
        hit_n = {k for k in keywords if re.search(rf"\b{re.escape(k)}\b", n)}
        return bool(hit_p) and bool(hit_n) and hit_p | hit_n == set(keywords)

    def lc_ok(row):
        return all(
            str(row[c]).strip().lower() not in discard_lc
            for c in ("Phase", "Status", "Life Cycle Stage", "Life Cycle Status")
        )

    def name_prefix_ok(name):
        return name.lower().startswith(f"[{sr_or_im.lower()} ")

    def update_commitments(orig, sched, rsp, rsl):
        out = []
        for line in str(orig).splitlines():
            if "RSP" in line:
                line = re.sub(r"RSP .*? P1-P4", f"RSP {sched} P1-P4", line)
                line = re.sub(r"P1-P4 .*?$", f"P1-P4 {rsp}", line)
            elif "RSL" in line:
                line = re.sub(r"RSL .*? P1-P4", f"RSL {sched} P1-P4", line)
                line = re.sub(r"P1-P4 .*?$", f"P1-P4 {rsl}", line)
            out.append(line)
        return "\n".join(out)

    def commit_block(cc):
        lines = [
            f"[{cc}] SLA SR RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}",
            f"[{cc}] OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
        return "\n".join(lines)

    for wb in src_dir.glob("ALL_Service_Offering_*.xlsx"):
        df = pd.read_excel(wb, sheet_name="Child SO lvl1")
        if any(c not in df.columns for c in need_cols):
            continue

        mask = (
            df.apply(row_keywords_ok, axis=1)
            & df["Name (Child Service Offering lvl 1)"]
                 .astype(str)
                 .apply(name_prefix_ok)
            & df.apply(lc_ok, axis=1)
            & (df["Service Commitments"]
                 .astype(str)
                 .str.strip()
                 .replace({"nan": ""}) != "-")
        )

        if require_corp:
            mask &= df["Name (Child Service Offering lvl 1)"]\
                       .str.contains(r"\bCORP\b", case=False)
        else:
            mask &= ~df["Name (Child Service Offering lvl 1)"]\
                        .str.contains(r"\bCORP\b", case=False)

        base_pool = df.loc[mask]
        if base_pool.empty:
            continue
        base_row = base_pool.iloc[0].to_frame().T.copy()

        country = wb.stem.split("_")[-1].upper()
        tag_hs = f"HS {country}"
        tag_ds = f"DS {country}"

        if require_corp:
            if country == "DE":
                receivers = ["DS DE", "HS DE"]
            elif country == "CY":
                receivers = ["HS CY", "DS CY"]
            elif country in {"UA", "MD"}:
                receivers = [f"DS {country}"]
            elif country == "PL":
                has_ds = "DS PL" in base_pool["Name (Child Service Offering lvl 1)"].str.cat(sep=" ")
                receivers = ["DS PL"] if has_ds else ["HS PL"]
            else:
                receivers = [f"DS {country}"]
        else:
            receivers = [""]

        parent_full = str(base_row.iloc[0]["Parent Offering"])
        m_inner = re.search(r"\[Parent\s+(.*?)\]", parent_full, re.I)
        parent_inner = m_inner.group(1).strip() if m_inner else ""
        parent_desc = parent_full.split("]", 1)[-1].strip()
        inner_tokens = parent_inner.split()

        for app in new_apps:
            for recv in receivers:
                tag_in = tag_ds if recv.startswith("DS") else tag_hs
                if require_corp:
                    head = f"[{sr_or_im} {delivering_tag} CORP {recv}"
                else:
                    head = f"[{sr_or_im} {tag_in}"

                filtered = [tok for tok in inner_tokens if tok not in recv.split()]
                rest = " ".join(filtered)

                name_head = f"{head} {rest}]".replace("  ", " ").replace(" ]", "]")
                new_name = f"{name_head} {parent_desc} {app} Prod {schedule_suffix}".replace("  ", " ")
                if new_name in seen:
                    continue
                seen.add(new_name)

                row = base_row.copy()
                row["Name (Child Service Offering lvl 1)"] = new_name
                row["Delivery Manager"] = delivery_manager
                row["Support group"] = support_group
                row["Managed by Group"] = managed_by_group
                for c in row.columns:
                    if "Aliases" in c:
                        row[c] = aliases_value
                if country == "DE":
                    row["Subscribed by Company"] = (
                        "DE Internal Patients\nDE External Patients"
                        if tag_in == tag_hs
                        else "DE IFLB Laboratories\nDE IMD Laboratories"
                    )
                elif country == "UA":
                    row["Subscribed by Company"] = "Сiнево Україна"
                else:
                    row["Subscribed by Company"] = tag_in
                orig_comm = str(row.iloc[0]["Service Commitments"]).strip()
                if not orig_comm or orig_comm == "-":
                    row["Service Commitments"] = commit_block(country)
                else:
                    row["Service Commitments"] = update_commitments(
                        orig_comm, schedule_suffix, rsp_duration, rsl_duration
                    )
                if require_corp:
                    depend_tag = f"{delivering_tag} Prod"
                elif global_prod or "servicenow" in app.lower():
                    depend_tag = "Global Prod"
                else:
                    depend_tag = f"{tag_in} Prod"
                row["Service Offerings | Depend On (Application Service)"] = f"[{depend_tag}] {app}"

                sheets.setdefault(country, pd.DataFrame())
                sheets[country] = pd.concat([sheets[country], row], ignore_index=True)

    if not sheets:
        raise ValueError(
            "No rows matched your criteria. Please check your Keywords, CORP filter, and that your template contains those entries."
        )

    out_dir.mkdir(parents=True, exist_ok=True)
    outfile = out_dir / f"Offerings_NEW_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with pd.ExcelWriter(outfile, engine="openpyxl") as writer:
        for cc, dfc in sheets.items():
            df_export = dfc.drop_duplicates(subset=["Name (Child Service Offering lvl 1)"])\
                           .drop(columns=["Number"], errors="ignore")
            df_export.to_excel(writer, sheet_name=cc, index=False)

    wb = load_workbook(outfile)
    for ws in wb.worksheets:
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = (
                max(len(str(c.value)) if c.value else 0 for c in col) + 2
            )
            for c in col:
                c.alignment = Alignment(wrap_text=True)
    wb.save(outfile)
    return outfile
