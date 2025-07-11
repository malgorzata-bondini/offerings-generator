import datetime as dt
import re
import warnings
from pathlib import Path
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

need_cols = [
    "Name (Child Service Offering lvl 1)",
    "Parent Offering",
    "Service Offerings | Depend On (Application Service)",
    "Service Commitments",
    "Delivery Manager",
    "Subscribed by Location",
    "Phase",
    "Status",
    "Life Cycle Stage",
    "Life Cycle Status",
    "Support group",
    "Managed by Group",
    "Subscribed by Company"
]

skip_hs_for = {"MD", "UA", "PL"}
discard_lc = {"retired", "retiring", "end of life", "end of support"}

def run_generator(
    *,
    keywords,new_apps,days,hours,
    delivery_manager,global_prod,
    rsp_duration,rsl_duration,
    sr_or_im,require_corp,delivering_tag,
    support_group,managed_by_group,aliases_on,
    src_dir:Path,out_dir:Path
):
    schedule_suffix = " ".join(f"{d} {h}" for d,h in zip(days,hours))
    sheets = {}
    seen = set()

    def row_keywords_ok(row):
        p = str(row["Parent Offering"]).lower()
        n = str(row["Name (Child Service Offering lvl 1)"]).lower()
        hit_p = {k for k in keywords if re.search(rf"\b{re.escape(k)}\b",p)}
        hit_n = {k for k in keywords if re.search(rf"\b{re.escape(k)}\b",n)}
        return bool(hit_p) and bool(hit_n) and hit_p|hit_n==set(keywords)

    def lc_ok(row):
        return all(
            str(row[c]).strip().lower() not in discard_lc
            for c in ("Phase","Status","Life Cycle Stage","Life Cycle Status")
        )

    def name_prefix_ok(name):
        return name.lower().startswith(f"[{sr_or_im.lower()} ")

    def commit_block(cc):
        lines = [
            f"[{cc}] SLA SR RSP {schedule_suffix} P1-P4 {rsp_duration}",
            f"[{cc}] SLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}",
            f"[{cc}] OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}"
        ]
        return "\n".join(lines)

    def update_commitments(orig):
        # update RSP/RSL and then always append OLA
        out = []
        for line in str(orig).splitlines():
            if line.startswith("RSP "):
                line = re.sub(r"P1-P4 .*?$",f"P1-P4 {rsp_duration}",line)
            elif line.startswith("RSL "):
                line = re.sub(r"P1-P4 .*?$",f"P1-P4 {rsl_duration}",line)
            out.append(line)
        out.append(f"OLA SR RSL {schedule_suffix} P1-P4 {rsl_duration}")
        return "\n".join(out)

    for wb in src_dir.glob("ALL_Service_Offering_*.xlsx"):
        df = pd.read_excel(wb,sheet_name="Child SO lvl1")
        if any(c not in df.columns for c in need_cols):
            continue

        mask = (
            df.apply(row_keywords_ok,axis=1)
            & df["Name (Child Service Offering lvl 1)"].astype(str).apply(name_prefix_ok)
            & df.apply(lc_ok,axis=1)
            & (df["Service Commitments"].astype(str).str.strip().replace({"nan":""})!="-")
        )

        if require_corp:
            mask &= df["Name (Child Service Offering lvl 1)"].str.contains(r"\bCORP\b",case=False)
        else:
            mask &= ~df["Name (Child Service Offering lvl 1)"].str.contains(r"\bCORP\b",case=False)

        base_pool = df.loc[mask]
        if base_pool.empty:
            continue
        row0 = base_pool.iloc[0].to_frame().T.copy()

        country = wb.stem.split("_")[-1].upper()
        tag_hs = f"HS {country}"
        tag_ds = f"DS {country}"

        if require_corp:
            if country=="DE":
                receivers=["DS DE","HS DE"]
            elif country=="CY":
                receivers=["HS CY","DS CY"]
            elif country in {"UA","MD"}:
                receivers=[f"DS {country}"]
            elif country=="PL":
                has_ds="DS PL" in base_pool["Name (Child Service Offering lvl 1)"].str.cat(sep=" ")
                receivers=["DS PL"] if has_ds else ["HS PL"]
            else:
                receivers=[f"DS {country}"]
        else:
            receivers=[""]

        parent_full = str(row0.iloc[0]["Parent Offering"])
        m = re.search(r"\[Parent\s+(.*?)\]",parent_full,re.I)
        parent_inner = m.group(1).strip() if m else ""
        parent_desc = parent_full.split("]",1)[-1].strip()
        toks = parent_inner.split()

        for app in new_apps:
            for recv in receivers:
                tag_in = tag_ds if recv.startswith("DS") else tag_hs
                if require_corp:
                    head = f"[{sr_or_im} {delivering_tag} CORP {recv}"
                else:
                    head = f"[{sr_or_im} {tag_in}"
                rest = " ".join(tok for tok in toks if tok not in recv.split())
                name_head = f"{head} {rest}]".replace("  "," ").replace(" ]","]")
                new_name = f"{name_head} {parent_desc} {app} Prod {schedule_suffix}".replace("  "," ")
                if new_name in seen:
                    continue
                seen.add(new_name)

                r = row0.copy()
                r["Name (Child Service Offering lvl 1)"] = new_name
                r["Delivery Manager"] = delivery_manager
                r["Support group"] = support_group
                r["Managed by Group"] = managed_by_group
                for c in r.columns:
                    if "Aliases" in c:
                        r[c] = "" 
                if country=="DE":
                    r["Subscribed by Company"] = (
                        "DE Internal Patients\nDE External Patients"
                        if tag_in==tag_hs else "DE IFLB Laboratories\nDE IMD Laboratories"
                    )
                elif country=="UA":
                    r["Subscribed by Company"] = "Сiнево Україна"
                else:
                    r["Subscribed by Company"] = tag_in

                orig_comm = str(r.iloc[0]["Service Commitments"]).strip()
                if orig_comm=="-" or not orig_comm:
                    r["Service Commitments"] = commit_block(country)
                else:
                    r["Service Commitments"] = update_commitments(orig_comm)

                if global_prod:
                    dtg = "Global Prod"
                elif require_corp:
                    dtg = f"{delivering_tag} Prod"
                else:
                    dtg = f"{tag_in} Prod"

                r["Service Offerings | Depend On (Application Service)"] = f"[{dtg}] {app}"

                sheets.setdefault(country,pd.DataFrame())
                sheets[country] = pd.concat([sheets[country],r],ignore_index=True)

    if not sheets:
        raise ValueError("No rows matched your criteria. Please check your Keywords, CORP filter, and that your template contains those entries.")

    out_dir.mkdir(parents=True,exist_ok=True)
    out = out_dir/f"Offerings_NEW_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
    with pd.ExcelWriter(out,engine="openpyxl") as w:
        for cc,dfc in sheets.items():
            df2 = (
                dfc.drop_duplicates(subset=["Name (Child Service Offering lvl 1)"])
                   .drop(columns=["Number"],errors="ignore")
                   .replace({True:"true",False:"false"})
            )
            df2.to_excel(w,sheet_name=cc,index=False)

    wb = load_workbook(out)
    for ws in wb.worksheets:
        ws.auto_filter.ref = ws.dimensions
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = max(len(str(c.value)) if c.value else 0 for c in col)+2
            for c in col:
                c.alignment = Alignment(wrap_text=True)
    wb.save(out)
    return out
