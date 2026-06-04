# -*- coding: utf-8 -*-
"""
合并 raw_data 下所有 CRF 表格为统一的长格式 Excel。
- 一行 = 一个患者 × 一天 × 一个时间点
- 同时输出"基本信息"工作表与"按天汇总"工作表
"""
import os
import re
import glob
import datetime as dt
from collections import OrderedDict

import openpyxl
import pandas as pd

ROOT = r"d:\AA_Lab\项目8：腹内压前瞻性入组项目192"
RAW_DIRS = [
    os.path.join(ROOT, "raw_data"),
    os.path.join(ROOT, "raw_data", "万博数据-2026.03.18", "万博数据"),
    os.path.join(ROOT, "raw_data", "腹内压数据-2026.03.12", "万博腹内压数据"),
]

TIME_POINTS = ["0小时", "2小时", "4小时", "6小时", "8小时",
               "10小时", "12小时", "14小时", "16小时",
               "18小时", "20小时", "22小时"]

VITAL_COLS = ["腹内压", "是否俯卧位", "体温", "心率", "呼吸",
              "血压", "血糖", "排便量", "肠内营养", "肠内类型", "肠内速度"]
FLUID_COLS = ["输液量_晶体", "输液量_胶体", "尿量", "胃肠减压_液体",
              "是否使用升压药", "腹腔引流量", "是否使用镇痛药物",
              "是否使用机械通气", "PEEP", "是否进行血液滤过", "血滤脱水量"]

# 表头规范化（去掉换行 / 全半角括号差异）
def norm_header(s):
    if s is None:
        return ""
    s = str(s).replace("\n", "").replace("\r", "").strip()
    s = s.replace("（", "(").replace("）", ")")
    return s

VITAL_HEADER_MAP = {
    "腹内压": "腹内压",
    "是否俯卧位": "是否俯卧位",
    "体温": "体温",
    "心率": "心率",
    "呼吸": "呼吸",
    "血压": "血压",
    "血糖": "血糖",
    "排便量": "排便量",
    "肠内营养": "肠内营养",
    "肠内类型": "肠内类型",
    "肠内速度": "肠内速度",
}
FLUID_HEADER_MAP = {
    "输液量(晶体)": "输液量_晶体",
    "输液量(胶体)": "输液量_胶体",
    "尿量": "尿量",
    "胃肠减压(液体)": "胃肠减压_液体",
    "是否使用升压药": "是否使用升压药",
    "腹腔引流量": "腹腔引流量",
    "是否使用镇痛药物": "是否使用镇痛药物",
    "是否使用机械通气": "是否使用机械通气",
    "PEEP": "PEEP",
    "是否进行血液滤过": "是否进行血液滤过",
    "血滤脱水量": "血滤脱水量",
}

INFO_RE_NAME = re.compile(r"姓名[:：_\s]*([^\s年_]+)")
INFO_RE_AGE = re.compile(r"年龄[:：_\s]*_?\s*(\d+)")
INFO_RE_HID = re.compile(r"住院号[:：_\s]*_*\s*([A-Za-z0-9]+)")
INFO_RE_DATE = re.compile(r"入组日期[:：_\s]*_?\s*([0-9]{4}[-/.年][0-9]{1,2}[-/.月][0-9]{1,2}日?)")

DAY_RE = re.compile(r"第\s*_?\s*(\d+)\s*_?\s*天")


def parse_info(s):
    if not s:
        return {}
    s = str(s)
    out = {}
    m = INFO_RE_NAME.search(s)
    if m:
        out["姓名"] = m.group(1).strip("_ ")
    m = INFO_RE_AGE.search(s)
    if m:
        out["年龄"] = int(m.group(1))
    m = INFO_RE_HID.search(s)
    if m:
        out["住院号"] = m.group(1).strip()
    m = INFO_RE_DATE.search(s)
    if m:
        out["入组日期"] = parse_date(m.group(1))
    return out


def parse_date(s):
    if s is None:
        return None
    if isinstance(s, (dt.datetime, dt.date)):
        return s.strftime("%Y-%m-%d")
    s = str(s).strip()
    s = s.replace("年", "-").replace("月", "-").replace("日", "").replace("/", "-").replace(".", "-")
    parts = s.split("-")
    if len(parts) == 3:
        try:
            y, m, d = (int(parts[0]), int(parts[1]), int(parts[2]))
            return f"{y:04d}-{m:02d}-{d:02d}"
        except ValueError:
            return s
    return s


def find_header_row(rows, marker_col1="时间点", marker_col2=None):
    """在 rows 中找到表头行，返回 (row_idx, header_dict col_index->norm_name)"""
    for i, r in enumerate(rows):
        cells = [norm_header(c) for c in r]
        if marker_col1 in cells:
            if marker_col2 is None or marker_col2 in cells:
                hdr = {idx: cells[idx] for idx in range(len(cells)) if cells[idx]}
                return i, hdr
    return None, None


def get_value(row, header_map, target_norm):
    """根据 header_map(col_idx -> norm_header)，取目标列的值。"""
    for idx, h in header_map.items():
        if h == target_norm:
            return row[idx]
    return None


# 录入错位修正：{姓名: [(原列名, 应为列名), ...]}
SWAP_FIXES = {
    "何绪钦": [("心率", "呼吸")],
    "叶海兵": [("心率", "呼吸")],
    "桂徽斌": [("心率", "呼吸")],
    "章新兵": [("心率", "呼吸")],
    "胡世仲": [("心率", "呼吸")],
}


def parse_sheet(ws, file_basename, sheet_name):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], {}

    info = {}
    day_num = None
    for r in rows[:5]:
        for c in r:
            if c is None:
                continue
            txt = str(c)
            if "姓名" in txt and not info:
                info = parse_info(txt)
            if day_num is None:
                m = DAY_RE.search(txt)
                if m:
                    day_num = int(m.group(1))
    if day_num is None:
        m = DAY_RE.search(sheet_name or "")
        if m:
            day_num = int(m.group(1))

    # 第一段：生命体征表
    v_idx, v_hdr = find_header_row(rows, "时间点", "腹内压")
    # 第二段：液体出入量表（在生命体征表之后再次出现"时间点"）
    f_idx, f_hdr = (None, None)
    if v_idx is not None:
        sub = rows[v_idx + 1:]
        f_off, f_hdr = find_header_row(sub, "时间点", "尿量")
        if f_off is not None:
            f_idx = v_idx + 1 + f_off

    if v_idx is None and f_idx is None:
        return [], info

    def collect_section(start_idx, end_idx, header_map, target_cols_map):
        """target_cols_map: {raw_norm_header: out_name}"""
        out = {}
        if start_idx is None:
            return out
        end = end_idx if end_idx is not None else len(rows)
        for r in rows[start_idx + 1:end]:
            tp = None
            for idx, h in header_map.items():
                if h == "时间点":
                    tp = norm_header(r[idx])
                    break
            if not tp or tp not in TIME_POINTS:
                continue
            d = {}
            for raw_h, out_name in target_cols_map.items():
                d[out_name] = get_value(r, header_map, raw_h)
            out[tp] = d
        return out

    vital_data = collect_section(v_idx, f_idx, v_hdr or {}, VITAL_HEADER_MAP)
    fluid_data = collect_section(f_idx, None, f_hdr or {}, FLUID_HEADER_MAP)

    # 录入错位修正：对该患者的指定列对调
    patient_name = (info.get("姓名") or "").strip("_ ")
    swap_pairs = SWAP_FIXES.get(patient_name, [])
    if swap_pairs:
        for tp, vd in vital_data.items():
            for a, b in swap_pairs:
                if a in vd and b in vd:
                    vd[a], vd[b] = vd[b], vd[a]

    records = []
    for tp in TIME_POINTS:
        rec = OrderedDict()
        rec["文件"] = file_basename
        rec["姓名"] = info.get("姓名")
        rec["住院号"] = info.get("住院号")
        rec["年龄"] = info.get("年龄")
        rec["入组日期"] = info.get("入组日期")
        rec["第几天"] = day_num
        rec["时间点"] = tp
        rec["小时"] = int(tp.replace("小时", ""))
        v = vital_data.get(tp, {})
        f = fluid_data.get(tp, {})
        for c in VITAL_COLS:
            rec[c] = v.get(c)
        for c in FLUID_COLS:
            rec[c] = f.get(c)
        # 仅保留至少一个数据非空的记录
        if any(rec[c] is not None for c in VITAL_COLS + FLUID_COLS):
            records.append(rec)
    return records, info


def collect_files():
    found = []
    seen = set()
    for d in RAW_DIRS:
        if not os.path.isdir(d):
            continue
        for f in glob.glob(os.path.join(d, "*.xlsx")):
            name = os.path.basename(f)
            if name.startswith("~$"):
                continue
            if name in seen:
                continue
            # 跳过 raw_data/ 顶层之外的子目录里的同名文件——按 basename 去重
            seen.add(name)
            found.append(f)
    return found


def patient_id_from_name(name):
    return name


def main():
    log_lines = []
    files = collect_files()
    log_lines.append(f"共发现 {len(files)} 个 CRF 文件")
    all_records = []
    info_records = []

    for f in files:
        base = os.path.basename(f)
        try:
            wb = openpyxl.load_workbook(f, data_only=True)
        except Exception as e:
            log_lines.append(f"[ERROR] 无法打开: {base} -> {e}")
            continue
        patient_info = {}
        for sn in wb.sheetnames:
            if not re.search(r"第\s*\d+\s*天", sn):
                continue
            ws = wb[sn]
            recs, info = parse_sheet(ws, base, sn)
            if info and not patient_info:
                patient_info = info
            elif info:
                # 多 sheet 信息合并（取第一个非空字段）
                for k, v in info.items():
                    if k not in patient_info or not patient_info[k]:
                        patient_info[k] = v
            all_records.extend(recs)
            log_lines.append(f"{base} | {sn} | rows={len(recs)}")
        info_records.append({
            "文件": base,
            "姓名": patient_info.get("姓名"),
            "住院号": patient_info.get("住院号"),
            "年龄": patient_info.get("年龄"),
            "入组日期": patient_info.get("入组日期"),
        })

    df = pd.DataFrame(all_records)
    info_df = pd.DataFrame(info_records).drop_duplicates(subset=["文件"]).reset_index(drop=True)

    # 排序
    if not df.empty:
        df.sort_values(["姓名", "第几天", "小时"], inplace=True, kind="stable")
        df.reset_index(drop=True, inplace=True)

    # ====== 数据合理性核查与自动修正 ======
    qa_records = []
    if not df.empty:
        # 先转数值
        for c in ["体温", "心率", "呼吸", "血糖", "血压_收缩", "血压_舒张"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce")
        df["体温"] = pd.to_numeric(df["体温"], errors="coerce")
        df["血糖"] = pd.to_numeric(df["血糖"], errors="coerce")

        def log_qa(姓名, 天, tp, 列, 原值, 新值, 操作, 说明):
            qa_records.append({
                "姓名": 姓名, "第几天": 天, "时间点": tp,
                "字段": 列, "原值": 原值, "修正后": 新值,
                "处理": 操作, "说明": 说明,
            })

        # 体温 > 100°C 视为缺小数点
        mask = df["体温"] > 100
        for i in df.index[mask]:
            old = df.at[i, "体温"]
            new = round(old / 10, 1)
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "体温", old, new, "自动修正", "体温>100°C，按缺少小数点处理 /10")
            df.at[i, "体温"] = new

        # 血糖 > 40 mmol/L 视为缺小数点
        mask = df["血糖"] > 40
        for i in df.index[mask]:
            old = df.at[i, "血糖"]
            new = round(old / 10, 2)
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "血糖", old, new, "自动修正", "血糖>40，按缺少小数点处理 /10")
            df.at[i, "血糖"] = new

        # 仅标记，不自动修改 ——————————————————————
        # 体温过低
        for i in df.index[(df["体温"].notna()) & ((df["体温"] < 34) | (df["体温"] > 42))]:
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "体温", df.at[i, "体温"], None, "仅标记", "体温超出 34–42°C 合理范围")

        # 心率/呼吸 同值且数值偏小或偏大（疑似录入复制错位）
        for i in df.index[(df["心率"].notna()) & (df["呼吸"].notna()) & (df["心率"] == df["呼吸"])]:
            v = df.at[i, "心率"]
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "心率/呼吸", f"{v}/{v}", None, "仅标记",
                   "心率与呼吸数值相同，疑似单元格复制错误")

        # 心率/呼吸 越界
        for i in df.index[(df["心率"].notna()) & ((df["心率"] < 30) | (df["心率"] > 200))]:
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "心率", df.at[i, "心率"], None, "仅标记", "心率超出 30–200 bpm 合理范围")
        for i in df.index[(df["呼吸"].notna()) & ((df["呼吸"] < 5) | (df["呼吸"] > 60))]:
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "呼吸", df.at[i, "呼吸"], None, "仅标记", "呼吸超出 5–60 次/分 合理范围")

        # PEEP 越界
        peep = pd.to_numeric(df.get("PEEP"), errors="coerce")
        for i in df.index[(peep.notna()) & ((peep < 0) | (peep > 20))]:
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "PEEP", df.at[i, "PEEP"], None, "仅标记", "PEEP 超出 0–20 cmH2O 合理范围")

        # 单次胃肠减压量 >1500
        gi = pd.to_numeric(df.get("胃肠减压_液体"), errors="coerce")
        for i in df.index[gi.notna() & (gi > 1500)]:
            log_qa(df.at[i, "姓名"], df.at[i, "第几天"], df.at[i, "时间点"],
                   "胃肠减压_液体", df.at[i, "胃肠减压_液体"], None, "仅标记",
                   "单次胃肠减压量 >1500ml，请核对")

    qa_df = pd.DataFrame(qa_records)

    # 按天汇总（每患者每天的均值/总和）
    daily = pd.DataFrame()
    if not df.empty:
        num_cols = ["腹内压", "体温", "心率", "呼吸", "血糖"]
        sum_cols = ["排便量", "肠内营养", "输液量_晶体", "输液量_胶体",
                    "尿量", "胃肠减压_液体", "腹腔引流量", "血滤脱水量"]
        for c in num_cols + sum_cols:
            df[c] = pd.to_numeric(df[c], errors="coerce")
        agg = {c: "mean" for c in num_cols}
        agg.update({c: "sum" for c in sum_cols})
        daily = (df.groupby(["姓名", "住院号", "第几天"], dropna=False)
                   .agg(agg).round(2).reset_index())
        daily.rename(columns={c: f"{c}_日均" for c in num_cols}, inplace=True)
        daily.rename(columns={c: f"{c}_日合计" for c in sum_cols}, inplace=True)

    # 宽格式：一个患者一行
    # 列结构：基本信息 + 监测总天数 + 全程汇总(均值/合计/最大值) + 按天汇总(D1..Dn) + 时点明细(D{d}_T{h}_变量)
    wide = pd.DataFrame()
    if not df.empty:
        all_num_cols = ["腹内压", "体温", "心率", "呼吸", "血糖"]
        all_sum_cols = ["排便量", "肠内营养", "输液量_晶体", "输液量_胶体",
                        "尿量", "胃肠减压_液体", "腹腔引流量", "血滤脱水量"]
        flag_cols = ["是否俯卧位", "是否使用升压药", "是否使用镇痛药物",
                     "是否使用机械通气", "是否进行血液滤过"]
        text_cols = ["血压", "肠内类型"]

        # 基本信息（每个患者一行）
        base = (df.groupby("姓名", dropna=False)
                  .agg(住院号=("住院号", "first"),
                       年龄=("年龄", "first"),
                       入组日期=("入组日期", "first"),
                       文件=("文件", "first"),
                       监测总天数=("第几天", "nunique"),
                       监测总时点数=("时间点", "count"))
                  .reset_index())

        # 全程汇总
        overall_agg = {c: "mean" for c in all_num_cols}
        overall_agg.update({c: "sum" for c in all_sum_cols})
        overall = (df.groupby("姓名", dropna=False)
                     .agg(overall_agg).round(2).reset_index())
        overall.rename(columns={c: f"{c}_全程均值" for c in all_num_cols}, inplace=True)
        overall.rename(columns={c: f"{c}_全程合计" for c in all_sum_cols}, inplace=True)
        # 腹内压最大值
        max_iap = (df.groupby("姓名", dropna=False)["腹内压"]
                     .max().reset_index().rename(columns={"腹内压": "腹内压_全程最大"}))
        # 是否曾出现腹内高压（腹内压 ≥12 / 20）
        iah = (df.groupby("姓名", dropna=False)["腹内压"]
                 .apply(lambda s: "是" if (s.dropna() >= 12).any() else "否")
                 .reset_index().rename(columns={"腹内压": "曾出现腹内高压(IAP≥12)"}))
        acs = (df.groupby("姓名", dropna=False)["腹内压"]
                 .apply(lambda s: "是" if (s.dropna() >= 20).any() else "否")
                 .reset_index().rename(columns={"腹内压": "曾出现ACS(IAP≥20)"}))

        wide = base.merge(overall, on="姓名", how="left") \
                   .merge(max_iap, on="姓名", how="left") \
                   .merge(iah, on="姓名", how="left") \
                   .merge(acs, on="姓名", how="left")

        # 按天汇总(横向展开 D1.. Dn)
        if not daily.empty:
            day_pivot = daily.copy()
            day_pivot["第几天"] = "D" + day_pivot["第几天"].astype(int).astype(str)
            day_metric_cols = [c for c in day_pivot.columns
                               if c not in ("姓名", "住院号", "第几天")]
            day_pivot = day_pivot.pivot(index="姓名",
                                        columns="第几天",
                                        values=day_metric_cols)
            # 列变成 (变量, Dx) -> 拼成 Dx_变量
            day_pivot.columns = [f"{d}_{m}" for m, d in day_pivot.columns]
            # 按天序号、再按变量名排序
            def _day_sort_key(c):
                d, _, m = c.partition("_")
                return (int(d[1:]), m)
            day_pivot = day_pivot.reindex(
                sorted(day_pivot.columns, key=_day_sort_key), axis=1
            ).reset_index()
            wide = wide.merge(day_pivot, on="姓名", how="left")

        # 时点明细横向展开：D{d}_T{h}_{变量}
        det_cols = all_num_cols + all_sum_cols + flag_cols + text_cols
        wide_long = df[["姓名", "第几天", "小时"] + det_cols].copy()
        wide_long["第几天"] = wide_long["第几天"].astype("Int64")
        wide_long = wide_long.dropna(subset=["第几天", "小时"])
        wide_long["dt"] = ("D" + wide_long["第几天"].astype(int).astype(str)
                          + "_T" + wide_long["小时"].astype(int).astype(str))
        det_pivot = wide_long.pivot_table(index="姓名",
                                          columns="dt",
                                          values=det_cols,
                                          aggfunc="first")
        det_pivot.columns = [f"{dt}_{m}" for m, dt in det_pivot.columns]

        def _dt_sort_key(c):
            head, _, m = c.partition("_")
            d_str, t_str = head.split("_") if "_" in head else (head, "T0")
            # head 已经是 "D1_T0" 形式，但前面 partition 错了，重新分
            return c
        # 正确解析 D{d}_T{h}_{m}
        def _dt_sort_key2(c):
            parts = c.split("_", 2)
            d = int(parts[0][1:]) if parts[0].startswith("D") else 99
            t = int(parts[1][1:]) if len(parts) > 1 and parts[1].startswith("T") else 99
            m = parts[2] if len(parts) > 2 else ""
            return (d, t, m)

        det_pivot = det_pivot.reindex(
            sorted(det_pivot.columns, key=_dt_sort_key2), axis=1
        ).reset_index()
        wide = wide.merge(det_pivot, on="姓名", how="left")

    # 输出
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M")
    out_dir = os.path.join(ROOT, "output", ts)
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"CRF合并_腹内压_{ts}.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        info_df.to_excel(w, sheet_name="患者基本信息", index=False)
        df.to_excel(w, sheet_name="时点明细", index=False)
        if not daily.empty:
            daily.to_excel(w, sheet_name="按天汇总", index=False)
        if not wide.empty:
            wide.to_excel(w, sheet_name="宽表_一人一行", index=False)
        if not qa_df.empty:
            qa_df.to_excel(w, sheet_name="数据核查记录", index=False)

    log_lines.append(f"输出文件: {out_path}")
    log_lines.append(f"时点明细行数: {len(df)}; 患者数: {info_df['姓名'].nunique()}")

    log_dir = os.path.join(ROOT, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, f"merge_crf_{ts}.log")
    with open(log_path, "w", encoding="utf-8") as fp:
        fp.write("\n".join(log_lines))

    print(out_path)
    print(log_path)
    print(f"records={len(df)} patients={info_df['姓名'].nunique()}")


if __name__ == "__main__":
    main()
