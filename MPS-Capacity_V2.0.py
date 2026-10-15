"""主计划约束表生成工具 - 合并版 (V2.0)

把原先两个脚本合并为一个, 用第一个命令行参数选择数据源, 口径逐行保留原实现:

  prelaser : 原 MPS-Capacity_core_V1.1.py
             <Capacity.xlsx> <PreLaserIOS.xlsx> <PPVT_IOS.xlsx> <Mapping.xlsx>
             制程段: Pre-laser / FWDL / Fresh (Capacity 的 PA 页取 W/H、UPH、效率、良率)

  fatp     : 原 MPS-Capacity-FATP-SUB-SMT_V1.5.py
             <SMT-Mapping.xlsx> <Sub-FATP-Mapping.xlsx> <FATP_Capacity.xlsx> <SMT&Sub_Capacity.xlsx>
             制程段: 由两个 Mapping 表决定 (FATP / SUB / PCM / SMT / Coating / SA ...)

输出契约两个模式完全一致: 固定文件名 主计划约束表-导入模板.csv、16 列固定列序、
utf-8-sig、写到源文件所在目录、主计划版本号 H36aV{MMDD}。

用法:
  MPS-Capacity_V2.0.py prelaser <Capacity.xlsx> <PreLaserIOS.xlsx> <PPVT_IOS.xlsx> <Mapping.xlsx>
  MPS-Capacity_V2.0.py fatp     <SMT-Mapping.xlsx> <Sub-FATP-Mapping.xlsx> <FATP_Capacity.xlsx> <SMT&Sub_Capacity.xlsx>

V2.0 合并与口径变更:
  - 合并原两个脚本为一个文件, 用第一个参数 prelaser / fatp 选择数据源
  - 日期起点改为"从源数据文件第一天的日期开始生成"(含历史数据):
    INCLUDE_TODAY = None, 两个模式统一生效
  - 输出 UPH 列改为源文件中该资源 UPH 行的 MAX(UPH) (与 目标UPH 同值),
    两个模式统一生效
  - 校验(UPH<300 / Input超限)仍基于源文件"当日原始 UPH", 不受上面改动影响

==========================================================================
原 MPS-Capacity_core_V1.1.py 版本记录:
V1.1 更新:
  - Input值: 四舍五入取整，DS/NS拆分确保总量一致（奇数时DS多1）
  - 目标UPH: 保留两位小数
  - UPH: 保留两位小数
V1.1.1 更新:
  - 取值起点改为"从当前日期开始": 只输出 今天及以后 的日期行(含今天),
    自动剔除源文件中早于今天的历史日期(Pre-Laser / PPVT 的 For Plan 页
    与 Capacity 的 PA 页使用同一起点, 保持按日期取值关系不变)
  - 起点由文件顶部 INCLUDE_TODAY 控制: True=含今天(默认), False=从明天起,
    None=不过滤(生成全部日期, 含历史)
  - 兜底: 若源文件所有日期都早于今天, 则取消过滤(生成全部), 避免输出空文件
  - 运行报告新增 日期起点 与 输出日期范围 两行, 便于核对
  - 工作时间(小时) 由"全天工时/2"改为按整班次向上取整并夹在 [8, 10]:
      < 8 -> 8, 8~9 -> 9, 9~10 -> 10, > 10 -> 10
    (由 NORM_WORK_HOURS 控制, 置 False 可保留原始值;
     原始全天工时 wt 不受影响, 仍用于 W/H全天>20 与 Input超限 校验)
  - 运行报告新增 工作时间班次取整统计 一段

原 MPS-Capacity-FATP-SUB-SMT_V1.5.py 版本记录:
V1.5.1 更新:
  - 日期起点改为"从当前日期开始": 只生成 今天及以后 的日期行(含今天),
    自动剔除源文件中早于今天的历史日期, 不再输出已过去的日期数据。
  - 起点由文件顶部 INCLUDE_TODAY 控制: True=含今天(默认), False=从明天起,
    None=不过滤(生成全部日期, 含历史)。
  - 日期过滤在 日期->列 映射阶段完成, For Plan(Input) 与产能页(W/H、UPH、
    效率、良率)使用同一个起点, 保持按日期匹配关系不变。
  - 取整余数平衡(累计误差调整到 Input 最大的一天)只在起点之后的日期内进行。
  - 兜底: 若源文件所有日期都早于今天, 则取消过滤(生成全部), 避免输出空文件。
  - 运行报告新增 日期起点 与 输出日期范围 两行, 便于核对。
V1.5 更新 (20260911):
  - 工作时间(小时) 填充时按整班次取整(单班工时 = 全天工时 / 2):
      < 8 小时 -> 8, 8~9 -> 9, 9~10 -> 10, > 10 -> 10 (ceil 后夹在 [8, 10])
    原始全天工时(wt)不受影响, 仍用于 'W/H全天>20' 与 'Input超限' 校验。
V1.4 更新:
  - 同时读取两个 Mapping 表(SMT + Sub FATP), 填充 Mapping 中出现的所有制程段
  - SMT Mapping 制程段(PCM/SMT/Coating/SA): 数据取自 SMT&Sub 产能文件
    (资源编码含 L/R 用 Buds Assumption, 否则 Case Assumption), 取数逻辑同 SUB
  - 制程段/资源类型 输出 Mapping 中的原始制程段值(如 SMT-L / Coating-FT / SA / PCM-L)
V1.3 更新 (20260811):
  - 移除 物品类型 生成 (FATP/SUB/SMT 均不需要), 输出该列置空
V1.2 更新 (20260811, 数据源由 4 文件改为 3 文件, 制程体系由 pre-laser/fwdl/fresh 改为 FATP/SUB):
  - FATP 行: Input 取 FATP 文件 'For Plan' 页(行号 = mapping 第7列, 日期行1自列14);
    W/H/UPH/效率/良率 取 Left/Right/Case 页(按资源编码含 -L/-R 选页, 无 L/R 用 Case),
    行号 = mapping 第25/27/29、31/33/35、37/39/41、43/45/47 列, 日期行2自列3
  - SUB 行: Input 取 SMT&Sub 文件 'For Plan' 页(行号 = mapping 第5列, 日期行1自列14);
    W/H/UPH/效率/良率 取 Buds/Case Assumption 页(资源编码含 L/R 用 Buds),
    行号 = mapping 第9/11、13/15、17/19、21/23 列, 日期行4自列4
  - 日期轴: 以产能页日期为准, For Plan 页按日期匹配取 Input
  - Input 值: 四舍五入取整, DS/NS 拆分确保总量一致(奇数时NS多1)
  - 目标UPH: 保留两位小数; 工作时间/效率/良率: 保留6位
==========================================================================
"""

import pandas as pd, numpy as np, re, os, sys, math
from datetime import datetime, date, timedelta

# ============================================================
# 日期起点设置
#   None  -> 不过滤: 从源数据文件中最早的日期开始生成 (含历史数据)  <- V2.0 默认
#   True  -> 只生成"今天及以后"的日期行 (含今天)
#   False -> 只生成"明天及以后"的日期行 (不含今天)
# ============================================================
INCLUDE_TODAY = None

# ============================================================
# 工作时间(小时) 班次取整设置 (仅 prelaser 模式使用; fatp 模式恒开启取整)
#   True  -> 单班工时按整班次向上取整并夹在 [8, 10] (默认):
#              < 8         -> 8
#              8 ~ 9       -> 9
#              9 ~ 10      -> 10
#              > 10        -> 10
#   False -> 保留原始值(全天工时 / 2, 不取整)
#   注: 原始全天工时 wt 不受影响, 仍用于 'W/H全天>20' 与 'Input超限' 校验
# ============================================================
NORM_WORK_HOURS = True


# ============================================================
# 公共函数
# ============================================================

def get_start_date():
    """返回本次生成允许的最早日期; None 表示不过滤(生成全部日期)"""
    if INCLUDE_TODAY is None:
        return None
    t = date.today()
    return t if INCLUDE_TODAY else t + timedelta(days=1)


def norm_wh(h):
    """单班工作时间(小时)按整班次取整:
      < 8   -> 8
      8~9   -> 9
      9~10  -> 10
      > 10  -> 10
    即向上取整后夹在 [8, 10] 区间 (先 round 6 位规避浮点误差)。"""
    if h is None: return None
    v = math.ceil(round(float(h), 6))
    return float(min(10, max(8, v)))


def pn(s):
    """从 '第441行' 等文本中提取行号"""
    if pd.isna(s): return None
    m = re.search(r'(\d+)', str(s).replace(',', ''))
    return int(m.group(1)) if m else None


# ============================================================
# 数据源一: Pre-laser / FWDL / Fresh
#   (原 MPS-Capacity_core_V1.1.py)
# ============================================================

def ed(df, min_date=None):
    """For Plan(Pre-Laser) 页: 列号 -> 日期。min_date 非空时跳过更早的日期"""
    dm = {}
    for j in range(len(df.columns)):
        for ri in [0, 1]:
            if ri < len(df):
                v = df.iloc[ri, j]
                if pd.notna(v):
                    try:
                        d = pd.Timestamp(v).date()
                        if min_date is not None and d < min_date:
                            break          # 该列已解析出日期且早于起点 -> 跳过本列
                        dm[j] = d
                        break
                    except:
                        pass
    return dm


def ed_ppvt(df, min_date=None):
    """For Plan(PPVT) 页: 列号 -> 日期。min_date 非空时跳过更早的日期"""
    dm = {}
    for j in range(14, len(df.columns)):
        v = df.iloc[0, j]
        if pd.notna(v):
            try:
                d = pd.Timestamp(v).date()
                if min_date is not None and d < min_date:
                    continue
                dm[j] = d
            except:
                pass
    return dm


def bld_pa_d2c(pa, min_date=None):
    """Capacity 的 PA 页: 日期 -> 列号 (取第2行日期)。min_date 非空时跳过更早的日期"""
    dm = {}
    for j in range(pa.shape[1]):
        v = pa.iloc[1, j]
        if pd.notna(v):
            try:
                d = pd.Timestamp(v).date()
                if d.year <= 2020: continue
                if min_date is not None and d < min_date:
                    continue
                dm[d] = j
            except: pass
    return dm


def pa_val(pa, pa_d2c, r, td):
    if r is None: return None
    c = pa_d2c.get(td)
    if c is not None:
        v = pa.iloc[r, c]
        if pd.notna(v) and isinstance(v, (int, float)) and v > 0:
            return float(v)
        return None
    for j in range(3, pa.shape[1]):
        v = pa.iloc[r, j]
        if pd.notna(v) and isinstance(v, (int, float)) and v > 0:
            return float(v)
    return None


def pa_maxr(pa, r):
    if r is None: return None
    b = 0.0
    for j in range(3, pa.shape[1]):
        v = pa.iloc[r, j]
        if pd.notna(v) and isinstance(v, (int, float)) and v > b:
            b = float(v)
    return b if b > 0 else None


def get_item_type(proc, fl):
    """Pre-laser: 根据线体关键字确定物品类型"""
    if proc != 'pre-laser':
        return ''
    fl_u = str(fl).upper()
    if ' C' in fl_u or fl_u.startswith('C'):
        return 'Case'
    if ' L' in fl_u or fl_u.startswith('L'):
        return 'L'
    if ' R' in fl_u or fl_u.startswith('R'):
        return 'R'
    return ''


def bld_pre_idx(df):
    idx, last = {}, ''
    for i in range(len(df)):
        c = df.iloc[i, 2]
        d = df.iloc[i, 3]
        if pd.notna(c): last = str(c).strip()
        if pd.notna(d) and 'Plan Input' in str(d):
            idx[f"{last} {str(d).strip()}"] = i
    return idx


def bld_ppvt_idx(df):
    idx, last = {}, ''
    for i in range(len(df)):
        m = df.iloc[i, 12]
        n = df.iloc[i, 13]
        if pd.notna(m): last = str(m).strip()
        if pd.notna(n) and 'Plan Input' in str(n):
            idx[f"{last} {str(n).strip()}"] = i
    return idx


def load_upu_map(path):
    """读取《目标UPHmapping》表: 机种 | 制程段 | 资源编码 | 目标UPH
    返回 {(机种, 制程段, 资源编码): 目标UPH}, 严格一致匹配(不忽略大小写与空格)"""
    df = pd.read_excel(path, engine='openpyxl')
    m = {}
    for _, r in df.iterrows():
        if pd.isna(r.iloc[0]) or pd.isna(r.iloc[1]) or pd.isna(r.iloc[2]) or pd.isna(r.iloc[3]):
            continue
        m[(str(r.iloc[0]), str(r.iloc[1]), str(r.iloc[2]))] = float(r.iloc[3])
    return m


def run_prelaser(path_capacity, path_pre, path_ppvt, path_mapping, path_upu_map=None):
    print("读取文件...")
    mapping = pd.read_excel(path_mapping, engine='openpyxl')
    fp_pre = pd.read_excel(path_pre, sheet_name='For Plan', engine='openpyxl')
    fp_ppvt = pd.read_excel(path_ppvt, sheet_name='For Plan', engine='openpyxl')
    pa = pd.read_excel(path_capacity, sheet_name='PA', engine='openpyxl')
    upu = load_upu_map(path_upu_map) if path_upu_map else None
    if upu is not None:
        print(f"目标UPHmapping: {os.path.basename(path_upu_map)} 共 {len(upu)} 条 (机种+制程段+资源编码 严格匹配)")
    else:
        print("目标UPHmapping: 未提供, 目标UPH 取源文件 MAX(UPH)")

    # 取值起点: 默认今天(含当天), 自动剔除历史日期
    sd = get_start_date()
    pa_d2c = bld_pa_d2c(pa, sd)

    dm_pre = ed(fp_pre, sd)
    dm_ppvt = ed_ppvt(fp_ppvt, sd)
    pre_idx = bld_pre_idx(fp_pre)
    ppvt_idx = bld_ppvt_idx(fp_ppvt)
    print(f"Pre-Laser: {len(pre_idx)}个输入标签, PPVT: {len(ppvt_idx)}个输入标签")
    print(f"日期轴(>=起点): Pre-Laser {len(dm_pre)}天, PPVT {len(dm_ppvt)}天, PA {len(pa_d2c)}天")
    if sd is not None:
        print(f"日期起点: {sd} (含当天, 已剔除该日期之前的历史数据)")
    else:
        print("日期起点: 不过滤 (从源数据文件最早日期开始, 含历史数据)")

    # 安全兜底: 若源文件全部为历史日期, 则不过滤, 避免输出空文件
    filtered = True
    if sd is not None and len(dm_pre) == 0 and len(dm_ppvt) == 0 and len(pa_d2c) == 0:
        print(f"[!] 源文件中没有 >= {sd} 的日期, 已取消日期过滤")
        pa_d2c = bld_pa_d2c(pa)
        dm_pre = ed(fp_pre)
        dm_ppvt = ed_ppvt(fp_ppvt)
        filtered = False

    EFF_R = 51
    ver = f"H36aV{datetime.now().strftime('%m%d')}"
    rows, iss = [], {'W/H全天>20': [], 'UPH<300': [], '良率>1': [], 'Input超限': []}
    upu_hit, upu_miss_seen, upu_miss = set(), set(), []
    wh_adj = {'<8->8': 0, '8~9->9': 0, '9~10->10': 0, '>10->10': 0, '不变': 0}

    for _, mr in mapping.iterrows():
        proc = str(mr.iloc[1]).strip().lower()
        res = str(mr.iloc[3]).strip()
        mdl = str(mr.iloc[0]).strip() if pd.notna(mr.iloc[0]) else ''
        fl = str(mr.iloc[4]).strip() if pd.notna(mr.iloc[4]) else ''
        fr, wr, ur, yr = pn(mr.iloc[5]), pn(mr.iloc[7]), pn(mr.iloc[9]), pn(mr.iloc[13])

        fm = dm_pre if proc == 'pre-laser' else dm_ppvt
        im = pre_idx if proc == 'pre-laser' else ppvt_idx
        src = fp_pre if proc == 'pre-laser' else fp_ppvt

        dr = None
        # Input值取源文件哪一行由 mapping 第5列决定: 指向 Plan Output 行则取 Output 值,
        # 指向 Plan Input 行则取 Input 值。第5列为 Excel 1-based 行号, pandas 读入时
        # 首行作表头被丢弃, 实际索引 = 行号 - 表头偏移; 用行标签校验偏移, 兼容表头行数变化
        if fr and 1 <= fr <= len(src):
            suf = fl[-2:] if len(fl) >= 2 else ''
            for off in (2, 1, 3):
                cand = fr - off
                if 0 <= cand < len(src):
                    lbl = str(src.iloc[cand, 3] if proc == 'pre-laser' else src.iloc[cand, 13]).strip()
                    if ('Plan Output' in lbl or 'Plan Input' in lbl) and (not suf or lbl.endswith(suf)):
                        dr = cand; break
        if dr is None:
            for lb, ri in im.items():
                if fl.replace(' ', '') in lb.replace(' ', ''):
                    dr = ri; break

        w0 = wr - 1 if wr else None
        u0 = ur - 1 if ur else None
        y0 = yr - 1 if yr else None

        if proc == 'fwdl':
            if wr == 184: w0 = 182
            if ur == 181: u0 = 179
            if yr == 182: y0 = 180
        elif proc == 'pre-laser':
            if ur:
                if ur >= 237: u0 = ur - 3
                else: u0 = ur - 2
            if yr:
                if yr >= 245: y0 = yr - 3
                else: y0 = yr - 2
            if wr == 249: w0 = 247
        elif proc == 'fresh':
            if wr == 25: w0 = 23
            if ur == 33: u0 = 31
            if yr == 49: y0 = 47

        for cj, td in sorted(fm.items()):
            inp = None
            if dr is not None:
                v = src.iloc[dr, cj]
                if pd.notna(v) and v != 0:
                    inp = float(v)

            wt = pa_val(pa, pa_d2c, w0, td)
            up = pa_val(pa, pa_d2c, u0, td)
            yl = pa_val(pa, pa_d2c, y0, td)
            ef = pa_val(pa, pa_d2c, EFF_R, td)
            mu = pa_maxr(pa, u0)

            if wt is None or wt == 0: continue
            if proc in ('fwdl', 'fresh'):
                if up is None or up == 0: continue
                if proc == 'fwdl' and (inp is None or inp == 0): continue
                if proc == 'fresh' and (inp is None or inp == 0): continue
            if inp is None: continue

            # 工作时间(小时): 全天工时/2 为单班工时, 再按整班次向上取整夹在 [8, 10]
            wh_raw = wt / 2.0
            wh_h = norm_wh(wh_raw) if NORM_WORK_HOURS else wh_raw
            if not NORM_WORK_HOURS or wh_h == wh_raw:
                wh_adj['不变'] += 1
            elif wh_raw < 8:
                wh_adj['<8->8'] += 1
            elif wh_raw < 9:
                wh_adj['8~9->9'] += 1
            elif wh_raw < 10:
                wh_adj['9~10->10'] += 1
            else:
                wh_adj['>10->10'] += 1

            if wt > 20: iss['W/H全天>20'].append((proc, res, td, wt))
            if up is not None and up < 300: iss['UPH<300'].append((proc, res, td, up))
            if yl is not None and yl > 1: iss['良率>1'].append((proc, res, td, yl))
            if inp > 0 and wt and up and up > 0:
                lm = wt * up
                if inp > lm * 1.02:
                    iss['Input超限'].append((proc, res, td, round(inp, 0), round(lm, 0)))

            # V1.1: Input值四舍五入取整，DS/NS拆分确保总量一致
            inp_rounded = round(inp)
            ds_val = inp_rounded // 2          # DS取一半（向下取整）
            ns_val = inp_rounded - ds_val      # NS取剩余（奇数时NS多1，确保总和=inp_rounded）

            # V2.1: 目标UPH 优先取《目标UPHmapping》(机种+制程段+资源编码 严格匹配), 未命中沿用 MAX(UPH)
            tgt = round(mu, 2) if mu else None
            if upu is not None:
                k = (mdl, str(mr.iloc[1]), res)
                if k in upu:
                    tgt = round(upu[k], 2)
                    upu_hit.add(k)
                elif k not in upu_miss_seen:
                    upu_miss_seen.add(k)
                    upu_miss.append((k[0], k[1], k[2], round(mu, 2) if mu else None))

            for s, inp_val in [('DS', ds_val), ('NS', ns_val)]:
                rows.append({
                    '主计划版本号': ver, '机种': mdl,
                    '制程段': mr.iloc[1], '物品类型': get_item_type(proc, fl),
                    '资源类型': mr.iloc[1], '尺寸': mr.iloc[2],
                    '资源编码': res, '白/夜班': s,
                    '日期': str(td), '工作时间(小时)': wh_h,
                    '目标UPH': tgt,                                # V1.1: 保留两位小数 / V2.1: mapping 命中取 mapping 值
                    '开线班数': 1,
                    'Input值': inp_val,                            # V1.1: 整数
                    'UPH': round(mu, 2) if mu else None,           # V2.0: 取源文件该行 MAX(UPH)
                    '效率': round(ef, 6) if ef else None,
                    '良率': round(yl, 6) if yl else None,
                })

    cols = ['主计划版本号', '机种', '制程段', '物品类型', '资源类型', '尺寸',
            '资源编码', '白/夜班', '日期', '工作时间(小时)', '目标UPH',
            '开线班数', 'Input值', 'UPH', '效率', '良率']
    df = pd.DataFrame(rows, columns=cols)
    # V1.1: 工作时间/效率/良率保留6位，目标UPH/UPH已在上方取2位，Input值为整数
    for c in ['工作时间(小时)', '效率', '良率']:
        df[c] = df[c].astype(float).round(6)

    output_dir = os.path.dirname(path_capacity)
    output_path = os.path.join(output_dir, '主计划约束表-导入模板.csv')
    try:
        df.to_csv(output_path, index=False, encoding='utf-8-sig', float_format='%.4f')
    except PermissionError:
        _ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        _alt = os.path.join(output_dir, f'主计划约束表-导入模板_{_ts}.csv')
        print(f"[WARN] 输出文件被占用 (Excel 未关闭或被其他程序锁定): {output_path}")
        print(f"       已改用安全文件名写入: {_alt}")
        df.to_csv(_alt, index=False, encoding='utf-8-sig', float_format='%.4f')
        output_path = _alt

    print()
    print("========== 生成报告 ==========")
    print(f"输出文件: {output_path}")
    print(f"总行数: {len(df)}")
    if len(df):
        _d = pd.to_datetime(df['日期'])
        print(f"日期范围: {_d.min().date()} ~ {_d.max().date()} (共 {_d.dt.normalize().nunique()} 天)")
    else:
        print("[!] 未生成任何数据行, 请检查源文件日期范围与系统日期是否匹配")
    print()
    for p in ['Pre-laser', 'FWDL', 'Fresh']:
        s = df[df['制程段'] == p]
        print(f"  {p}: {len(s)}行  "
              f"Inp>0={(s['Input值'] > 0).sum()}  "
              f"UPH={s['UPH'].notna().sum()}  "
              f"Yld={s['良率'].notna().sum()}")
    print()
    print("--- 工作时间(小时)班次取整 ---")
    if not NORM_WORK_HOURS:
        print("  (已关闭: NORM_WORK_HOURS = False, 保留原始 全天工时/2)")
    for k, n in wh_adj.items():
        print(f"  {k}: {n}条")
    if len(df):
        vals = sorted(df['工作时间(小时)'].dropna().unique().tolist())
        print(f"  输出取值范围: {[round(v, 6) for v in vals]}")
    print()
    if upu is not None:
        print("--- 目标UPHmapping (机种+制程段+资源编码 严格匹配) ---")
        print(f"  mapping条数: {len(upu)}  命中组合: {len(upu_hit)}  "
              f"未使用组合: {len(upu) - len(upu_hit)}  未命中组合: {len(upu_miss)}")
        if upu_miss:
            print("  未命中清单 (可复制到 目标UPHmapping.xlsx 补充):")
            print("    机种\t制程段\t资源编码\t当前目标UPH")
            for r in upu_miss[:50]:
                print(f"    {r[0]}\t{r[1]}\t{r[2]}\t{r[3]}")
            if len(upu_miss) > 50:
                print(f"    ... 另有 {len(upu_miss) - 50} 个组合未列出")
        print()
    print("--- 校验 ---")
    for k, its in iss.items():
        if its: print(f"  [!] {k}: {len(its)}条")
        else:   print(f"  [OK] {k}: 无问题")
    return output_path


# ============================================================
# 数据源二: FATP / SUB / SMT
#   (原 MPS-Capacity-FATP-SUB-SMT_V1.5.py)
# ============================================================

def bld_d2c(df, date_row, start_col, min_date=None):
    """建立 日期->列号 映射 (min_date 非空时只保留 >= min_date 的日期)"""
    dm = {}
    for j in range(start_col, df.shape[1]):
        v = df.iloc[date_row, j]
        if pd.notna(v):
            try:
                d = pd.Timestamp(v).date()
                if d.year > 2020:
                    if min_date is not None and d < min_date:
                        continue
                    dm[d] = j
            except:
                pass
    return dm


def val_at(df, dm, r, td):
    """按日期取某行数值(仅 >0 有效)"""
    if r is None: return None
    c = dm.get(td)
    if c is not None:
        v = df.iloc[r, c]
        if pd.notna(v) and isinstance(v, (int, float)) and v > 0:
            return float(v)
        return None
    for j in range(df.shape[1]):
        v = df.iloc[r, j]
        if pd.notna(v) and isinstance(v, (int, float)) and v > 0:
            return float(v)
    return None


def row_max(df, r, start_col):
    """取某行整行最大值 (用于目标UPH)"""
    if r is None: return None
    b = 0.0
    for j in range(start_col, df.shape[1]):
        v = df.iloc[r, j]
        if pd.notna(v) and isinstance(v, (int, float)) and v > b:
            b = float(v)
    return b if b > 0 else None


def is_lr(proc, res):
    """制程段名或资源编码是否带 L/R 标识 (用于选 Buds/Case Assumption 页)。
    同时兼容 'H36a-L-FT' 与 'H36aL-P2i' 两种 L/R 命名 (后者 L/R 前无连字符)。"""
    pu, ru = str(proc).upper(), str(res).upper()
    return ('-L' in pu) or ('-R' in pu) or ('-L' in ru) or ('-R' in ru)


def run_fatp_sub_smt(path_smt_map, path_sub_map, path_fatp, path_smt, path_upu_map=None):
    print("读取文件...")
    sub_map = pd.read_excel(path_sub_map, engine='openpyxl')
    smt_map = pd.read_excel(path_smt_map, engine='openpyxl')
    fatp_fp = pd.read_excel(path_fatp, sheet_name='For Plan', header=None)
    smt_fp  = pd.read_excel(path_smt,  sheet_name='For Plan', header=None)
    upu = load_upu_map(path_upu_map) if path_upu_map else None
    if upu is not None:
        print(f"目标UPHmapping: {os.path.basename(path_upu_map)} 共 {len(upu)} 条 (机种+制程段+资源编码 严格匹配)")
    else:
        print("目标UPHmapping: 未提供, 目标UPH 取源文件 MAX(UPH)")

    # For Plan 日期列: 行1(0-based) 自列14起
    # 只保留 >= 起始日期 的日期 (起始日期 = 今天, 见文件顶部 INCLUDE_TODAY)
    sd = get_start_date()
    fatp_fp_d2c = bld_d2c(fatp_fp, 1, 14, sd)
    smt_fp_d2c  = bld_d2c(smt_fp, 1, 14, sd)
    print(f"FATP For Plan 日期: {len(fatp_fp_d2c)}天, "
          f"SMT&Sub For Plan 日期: {len(smt_fp_d2c)}天")
    if sd is not None:
        print(f"日期起点: {sd} (含当天, 已剔除该日期之前的历史数据)")
    else:
        print("日期起点: 不过滤 (从源数据文件最早日期开始, 含历史数据)")

    # 安全兜底: 若源文件全部为历史日期, 则不过滤, 避免输出空文件
    if sd is not None:
        if len(fatp_fp_d2c) == 0 and len(smt_fp_d2c) == 0:
            print(f"[!] 源文件中没有 >= {sd} 的日期, 已取消日期过滤")
            fatp_fp_d2c = bld_d2c(fatp_fp, 1, 14)
            smt_fp_d2c  = bld_d2c(smt_fp, 1, 14)

    # 产能页缓存 (按 文件+页面)
    cap_cache = {}

    def get_cap(f, sheet):
        key = (f, sheet)
        if key not in cap_cache:
            cap_cache[key] = pd.read_excel(f, sheet_name=sheet, header=None)
        return cap_cache[key]

    ver = f"H36aV{datetime.now().strftime('%m%d')}"
    rows, iss = [], {'W/H全天>20': [], 'UPH<300': [], '良率>1': [], 'Input超限': []}
    upu_hit, upu_miss_seen, upu_miss = set(), set(), []
    wh_adj = {'<8->8': 0, '8~9->9': 0, '9~10->10': 0, '>10->10': 0, '不变': 0}

    def handle(mr, src):
        proc = str(mr.iloc[1]).strip()
        mdl = str(mr.iloc[0]).strip() if pd.notna(mr.iloc[0]) else ''
        res = str(mr.iloc[3]).strip()
        sz  = str(mr.iloc[2]).strip() if pd.notna(mr.iloc[2]) else ''
        res_u = res.upper()

        if src == 'sub':
            if proc == 'FATP':
                # 资源编码含 -L/-R 选 Left/Right 页, 否则 Case 页
                if '-L' in res_u:
                    pg, wc, uc, ec, yc, in_c = 'Left',  25, 31, 37, 43, 7
                elif '-R' in res_u:
                    pg, wc, uc, ec, yc, in_c = 'Right', 27, 33, 39, 45, 7
                else:
                    pg, wc, uc, ec, yc, in_c = 'Case',  29, 35, 41, 47, 7
                fp, fp_d2c, capf = fatp_fp, fatp_fp_d2c, path_fatp
                drow, scol, vcol = 2, 3, 3              # Left/Right/Case: 日期行2 自列3
            elif proc == 'SUB':
                # 资源编码含 L/R 用 Buds Assumption, 否则 Case Assumption
                if is_lr(proc, res):
                    pg, wc, uc, ec, yc, in_c = 'Buds Assumption', 9, 13, 17, 21, 5
                else:
                    pg, wc, uc, ec, yc, in_c = 'Case Assumption', 11, 15, 19, 23, 5
                fp, fp_d2c, capf = smt_fp, smt_fp_d2c, path_smt
                drow, scol, vcol = 4, 4, 4              # Assumption: 日期行4 自列4
            else:
                return                                  # 非 FATP/SUB 跳过
        else:  # src == 'smt'  (SMT Mapping 表内所有制程段)
            # 制程段名/资源编码含 L/R 用 Buds Assumption, 否则 Case Assumption
            # (列号同 SUB 逻辑左移2列; 兼容 A5-3F-H36aL-P2i 这类 L/R 不带连字符的资源编码)
            if is_lr(proc, res):
                pg, wc, uc, ec, yc, in_c = 'Buds Assumption', 7, 11, 15, 19, 5
            else:
                pg, wc, uc, ec, yc, in_c = 'Case Assumption', 9, 13, 17, 21, 5
            fp, fp_d2c, capf = smt_fp, smt_fp_d2c, path_smt
            drow, scol, vcol = 4, 4, 4

        in_r = pn(mr.iloc[in_c])
        cap = get_cap(capf, pg)
        d2c = bld_d2c(cap, drow, scol, sd)
        if not d2c and sd is not None:
            d2c = bld_d2c(cap, drow, scol)

        wr, ur, er, yr = pn(mr.iloc[wc]), pn(mr.iloc[uc]), pn(mr.iloc[ec]), pn(mr.iloc[yc])
        w0 = wr - 1 if wr else None
        u0 = ur - 1 if ur else None
        e0 = er - 1 if er else None
        y0 = yr - 1 if yr else None
        i0 = in_r - 1 if in_r else None

        mu = row_max(cap, u0, vcol)

        # 先收集该线所有有效日期数据, 再统一取整 + 余数平衡
        day_data = []
        for td in sorted(d2c):
            # Input: For Plan 页按日期匹配
            inp = None
            c_in = fp_d2c.get(td)
            if i0 is not None and c_in is not None:
                v = fp.iloc[i0, c_in]
                if pd.notna(v) and v != 0:
                    inp = float(v)

            wt = val_at(cap, d2c, w0, td)
            up = val_at(cap, d2c, u0, td)
            yl = val_at(cap, d2c, y0, td)
            ef = val_at(cap, d2c, e0, td)

            if inp is None: continue
            if wt is None or wt == 0: continue      # 工时为0不输出

            # 校验(基于原始值)
            if wt > 20: iss['W/H全天>20'].append((proc, res, td, wt))
            if up is not None and up < 300: iss['UPH<300'].append((proc, res, td, up))
            if yl is not None and yl > 1: iss['良率>1'].append((proc, res, td, yl))
            if inp > 0 and wt and up and up > 0:
                lm = wt * up
                if inp > lm * 1.02:
                    iss['Input超限'].append((proc, res, td, round(inp, 0), round(lm, 0)))

            day_data.append([td, inp, wt, up, yl, ef])

        # 取整余数平衡 - 逐日 round 的累计误差集中调整到该线 Input 最大的一天
        if day_data:
            delta = round(sum(d[1] for d in day_data)) - sum(round(d[1]) for d in day_data)
            if delta:
                jmax = max(range(len(day_data)), key=lambda j: day_data[j][1])
                if day_data[jmax][1] + delta >= 0:
                    day_data[jmax][1] += delta

        for td, inp, wt, up, yl, ef in day_data:
            # 单班工时 = 全天工时/2, 再按整班次取整(<8按8, 8~9按9, 9~10按10, >10按10)
            wh_raw = wt / 2.0
            wh_h = norm_wh(wh_raw)
            if wh_h == wh_raw:
                wh_adj['不变'] += 1
            elif wh_raw < 8:
                wh_adj['<8->8'] += 1
            elif wh_raw < 9:
                wh_adj['8~9->9'] += 1
            elif wh_raw < 10:
                wh_adj['9~10->10'] += 1
            else:
                wh_adj['>10->10'] += 1

            # Input 取整, DS/NS 拆分确保总量一致
            inp_rounded = round(inp)
            ds_val = inp_rounded // 2          # DS取一半(向下取整)
            ns_val = inp_rounded - ds_val      # NS取剩余(奇数时NS多1)

            # V2.1: 目标UPH 优先取《目标UPHmapping》(机种+制程段+资源编码 严格匹配), 未命中沿用 MAX(UPH)
            tgt = round(mu, 2) if mu else None
            if upu is not None:
                k = (mdl, proc, res)
                if k in upu:
                    tgt = round(upu[k], 2)
                    upu_hit.add(k)
                elif k not in upu_miss_seen:
                    upu_miss_seen.add(k)
                    upu_miss.append((k[0], k[1], k[2], round(mu, 2) if mu else None))

            for s, inp_val in [('DS', ds_val), ('NS', ns_val)]:
                rows.append({
                    '主计划版本号': ver, '机种': mdl,
                    '制程段': proc, '物品类型': '',
                    '资源类型': proc, '尺寸': sz,
                    '资源编码': res, '白/夜班': s,
                    '日期': str(td), '工作时间(小时)': wh_h,
                    '目标UPH': tgt,                                # V2.1: mapping 命中取 mapping 值
                    '开线班数': 1,
                    'Input值': inp_val,
                    'UPH': round(mu, 2) if mu else None,   # V2.0: 取源文件该行 MAX(UPH)
                    '效率': round(ef, 6) if ef else None,
                    '良率': round(yl, 6) if yl else None,
                })

    for _, mr in sub_map.iterrows():
        handle(mr, 'sub')
    for _, mr in smt_map.iterrows():
        handle(mr, 'smt')

    cols = ['主计划版本号', '机种', '制程段', '物品类型', '资源类型', '尺寸',
            '资源编码', '白/夜班', '日期', '工作时间(小时)', '目标UPH',
            '开线班数', 'Input值', 'UPH', '效率', '良率']
    df = pd.DataFrame(rows, columns=cols)
    for c in ['工作时间(小时)', '效率', '良率']:
        df[c] = df[c].astype(float).round(6)

    output_dir = os.path.dirname(path_sub_map)
    output_path = os.path.join(output_dir, '主计划约束表-导入模板.csv')
    try:
        df.to_csv(output_path, index=False, encoding='utf-8-sig', float_format='%.4f')
    except PermissionError:
        _ts = datetime.now().strftime('%Y%m%d_%H%M%S')
        _alt = os.path.join(output_dir, f'主计划约束表-导入模板_{_ts}.csv')
        print(f"[WARN] 输出文件被占用 (Excel 未关闭或被其他程序锁定): {output_path}")
        print(f"       已改用安全文件名写入: {_alt}")
        df.to_csv(_alt, index=False, encoding='utf-8-sig', float_format='%.4f')
        output_path = _alt

    print()
    print("========== 生成报告 ==========")
    print(f"输出文件: {output_path}")
    print(f"总行数: {len(df)}")
    if len(df):
        _d = pd.to_datetime(df['日期'])
        print(f"日期范围: {_d.min().date()} ~ {_d.max().date()} (共 {_d.dt.normalize().nunique()} 天)")
    else:
        print("[!] 未生成任何数据行, 请检查源文件日期范围与系统日期是否匹配")
    print()
    for p in sorted(df['制程段'].unique()):
        s = df[df['制程段'] == p]
        print(f"  {p}: {len(s)}行  "
              f"Inp>0={(s['Input值'] > 0).sum()}  "
              f"UPH={s['UPH'].notna().sum()}  "
              f"Yld={s['良率'].notna().sum()}")
    print()
    print("--- 工作时间(小时)班次取整 (V1.5) ---")
    for k, n in wh_adj.items():
        print(f"  {k}: {n}条")
    print()
    if upu is not None:
        print("--- 目标UPHmapping (机种+制程段+资源编码 严格匹配) ---")
        print(f"  mapping条数: {len(upu)}  命中组合: {len(upu_hit)}  "
              f"未使用组合: {len(upu) - len(upu_hit)}  未命中组合: {len(upu_miss)}")
        if upu_miss:
            print("  未命中清单 (可复制到 目标UPHmapping.xlsx 补充):")
            print("    机种\t制程段\t资源编码\t当前目标UPH")
            for r in upu_miss[:50]:
                print(f"    {r[0]}\t{r[1]}\t{r[2]}\t{r[3]}")
            if len(upu_miss) > 50:
                print(f"    ... 另有 {len(upu_miss) - 50} 个组合未列出")
        print()
    print("--- 校验 ---")
    for k, its in iss.items():
        if its: print(f"  [!] {k}: {len(its)}条")
        else:   print(f"  [OK] {k}: 无问题")
    return output_path


# ============================================================
# 命令行入口
# ============================================================

USAGE = """用法:
  MPS-Capacity_V2.0.py prelaser <Capacity.xlsx> <PreLaserIOS.xlsx> <PPVT_IOS.xlsx> <Mapping.xlsx> [目标UPHmapping.xlsx]
  MPS-Capacity_V2.0.py fatp     <SMT-Mapping.xlsx> <Sub-FATP-Mapping.xlsx> <FATP_Capacity.xlsx> <SMT&Sub_Capacity.xlsx> [目标UPHmapping.xlsx]

  prelaser : Pre-laser / FWDL / Fresh 数据源 (原 MPS-Capacity_core_V1.1.py)
  fatp     : FATP / SUB / SMT 数据源    (原 MPS-Capacity-FATP-SUB-SMT_V1.5.py)

  第5个参数《目标UPHmapping.xlsx》可选 (Sheet1: 机种 | 制程段 | 资源编码 | 目标UPH):
    提供时按 机种+制程段+资源编码 严格一致匹配, 命中行的 目标UPH 取 mapping 值;
    未命中行沿用源文件 MAX(UPH), 并在生成报告的"目标UPHmapping"区列出未命中清单"""


def main(argv):
    if len(argv) < 2:
        print(USAGE)
        sys.exit(1)

    mode = argv[1].strip().lower()
    args = argv[2:]

    if len(args) == 5 and not os.path.exists(args[4]):
        print(f"目标UPHmapping 文件不存在: {args[4]}")
        sys.exit(1)

    if mode == 'prelaser':
        if len(args) not in (4, 5):
            print("prelaser 模式需要 4 个文件参数 (可选第5个: 目标UPHmapping.xlsx):")
            print(USAGE)
            print(f"传入参数数量: {len(args)}")
            sys.exit(1)
        run_prelaser(*args)
    elif mode == 'fatp':
        if len(args) not in (4, 5):
            print("fatp 模式需要 4 个文件参数 (可选第5个: 目标UPHmapping.xlsx):")
            print(USAGE)
            print(f"传入参数数量: {len(args)}")
            sys.exit(1)
        run_fatp_sub_smt(*args)
    else:
        print(f"未知模式: {mode}")
        print(USAGE)
        sys.exit(1)


if __name__ == '__main__':
    main(sys.argv)
