#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""生成瓷砖行业景气分析项目的三个 Excel 工作簿（可重复运行，覆盖生成）。"""

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent

# ---------- 统一样式 ----------
HDR_FILL = PatternFill("solid", fgColor="1F4E79")
HDR_FONT = Font(bold=True, color="FFFFFF", size=11)
TITLE_FONT = Font(bold=True, size=14, color="1F4E79")
SUB_FONT = Font(bold=True, size=11, color="1F4E79")
NOTE_FONT = Font(size=9, color="808080", italic=True)
BAND_FILL = PatternFill("solid", fgColor="EDF2F7")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header(ws, row: int, ncols: int):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HDR_FILL
        cell.font = HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def body(ws, first_row: int, ncols: int, band: bool = True):
    for r in range(first_row, ws.max_row + 1):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if band and (r - first_row) % 2 == 1:
                cell.fill = BAND_FILL


# =====================================================================
# 1. indicator_tree.xlsx
# =====================================================================
INDICATORS = [
    ("需求", "地产竣工面积同比", "国家统计局", "月", "正", 0.15, "直接", "领先瓷砖需求约 1-2 个季度，核心变量"),
    ("需求", "商品房销售面积同比", "国家统计局", "月", "正", 0.05, "直接", "领先竣工，链条更长"),
    ("需求", "瓷砖产量同比", "统计局 / 协会", "月/季", "正", 0.15, "需核实", "核心指标，注意统计口径调整"),
    ("需求", "瓷砖出口金额同比", "海关总署", "月", "正", 0.10, "直接", "外需，占比小但边际敏感"),
    ("供给", "行业开工率", "协会 / 新闻", "季", "正", 0.05, "代理", "难直接获取，用产区新闻+公司表述代理"),
    ("供给", "库存天数", "公司财报 / 调研", "季", "负", 0.05, "代理", "负向指标：库存高企利空"),
    ("价格", "瓷砖价格指数", "协会 / 新闻", "月", "正", 0.10, "代理", "协会指数可能付费，先用公开摘要"),
    ("成本", "能源成本（天然气）", "统计局 / 发改委", "月", "负", 0.05, "直接", "负向指标：成本上升利空"),
    ("盈利", "上市公司毛利率", "巨潮资讯（东鹏/蒙娜丽莎/帝欧）", "季", "正", 0.10, "直接", "行业盈利能力的可观测代理"),
    ("信用", "应收账款周转天数", "巨潮资讯", "季", "负", 0.05, "直接", "负向指标：账期恶化是景气下行的早期信号"),
    ("政策", "地产政策强度", "手工打分", "月", "正", 0.10, "手工", "需自建 0-3 分打分规则，见口径说明"),
    ("政策", "环保/能耗政策", "手工打分", "季", "负", 0.05, "手工", "负向指标：停限产压制供给"),
]

MAPPING = [
    ("瓷砖产量同比", "磨边/抛光耗材需求量", "同步或略滞后", "高", "耗材按产量消耗，产量直接决定用量", "奔朗新材营收增速"),
    ("行业开工率", "耗材消耗强度", "同步", "中", "开工率高位=>设备满负荷=>耗材更换频率上升", "产区开工率、公司表述"),
    ("瓷砖企业资本开支", "压机模具订单", "领先 1-2 季", "中", "新增产线/技改才买模具，属资本开支驱动", "科达制造陶瓷机械收入"),
    ("瓷砖毛利率", "磨具降价压力 / 账期", "滞后 1-2 季", "中", "瓷砖打价格战=>向上游磨具压价+拉长账期", "磨具公司毛利率、应收周转"),
    ("瓷砖出口", "磨具出口 / 海外产能", "滞后", "低-中", "出口带动海外产能，间接带动磨具出口", "磨具出口金额"),
    ("大规格/岩板占比", "高端磨具单位消耗", "结构变量", "中", "大规格砖对磨边轮精度要求高，单耗与单价均升", "协会产品结构数据"),
]


def build_indicator_tree():
    wb = Workbook()

    # --- Sheet 1 指标树 ---
    ws = wb.active
    ws.title = "指标树"
    ws["A1"] = "瓷砖行业景气指标树"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "权重合计必须 = 100%；新增指标前先问：这个数据我未来 8 周能持续拿到吗？拿不到就标「代理」，不要硬凑。"
    ws["A2"].font = NOTE_FONT

    headers = ["类别", "指标", "来源", "频率", "方向", "权重", "可获取性", "备注"]
    ws.append([])
    ws.append(headers)
    hdr_row = ws.max_row
    style_header(ws, hdr_row, len(headers))

    for row in INDICATORS:
        cat, name, src, freq, direction, weight, avail, note = row
        ws.append([cat, name, src, freq, direction, weight, avail, note])

    body(ws, hdr_row + 1, len(headers))
    for r in range(hdr_row + 1, ws.max_row + 1):
        ws.cell(row=r, column=6).number_format = "0%"
        ws.cell(row=r, column=6).alignment = Alignment(horizontal="center")
        for c in (4, 5, 7):
            ws.cell(row=r, column=c).alignment = Alignment(horizontal="center")

    total = hdr_row + len(INDICATORS) + 1
    ws.cell(row=total, column=5, value="合计").font = Font(bold=True)
    ws.cell(row=total, column=5).alignment = Alignment(horizontal="right")
    ws.cell(row=total, column=6, value=f"=SUM(F{hdr_row + 1}:F{total - 1})")
    ws.cell(row=total, column=6).font = Font(bold=True, color="C00000")
    ws.cell(row=total, column=6).number_format = "0%"
    ws.cell(row=total, column=6).alignment = Alignment(horizontal="center")

    ws.cell(row=total + 2, column=1, value="方向说明：「正」= 指标上行代表景气改善；「负」= 指标上行代表景气恶化（打分时要取反）。").font = NOTE_FONT
    ws.cell(row=total + 3, column=1, value="可获取性说明：直接=公开数据可直接下载；代理=需用替代指标估算；手工=需按规则自行打分；需核实=口径可能变化，使用前先确认。").font = NOTE_FONT

    set_widths(ws, [8, 24, 30, 8, 8, 8, 10, 46])
    ws.freeze_panes = f"A{hdr_row + 1}"
    ws.auto_filter.ref = f"A{hdr_row}:H{total - 1}"

    # --- Sheet 2 磨具映射 ---
    ws2 = wb.create_sheet("磨具映射")
    ws2["A1"] = "瓷砖景气 → 磨具需求 映射表"
    ws2["A1"].font = TITLE_FONT
    ws2["A2"] = "磨具分两类，传导逻辑完全不同：耗材跟「产量」，模具跟「资本开支」。不要把两者合成一个分。"
    ws2["A2"].font = NOTE_FONT

    h2 = ["瓷砖指标", "磨具含义", "时滞", "弹性", "传导逻辑", "验证数据"]
    ws2.append([])
    ws2.append(h2)
    hdr2 = ws2.max_row
    style_header(ws2, hdr2, len(h2))
    for row in MAPPING:
        ws2.append(list(row))
    body(ws2, hdr2 + 1, len(h2))
    for r in range(hdr2 + 1, ws2.max_row + 1):
        for c in (3, 4):
            ws2.cell(row=r, column=c).alignment = Alignment(horizontal="center")

    ws2.cell(row=ws2.max_row + 2, column=1,
             value="⚠️ 非瓷砖下游提醒：石材、玻璃、日用陶瓷同样消耗磨具，归因时不能全部算到瓷砖头上。").font = Font(size=9, color="C00000")
    set_widths(ws2, [20, 26, 14, 10, 46, 26])
    ws2.freeze_panes = f"A{hdr2 + 1}"

    # --- Sheet 3 口径说明 ---
    ws3 = wb.create_sheet("口径说明")
    ws3["A1"] = "口径与规则说明（写清楚，比写好看重要）"
    ws3["A1"].font = TITLE_FONT
    lines = [
        ("一、行业范围", ""),
        ("", "行业 = 建筑陶瓷 / 瓷砖，不含卫生洁具。"),
        ("", "下游拆三块：工程（地产集采）、零售（经销）、出口。"),
        ("", "磨具映射只做两类：磨边/抛光耗材（跟产量）、压机模具（跟资本开支）。"),
        ("二、打分规则", ""),
        ("", "每个指标打 -2 ~ +2："),
        ("", "+2 显著改善：同比转正且环比上行，或处于历史低分位"),
        ("", "+1 温和改善"),
        ("", " 0 中性 / 无数据"),
        ("", "-1 温和恶化"),
        ("", "-2 显著恶化"),
        ("", "负向指标（库存天数、应收账款周转天数、能源成本、环保政策）恶化时打负分，"),
        ("", "最省事的办法：在「评分卡」的得分列直接填带符号的分数，不要靠公式取反。"),
        ("三、合成规则", ""),
        ("", "景气分 = Σ(指标得分 × 权重)，理论区间 -2 ~ +2"),
        ("", "> 0.5 景气上行 ； -0.5 ~ 0.5 中性震荡 ； < -0.5 景气下行"),
        ("", "磨具耗材代理分、压机模具代理分 单独算，不与瓷砖景气分混用。"),
        ("四、统计口径风险", ""),
        ("", "瓷砖产量统计口径历史上调整过，跨年同比前必须先确认口径一致，"),
        ("", "口径不一致就在「备注」写清楚，宁可标注「不可比」也不要强行算同比。"),
        ("五、数据纪律", ""),
        ("", "1. 所有数字必须能追到来源（来源 + 链接 + 日期三件套）。"),
        ("", "2. 拿不到就写「缺失」，不要估算、不要用研报数字顶替。"),
        ("", "3. 研报只当观点样本，不当数据。"),
        ("", "4. 不爬付费研报，不碰内幕信息。"),
    ]
    r = 3
    for head, text in lines:
        if head:
            ws3.cell(row=r, column=1, value=head).font = SUB_FONT
            r += 1
        if text:
            ws3.cell(row=r, column=1, value=text)
            r += 1
    set_widths(ws3, [100])

    # --- Sheet 4 使用说明 ---
    ws4 = wb.create_sheet("使用说明")
    ws4["A1"] = "三个工作簿怎么配合用"
    ws4["A1"].font = TITLE_FONT
    guide = [
        ("工作簿", "作用", "什么时候改"),
        ("indicator_tree.xlsx", "定义口径、指标、权重、磨具映射逻辑", "第 1 周定稿后基本不动；只在发现指标拿不到时替换"),
        ("scorecard.xlsx", "每月/每季把数据填进去，自动算景气分", "每月更新一次原始数据，判断方向后填「得分」"),
        ("forecast_log.xlsx", "记录每次判断与事后验证，防止自欺欺人", "每次给出判断时登记；下季度回来填「实际结果」"),
    ]
    for i, row in enumerate(guide, start=3):
        for j, v in enumerate(row, start=1):
            cell = ws4.cell(row=i, column=j, value=v)
            if i == 3:
                cell.fill = HDR_FILL
                cell.font = HDR_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws4.cell(row=8, column=1, value="月度节奏建议（每次约 2 小时）").font = SUB_FONT
    cadence = [
        "1) 打开 scorecard.xlsx，把本月新出的数据填进「原始数据」Sheet",
        "2) 在「评分卡」更新最新值 / 上期值，系统自动算同比",
        "3) 逐个指标填得分（-2~+2），写一句打分依据 —— 这句依据是后面复盘的唯一线索",
        "4) 看「汇总」Sheet 的景气分和磨具两个代理分",
        "5) 如果景气分跨越 ±0.5 边界，或连续两月同向变化，在 forecast_log 登记一次判断",
    ]
    for i, t in enumerate(cadence, start=9):
        ws4.cell(row=i, column=1, value=t)
    set_widths(ws4, [24, 52, 52])

    out = ROOT / "indicator_tree.xlsx"
    wb.save(out)
    return out


# =====================================================================
# 2. scorecard.xlsx
# =====================================================================
def build_scorecard():
    wb = Workbook()
    wss = wb.active
    wss.title = "汇总"

    # --- Sheet 2 原始数据 ---
    ws = wb.create_sheet("原始数据")
    ws["A1"] = "原始数据登记表（一行一个数据点，只增不改）"
    ws["A1"].font = TITLE_FONT
    h = ["日期", "指标", "值", "单位", "来源", "链接", "备注"]
    ws.append([])
    ws.append(h)
    hdr = ws.max_row
    style_header(ws, hdr, len(h))
    sample = [
        ("2026-09-30", "地产竣工面积同比", -12.5, "%", "国家统计局", "", "示例行，请替换为真实数据"),
        ("2026-09-30", "瓷砖产量同比", -8.0, "%", "示例来源", "", "示例行，请替换为真实数据"),
    ]
    for row in sample:
        ws.append(list(row))
    body(ws, hdr + 1, len(h))
    for r in range(hdr + 1, ws.max_row + 1):
        ws.cell(row=r, column=3).number_format = "0.00"
        if "示例" in str(ws.cell(row=r, column=5).value):
            for c in range(1, len(h) + 1):
                ws.cell(row=r, column=c).fill = WARN_FILL
    set_widths(ws, [12, 26, 12, 8, 26, 34, 40])
    ws.freeze_panes = f"A{hdr + 1}"
    ws.auto_filter.ref = f"A{hdr}:G{max(ws.max_row, hdr + 1)}"

    # --- Sheet 2 评分卡 ---
    wsc = wb.create_sheet("评分卡")
    wsc["A1"] = "景气评分卡（得分列手工填 -2 ~ +2，其余自动算）"
    wsc["A1"].font = TITLE_FONT
    wsc["A2"] = "负向指标已标注方向=负，请直接填带符号的分数（例如库存天数上升 => 填 -1）。"
    wsc["A2"].font = NOTE_FONT

    hc = ["类别", "指标", "方向", "权重", "最新值", "上期值", "同比/变化", "得分", "加权分", "打分依据（必填）", "数据日期"]
    wsc.append([])
    wsc.append(hc)
    hdrc = wsc.max_row
    style_header(wsc, hdrc, len(hc))

    start = hdrc + 1
    for cat, name, _src, _freq, direction, weight, _avail, _note in INDICATORS:
        wsc.append([cat, name, direction, weight, None, None, None, None, None, "", None])

    end = wsc.max_row
    for r in range(start, end + 1):
        wsc.cell(row=r, column=4).number_format = "0%"
        wsc.cell(row=r, column=7, value=f"=IFERROR((E{r}-F{r})/ABS(F{r}),\"\")")
        wsc.cell(row=r, column=7).number_format = "0.0%"
        wsc.cell(row=r, column=9, value=f"=IF(H{r}=\"\",\"\",D{r}*H{r})")
        wsc.cell(row=r, column=9).number_format = "0.00"
        wsc.cell(row=r, column=11).number_format = "yyyy-mm"
        for c in (1, 3, 4, 8):
            wsc.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
    body(wsc, start, len(hc), band=False)

    dv = DataValidation(type="list", formula1='"-2,-1,0,1,2"', allow_blank=True,
                        error="只能填 -2 / -1 / 0 / 1 / 2", errorTitle="得分超出范围")
    wsc.add_data_validation(dv)
    dv.add(f"H{start}:H{end}")

    tot = end + 1
    wsc.cell(row=tot, column=2, value="景气分（加权合计）").font = Font(bold=True)
    wsc.cell(row=tot, column=9, value=f"=SUM(I{start}:I{end})")
    wsc.cell(row=tot, column=9).font = Font(bold=True, size=12, color="C00000")
    wsc.cell(row=tot, column=9).number_format = "0.00"
    wsc.cell(row=tot, column=10, value="← 结果在「汇总」Sheet")
    wsc.cell(row=tot, column=10).font = NOTE_FONT

    wsc.cell(row=tot + 2, column=2,
             value="未填得分的指标按 0 计，会稀释景气分 —— 数据缺失时请在依据列写明「缺数据，按 0 处理」。").font = NOTE_FONT

    set_widths(wsc, [8, 26, 8, 8, 12, 12, 12, 8, 10, 44, 12])
    wsc.freeze_panes = f"A{hdrc + 1}"

    # --- Sheet 3 磨具代理 ---
    wsm = wb.create_sheet("磨具代理")
    wsm["A1"] = "磨具需求代理分（两个分分开算，不要混）"
    wsm["A1"].font = TITLE_FONT
    wsm["A2"] = "得分列自动引用「评分卡」对应指标；标记「手工」的行请自己填 -2 ~ +2。"
    wsm["A2"].font = NOTE_FONT

    hm = ["归属", "驱动因素", "映射权重", "得分", "加权分", "说明"]
    wsm.append([])
    wsm.append(hm)
    hdrm = wsm.max_row
    style_header(wsm, hdrm, len(hm))

    # (归属, 驱动, 权重, 得分公式或None, 说明)
    rows_m = [
        ("磨边/抛光耗材", "瓷砖产量同比", 0.40, f"=评分卡!H{start + 2}", "同步，权重最高"),
        ("磨边/抛光耗材", "行业开工率", 0.25, f"=评分卡!H{start + 4}", "消耗强度"),
        ("磨边/抛光耗材", "大规格/岩板占比", 0.20, None, "手工：结构升级打分，缺数据按 0"),
        ("磨边/抛光耗材", "瓷砖出口金额同比", 0.15, f"=评分卡!H{start + 3}", "间接"),
        ("压机模具", "瓷砖企业资本开支", 0.55, None, "手工：从财报/公告判断增减，缺数据按 0"),
        ("压机模具", "新增产线 / 技改", 0.30, None, "手工：从公告/新闻数项目"),
        ("压机模具", "上市公司毛利率", 0.15, f"=评分卡!H{start + 8}", "盈利支撑扩产意愿"),
    ]
    for belong, driver, w, formula, note in rows_m:
        wsm.append([belong, driver, w, formula, None, note])

    m_start, m_end = hdrm + 1, wsm.max_row
    for r in range(m_start, m_end + 1):
        wsm.cell(row=r, column=3).number_format = "0%"
        wsm.cell(row=r, column=5, value=f"=IF(D{r}=\"\",\"\",C{r}*D{r})")
        wsm.cell(row=r, column=5).number_format = "0.00"
        for c in (1, 3, 4):
            wsm.cell(row=r, column=c).alignment = Alignment(horizontal="center", vertical="center")
    body(wsm, m_start, len(hm), band=False)

    dv2 = DataValidation(type="list", formula1='"-2,-1,0,1,2"', allow_blank=True)
    wsm.add_data_validation(dv2)
    dv2.add(f"D{m_start}:D{m_end}")

    mc = wsm.max_row + 2
    wsm.cell(row=mc, column=2, value="磨边/抛光耗材 代理分").font = Font(bold=True)
    wsm.cell(row=mc, column=5, value=f"=SUMIF(A{m_start}:A{m_end},\"磨边/抛光耗材\",E{m_start}:E{m_end})")
    wsm.cell(row=mc, column=5).font = Font(bold=True, size=12, color="C00000")
    wsm.cell(row=mc, column=5).number_format = "0.00"
    wsm.cell(row=mc + 1, column=2, value="压机模具 代理分").font = Font(bold=True)
    wsm.cell(row=mc + 1, column=5, value=f"=SUMIF(A{m_start}:A{m_end},\"压机模具\",E{m_start}:E{m_end})")
    wsm.cell(row=mc + 1, column=5).font = Font(bold=True, size=12, color="C00000")
    wsm.cell(row=mc + 1, column=5).number_format = "0.00"

    set_widths(wsm, [18, 24, 10, 8, 10, 40])
    wsm.freeze_panes = f"A{hdrm + 1}"

    # --- Sheet 4 汇总（已作为首个 Sheet 创建） ---
    wss["A1"] = "景气与磨具需求 汇总看板"
    wss["A1"].font = TITLE_FONT
    wss["A2"] = "改这里没用，数值全部来自「评分卡」和「磨具代理」。"
    wss["A2"].font = NOTE_FONT

    blocks = [
        (4, "瓷砖行业景气分", f"=评分卡!I{tot}", "0.00"),
        (5, "景气分级", f"=IF(评分卡!I{tot}>0.5,\"景气上行\",IF(评分卡!I{tot}<-0.5,\"景气下行\",\"中性震荡\"))", "General"),
        (7, "磨边/抛光耗材 代理分", f"='磨具代理'!E{mc}", "0.00"),
        (8, "压机模具 代理分", f"='磨具代理'!E{mc + 1}", "0.00"),
    ]
    for r, label, formula, fmt in blocks:
        wss.cell(row=r, column=1, value=label).font = SUB_FONT
        cell = wss.cell(row=r, column=2, value=formula)
        cell.font = Font(bold=True, size=13, color="C00000")
        cell.number_format = fmt
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER

    wss.cell(row=10, column=1, value="判读规则").font = SUB_FONT
    rules = [
        "景气分 > 0.5 且耗材代理分 > 0：耗材需求向上，重点看产量能否持续",
        "景气分 < -0.5 且压机模具分 < 0：扩产停滞，模具端承压更早更久",
        "景气分与耗材分背离：多半是出口或结构在抵消国内产量下滑，去查明细",
        "两个磨具分背离：说明「产量」与「资本开支」周期错位，属正常，要分开表述",
    ]
    for i, t in enumerate(rules, start=11):
        wss.cell(row=i, column=1, value="· " + t)
    set_widths(wss, [30, 26])

    out = ROOT / "scorecard.xlsx"
    wb.save(out)
    return out


# =====================================================================
# 3. forecast_log.xlsx
# =====================================================================
def build_forecast_log():
    wb = Workbook()

    ws = wb.active
    ws.title = "预测日志"
    ws["A1"] = "预测日志（每给一次判断就登记，下季度回来结算）"
    ws["A1"].font = TITLE_FONT
    ws["A2"] = "没有日志的判断不算判断。对错不重要，「错因分类」才是资产。"
    ws["A2"].font = NOTE_FONT

    h = ["登记日期", "预测期间", "景气方向判断", "置信度", "核心依据（3 条以内）",
         "验证指标", "验证日期", "实际结果", "对错", "错因分类", "备注"]
    ws.append([])
    ws.append(h)
    hdr = ws.max_row
    style_header(ws, hdr, len(h))

    for _ in range(12):
        ws.append([None] * len(h))
    body(ws, hdr + 1, len(h), band=False)

    for r in range(hdr + 1, ws.max_row + 1):
        ws.cell(row=r, column=1).number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=7).number_format = "yyyy-mm-dd"
        ws.cell(row=r, column=4).number_format = "0%"
        ws.row_dimensions[r].height = 34

    dv_dir = DataValidation(type="list", formula1='"景气上行,中性震荡,景气下行"', allow_blank=True)
    dv_conf = DataValidation(type="list", formula1='"高,中,低"', allow_blank=True)
    dv_ok = DataValidation(type="list", formula1='"对,部分对,错,无法验证"', allow_blank=True)
    dv_reason = DataValidation(
        type="list",
        formula1='"数据问题（口径/缺失）,逻辑问题（传导链断裂）,时点问题（领先滞后判断错）,证据问题（关键信息漏读）,运气"',
        allow_blank=True,
    )
    for dv, col in ((dv_dir, "C"), (dv_conf, "D"), (dv_ok, "I"), (dv_reason, "J")):
        ws.add_data_validation(dv)
        dv.add(f"{col}{hdr + 1}:{col}{ws.max_row}")

    set_widths(ws, [12, 10, 14, 8, 40, 26, 12, 26, 8, 26, 26])
    ws.freeze_panes = f"A{hdr + 1}"

    # --- 复盘清单 ---
    ws2 = wb.create_sheet("复盘清单")
    ws2["A1"] = "每季度复盘必答的 5 个问题"
    ws2["A1"].font = TITLE_FONT
    qs = [
        ("错在数据？", "统计口径变了？数据缺失被当成 0？用了不可比的同比？"),
        ("错在逻辑？", "传导链在哪一环断了？产量真的决定磨具耗材吗？"),
        ("错在时点？", "领先/滞后判断错了几个季度？"),
        ("错在证据？", "LLM 抽取有没有漏掉关键的反面证据？"),
        ("错在仓位/权重？", "是不是某个高权重指标本身就不该给这么高权重？"),
    ]
    ws2.append([])
    ws2.append(["复盘问题", "追问细节", "本季结论"])
    hdr2 = ws2.max_row
    style_header(ws2, hdr2, 3)
    for q, detail in qs:
        ws2.append([q, detail, ""])
    body(ws2, hdr2 + 1, 3)
    for r in range(hdr2 + 1, ws2.max_row + 1):
        ws2.row_dimensions[r].height = 46
        ws2.cell(row=r, column=3).alignment = Alignment(vertical="top", wrap_text=True)
    set_widths(ws2, [18, 46, 50])

    ws2.cell(row=ws2.max_row + 2, column=1,
             value="⚠️ 样本只有几个季度，禁止为了「历史全对」反复调权重。重点看逻辑是否成立、是否可证伪。").font = Font(size=9, color="C00000")

    out = ROOT / "forecast_log.xlsx"
    wb.save(out)
    return out


if __name__ == "__main__":
    ROOT.mkdir(parents=True, exist_ok=True)
    for f in (build_indicator_tree(), build_scorecard(), build_forecast_log()):
        print("已生成:", f)
