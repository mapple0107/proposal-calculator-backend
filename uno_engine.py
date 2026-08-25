"""
LibreOffice UNO 重算引擎
負責：開啟商品範本 xlsx -> 寫入客戶輸入 -> calculateAll() -> 讀出結果表 -> (可選)匯出 PDF
"""
import os
import shutil
import tempfile
import time
import uno
from com.sun.star.beans import PropertyValue

UNO_HOST = os.environ.get("UNO_HOST", "localhost")
UNO_PORT = os.environ.get("UNO_PORT", "2002")

TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), "templates")

# 各商品代碼對應的範本檔案（PFA3/PFA6 共用同一份範本，用繳費年期欄位區分）
PRODUCT_TEMPLATE = {
    "PFA3": "PFA.xlsx",
    "PFA6": "PFA.xlsx",
    "CLZ": "CLZ.xlsx",
    "CLX": "CLX.xlsx",
    "LTS": "LTS.xlsx",
}

# 商品家族：決定要用哪一套計算邏輯（輸入欄位對照、輸出表格結構都不同）
PRODUCT_FAMILY = {
    "PFA3": "PFA",
    "PFA6": "PFA",
    "CLZ": "CANCER",
    "CLX": "CANCER",
    "LTS": "LTS",
}

# 注意：三個情境表格的代號雖是 H/M/L，但不是「紅利由高到低」的意思，
# 而是台灣保險局要求分紅保單揭露的三種法定情境（依 OP!B173:B175 定義）：
#   H (總表_分紅_H) = 假設分紅金額可能為零（最保守揭露情境，早年通常為0）
#   M (總表_分紅_M) = 最可能紅利（業務員一般引用的主情境）
#   L (總表_分紅_L) = 較低紅利
# 三張表一律會同時算出，跟 輸入頁!E23（計算預估紅利）的選擇無關；
# E23 只決定「列印頁」PDF 上要標示/主打哪一個情境。
RESULT_SHEETS = {
    "zero_possible": {"sheet": "總表_分紅_H", "label": "假設分紅金額可能為零"},
    "most_likely": {"sheet": "總表_分紅_M", "label": "最可能紅利"},
    "lower": {"sheet": "總表_分紅_L", "label": "較低紅利"},
}

# 總表_分紅_x 的欄位對照（依實際 xlsx 表頭確認）
COLUMN_MAP = [
    ("policy_year", "B"),      # 保單年度
    ("age", "C"),               # 保險年齡
    ("annual_premium", "D"),    # 年度實繳保費(已扣除年度保單紅利)
    ("cum_premium", "E"),       # 累計實繳保費
    ("survival_benefit", "F"),  # 生存保險金
    ("cum_survival_benefit", "G"),  # 累計已領生存保險金
    ("death_benefit", "H"),     # 年度末 身故/完全失能保障
    ("cash_value", "I"),        # 年度末 解約金
    ("reduced_paid_up", "J"),   # 年度末 減額繳清保險金額
    ("annual_dividend", "K"),   # 年度末 年度保單紅利
    ("paid_up_amount", "M"),    # 年度末 繳清保險金額
    ("total_amount", "N"),      # 年度末 保險金額
]

DATA_ROW_START = 5
DATA_ROW_END = 115  # 對應 named range TOV_BONU_M!B5:AG115


def _mkprop(name, value):
    p = PropertyValue()
    p.Name = name
    p.Value = value
    return p


def _connect(retries=10, delay=1.0):
    local_ctx = uno.getComponentContext()
    resolver = local_ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.bridge.UnoUrlResolver", local_ctx)
    last_err = None
    for _ in range(retries):
        try:
            ctx = resolver.resolve(
                f"uno:socket,host={UNO_HOST},port={UNO_PORT};urp;StarOffice.ComponentContext")
            smgr = ctx.ServiceManager
            desktop = smgr.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
            return desktop
        except Exception as e:
            last_err = e
            time.sleep(delay)
    raise RuntimeError(f"無法連線 LibreOffice UNO ({UNO_HOST}:{UNO_PORT}): {last_err}")


def _roc_birth_int(birth_year, birth_month, birth_day):
    """把民國生日組成輸入頁!E8 需要的整數格式，例如 110年01月01日 -> 1100101"""
    return int(f"{birth_year}{birth_month:02d}{birth_day:02d}")


def _read_error(cell):
    try:
        return cell.getError()
    except Exception:
        return 0


class CalcError(Exception):
    pass


def calculate(product_code: str, inputs: dict, want_pdf: bool = False):
    """依商品家族分派到對應的計算函式。"""
    family = PRODUCT_FAMILY.get(product_code.upper())
    if family == "CANCER":
        return calculate_cancer(product_code, inputs, want_pdf=want_pdf)
    if family == "LTS":
        return calculate_lts(product_code, inputs, want_pdf=want_pdf)
    return calculate_pfa(product_code, inputs, want_pdf=want_pdf)


def calculate_pfa(product_code: str, inputs: dict, want_pdf: bool = False):
    """
    inputs 需包含：
      name (str)
      gender ("男"/"女")
      birth_year / birth_month / birth_day  (民國年/月/日)
      payment_term (3 或 6)
      payment_freq ("年繳"/"半年繳"/"季繳"/"月繳")
      dividend_scenario ("高分紅"/"中分紅"/"低分紅")
      input_mode ("face_amount" 或 "premium")
      face_amount_wan (input_mode=face_amount 時必填，單位：萬元)
      premium_amount (input_mode=premium 時必填，單位：元)
      death_benefit_pct (0~100，預設0)
      installment_period (10 或 20，預設20)
      relationship (預設 "同被保險人")
      discount (預設 "無")
    """
    template_name = PRODUCT_TEMPLATE.get(product_code)
    if not template_name:
        raise CalcError(f"未支援的商品代碼: {product_code}")

    src_path = os.path.join(TEMPLATE_DIR, template_name)
    if not os.path.exists(src_path):
        raise CalcError(f"找不到範本檔案: {template_name}")

    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir="/tmp")
    os.close(fd)
    shutil.copyfile(src_path, tmp_path)

    desktop = _connect()
    url = "file://" + tmp_path
    props = [_mkprop("Hidden", True)]
    doc = desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))

    try:
        sheets = doc.getSheets()
        ip = sheets.getByName("輸入頁")
        op = sheets.getByName("OP")

        ip.getCellRangeByName("E5").setString(inputs.get("name", "客戶"))
        ip.getCellRangeByName("K5").setString(inputs["gender"])
        ip.getCellRangeByName("E8").setValue(
            _roc_birth_int(int(inputs["birth_year"]), int(inputs["birth_month"]), int(inputs["birth_day"]))
        )
        ip.getCellRangeByName("E37").setValue(int(inputs["payment_term"]))
        ip.getCellRangeByName("E17").setString(inputs.get("payment_freq", "年繳"))
        ip.getCellRangeByName("E23").setString(inputs.get("dividend_scenario", "中分紅"))
        ip.getCellRangeByName("E12").setString(inputs.get("relationship", "同被保險人"))
        ip.getCellRangeByName("E21").setString(inputs.get("discount", "無"))

        death_pct = float(inputs.get("death_benefit_pct", 0) or 0) / 100.0
        ip.getCellRangeByName("G28").setValue(death_pct)
        if death_pct > 0:
            ip.getCellRangeByName("G32").setValue(int(inputs.get("installment_period", 20)))

        input_mode = inputs.get("input_mode", "face_amount")
        if input_mode == "premium":
            # 第一輪：用保費試算換算保額（OP!C62），再用換算後的保額做正式試算
            premium = float(inputs["premium_amount"])
            ip.getCellRangeByName("E55").setValue(premium)
            # 先給一個暫定保額讓公式鏈不報錯，再重算讀出換算值
            ip.getCellRangeByName("J39").setValue(30)
            doc.calculateAll()
            converted = op.getCellRangeByName("C62").getValue()
            if not converted or converted <= 0:
                raise CalcError("保費換算保額失敗，請確認輸入的保費金額")
            ip.getCellRangeByName("J39").setValue(converted)
        else:
            face_amount = float(inputs["face_amount_wan"])
            ip.getCellRangeByName("J39").setValue(face_amount)

        doc.calculateAll()

        # 檢核錯誤
        err_cell = ip.getCellRangeByName("E9")
        check_text = err_cell.getString()

        premiums = {
            "first_period_premium": op.getCellRangeByName("C48").getValue(),
            "renewal_period_premium": op.getCellRangeByName("C49").getValue(),
            "first_year_premium": op.getCellRangeByName("C51").getValue(),
            "renewal_year_premium": op.getCellRangeByName("C52").getValue(),
            "final_face_amount_wan": op.getCellRangeByName("C28").getValue(),
        }

        tables = {}
        for key, meta in RESULT_SHEETS.items():
            sheet_name = meta["sheet"]
            sheet = sheets.getByName(sheet_name)
            rows = []
            for r in range(DATA_ROW_START, DATA_ROW_END + 1):
                b_val = sheet.getCellRangeByName(f"B{r}").getString()
                if b_val == "":
                    break
                row = {}
                for field, col in COLUMN_MAP:
                    c = sheet.getCellRangeByName(f"{col}{r}")
                    row[field] = c.getValue() if c.getString() != "" else None
                rows.append(row)
            tables[key] = {"label": meta["label"], "rows": rows}

        pdf_path = None
        if want_pdf:
            fd2, pdf_path = tempfile.mkstemp(suffix=".pdf", dir="/tmp")
            os.close(fd2)
            controller = doc.getCurrentController()
            print_sheet = sheets.getByName("列印頁")
            controller.setActiveSheet(print_sheet)
            # 跳過第1頁（封面/基本資料摘要頁），PDF 直接從分紅保單利益彙總表開始
            filter_data = uno.Any(
                "[]com.sun.star.beans.PropertyValue",
                (_mkprop("PageRange", "2-"),),
            )
            export_props = [
                _mkprop("FilterName", "calc_pdf_Export"),
                _mkprop("SelectionOnly", False),
                _mkprop("FilterData", filter_data),
            ]
            doc.storeToURL("file://" + pdf_path, tuple(export_props))

        return {
            "check_message": check_text,
            "premiums": premiums,
            "tables": tables,
        }, pdf_path
    finally:
        doc.close(False)
        try:
            os.remove(tmp_path)
        except OSError:
            pass


# ============================================================
# CLZ / CLX 防癌險（無分紅，單純定期還本型健康險）
# ============================================================

# 列印頁的保單年度明細表分成左右兩個並排區塊（省版面用的排版方式）：
# 左區塊 B:I 從第1年開始，滿58列後接續右區塊 L:S 到保障期滿為止。
# 兩區塊要依序串接，才是完整、連續的保單年度明細。
CANCER_COLUMN_MAP_LEFT = [
    ("policy_year", "B"),
    ("age", "C"),
    ("annual_premium", "D"),
    ("cum_premium", "F"),
    ("death_benefit", "I"),
]
CANCER_COLUMN_MAP_RIGHT = [
    ("policy_year", "L"),
    ("age", "M"),
    ("annual_premium", "N"),
    ("cum_premium", "P"),
    ("death_benefit", "S"),
]
CANCER_TABLE_ROW_START = 49


def calculate_cancer(product_code: str, inputs: dict, want_pdf: bool = False):
    """
    inputs 需包含：
      name (str)
      birth_year / birth_month / birth_day（民國年/月/日）
      relationship（預設 "同被保險人"，選項：同被保險人/與被保險人不同）
      payment_freq（預設 "年繳"，選項：年繳/半年繳/季繳/月繳）
      first_payment_method（預設 "金融機構轉帳"，選項：匯款/金融機構轉帳/一般信用卡/富邦信用卡）
      renewal_payment_method（預設 "金融機構轉帳"，選項：金融機構轉帳/一般信用卡/富邦信用卡/自行繳費）
      discount（預設 "無"，選項：無/員工轉帳件）
      payment_term（10 或 20）
      face_amount_wan（保額，單位：萬元）

    注意：性別（K5）由範本本身固定（CLZ=男 / CLX=女），不開放輸入。
    """
    template_name = PRODUCT_TEMPLATE.get(product_code)
    if not template_name:
        raise CalcError(f"未支援的商品代碼: {product_code}")

    src_path = os.path.join(TEMPLATE_DIR, template_name)
    if not os.path.exists(src_path):
        raise CalcError(f"找不到範本檔案: {template_name}")

    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir="/tmp")
    os.close(fd)
    shutil.copyfile(src_path, tmp_path)

    desktop = _connect()
    url = "file://" + tmp_path
    props = [_mkprop("Hidden", True)]
    doc = desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))

    try:
        sheets = doc.getSheets()
        ip = sheets.getByName("輸入頁")
        op = sheets.getByName("OP")

        ip.getCellRangeByName("E5").setString(inputs.get("name", "客戶"))
        ip.getCellRangeByName("E7").setValue(
            _roc_birth_int(int(inputs["birth_year"]), int(inputs["birth_month"]), int(inputs["birth_day"]))
        )
        ip.getCellRangeByName("E12").setString(inputs.get("relationship", "同被保險人"))
        ip.getCellRangeByName("E16").setString(inputs.get("payment_freq", "年繳"))
        ip.getCellRangeByName("Q16").setString(inputs.get("first_payment_method", "金融機構轉帳"))
        ip.getCellRangeByName("Q18").setString(inputs.get("renewal_payment_method", "金融機構轉帳"))
        ip.getCellRangeByName("E19").setString(inputs.get("discount", "無"))
        ip.getCellRangeByName("D22").setValue(int(inputs["payment_term"]))
        ip.getCellRangeByName("J24").setValue(float(inputs["face_amount_wan"]))

        doc.calculateAll()

        premiums = {
            "first_period_premium": op.getCellRangeByName("C48").getValue(),
            "renewal_period_premium": op.getCellRangeByName("C49").getValue(),
            "first_year_premium": op.getCellRangeByName("C51").getValue(),
            "renewal_year_premium": op.getCellRangeByName("C52").getValue(),
            "final_face_amount_wan": op.getCellRangeByName("C28").getValue(),
            # 折扣前的應繳保費（依所選繳別換算，年繳時即為年繳保費），對應官方 PDF 上方
            # 「＿繳應繳保費/元」欄位，與已扣除折扣的 first_period_premium 不同。
            "annual_premium_gross": op.getCellRangeByName("C35").getValue(),
        }

        print_sheet = sheets.getByName("列印頁")

        def is_data_cell(cell):
            # 資料列的儲存格可能是純數字（VALUE）或公式算出數字（FORMULA），
            # 表格結束後接的是文字註腳（TEXT，非公式）或真正空白（EMPTY）、
            # 或是公式算出空字串（FORMULA 但字串為空）——這幾種都視為「結束」。
            return cell.getType().value != "TEXT" and cell.getString().strip() != ""

        def read_block(column_map, first_col):
            block_rows = []
            r = CANCER_TABLE_ROW_START
            while True:
                first_cell = print_sheet.getCellRangeByName(f"{first_col}{r}")
                if not is_data_cell(first_cell):
                    break
                row = {}
                for field, col in column_map:
                    c = print_sheet.getCellRangeByName(f"{col}{r}")
                    row[field] = c.getValue() if is_data_cell(c) else None
                block_rows.append(row)
                r += 1
                if r > CANCER_TABLE_ROW_START + 100:  # 安全上限，避免異常時無窮迴圈
                    break
            return block_rows

        rows = read_block(CANCER_COLUMN_MAP_LEFT, "B") + read_block(CANCER_COLUMN_MAP_RIGHT, "L")

        pdf_path = None
        if want_pdf:
            fd2, pdf_path = tempfile.mkstemp(suffix=".pdf", dir="/tmp")
            os.close(fd2)
            controller = doc.getCurrentController()
            controller.setActiveSheet(print_sheet)
            # 原始模板在 100% 縮放下欄位寬度超出單頁可印範圍，導致最右側欄位（如「說明」）
            # 文字被裁切、內容卡到旁邊的頁面。改為「縮放至頁寬 1 頁」，高度不限頁數，
            # 讓整張表自動等比縮小以完整印在一頁寬度內。
            style_name = print_sheet.PageStyle
            page_style = doc.getStyleFamilies().getByName("PageStyles").getByName(style_name)
            page_style.ScaleToPagesX = 1
            page_style.ScaleToPagesY = 0
            # 跳過第1頁（封面/基本資料摘要頁），PDF 直接從投保利益表開始
            filter_data = uno.Any(
                "[]com.sun.star.beans.PropertyValue",
                (_mkprop("PageRange", "2-"),),
            )
            export_props = [
                _mkprop("FilterName", "calc_pdf_Export"),
                _mkprop("SelectionOnly", False),
                _mkprop("FilterData", filter_data),
            ]
            doc.storeToURL("file://" + pdf_path, tuple(export_props))

        return {
            "premiums": premiums,
            "tables": {"main": {"label": "保單年度明細", "rows": rows}},
        }, pdf_path
    finally:
        doc.close(False)
        try:
            os.remove(tmp_path)
        except OSError:
            pass


# ============================================================
# LTS（詠馨久久長期照顧終身壽險）— 單一被保險人、終身（保險年齡屆滿99歲）、
# 無分紅，年度表含解約金／減額繳清欄位。
# ============================================================
LTS_COLUMN_MAP = [
    ("policy_year", "B"),
    ("age", "C"),
    ("annual_premium", "D"),
    ("cum_premium", "G"),
    ("death_benefit", "K"),
    ("cash_value", "O"),          # 年度末解約金
    ("reduced_paid_up", "S"),     # 年度末減額繳清保險金額
]
LTS_TABLE_ROW_START = 56
LTS_TERMINAL_AGE = 99  # 保障至保險年齡屆滿99歲（終身）


def calculate_lts(product_code: str, inputs: dict, want_pdf: bool = False):
    """
    inputs 需包含：
      name (str)
      gender（"男"/"女"）
      birth_year / birth_month / birth_day（民國年/月/日）
      relationship（預設 "同被保險人"，選項：同被保險人/與被保險人不同）
      payment_freq（預設 "年繳"，選項：年繳/半年繳/季繳/月繳）
      first_payment_method / renewal_payment_method（預設 "金融機構轉帳"）
      discount（預設 "無"）
      payment_term（10 / 20 / 30）
      face_amount_wan（保額，單位：萬元，範圍依投保年齡約 0.5~5 萬）

    列印頁的年度明細表為單一連續表格（非左右兩欄），但每印表頁會重複表頭列
    （文字列，非資料列），所以用「已讀到的有效資料列數 == 99 - 保險年齡 + 1」
    作為停止條件，而不是「遇到非資料列就停止」，否則會被表頭重複列誤判為表尾。
    """
    template_name = PRODUCT_TEMPLATE.get(product_code)
    if not template_name:
        raise CalcError(f"未支援的商品代碼: {product_code}")

    src_path = os.path.join(TEMPLATE_DIR, template_name)
    if not os.path.exists(src_path):
        raise CalcError(f"找不到範本檔案: {template_name}")

    fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", dir="/tmp")
    os.close(fd)
    shutil.copyfile(src_path, tmp_path)

    desktop = _connect()
    url = "file://" + tmp_path
    props = [_mkprop("Hidden", True)]
    doc = desktop.loadComponentFromURL(url, "_blank", 0, tuple(props))

    try:
        sheets = doc.getSheets()
        ip = sheets.getByName("輸入頁")
        op = sheets.getByName("OP")

        ip.getCellRangeByName("E5").setString(inputs.get("name", "客戶"))
        ip.getCellRangeByName("K5").setString(inputs.get("gender", "女"))
        ip.getCellRangeByName("E7").setValue(
            _roc_birth_int(int(inputs["birth_year"]), int(inputs["birth_month"]), int(inputs["birth_day"]))
        )
        ip.getCellRangeByName("E12").setString(inputs.get("relationship", "同被保險人"))
        ip.getCellRangeByName("E16").setString(inputs.get("payment_freq", "年繳"))
        ip.getCellRangeByName("Q16").setString(inputs.get("first_payment_method", "金融機構轉帳"))
        ip.getCellRangeByName("Q18").setString(inputs.get("renewal_payment_method", "金融機構轉帳"))
        ip.getCellRangeByName("E19").setString(inputs.get("discount", "無"))
        ip.getCellRangeByName("D22").setValue(int(inputs["payment_term"]))
        ip.getCellRangeByName("J24").setValue(float(inputs["face_amount_wan"]))

        doc.calculateAll()

        insurance_age = op.getCellRangeByName("D5").getValue()

        premiums = {
            "first_period_premium": op.getCellRangeByName("C48").getValue(),
            "renewal_period_premium": op.getCellRangeByName("C49").getValue(),
            "first_year_premium": op.getCellRangeByName("C51").getValue(),
            "renewal_year_premium": op.getCellRangeByName("C52").getValue(),
            "final_face_amount_wan": op.getCellRangeByName("C28").getValue(),
            "annual_premium_gross": op.getCellRangeByName("C35").getValue(),
            "insurance_age": insurance_age,
        }

        print_sheet = sheets.getByName("列印頁")

        def is_data_cell(cell):
            return cell.getType().value != "TEXT" and cell.getString().strip() != ""

        rows = []
        if insurance_age and insurance_age > 0:
            expected_years = int(LTS_TERMINAL_AGE - insurance_age + 1)
            r = LTS_TABLE_ROW_START
            scanned = 0
            # 安全上限：正常情況不會超過約 (99-最低承保年齡+1) 列，多留一倍緩衝
            while len(rows) < expected_years and scanned < 400:
                first_cell = print_sheet.getCellRangeByName(f"B{r}")
                if is_data_cell(first_cell):
                    row = {}
                    for field, col in LTS_COLUMN_MAP:
                        c = print_sheet.getCellRangeByName(f"{col}{r}")
                        row[field] = c.getValue() if is_data_cell(c) else None
                    rows.append(row)
                r += 1
                scanned += 1

        pdf_path = None
        if want_pdf:
            fd2, pdf_path = tempfile.mkstemp(suffix=".pdf", dir="/tmp")
            os.close(fd2)
            controller = doc.getCurrentController()
            controller.setActiveSheet(print_sheet)
            # 同 CLZ/CLX：原始模板欄寬在 100% 縮放下會超出單頁可印範圍，改為縮放至頁寬 1 頁
            style_name = print_sheet.PageStyle
            page_style = doc.getStyleFamilies().getByName("PageStyles").getByName(style_name)
            page_style.ScaleToPagesX = 1
            page_style.ScaleToPagesY = 0
            # 跳過第1頁（封面/基本資料摘要頁），PDF 直接從投保利益表開始
            filter_data = uno.Any(
                "[]com.sun.star.beans.PropertyValue",
                (_mkprop("PageRange", "2-"),),
            )
            export_props = [
                _mkprop("FilterName", "calc_pdf_Export"),
                _mkprop("SelectionOnly", False),
                _mkprop("FilterData", filter_data),
            ]
            doc.storeToURL("file://" + pdf_path, tuple(export_props))

        return {
            "premiums": premiums,
            "tables": {"main": {"label": "保單年度明細", "rows": rows}},
        }, pdf_path
    finally:
        doc.close(False)
        try:
            os.remove(tmp_path)
        except OSError:
            pass
