import os
import pandas as pd
import re
import xlrd  # noqa: F401

def to_halfwidth(s: str) -> str:
    """將全形字元轉為半形"""
    return s.translate(
        {ord(f): ord(t) for f, t in zip(
            u'！＂＃＄％＆＇（）＊＋，－．／０１２３４５６７８９：；＜＝＞？＠'
            u'ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺ［＼］＾＿'
            u'｀ａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ｛｜｝～　',
            u'!"#$%&\'()*+,-./0123456789:;<=>?@'
            u'ABCDEFGHIJKLMNOPQRSTUVWXYZ[\\]^_'
            u'`abcdefghijklmnopqrstuvwxyz{|}~ '
        )}
    )


# 取得表頭資訊
def extract_header_info(df: pd.DataFrame) -> dict:
    header = {"對象": None, "日期": None, "單號": None}
    for _, row in df.iterrows():
        row_list = list(row)
        for i, cell in enumerate(row_list):
            cell_str = to_halfwidth(str(cell).strip())
            if not cell_str or cell_str.lower() == "nan":
                continue

            def search_right(start, max_offset=5):
                for offset in range(1, max_offset + 1):
                    if start + offset < len(row_list):
                        val = str(row_list[start + offset]).strip()
                        val = to_halfwidth(val).replace(" ", "")
                        if val and val.lower() != "nan":
                            return val
                return None

            # 日期判定 + 對象抓取（在同一 row 的左邊）
            if "date" in cell_str.lower() and not header["日期"]:
                header["日期"] = search_right(i)

                # 往左找第一個合理非空字串當作對象
                for j in range(i - 1, -1, -1):
                    left_val = to_halfwidth(str(row_list[j]).strip())
                    if left_val and left_val.lower() not in ["nan", "date", ":"]:
                        header["對象"] = left_val
                        break

            if "no" in cell_str.lower() and not header["單號"]:
                header["單號"] = search_right(i)

        if all(header.values()):
            break
    return header


def find_product_start_row(df: pd.DataFrame):
    for i, row in df.iterrows():
        cells = [str(c).strip() for c in row if pd.notna(c)]
        price_like = [c for c in cells if re.fullmatch(r"\d{4,6}", c)]
        unit = any(c.lower() in ["ea", "set", "pcs"] for c in cells)
        has_dash = any("-" in c for c in cells)
        if unit and len(price_like) >= 2 and has_dash:
            return i
    return 0

def parse_product_rows(df: pd.DataFrame):
    print("Parsing product rows...")
    products = []
    current = None

    start_row = find_product_start_row(df)
    print(f"自動偵測產品起始列：{start_row}")

    for i, row in df.iterrows():
        if i < start_row:
            continue

        cells = [str(c).strip() for c in row[1:] if pd.notna(c)]
        if not cells:
            continue

        price_like = [c for c in cells if re.fullmatch(r"\d{4,6}", c)]

        qty = None
        unit = None
        matched_unit_qty = None  # 用於完整字串移除

        for c in cells:
            if match := re.match(r"^(\d+)\s*(ea|set|pcs)$", c.lower()):
                qty = int(match.group(1))
                unit = match.group(2)
                matched_unit_qty = c
                break
            elif c.isdigit() and qty is None:
                qty = int(c)
            elif c.lower() in ["ea", "set", "pcs"] and unit is None:
                unit = c.lower()
        code = next((c for c in cells if re.fullmatch(r"[0-9A-Za-z\\-]+", c) and "-" in c), None)

        if code and unit and qty and len(price_like) >= 2:
            if current:
                products.append(current)
            current = {
                "貨號": code,
                "描述": "",
                "數量": int(qty),
                "單位": unit,
                "單價": int(price_like[0]),
                "金額": int(price_like[1]),
            }
            desc_parts = []
            for c in cells:
                if c == code or c.lower() == unit or c == matched_unit_qty:
                    continue
                if str(qty) == c:
                    continue
                if c in price_like:
                    continue
                desc_parts.append(c)

            current["描述"] = " ".join(desc_parts)
        elif current and any(re.search(r"[\u4e00-\u9fa5]", c) for c in cells):
            current["描述"] += "；" + " ".join(cells)

    if current:
        products.append(current)

    return products



# file_path = "data/三重0902201.xls"
# df = pd.read_excel(file_path, engine='xlrd', header=None)
# # print(type(df))
# results = parse_product_rows(df)
# #
# for item in results:
#     print(item)