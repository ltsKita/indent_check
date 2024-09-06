import re
from lxml import etree as ET

# WordprocessingMLの名前空間を定義
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# レベル1〜4の正規表現パターン
level_patterns = {
    1: r"^\d+\.(?!\d)",               # レベル1: 例 "1. "
    2: r"^\d+\.\d+(?!\.\d)",          # レベル2: 例 "1.1 "
    3: r"^\d+\.\d+\.\d+(?!\.\d)",     # レベル3: 例 "1.1.1 "
    4: r"^\d+\.\d+\.\d+\.\d+",        # レベル4: 例 "1.1.1.1 "
    5: r"^\(\d+\)",                # レベル5: 例 "(1)", "(10)"
    6: r"^[ａ-ｚ]\.",              # レベル6: 例 "ａ．"
    7: r"^\([a-z]\)",              # レベル7: 例 "(a)"
    8: r"^\([a-z]-\d+\)",          # レベル8: 例 "(a-1)"
    9: r"^\([a-z]-\d+-\d+\)"       # レベル9: 例 "(a-1-1)"
}

# カッコを補完するパターン（レベル5,7,8,9）
bracket_patterns = {
    5: (r"^\(\d+\)", "(", ")"),              # レベル5: 例 "(1)"
    7: (r"^\([a-z]\)", "(", ")"),            # レベル7: 例 "(a)"
    8: (r"^\([a-z]-\d+\)", "(", ")"),        # レベル8: 例 "(a-1)"
    9: (r"^\([a-z]-\d+-\d+\)", "(", ")")     # レベル9: 例 "(a-1-1)"
}

def highlight_text(run):
    """指定された<w:r>要素にハイライトを追加する。"""
    rPr = run.find('.//w:rPr', namespaces=ns)
    if rPr is None:
        rPr = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    
    highlight = rPr.find('.//w:highlight', namespaces=ns)
    if highlight is not None:
        rPr.remove(highlight)

    new_highlight = ET.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
    new_highlight.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'yellow')

def process_numbering(paragraph, log):
    """レベル1〜4の項目番号の後に全角スペースを1つだけ挿入。既に全角スペースがある場合は何もしない。"""
    runs = paragraph.findall(".//w:r", namespaces=ns)
    
    for run in runs:
        w_t = run.find(".//w:t", namespaces=ns)
        if w_t is None or not w_t.text:
            continue

        original_text = w_t.text.strip()

        # レベル1〜4の処理
        for level, pattern in level_patterns.items():
            if level in [1, 2, 3, 4]:
                match = re.match(pattern, original_text)
                if match:
                    item_number = match.group(0)
                    after_item_number = original_text[len(item_number):]

                    # 全角スペースまたは半角スペースが2つ以上ある場合は削除して全角スペース1つにする
                    after_item_number = re.sub(r"^[ 　]+", "", after_item_number)
                    new_text = item_number + "　" + after_item_number
                    if new_text != w_t.text:
                        w_t.text = new_text
                        log_entry = f"レベル{level}: 元のテキスト: '{original_text}' -> 補完後のテキスト: '{new_text}'"
                        log.append(log_entry)
                        highlight_text(run)

def complete_brackets(paragraph, log):
    """レベル5,7,8,9のカッコを補完する処理。項目番号がある場合のみ片方のカッコを補完"""
    runs = paragraph.findall(".//w:r", namespaces=ns)
    
    for run in runs:
        w_t = run.find(".//w:t", namespaces=ns)
        if w_t is None or not w_t.text:
            continue

        original_text = w_t.text.strip()

        # カッコが片方欠けている項目番号の検出
        item_number_match = re.match(r"^(\()?([a-zA-Z0-9]+)(\))?", original_text)

        # 項目番号が見つからなかった場合はスキップ
        if not item_number_match:
            continue

        # 項目番号部分を抽出
        open_bracket = item_number_match.group(1)  # 開始カッコ
        number = item_number_match.group(2)        # 項目番号
        close_bracket = item_number_match.group(3) # 終了カッコ

        # 欠けているカッコを補完するロジック
        new_text = original_text
        if open_bracket is None and close_bracket is not None:
            # 開始カッコを補完する
            new_text = f"({number}{original_text[len(number):]}"
        elif open_bracket is not None and close_bracket is None:
            # 終了カッコを補完する
            new_text = f"{original_text[:len(number)+1]}){original_text[len(number)+1:]}"

        # 元のカッコを保持しつつ、欠けた方のみ補完する
        if new_text != original_text:
            w_t.text = new_text
            log_entry = f"カッコ補完: 元のテキスト: '{original_text}' -> 補完後のテキスト: '{new_text}'"
            log.append(log_entry)
            highlight_text(run)  # ハイライトを適用



        # スペース処理
        for level, (pattern, open_bracket, close_bracket) in bracket_patterns.items():
            if re.match(pattern, original_text):
                item_number = re.match(pattern, original_text).group(0)
                after_item_number = original_text[len(item_number):]

                # 項目番号の後ろに余分なスペースがある場合は削除して半角スペース1つだけにする
                if re.match(r"[ 　]{2,}", after_item_number) or re.match(r"^[　]+", after_item_number):
                    after_item_number = re.sub(r"[ 　]+", " ", after_item_number).strip()
                    new_text = f"{item_number} {after_item_number}"

                    if new_text != original_text:
                        w_t.text = new_text
                        log_entry = f"スペース修正: '{original_text}' -> '{new_text}'"
                        log.append(log_entry)
                        highlight_text(run)  # ハイライト適用

        # レベル6のスペース削除処理
        if re.match(level_patterns[6], original_text):
            # ピリオドの後ろのスペース削除
            new_text = re.sub(r"\.\s+[ 　]*", ".", original_text)
            if new_text != original_text:
                w_t.text = new_text
                log_entry = f"レベル6スペース削除: '{original_text}' -> '{new_text}'"
                log.append(log_entry)
                highlight_text(run)  # ハイライト適用

def remove_leading_spaces(paragraph, log):
    """<w:t>要素の先頭に全角・半角スペースがある場合、それを削除する処理"""
    runs = paragraph.findall(".//w:r", namespaces=ns)
    
    for run in runs:
        w_t = run.find(".//w:t", namespaces=ns)
        if w_t is None or not w_t.text:
            continue

        original_text = w_t.text
        new_text = original_text.lstrip(" 　")  # 先頭の全角・半角スペースを削除

        if new_text != original_text:
            w_t.text = new_text
            log_entry = f"先頭スペース削除: 元のテキスト: '{original_text}' -> 修正後のテキスト: '{new_text}'"
            log.append(log_entry)
            highlight_text(run)

def process_brackets_in_xml(xml_content):
    """XML文書を解析し、全体の補完処理を実行"""
    try:
        parser = ET.XMLParser(remove_blank_text=True)
        tree = ET.fromstring(xml_content.encode('utf-8'), parser)
    except ET.XMLSyntaxError as e:
        raise ValueError(f"XMLの解析中にエラーが発生しました: {e}")

    log = []

    for paragraph in tree.findall(".//w:p", namespaces=ns):
        complete_brackets(paragraph, log)  # レベル5,7,8,9のカッコ補完とスペース処理
        process_numbering(paragraph, log)  # レベル1〜4の処理（変更なし）

    return ET.tostring(tree, encoding='unicode', pretty_print=True), log

# XMLファイルの読み込みと処理
xml_file_path = "xml_new/word/document.xml"
with open(xml_file_path, "r", encoding="utf-8") as file:
    xml_content = file.read()

# XML解析と補完処理の実行
updated_xml, log = process_brackets_in_xml(xml_content)

# 修正されたXMLを保存
with open(xml_file_path, "w", encoding="utf-8") as file:
    file.write(updated_xml)

# ログの出力（変換された部分のみ）
log_file_path = "bracket_completion_log.txt"
with open(log_file_path, "w", encoding="utf-8") as file:
    for entry in log:
        file.write(entry + "\n")

print(f"処理が完了しました。ログは {log_file_path} に保存されました。")