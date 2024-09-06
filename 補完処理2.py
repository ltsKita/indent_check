import re
from lxml import etree as ET

# WordprocessingMLの名前空間を定義
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# 補完するパターン
bracket_patterns = {
    1: (r"^\d+\.", "(", ")"),       # レベル1: 例 "1."
    2: (r"^\d+\.\d+", "(", ")"),    # レベル2: 例 "1.1"
    3: (r"^\d+\.\d+\.\d+", "(", ")"),  # レベル3: 例 "1.1.1"
    4: (r"^\d+\.\d+\.\d+\.\d+", "(", ")"),  # レベル4: 例 "1.1.1.1"
    5: (r"^\d+", "(", ")"),         # レベル5: 例 "(1)"
    7: (r"^[a-z]", "(", ")"),       # レベル7: 例 "(a)"
    8: (r"^[a-z]-\d+", "(", ")"),   # レベル8: 例 "(a-1)"
    9: (r"^[a-z]-\d+-\d+", "(", ")")# レベル9: 例 "(a-1-1)"
}

# 正規表現パターン（補完後のテキストが引っかかるか確認する用）
validation_patterns = {
    1: r"^\d+\.\s*.*$",                # レベル1: 例 "1. "
    2: r"^\d+\.\d+\s*.*$",             # レベル2: 例 "1.1 "
    3: r"^\d+\.\d+\.\d+\s*.*$",        # レベル3: 例 "1.1.1 "
    4: r"^\d+\.\d+\.\d+\.\d+\s*.*$",   # レベル4: 例 "1.1.1.1 "
    5: r"^\(\d+\)\s*.*$",              # レベル5: 例 "(1)"
    7: r"^\([a-z]\)\s*.*$",            # レベル7: 例 "(a)"
    8: r"^\([a-z]-\d+\)\s*.*$",        # レベル8: 例 "(a-1)"
    9: r"^\([a-z]-\d+-\d+\)\s*.*$"     # レベル9: 例 "(a-1-1)"
}

def highlight_text(run):
    """
    指定された<w:r>要素にハイライトを追加する。
    """
    rPr = run.find('.//w:rPr', namespaces=ns)
    
    if rPr is None:
        # rPrが存在しない場合は新規作成
        rPr = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    
    # 既存のハイライトを削除
    highlight = rPr.find('.//w:highlight', namespaces=ns)
    if highlight is not None:
        rPr.remove(highlight)
    
    # ハイライトの追加
    new_highlight = ET.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
    new_highlight.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'yellow')

def remove_leading_full_width_space(run):
    """
    <w:t>の先頭が全角スペースで始まる場合、その全角スペースを全て除去する。
    """
    w_t = run.find(".//w:t", namespaces=ns)
    if w_t is not None and w_t.text:
        # 全角スペースを除去
        new_text = w_t.text.lstrip("　")
        if new_text != w_t.text:
            w_t.text = new_text

def delete_preserved_space(paragraph):
    """
    項目番号の次に<w:t xml:space="preserve"> </w:t>がある場合、その要素を削除する。
    """
    runs = paragraph.findall(".//w:r", namespaces=ns)
    for i, run in enumerate(runs[:-1]):  # 最後の要素は確認しない
        next_run = runs[i + 1]
        w_t = run.find(".//w:t", namespaces=ns)
        next_w_t = next_run.find(".//w:t", namespaces=ns)

        if w_t is not None and next_w_t is not None:
            # 項目番号の次の<w:t>がスペースのみである場合、削除
            if next_w_t.get("{http://www.w3.org/XML/1998/namespace}space") == "preserve" and next_w_t.text == " ":
                paragraph.remove(next_run)

def complete_brackets(paragraph, log):
    """
    段落内の<w:t>のテキストを解析し、カッコや全角スペースが不足している場合に補完する。
    必要な場合のみ変更を行い、ログに記録する。また、変更箇所にハイライトを追加する。
    """
    runs = paragraph.findall(".//w:r", namespaces=ns)

    for i, run in enumerate(runs):
        w_t = run.find(".//w:t", namespaces=ns)
        if w_t is None:
            continue
        
        original_text = w_t.text
        if original_text is None:
            continue

        text = original_text.strip()

        # レベル1〜4の項目番号に全角スペースを補完
        modified = False  # 変更があったかどうかを確認
        for level, (pattern, open_bracket, close_bracket) in bracket_patterns.items():
            if level in [1, 2, 3, 4]:  # レベル1〜4に限定
                if re.match(pattern, text):
                    log_entry = f"レベル{level}: 元のテキスト: '{original_text}'"

                    # 条件: 次の<w:t>がない場合、全角スペースを追加しない
                    if i + 1 >= len(runs):  # 次の<w:t>がない場合
                        continue

                    # 先に次の<w:t>が全角スペースで始まる場合、それを削除
                    next_run = runs[i + 1]
                    remove_leading_full_width_space(next_run)

                    # 条件: 次の<w:t>が全角スペースのみである場合、または先頭が全角スペースの場合、追記しない
                    next_w_t = next_run.find(".//w:t", namespaces=ns)
                    if next_w_t is not None:
                        if next_w_t.text == "　" or (next_w_t.text and next_w_t.text.startswith("　")):
                            continue
                    
                    # 全角スペースがない場合、追記してハイライトを追加
                    if not text.endswith("　"):
                        text += "　"  # 全角スペースを追記
                        modified = True

                    # 正規表現に該当する場合のみ補完し、ログに記録
                    if modified and re.match(validation_patterns[level], text):
                        log_entry += f" -> 補完後のテキスト: '{text}'"
                        log.append(log_entry)

                        # テキストの更新
                        w_t.text = text

                        # 変更箇所にハイライトを追加
                        highlight_text(run)

        # 不要な<w:t xml:space="preserve"> </w:t>要素を削除
        delete_preserved_space(paragraph)

def process_brackets_in_xml(xml_content):
    """
    XML文書を解析し、カッコやスペースを補完する処理を実行し、処理結果をログに記録する。
    変更が必要な箇所のみ変更を行い、ハイライトを追加する。
    """
    try:
        parser = ET.XMLParser(remove_blank_text=True)
        tree = ET.fromstring(xml_content.encode('utf-8'), parser)
    except ET.XMLSyntaxError as e:
        raise ValueError(f"XMLの解析中にエラーが発生しました: {e}")
    
    log = []

    for paragraph in tree.findall(".//w:p", namespaces=ns):
        complete_brackets(paragraph, log)

    return ET.tostring(tree, encoding='unicode', pretty_print=True), log

# XMLファイルの読み込みと処理
xml_file_path = "xml_new/word/document.xml"
with open(xml_file_path, "r", encoding="utf-8") as file:
    xml_content = file.read()

# XML解析とカッコ補完処理の実行
updated_xml, log = process_brackets_in_xml(xml_content)

# 修正されたXMLを保存
with open(xml_file_path, "w", encoding="utf-8") as file:
    file.write(updated_xml)

# ログの出力
log_file_path = "bracket_completion_log.txt"
with open(log_file_path, "w", encoding="utf-8") as file:
    for entry in log:
        file.write(entry + "\n")

print(f"カッコとスペースの補完処理が完了しました。ログは {log_file_path} に保存されました。")