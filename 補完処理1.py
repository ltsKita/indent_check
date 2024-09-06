import re
from lxml import etree as ET

# WordprocessingMLの名前空間を定義
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

# 補完するパターン
bracket_patterns = {
    5: (r"^\d+", "(", ")"),         # レベル5: 例 "(1)"
    7: (r"^[a-z]", "(", ")"),       # レベル7: 例 "(a)"
    8: (r"^[a-z]-\d+", "(", ")"),   # レベル8: 例 "(a-1)"
    9: (r"^[a-z]-\d+-\d+", "(", ")")# レベル9: 例 "(a-1-1)"
}

# 正規表現パターン（補完後のテキストが引っかかるか確認する用）
validation_patterns = {
    5: r"^\(\d+\)\s*.*$",                # レベル5: 例 "(1)"
    7: r"^\([a-z]\)\s*.*$",              # レベル7: 例 "(a)"
    8: r"^\([a-z]-\d+\)\s*.*$",          # レベル8: 例 "(a-1)"
    9: r"^\([a-z]-\d+-\d+\)\s*.*$"       # レベル9: 例 "(a-1-1)"
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

def complete_brackets(paragraph, log):
    """
    段落内の<w:t>のテキストを解析し、カッコが不足している場合に補完する。
    必要な場合のみ変更を行い、ログに記録する。また、変更箇所にハイライトを追加する。
    """
    for run in paragraph.findall(".//w:r", namespaces=ns):
        w_t = run.find(".//w:t", namespaces=ns)
        if w_t is None:
            continue
        
        original_text = w_t.text
        if original_text is None:
            continue

        text = original_text.strip()

        # カッコが欠けているか確認し、補完が必要な場合のみ変更を行う
        modified = False  # 変更があったかどうかを確認
        if "(" in text or ")" in text:
            for level, (pattern, open_bracket, close_bracket) in bracket_patterns.items():
                if re.match(pattern, text):
                    log_entry = f"レベル{level}: 元のテキスト: '{original_text}'"

                    # カッコの補完: 前カッコがない場合後ろカッコ、後ろカッコがない場合前カッコを補完
                    if open_bracket in text and close_bracket not in text:
                        text = text + close_bracket
                        modified = True
                    elif close_bracket in text and open_bracket not in text:
                        text = open_bracket + text
                        modified = True

                    # 正規表現に該当する場合のみ補完し、ログに記録
                    if modified and re.match(validation_patterns[level], text):
                        log_entry += f" -> 補完後のテキスト: '{text}'"
                        log.append(log_entry)

                        # テキストの更新
                        w_t.text = text

                        # 変更箇所にハイライトを追加
                        highlight_text(run)

def process_brackets_in_xml(xml_content):
    """
    XML文書を解析し、カッコを補完する処理を実行し、処理結果をログに記録する。
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

print(f"カッコ補完処理が完了しました。ログは {log_file_path} に保存されました。")