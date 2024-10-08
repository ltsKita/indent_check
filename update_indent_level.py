"""
このファイルではインデントレベルの修正を行います。
インデントルールの適用にあたり、全てのインデントの再設定を実施する処理となるため文書全体が変更範囲とみなされます。
そのため、ここではハイライト処理を行っていません。
"""

import re
import xml.etree.ElementTree as ET

# WordprocessingMLの名前空間を定義
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
ET.register_namespace('w', ns['w'])  # 'w' 名前空間を再登録

# 正規表現パターン
patterns = {
    1: r"^\d+\.(?!\d)\s*.*$",         # レベル1: 例 "1.", "10." (ピリオドの後に数字が続かない)
    2: r"^\d+\.\d+(?!\.\d)\s*.*$",    # レベル2: 例 "1.1", "1.6", "10.2" (ピリオド+数字が続く)
    3: r"^\d+\.\d+\.\d+(?!\.\d)\s*.*$",  # レベル3: 例 "1.1.1", "10.2.3"
    4: r"^\d+\.\d+\.\d+\.\d+\s*.*$",     # レベル4: 例 "1.1.1.1", "10.2.3.4"
    5: r"^\(\d+\)\s*.*$",                # レベル5: 例 "(1)", "(10)"
    6: r"^[ａ-ｚ]\．\s*.*$",              # レベル6: 例 "ａ．"
    7: r"^\([a-z]\)\s*.*$",              # レベル7: 例 "(a)"
    8: r"^\([a-z]-\d+\)\s*.*$",          # レベル8: 例 "(a-1)"
    9: r"^\([a-z]-\d+-\d+\)\s*.*$"       # レベル9: 例 "(a-1-1)"
}

# 項目番号に対するインデント設定（レベル5以上のみ）
indent_settings_numbers = {
    5: {"w:leftChars": "100", "w:left": "240"},
    6: {"w:leftChars": "200", "w:left": "480"},
    7: {"w:leftChars": "300", "w:left": "720"},
    8: {"w:leftChars": "400", "w:left": "960"},
    9: {"w:leftChars": "500", "w:left": "1200"}
}

# 段落に対するインデント設定
indent_settings_paragraphs = {
    1: {"w:leftChars": "100", "w:left": "240", "w:firstLineChars": "100", "w:firstLine": "240"},
    2: {"w:leftChars": "100", "w:left": "240", "w:firstLineChars": "100", "w:firstLine": "240"},
    3: {"w:leftChars": "100", "w:left": "240", "w:firstLineChars": "100", "w:firstLine": "240"},
    4: {"w:leftChars": "100", "w:left": "240", "w:firstLineChars": "100", "w:firstLine": "240"},
    5: {"w:leftChars": "100", "w:left": "240", "w:firstLineChars": "100", "w:firstLine": "240"},
    6: {"w:leftChars": "200", "w:left": "480", "w:firstLineChars": "100", "w:firstLine": "240"},
    7: {"w:leftChars": "300", "w:left": "720", "w:firstLineChars": "100", "w:firstLine": "240"},
    8: {"w:leftChars": "400", "w:left": "960", "w:firstLineChars": "100", "w:firstLine": "240"},
    9: {"w:leftChars": "500", "w:left": "1200", "w:firstLineChars": "100", "w:firstLine": "240"}
}
def has_previous_paragraph_drawing(paragraph, root):
    """
    現在の段落の一つ前の段落に<w:drawing>タグがあり、その段落に「図」または「表」のキーワードが
    含まれている場合にTrueを返す。含まれていない場合はFalseを返す。
    """
    # 現在の段落 (w:p) の前の段落を取得
    paragraphs = list(root.findall(".//w:p", namespaces=ns))
    current_index = paragraphs.index(paragraph)

    # 最初の段落の場合、前の段落は存在しないのでFalseを返す
    if current_index == 0:
        return False

    # 一つ前の段落を取得
    previous_paragraph = paragraphs[current_index - 1]
    current_paragraph = paragraph

    # その段落内に<w:drawing>タグがあるかを確認
    drawing = previous_paragraph.find(".//w:drawing", namespaces=ns)
    
    # <w:drawing>がなければFalseを返す
    if drawing is None:
        return False

    # <w:t>のテキストをすべて取得して結合
    current_paragraph_text = extract_text_from_paragraph(current_paragraph)

    # 「図」または「表」が含まれていればTrueを返す
    if "図" in current_paragraph_text or "表" in current_paragraph_text:
        if drawing is not None:
            print(current_paragraph_text)
        return True

    return False

def extract_text_from_paragraph(paragraph):
    """
    段落 (w:p) の中にある<w:t>要素のテキストをすべて結合して返す。
    """
    return "".join([t.text for t in paragraph.findall(".//w:t", namespaces=ns) if t.text])



def parse_paragraph(paragraph):
    """
    段落から項目番号を抽出し、適切なレベルを返す。
    項目番号がない場合は、Noneを返す。
    """
    # paragraph内の<w:t>のテキストを全て結合し、前後の空白を削除
    text = "".join([t.text for t in paragraph.findall(".//w:t", namespaces=ns)]).strip()

    # テキスト全体をチェックして、項目番号のパターンが一致するか確認
    for level, pattern in patterns.items():
        if re.match(pattern, text):
            return level, text

    return None, text

def update_indent(paragraph, level, is_number):
    """
    段落のインデントを更新する。
    is_numberがTrueの場合は項目番号、Falseの場合は通常段落のインデントを適用する。
    ただし、<w:t>が存在しない、または空の場合はインデントを適用しない。
    """
    # <w:t> 要素を確認し、空や存在しない場合はインデントを適用しない
    texts = paragraph.findall(".//w:t", namespaces=ns)
    if not texts or all(t.text.strip() == "" for t in texts if t.text):
        return  # <w:t> がないか、空であれば何もしない

    pPr = paragraph.find(".//w:pPr", namespaces=ns)
    if pPr is None:
        pPr = ET.SubElement(paragraph, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr")
    
    # 既存の <w:ind> を削除
    ind = pPr.find(".//w:ind", namespaces=ns)
    if ind is not None:
        pPr.remove(ind)

    # is_number に基づいてインデント設定を取得
    settings = indent_settings_numbers.get(level) if is_number else indent_settings_paragraphs.get(level)

    # 新しいインデント設定を追加
    if settings:
        new_ind = ET.SubElement(pPr, "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}ind")
        for attr, value in settings.items():
            new_ind.set(attr, value)

def update_indent_level(xml_content):
    """
    XML文書を解析し、各段落に対して項目番号やインデントを適用する。
    処理結果をXMLとして返し、処理ログも返す。
    """
    # XMLの読み込み
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    current_level = 1
    current_numbers = {i: None for i in range(1, 10)}
    log = []

    for paragraph in root.findall(".//w:p", namespaces=ns):
        if has_previous_paragraph_drawing(paragraph, root):
            # 今の段落に含まれる<w:t>タグのテキストを取得してログに追加
            continue #インデント処理を行わない
        level, item_number = parse_paragraph(paragraph)

        if level is not None:
            # 項目番号に対するインデントの更新
            update_indent(paragraph, level, is_number=True)

            # 順序のチェック
            if current_numbers.get(level):
                previous_number = current_numbers[level]
                if item_number != previous_number:
                    log.append(f"レベル{level}で項目番号の順序が不正です: {previous_number} -> {item_number}. 内容: '{item_number}'")

            current_numbers[level] = item_number
            current_level = level

            log.append(f"レベル{level}: {item_number} - インデント更新 (項目番号). 内容: '{item_number}'")
        else:
            # 通常段落の場合は現在のレベルのインデントを適用
            text = item_number  # ここではitem_numberが通常段落のテキスト
            update_indent(paragraph, current_level, is_number=False)
            log.append(f"レベル{current_level}: 通常段落 - インデント適用. 内容: '{text}'")

    # 修正済みのXMLを返す
    return ET.tostring(root, encoding='unicode'), log

# このスクリプトが直接実行された場合のみ、以下のコードが動作するようにする
if __name__ == "__main__":
    # XMLファイルの内容を読み込み
    xml_file_path = "xml_new/word/document.xml"
    with open(xml_file_path, "r", encoding="utf-8") as file:
        xml_content = file.read()

    # XML解析の実行
    updated_xml, log = update_indent_level(xml_content)

    # 修正されたXMLを保存
    with open(xml_file_path, "w", encoding="utf-8") as file:
        file.write(updated_xml)

    # ログの出力
    log_file_path = "indentation_log.txt"
    with open(log_file_path, "w", encoding="utf-8") as file:
        file.write("\n".join(log))

    print(f"処理が完了しました。ログは {log_file_path} に保存されました。")