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
    6: r"^[ａ-ｚ]\.\s*.*$",              # レベル6: 例 "ａ．"
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

def extract_number(text, level):
    """
    テキストから項目番号を抽出し、リストとして返す。
    """
    if level in [1, 2, 3, 4]:
        return list(map(int, re.findall(r"\d+", text)))
    return []

def is_sequential(current, previous):
    """
    項目番号が連番かどうかを確認する。
    """
    if not previous:
        return True  # 最初の項目の場合は常にTrue
    for c, p in zip(current, previous):
        if c != p + 1:
            return False
    return True

def increment_number(number, level):
    """
    項目番号をインクリメントして、次の番号に修正する。
    レベルに応じて番号を修正し、それ以降のレベルの番号をリセット。
    """
    if len(number) < level:
        # レベルが足りない場合は追加
        number.append(1)
    else:
        # 指定レベルの番号をインクリメント
        number[level - 1] += 1
        # それ以降のレベルの番号をリセット
        number = number[:level]

    return number

def format_number(number, level):
    """
    項目番号をフォーマットして文字列に戻す。
    指定されたレベルに応じて、必要な部分のみを表示。
    """
    return ".".join(map(str, number[:level]))

def update_text(paragraph, new_number):
    """
    段落内のテキストを新しい項目番号に置き換える。
    """
    for t in paragraph.findall(".//w:t", namespaces=ns):
        text = t.text.strip()
        # 古い項目番号部分を新しい番号に置き換える
        new_text = re.sub(r"^\d+(\.\d+)*", new_number, text, 1)
        t.text = new_text

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

def process_xml(xml_content):
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    current_numbers = {i: [] for i in range(1, 5)}  # レベルごとに項目番号を保存
    log = []

    for paragraph in root.findall(".//w:p", namespaces=ns):
        level, text = parse_paragraph(paragraph)

        if level is not None and level <= 4:
            number = extract_number(text, level)
            previous_number = current_numbers[level]

            if not is_sequential(number, previous_number):
                # 順序が不正な場合、次の正しい番号に修正
                correct_number = increment_number(previous_number.copy(), level)
                new_number = format_number(correct_number, level)
                update_text(paragraph, new_number)
                log.append(f"レベル{level}で項目番号の修正: {text} -> {new_number}")

                # 修正された番号を現在の番号として保存
                current_numbers[level] = correct_number
            else:
                # 正しい順番の場合、そのまま保存
                current_numbers[level] = number
                log.append(f"レベル{level}で正しい順序: {text}")

    # 修正済みのXMLを返す
    return ET.tostring(root, encoding='unicode'), log

# XMLファイルの内容を読み込み
xml_file_path = "xml_new/word/document.xml"
with open(xml_file_path, "r", encoding="utf-8") as file:
    xml_content = file.read()

# XML解析の実行
updated_xml, log = process_xml(xml_content)

# 修正されたXMLを保存
with open(xml_file_path, "w", encoding="utf-8") as file:
    file.write(updated_xml)

# ログの出力
log_file_path = "indentation_log.txt"
with open(log_file_path, "w", encoding="utf-8") as file:
    for entry in log:
        file.write(f"{entry}\n")

print(f"処理が完了しました。ログは {log_file_path} に保存されました。")