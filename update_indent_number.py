import re
import xml.etree.ElementTree as ET
from collections import Counter

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

def find_majority_number(numbers):
    """
    全ての番号の最初の値を集計し、一番メジャーな値を返す。
    """
    first_digits = [number[0] for number in numbers if number]
    counter = Counter(first_digits)
    # 最も出現頻度の高い数字を取得
    majority_value = counter.most_common(1)[0][0]
    return majority_value

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

def process_xml(xml_content):
    """
    XML文書を解析し、各段落に対して項目番号やインデントを適用する。
    順番飛ばしがある場合は修正し、処理結果とログを返す。
    """
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    all_numbers = []  # 全ての番号を保存
    current_numbers = {i: [] for i in range(1, 5)}  # レベルごとに項目番号を保存
    log = []

    # まず全ての番号を収集
    for paragraph in root.findall(".//w:p", namespaces=ns):
        level, text = parse_paragraph(paragraph)

        if level is not None and level <= 4:
            number = extract_number(text, level)
            if number:
                all_numbers.append(number)

    # メジャーな最初の数字を決定
    majority_value = find_majority_number(all_numbers)

    # 再度全ての段落を処理して番号を修正
    for paragraph in root.findall(".//w:p", namespaces=ns):
        level, text = parse_paragraph(paragraph)

        if level is not None and level <= 4:
            number = extract_number(text, level)
            previous_number = current_numbers[level]

            if number and number[0] != majority_value:
                # メジャーな値に基づいて番号を修正
                number[0] = majority_value
                new_number = format_number(number, level)
                update_text(paragraph, new_number)
                log.append(f"レベル{level}で項目番号の修正: {text} -> {new_number}")
                current_numbers[level] = number
            else:
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