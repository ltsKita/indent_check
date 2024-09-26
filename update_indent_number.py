"""
このファイルでは項目番号が連番になるように修正します。
正規表現ルールで項目番号の値を取得します。
各項目番号レベルにおいて現在の値が、前回のものを1だけ進めた値と異なってれば連番ではないとして修正します。
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
    6: r"^[ａ-ｚ]．.*$",               # レベル6: 例 "ａ．"
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

def check_tab_before_t(paragraph):
    """
    <w:tab />が<w:t>要素の前にあるかどうかを確認する。
    存在する場合、それは項目番号ではないので処理をスキップする
    """
    runs = paragraph.findall(".//w:r", namespaces=ns)
    return len(runs) > 1


def highlight_text(run):
    """
    指定された<w:r>要素にハイライトを追加する。
    """
    rPr = run.find('.//w:rPr', namespaces=ns)
    if rPr is None:
        rPr = ET.SubElement(run, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr')
    
    highlight = rPr.find('.//w:highlight', namespaces=ns)
    if highlight is not None:
        rPr.remove(highlight)

    new_highlight = ET.SubElement(rPr, '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}highlight')
    new_highlight.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', 'yellow')

def extract_number(text, level):
    """
    テキストから項目番号を抽出し、リストとして返す。
    """
    match = re.match(patterns.get(level, ''), text)
    if not match:
        return []

    if level in [1, 2, 3, 4]:
        return list(map(int, re.findall(r"\d+", match.group())))
    elif level == 5:
        return [int(re.search(r"\d+", match.group()).group())]
    elif level == 6:
        return [re.search(r"[ａ-ｚ]", match.group()).group()]
    elif level == 7:
        return [re.search(r"[a-z]", match.group()).group()]
    elif level == 8:
        parts = re.match(r"\(([a-z])-(\d+)\)", match.group())
        return [parts.group(1), int(parts.group(2))]
    elif level == 9:
        parts = re.match(r"\(([a-z])-(\d+)-(\d+)\)", match.group())
        return [parts.group(1), int(parts.group(2)), int(parts.group(3))]

    return []

def update_text(paragraph, new_number):
    """
    段落内のテキストを新しい項目番号に置き換える。
    """
    for t in paragraph.findall(".//w:t", namespaces=ns):
        text = t.text.strip()
        new_text = re.sub(r"^\d+(\.\d+)*", new_number, text, 1)
        t.text = new_text

def parse_paragraph(paragraph):
    """
    段落から項目番号を抽出し、適切なレベルを返す。
    項目番号がない場合は、Noneを返す。
    """
    text = "".join([t.text for t in paragraph.findall(".//w:t", namespaces=ns)]).strip()

    for level, pattern in patterns.items():
        if re.match(pattern, text):
            return level, text

    return None, text

def increment_number(number, level, previous_numbers, current_text, is_first_item):
    """
    項目番号をインクリメント(連番処理)して、次の番号に修正する。
    最初の項目は文書内の番号をそのまま使用し、インクリメントしない。
    """
    existing_sub_level_number = extract_lowest_sub_number(current_text, level)

    print(f"Before incrementing - Level: {level}, Previous numbers: {previous_numbers}, Extracted sub level number: {existing_sub_level_number}")

    # 上位レベルの番号は直前の番号をそのまま引き継ぐ
    for i in range(level - 1):
        number[i] = previous_numbers[i]
    # 項目番号は初回のみインクリメントしない。2回目以降は順にインクリメント
    if is_first_item:
        number[level - 1] = existing_sub_level_number
    else:
        if number[level - 1] == 0:
            # 項目番号がリセットされた場合は初回のみインクリメントしない
            number[level - 1] = existing_sub_level_number
        else:
            # 通常時はインクリメント
            number[level - 1] += 1

    print(f"After incrementing - New numbers: {number}")

    return number

def extract_lowest_sub_number(text, level):
    """
    項目番号の数字をリスト形式で保持。
    その中で最も内側にある数字を取得し、必要な場合その数字で初期化する。
    取得できない場合は1で初期化する。
    ※今回の文書が9.1や9.2がなく、9.3から始まっていたためこの関数を作成
    """
    matches = re.findall(r"\d+", text)
    print(f"Extracted numbers from text '{text}': {matches}")
    
    if matches and len(matches) >= level:
        return int(matches[level - 1])
    else:
        print(f"Could not find a valid number for level {level} in text: '{text}'")
        return 1

def process_xml_level_1_to_4(xml_content):
    """
    レベル1〜4の処理を行う。
    項目番号が連番かチェックし、そうでない場合は修正
    """
    # XMLコンテンツをElementTree形式に変換し、ルート要素を取得
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    # base_numbersは階層ごとの項目番号を保持する配列。初期値は全て0
    base_numbers = [0, 0, 0, 0]
    log = []
    is_first_item = True

    # ログファイルを作成し、処理の結果を追跡
    with open("indentation_log_level_1_to_4.txt", "w", encoding="utf-8") as log_file:

        # すべての段落 (<w:p> 要素) を精査
        for paragraph in root.findall(".//w:p", namespaces=ns):
            # <w:tab />が <w:t> の前に存在する場合、その段落の処理をスキップ
            if check_tab_before_t(paragraph):
                continue
            
             # 段落から項目番号のレベルとテキストを取得
            level, text = parse_paragraph(paragraph)

            # 項目番号ではない場合（正規表現にマッチしない場合）は処理をスキップ
            if level is None:
                continue
            # レベル1〜4の項目番号のみ処理を行う
            if level is not None and level <= 4:
                number = extract_number(text, level)
                current_text = "".join([t.text for t in paragraph.findall(".//w:t", namespaces=ns)]).strip()
                # レベル1の場合、base_numbersの最初の要素に項目番号を設定し、初期化
                if level == 1:
                    base_numbers[0] = number[0]
                # 上位レベルがまだ初期化（設定）されていない場合、現在のレベルの項目番号の処理を行わない
                else:
                    if base_numbers[level - 2] == 0:
                        continue
                    # document_number で取得した現在の項目番号を previous_numbers として保持し、変更前の状態を保存
                    document_number = extract_number(current_text, level)
                    previous_numbers = document_number.copy()
                    # base_numbersはドキュメント全体における現在の項目番号の状態を保持するリスト。項目番号の階層（レベル1〜4）の状態を追跡している。
                    # ここで項目番号をインクリメント
                    base_numbers = increment_number(base_numbers, level, base_numbers, text, is_first_item)

                    # インクリメント前後の番号を比較して、変更があるか確認
                    previous_list = previous_numbers[:level]
                    current_list = base_numbers[:level]

                    log_file.write(f"Comparing: previous_list={previous_list} with current_list={current_list}\n")

                    # 番号が異なる場合、変更があったと判断してハイライトを追加
                    if previous_list != current_list:
                        log_file.write(f"Highlighting change: {previous_list} -> {current_list}\n")
                        for run in paragraph.findall(".//w:r", namespaces=ns):
                            highlight_text(run)
                    else:
                        log_file.write(f"No change detected: {previous_list} == {current_list}\n")

            # 番号をフォーマットし、現在の段落に対して更新
            formatted_number = format_number(base_numbers, level, add_period=(level == 1))
            update_text(paragraph, formatted_number)

            if level != None:
                log.append(f"レベル{level}で項目番号の修正: {text} -> {formatted_number}")

            is_first_item = False

    return ET.tostring(root, encoding='unicode'), log

def format_number(number, level, add_period=True):
    """
    数字リストを整形して、項目番号形式にする関数。
    """
    if level == 5:
        return f"({number[0]})"  # 例: (1)
    elif level == 6:
        return f"{number[0]}．"  # 全角文字 + 全角ピリオド 例: ａ．
    elif level == 7:
        return f"({number[0]})"  # 例: (a)
    elif level == 8:
        return f"({number[0]}-{number[1]})"  # 例: (a-1)
    elif level == 9:
        return f"({number[0]}-{number[1]}-{number[2]})"  # 例: (a-1-1)
    else:
        return ".".join(map(str, number[:level])) + ("." if add_period and level > 1 else "")
    
def increment_level_number(current_number, level):
    """
    レベルに応じたインクリメントを定義する関数。
    """
    if level == 5:
        # レベル5は数字をインクリメント
        return [current_number[0] + 1]
    elif level == 6:
        # レベル6は全角アルファベットをインクリメント（例：ａ -> ｂ）
        next_char = chr(ord(current_number[0]) + 1)
        return [next_char]
    elif level == 7:
        # レベル7は半角アルファベットをインクリメント
        next_char = chr(ord(current_number[0]) + 1)
        return [next_char]
    elif level == 8:
        # レベル8はレベル7の値を参照しつつ、最後の数字部分をインクリメント
        return [current_number[0], current_number[1] + 1]
    elif level == 9:
        # レベル9はレベル8の値を参照しつつ、最後の数字部分をインクリメント
        return [current_number[0], current_number[1], current_number[2] + 1]

def process_xml_level_5_to_9(xml_content):
    """
    レベル5〜9の処理を行い、番号が連番かどうかをチェックする。
    番号が非連番の場合は連番となるように修正し、ハイライトを追加する。
    """
    # XMLコンテンツをElementTree形式に変換し、ルート要素を取得
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()

    previous_numbers_for_levels_5_to_9 = {}  # 各レベルの前回の番号を保持する辞書
    log = []

    with open("indentation_log_level_5_to_9.txt", "w", encoding="utf-8") as log_file:

        for paragraph in root.findall(".//w:p", namespaces=ns):
            # <w:tab />が <w:t> の前に存在する場合、その段落の処理をスキップ
            if check_tab_before_t(paragraph):
                continue 
            level, text = parse_paragraph(paragraph)

            # 項目番号ではない場合は処理をスキップ
            if level is None:
                continue

            # レベル1〜4はスキップする
            if level is None or level < 5:
                continue 
            
            # 前回の段落の項目番号を辞書から取得。初期値は[1]、['ａ']などを使用
            previous_number = previous_numbers_for_levels_5_to_9.get(
                level, 
                [1] if level == 5 else 
                ['ａ'] if level == 6 else 
                ['a'] if level == 7 else 
                ['a', 1] if level == 8 else 
                ['a', 1, 1]
            )

            # 現在の段落の項目番号を抽出する（例: (1), (a-1) など）
            current_number = extract_number(text, level)

            # 初期値と一致するかどうかを確認
            initial_value = [1] if level == 5 else ['ａ'] if level == 6 else ['a'] if level == 7 else ['a', 1] if level == 8 else ['a', 1, 1]

            # current_number がリセットされ、初期値である場合はインクリメントを行わない
            # 次回以降は順にインクリメントする
            if current_number == initial_value:
                next_expected_number = current_number  # 初期値でリセット
                log_file.write(f"Initial value detected at level {level}. Resetting sequence to {current_number}\n")
            else:
                # 初期値でなければ、インクリメントして次の連番を設定
                next_expected_number = increment_level_number(previous_number, level)

            # 前回の番号をインクリメントしたものと現在の番号を比較
            # 一致していなければ連番でないと判断
            if current_number != next_expected_number:
                log_file.write(f"Non-sequential detected: expected {next_expected_number} but got {current_number}\n")
                
                # インクリメントされた番号に修正
                formatted_number = format_number(next_expected_number, level, add_period=False)
                for t in paragraph.findall(".//w:t", namespaces=ns):
                    original_text = t.text.strip()
                    if level == 5:
                        new_text = re.sub(r"\(\d+\)", formatted_number, original_text, count=1)
                    elif level == 6:
                        new_text = re.sub(r"^[ａ-ｚ]．", formatted_number, original_text, count=1)
                    elif level == 7:
                        new_text = re.sub(r"\([a-z]\)", formatted_number, original_text, count=1)
                    elif level == 8:
                        new_text = re.sub(r"\([a-z]-\d+\)", formatted_number, original_text, count=1)
                    elif level == 9:
                        new_text = re.sub(r"\([a-z]-\d+-\d+\)", formatted_number, original_text, count=1)

                    if new_text != original_text:
                        t.text = new_text
                        log_file.write(f"Updated <w:t> from {original_text} to {new_text}\n")
                    else:
                        log_file.write(f"No update: {original_text} remains unchanged\n")

                # ハイライトを追加
                for run in paragraph.findall(".//w:r", namespaces=ns):
                    highlight_text(run)
            else:
                log_file.write(f"Sequential order confirmed for level {level}: {next_expected_number} == {current_number}\n")

            # 今回の番号を辞書に保存し、次回の比較に使用
            previous_numbers_for_levels_5_to_9[level] = next_expected_number
            next_expected_number = increment_level_number(next_expected_number, level)  # 次の番号をインクリメント
            log_file.write(f"Next expected_number for level {level}: {next_expected_number}\n")
            log_file.write(f"-"*50+"\n")

    return ET.tostring(root, encoding='unicode'), log

# このスクリプトが直接実行された場合のみ、以下のコードが動作するようにする
if __name__ == "__main__":
    # XMLファイルの内容を読み込み
    xml_file_path = "xml_new/word/document.xml"
    with open(xml_file_path, "r", encoding="utf-8") as file:
        xml_content = file.read()

    # レベル1〜4のXML解析の実行
    updated_xml_1_to_4, log_1_to_4 = process_xml_level_1_to_4(xml_content)

    # レベル5〜9のXML解析の実行
    updated_xml_5_to_9, log_5_to_9 = process_xml_level_5_to_9(updated_xml_1_to_4)

    # 修正されたXMLを保存
    with open(xml_file_path, "w", encoding="utf-8") as file:
        file.write(updated_xml_5_to_9)