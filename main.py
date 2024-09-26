# docx_processing.py から関数をインポート
from make_xml_from_wordfile import get_docx_file, extract_docx_to_xml
from retuouch_indent_number import process_brackets_in_xml
from update_indent_number import process_xml_level_1_to_4, process_xml_level_5_to_9
from update_indent_level import update_indent_level
from remake_wordfile_from_xml import create_docx
import os


# .docx ファイルのパス取得
docx_file = get_docx_file("data")  # ディレクトリを指定

# XMLへ変換
extract_docx_to_xml(docx_file, "xml/")
extract_docx_to_xml(docx_file, "xml_new/")  # 別ディレクトリへの変換


"""
項目番号の形式に誤りがあった場合に修正する処理
"""
# 対象のxmlファイルを開く
xml_file_path = 'xml_new/word/document.xml'
with open(xml_file_path, "r", encoding="utf-8") as file:
    xml_content = file.read()

# XML解析と補完処理の実行
updated_xml, log_brackets = process_brackets_in_xml(xml_content)

# 修正されたXMLを保存(中間保存)
with open(xml_file_path, "w", encoding="utf-8") as file:
    file.write(updated_xml)

# ログの出力（補完処理部分のみ）
log_file_path = "bracket_completion_log.txt"
with open(log_file_path, "w", encoding="utf-8") as file:
    for entry in log_brackets:
        file.write(entry + "\n")

"""
項目番号の連番に誤りがあった場合に修正する処理
"""
# 対象のxmlファイルを開く
xml_file_path = 'xml_new/word/document.xml'
with open(xml_file_path, "r", encoding="utf-8") as file:
    xml_content = file.read()

# レベル1〜4のXML解析の実行
updated_xml_1_to_4, log_1_to_4 = process_xml_level_1_to_4(xml_content)

# レベル5〜9のXML解析の実行
updated_xml_5_to_9, log_5_to_9 = process_xml_level_5_to_9(updated_xml_1_to_4)

# 修正されたXMLを保存(中間保存)
with open(xml_file_path, "w", encoding="utf-8") as file:
    file.write(updated_xml_5_to_9)

"""
インデントレベルを修正する処理
"""
# 対象のxmlファイルを開く
xml_file_path = 'xml_new/word/document.xml'
with open(xml_file_path, "r", encoding="utf-8") as file:
    xml_content = file.read()

# XML解析の実行
updated_xml, log_level = update_indent_level(xml_content)

# 修正されたXMLを保存
with open(xml_file_path, "w", encoding="utf-8") as file:
    file.write(updated_xml)

# ログの出力
log_file_path = "indentation_log.txt"
with open(log_file_path, "w", encoding="utf-8") as file:
    file.write("\n".join(log_level))

# 校閲後のXMLファイルをWordファイルに再構成
core_filename = os.path.splitext(os.path.basename(docx_file))[0]
output_docx = f"【校閲ずみ】{core_filename}.docx"
create_docx("xml_new", output_docx)