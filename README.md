【操作方法】
1.dataディレクトリに校閲したいファイルを投入
2.make_xml_from_wordfile.pyを実行
    .docxをxml形式に解凍し、解析が行える状態にします。
3.retuouch_indent_number.pyを実行
    項目番号の記法が誤っている場合に修正。
4.update_indent_number.pyを実行
    項目番号が明らかに連番でない場合に修正。
6.update_indent_level.pyを実行
    インデントレベルの調整。
6.remake_wordfile_from_xml.pyを実行
    解析済みのxmlを.docxに再圧縮します。
7.delete_files.pyを実行
    実行の結果出力されたファイルを削除し、次の解析が行える状態とします。