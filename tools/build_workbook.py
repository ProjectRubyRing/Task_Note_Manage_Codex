"""Build the distributable workbooks without starting Excel or changing security settings."""

from __future__ import annotations

import argparse
from datetime import datetime
import io
from pathlib import Path
import re
import struct
import sys
from zipfile import ZIP_DEFLATED, ZipFile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / ".build" / "deps"))

import xlsxwriter
from pyopenvba import ExcelFile

FONT = "Meiryo UI"
STATUSES = ("起票済み", "対応中", "対応済み", "連携済み", "完了")
COLORS = ("#FFEFA6", "#C8E6FF", "#D2EDD6", "#E7DAF9", "#E2E5E9")
INK = "#243140"
MUTED = "#68707A"
PAPER = "#F5F3EF"
SHEETS = (
    ("タスクボード", "ActiveBoard"),
    ("完了", "CompletedBoard"),
    ("付箋の編集", "TaskEditor"),
    ("タスク一覧", "TaskData"),
    ("使い方", "Guide"),
    ("設定", "Settings"),
)
SAMPLE_TIME = datetime(2026, 10, 3, 9, 0)
SAMPLES = [
    ["T000001", "【サンプル】仕様を整理する", "目的と必要な画面を整理する。\n不明点を洗い出して確認する。", "関連：要件定義書・画面仕様書", "起票済み", SAMPLE_TIME, SAMPLE_TIME, None],
    ["T000002", "【サンプル】入力画面を実装する", "タイトル・内容・備考の入力を実装する。\n未入力時の動作も確認する。", "関連実装：入力フォーム\n確認点：必須項目・改行", "対応中", SAMPLE_TIME, SAMPLE_TIME, None],
    ["T000003", "【サンプル】動作確認を終える", "主要な操作とエラー時の動作を確認する。\n確認結果を記録する。", "確認内容：新規登録・更新・再表示", "対応済み", SAMPLE_TIME, SAMPLE_TIME, None],
    ["T000004", "【サンプル】関係者へ連携する", "実装内容と確認結果を共有する。\n受領・確認が済んだら完了へ変更する。", "連携方法や確認待ちの内容を記録", "連携済み", SAMPLE_TIME, SAMPLE_TIME, None],
    ["T000005", "【サンプル】完了したタスク", "この付箋は完了シートに保管されます。\n内容は削除されません。", "再開する場合はステータスを戻して登録・更新", "完了", SAMPLE_TIME, SAMPLE_TIME, SAMPLE_TIME],
]


def add_vba(path: Path) -> None:
    with ExcelFile(path) as book:
        project = book.add_vba_project()
        # New projects originate from an Office-authored template. The project
        # source encoding and its dir record must agree for Japanese VBA strings.
        old_record = struct.pack("<HIH", 3, 2, project.code_page)
        new_record = struct.pack("<HIH", 3, 2, 932)
        if project.dir_raw.count(old_record) != 1:
            raise ValueError("PROJECTCODEPAGE record must occur exactly once")
        project.dir_raw = project.dir_raw.replace(old_record, new_record, 1)
        project.code_page = 932
        project.dir_structure_dirty = True
        for name in ("TaskRules", "TaskBoard"):
            source = (ROOT / "src" / f"{name}.bas").read_text(encoding="utf-8")
            source.encode("cp932")  # Fail on characters VBA cannot preserve.
            project.add_module(name, source.replace("\n", "\r\n"))
        source = (ROOT / "src" / "ThisWorkbook.cls").read_text(encoding="utf-8")
        book.set_module("ThisWorkbook", source.replace("\n", "\r\n"))
        book.save()


def bind_shapes(path: Path) -> None:
    """Assign macros to DrawingML text boxes and mark the OOXML host as XLSM."""
    from xml.etree import ElementTree as ET

    ns = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}
    with ZipFile(path) as archive:
        entries = [(info, archive.read(info.filename)) for info in archive.infolist()]
    output = io.BytesIO()
    with ZipFile(output, "w", ZIP_DEFLATED) as archive:
        for info, data in entries:
            if re.fullmatch(r"xl/drawings/drawing\d+\.xml", info.filename):
                # Preserve the original namespace serialization; only mutate the
                # shape's start tag and nonvisual name, using parsed metadata.
                root = ET.fromstring(data)
                actions = {}
                for shape in root.findall(".//xdr:sp", ns):
                    nonvisual = shape.find("xdr:nvSpPr/xdr:cNvPr", ns)
                    description = nonvisual.get("descr", "")
                    if description.startswith("macro:"):
                        action = description.split(":", 1)[1]
                        actions[nonvisual.get("id")] = ("btn_" + action, action)
                    elif description.startswith("task:"):
                        task_id = description.split(":", 1)[1]
                        actions[nonvisual.get("id")] = ("card_" + task_id, "EditCard")
                text = data.decode("utf-8")

                def update_shape(match: re.Match) -> str:
                    block = match.group()
                    identity = re.search(r'<xdr:cNvPr\b[^>]*\bid="(\d+)"', block)
                    if identity and identity.group(1) in actions:
                        name, action = actions[identity.group(1)]
                        if "macro=" in block.split(">", 1)[0]:
                            block = re.sub(r'(<xdr:sp\b[^>]*\bmacro=")[^"]*', rf'\g<1>[0]!{action}', block, count=1)
                        elif block.startswith("<xdr:sp "):
                            block = block.replace("<xdr:sp ", f'<xdr:sp macro="[0]!{action}" ', 1)
                        else:
                            block = block.replace("<xdr:sp>", f'<xdr:sp macro="[0]!{action}">', 1)
                        block = re.sub(r'(<xdr:cNvPr\b[^>]*\bname=")[^"]*', rf'\g<1>{name}', block, count=1)
                    return block

                data = re.sub(r"<xdr:sp(?:\s[^>]*)?>.*?</xdr:sp>", update_shape, text, flags=re.S).encode("utf-8")
                data = data.replace(b'<a:latin typeface="Meiryo UI"/>',
                                    b'<a:latin typeface="Meiryo UI"/><a:ea typeface="Meiryo UI"/>')
            elif info.filename == "[Content_Types].xml":
                data = data.replace(
                    b"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
                    b"application/vnd.ms-excel.sheet.macroEnabled.main+xml",
                )
            archive.writestr(info, data)
    path.write_bytes(output.getvalue())


def preview(value: str, limit: int) -> str:
    value = value.replace("\r\n", " / ").replace("\n", " / ").replace("\r", " / ")
    if not value:
        return "（未記入）"
    return value if len(value) <= limit else value[:limit] + "…"


def build(path: Path, *, sample: bool = False) -> None:
    records = SAMPLES if sample else []
    path.parent.mkdir(parents=True, exist_ok=True)
    # in_memory avoids temporary-directory ACL differences in the build host.
    with xlsxwriter.Workbook(path, {"in_memory": True, "strings_to_formulas": False, "strings_to_urls": False}) as book:
        book.formats[0].set_font_name(FONT)
        book.formats[0].set_font_size(11)
        book.set_vba_name("ThisWorkbook")
        book.set_properties({"title": "付箋タスク管理", "subject": "ステータスで色が変わる付箋ボード", "author": "", "comments": "VBA source is included in the accompanying src folder."})

        def fmt(**options):
            return book.add_format({"font_name": FONT, "font_color": INK, "font_size": 11, **options})

        formats = {
            "paper": fmt(bg_color=PAPER),
            "title": fmt(font_size=27, bold=True, bg_color=PAPER, valign="vcenter"),
            "subtitle": fmt(font_size=11, font_color=MUTED, bg_color=PAPER, valign="vcenter"),
            "label": fmt(font_size=11, bold=True, bg_color=PAPER, valign="vcenter"),
            "input": fmt(bg_color="#FFFFFF", border=1, border_color="#C8CEC9", valign="top", text_wrap=True, num_format="@", locked=False),
            "text": fmt(num_format="@", valign="top", text_wrap=True),
            "date": fmt(num_format="yyyy/mm/dd hh:mm", valign="top"),
            "hint": fmt(bg_color=PAPER, font_color=MUTED, text_wrap=True, valign="top"),
            "head": fmt(bg_color=INK, font_color="#FFFFFF", bold=True, valign="vcenter"),
        }
        sheets = {}
        for name, code in SHEETS:
            sheet = book.add_worksheet(name)
            sheet.set_vba_name(code)
            sheet.hide_gridlines(2)
            sheet.set_zoom(85)
            sheet.set_default_row(18)
            sheet.set_tab_color(INK)
            sheet.protect("", {"autofilter": True})
            sheets[code] = sheet

        def background(sheet, last_row=45, last_col=56):
            sheet.set_column(0, last_col, 3)
            for row in range(last_row):
                sheet.set_row(row, 18)
                for col in range(last_col + 1):
                    sheet.write_blank(row, col, None, formats["paper"])

        def button(sheet, cell, label, macro, width=154, primary=False):
            sheet.insert_textbox(cell, label, {
                "width": width, "height": 36,
                "font": {"name": FONT, "size": 11, "color": "#FFFFFF" if primary else INK, "bold": True},
                "fill": {"color": INK if primary else "#FFFFFF"},
                "line": {"color": INK if primary else "#D5D7D6"},
                "align": {"horizontal": "center", "vertical": "middle"},
                "description": "macro:" + macro,
                "object_position": 3,
            })

        for code, title in (("ActiveBoard", "付箋タスクボード"), ("CompletedBoard", "完了した付箋")):
            sheet = sheets[code]
            background(sheet)
            sheet.merge_range("B2:BD3", title + ("  /  サンプル" if sample else ""), formats["title"])
            total = len([r for r in records if (r[4] == "完了") == (code == "CompletedBoard")])
            formula = '=COUNTIF(\'タスク一覧\'!E:E,"完了")' if code == "CompletedBoard" else '=COUNTA(\'タスク一覧\'!A7:A1048576)-COUNTIF(\'タスク一覧\'!E:E,"完了")'
            sheet.merge_range("B4:BD4", "", formats["subtitle"])
            sheet.write_formula("B4", formula + '&" 件の付箋  /  クリックして編集・ステータスを変更"', formats["subtitle"], f"{total} 件の付箋  /  クリックして編集・ステータスを変更")
            button(sheet, "B6", "＋ 新しい付箋", "NewTask", primary=True)
            button(sheet, "I6", "タスクボード" if code == "CompletedBoard" else "完了した付箋", "ShowBoard" if code == "CompletedBoard" else "ShowCompleted")
            button(sheet, "P6", "入力中の付箋", "ShowEditor")
            button(sheet, "W6", "再描画", "RefreshBoard", width=110)
            hint = "付箋をクリックして編集。ステータスを変更し［登録・更新］を押すと、色と配置が変わります。"
            if not records:
                hint = "［＋ 新しい付箋］から、最初のタスクを登録しましょう。" if code == "ActiveBoard" else "完了した付箋はここに移動します。内容・備考はそのまま残ります。"
            sheet.merge_range("B8:BD8", hint, formats["subtitle"])
            for lane in range(4):
                col = 1 + 14 * lane
                color = COLORS[4] if code == "CompletedBoard" else COLORS[lane]
                header = fmt(bg_color=color, bold=True, font_size=13, valign="vcenter", indent=1)
                label = "完了 / アーカイブ" if code == "CompletedBoard" else STATUSES[lane]
                sheet.merge_range(8, col, 9, col + 12, label, header)
            sheet.freeze_panes(10, 0)
            sheet.set_landscape()
            sheet.fit_to_pages(1, 0)
            sheet.set_paper(9)
            sheet.repeat_rows(0, 9)
            sheet.print_area("A1:BE45")

        counts = [0, 0, 0, 0, 0]
        for record in records:
            index = STATUSES.index(record[4])
            sheet = sheets["CompletedBoard" if index == 4 else "ActiveBoard"]
            lane = counts[index] % 4 if index == 4 else index
            slot = counts[index] // 4 if index == 4 else counts[index]
            counts[index] += 1
            text = (f"{record[4]}  /  {record[0]}\n{preview(record[1], 40)}\n\n"
                    f"内容\n{preview(record[2], 96)}\n\n備考・関連する実装\n{preview(record[3], 62)}\n\n"
                    "クリックして全文を表示・編集")
            sheet.insert_textbox(10, 1 + 14 * lane, text, {
                "width": 13 * 26, "height": 336, "y_offset": slot * 360,
                "font": {"name": FONT, "size": 11, "color": MUTED if index == 4 else INK},
                "fill": {"color": COLORS[index]}, "line": {"none": True},
                "align": {"vertical": "top"}, "margin": {"left": 18, "right": 18, "top": 16, "bottom": 16},
                "description": "task:" + record[0], "object_position": 3,
            })

        editor = sheets["TaskEditor"]
        background(editor, 39, 29)
        editor.set_column("A:A", 3)
        editor.set_column("B:P", 5)
        editor.set_column("Q:AD", 4)
        editor.set_zoom(95)
        editor.merge_range("B2:AD3", "付箋の編集", formats["title"])
        editor.merge_range("B4:F4", "新規", formats["label"])
        editor.merge_range("H4:P4", "登録・更新すると自動保存します", formats["subtitle"])
        editor.merge_range("B5:P5", "ステータス", formats["label"])
        editor.merge_range("B6:F6", "起票済み", formats["input"])
        editor.data_validation("B6", {"validate": "list", "source": "=StatusList", "error_type": "stop", "error_title": "ステータスを選択", "error_message": "一覧の5種類から選択してください。", "ignore_blank": False})
        editor.merge_range("H6:P6", "完了にすると「完了」シートへ移動", formats["subtitle"])
        for status, color in zip(STATUSES, COLORS):
            editor.conditional_format("B6:F6", {"type": "formula", "criteria": f'$B$6="{status}"', "format": fmt(bg_color=color)})
        editor.merge_range("B8:P8", "タイトル（必須・120文字以内）", formats["label"])
        editor.merge_range("B9:P11", "", formats["input"])
        editor.merge_range("B13:P13", "タスクの内容（8000文字以内）", formats["label"])
        editor.merge_range("B14:P21", "", formats["input"])
        editor.merge_range("B23:P23", "備考・関連する実装（8000文字以内）", formats["label"])
        editor.merge_range("B24:P31", "", formats["input"])
        for cell, maximum in (("B9", 120), ("B14", 8000), ("B24", 8000)):
            editor.data_validation(cell, {"validate": "length", "criteria": "<=", "value": maximum, "error_type": "stop", "error_message": f"{maximum}文字以内で入力してください。"})
        button(editor, "R6", "登録・更新", "SaveTask", width=196, primary=True)
        button(editor, "R9", "＋ 新しい付箋", "NewTask", width=196)
        button(editor, "R12", "タスクボードへ", "ShowBoard", width=196)
        editor.merge_range("R16:AC25", "入力のヒント\n\n・改行は Alt+Enter\n・F2 またはダブルクリックで編集\n・ステータスを選び［登録・更新］\n・長文は数式バーを広げて全文確認\n・入力中の内容はこのシートに残ります", formats["hint"])
        editor.merge_range("B33:P35", "付箋には内容・備考の冒頭を表示します。全文はこの編集欄とタスク一覧に保持します。\n登録したデータを破棄する操作はありません。完了後も編集・再開できます。", formats["hint"])
        editor.set_selection("B9")

        data = sheets["TaskData"]
        data.set_column("A:A", 13, formats["text"])
        data.set_column("B:B", 30, formats["text"])
        data.set_column("C:D", 54, formats["text"])
        data.set_column("E:E", 14, formats["text"])
        data.set_column("F:H", 21, formats["date"])
        data.merge_range("A2:H3", "タスク一覧", fmt(font_size=24, bold=True))
        data.merge_range("A4:H4", "行をダブルクリックすると付箋を編集できます。登録した内容の全文を保持する一覧です。", fmt(font_color=MUTED))
        for col, title in enumerate(("ID", "タイトル", "内容", "備考・関連する実装", "ステータス", "作成日時", "更新日時", "完了日時")):
            data.write(5, col, title, formats["head"])
        data.set_row(5, 28)
        for row_index, record in enumerate(records, 6):
            data.set_row(row_index, 64)
            for col, value in enumerate(record):
                if isinstance(value, datetime):
                    data.write_datetime(row_index, col, value, formats["date"])
                elif value is not None:
                    data.write_string(row_index, col, value, formats["text"])
        data.autofilter(5, 0, max(6, 5 + len(records)), 7)
        data.freeze_panes(6, 2)
        for status, color in zip(STATUSES, COLORS):
            data.conditional_format("A7:H1048576", {"type": "formula", "criteria": f'$E7="{status}"', "format": fmt(bg_color=color, font_color=MUTED if status == "完了" else INK)})

        guide = sheets["Guide"]
        guide.set_column("A:A", 3)
        guide.set_column("B:B", 23)
        guide.set_column("C:H", 14)
        guide.merge_range("B2:H3", "付箋タスク管理 / 使い方", fmt(font_size=24, bold=True))
        help_rows = [
            ("1. 最初に", "Windowsのデスクトップ版Excelでこの .xlsm を開きます。このファイルのマクロを有効にします。"),
            ("2. 付箋を作る", "［＋ 新しい付箋］→タイトル・内容・備考を入力→［登録・更新］。タスクボードに付箋が現れます。"),
            ("3. 進捗を変更", "付箋をクリック→ステータスを選択→［登録・更新］。付箋の色と列が自動で変わります。"),
            ("4. 完了する", "ステータスを「完了」にして［登録・更新］。灰色の付箋として「完了」シートへ移動します。"),
            ("5. 再開する", "「完了」シートの付箋をクリックし、別のステータスにして［登録・更新］。ボードへ戻ります。"),
            ("内容・備考", "改行はAlt+Enter。長文は編集欄または数式バーで確認できます。付箋上では冒頭のみ表示します。"),
            ("保存", "［登録・更新］でブックを自動保存します。入力途中の下書きを残して閉じる場合はCtrl+Sで保存します。"),
            ("色・文字", "起票済み：黄 / 対応中：青 / 対応済み：緑 / 連携済み：紫 / 完了：灰。フォントはMeiryo UIです。"),
            ("並び順", "各列は登録順です。完了シートも登録順に左から並びます。付箋の位置は自動整列されます。"),
            ("一覧・検索", "「タスク一覧」でフィルターやCtrl+Fを利用できます。行のダブルクリックで編集画面を開きます。"),
            ("ボタンが動かない", "デスクトップ版Excelで開き、マクロの許可を確認してください。ブラウザー版Excelではマクロは実行されません。"),
            ("画面の復旧", "［再描画］で一覧から付箋を作り直します。元データは「タスク一覧」に保持しています。"),
            ("入力上限", "タイトル120文字、内容・備考はそれぞれ8000文字（VBAのUTF-16文字数）。内容・備考は空欄でも登録できます。"),
            ("注意", "個人での編集を想定しています。同時編集には対応しません。ファイル形式は .xlsm のまま保存してください。"),
        ]
        for row, (label, description) in enumerate(help_rows, 5):
            guide.set_row(row, 46)
            guide.write(row, 1, label, fmt(bold=True, valign="top", bg_color=PAPER))
            guide.merge_range(row, 2, row, 7, description, fmt(text_wrap=True, valign="top", bg_color=PAPER))

        settings = sheets["Settings"]
        for row, status in enumerate(STATUSES):
            settings.write_string(row, 0, status)
        book.define_name("StatusList", "='設定'!$A$1:$A$5")
        settings.very_hidden()
        sheets["CompletedBoard"].set_tab_color(COLORS[4])
        editor.set_tab_color("#3D7A6A")
        sheets["ActiveBoard"].activate()
        sheets["ActiveBoard"].set_selection("B2")

    add_vba(path)
    bind_shapes(path)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--out-dir", type=Path, default=ROOT / ".build" / "output")
    args = parser.parse_args()
    # A rebuild never overwrites the user's working copy in the project root.
    for filename, sample in (("付箋タスク管理.xlsm", False), ("付箋タスク管理_サンプル.xlsm", True)):
        target = args.out_dir / filename
        build(target, sample=sample)
        print(f"Built: {target}")


if __name__ == "__main__":
    main()
