import csv
import logging
import os
import tkinter as tk
from tempfile import TemporaryDirectory
from tkinter import filedialog

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextContainer
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage
from pypdf import PdfReader, PdfWriter


def ask_foldername() -> str:
    # ルートウィンドウ作成
    root = tk.Tk()
    # ルートウィンドウの非表示
    root.withdraw()

    return filedialog.askdirectory()


def remove_copy_protections(source: str, destination: str) -> None:
    """
    コピープロテクトを解除する関数。プロテクトを解除したPDFファイルを出力する

    Args:
    - source (str): 元のPDFファイルのパス
    - destination (str): コピープロテクトを解除したPDFの出力先パス

    Returns:
    - None
    """
    # 元のPDFファイルを読み込む
    pdf_reader = PdfReader(source)

    # 空のPDFファイルを作成
    pdf_writer = PdfWriter()

    # 全てのページをコピーする
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        pdf_writer.add_page(page)

    # 新しいPDFを保存
    with open(destination, "wb") as output_file:
        pdf_writer.write(output_file)


def extract_text_with_positions(
    pdf_file_path: str,
) -> list[list[str, float, float]]:
    """
    PDFファイルからテキストと座標情報を抽出する関数

    Args:
    - pdf_file_path (str): PDFファイルのパス

    Returns:
    - list[list[str, float, float]]: テキストと座標情報のリスト
    """

    manager = PDFResourceManager()

    lprms = LAParams(line_margin=0.01, word_margin=0.1, char_margin=0.1)

    text_and_coordinates = []

    with open(pdf_file_path, "rb") as input:
        with PDFPageAggregator(manager, laparams=lprms) as device:
            # PDFPageInterpreterオブジェクトの取得
            iprtr = PDFPageInterpreter(manager, device)

            # ページごとで処理を実行
            page_num = 100
            for page in PDFPage.get_pages(input):
                iprtr.process_page(page)
                # ページ内の各テキストのレイアウト
                layouts = device.get_result()
                page_num -= 1
                for layout in layouts:
                    if isinstance(layout, LTTextContainer):
                        # テキストと座標情報をリストに追加
                        text_and_coordinates.append(
                            [
                                layout.get_text(),
                                round(layout.x0, 2),
                                round(layout.y0, 2) + page_num * 1000,
                            ]
                        )

    text_and_coordinates.sort(key=lambda x: x[2], reverse=True)
    return text_and_coordinates


def format_data(data: list[list[str, float, float]]) -> list[list[str]]:
    """
    特定の座標にあるテキストのみを抽出する関数

    Args:
    - data (list[list[str, float, float]]): テキストと座標を含む入力データ。

    Returns:
    - list[list[str]]: 必要な情報のみを持つ、整形されたテーブル。
    """
    targets_coordinates_x0 = [
        385.92,  # "出荷年月日"
        66.24,  # "品目"
        198,  # "取引先"
        447.12,  # "個体識別番号"
        498.96,  # "生年月日"
        561.6,  # "［本体］品代", "［消費税］品代"
        622.44,  # "［本体］心肝（内臓）", "［消費税］心肝（内臓）"
        681.84,  # "［本体］原皮", "［消費税］原皮"
    ]

    flg = -1
    head_coordinate_y0 = -1
    table = []

    for item in data:
        text, x0, y0 = item
        if x0 == 385.92:
            head_coordinate_y0 = y0
            flg = 0
            row = []

        if flg in [0, 1] and x0 in targets_coordinates_x0 and y0 <= head_coordinate_y0:
            word = (
                text.replace(" ", "")
                .replace(",", "")
                .replace("\n", "/")
                .replace("\u3000", "")
                .replace("【", "")
                .replace("】", "")
                .replace("生年月日", "")
                .rstrip("/")
            )
            if word != "":
                row.append(word)

        if flg in [0, 1] and x0 == 681.84 and y0 < head_coordinate_y0:
            flg += 1

        if flg == 2:
            flg = -1
            # データの整形と追加
            row = [len(table) + 1] + row[2:6] + [row[0]] + row[6:]
            price_sum = sum(map(int, row[6:]))
            row.append(str(price_sum))
            table.append(row)

    col_label = [
        "番号",
        "品目",
        "取引先",
        "個体識別番号",
        "生年月日",
        "出荷年月日",
        "［本体］品代",
        "［本体］心肝（内臓）",
        "［本体］原皮",
        "［消費税］品代",
        "［消費税］心肝（内臓）",
        "［消費税］原皮",
        "［合計］",
    ]

    return [col_label] + table


def main():
    # ユーザーにフォルダを選択させる
    target_folder = ask_foldername()
    # print("target_folder is :", target_folder)

    # 一時ディレクトリを作成して、PDF ファイルの処理を行う
    with TemporaryDirectory() as temp_dir:
        # 選択されたフォルダ内の各ファイルについて処理を実行する
        for filename in os.listdir(target_folder):
            # 拡張子が ".pdf" のファイルに対して処理を行う
            if not filename.endswith(".pdf"):
                continue

            # ファイルパスの設定
            target_file = os.path.join(target_folder, filename)
            output_file = target_file.replace(".pdf", ".csv")

            # PDF ファイルからコピープロテクトを取り除くために一時ディレクトリにコピーする
            temp_pdf = os.path.join(temp_dir, filename)
            remove_copy_protections(target_file, temp_pdf)

            # PDF からテキストと座標情報を取得する
            text_and_coordinates = extract_text_with_positions(temp_pdf)

            # 取得したテキストと座標情報を整形する
            formatted_data = format_data(text_and_coordinates)

            # 整形されたデータを CSV ファイルに書き込む
            with open(
                output_file,
                mode="w",
                errors="ignore",
                newline="",
            ) as f:
                csv.writer(f).writerows(formatted_data)
    # print("process has completed.")


if __name__ == "__main__":
    # ログを作成
    try:
        main()
    except:
        log_file = "logger.log"
        if not os.path.isfile(log_file):
            f = open(log_file, "x")
            f.close()

        logging.basicConfig(filename=log_file)
        logging.exception("What is doing when exception happens.")
