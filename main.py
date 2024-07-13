import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox

from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LAParams, LTTextContainer
from pdfminer.pdfinterp import PDFPageInterpreter, PDFResourceManager
from pdfminer.pdfpage import PDFPage
from pypdf import PdfReader, PdfWriter


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


class App:
    def __init__(self, root):
        self.root = root
        self.root.geometry("400x150")
        self.root.title("畜産生産実績一覧表 to CSV")

        self.label = tk.Label(root, text="処理するフォルダを選択してください")
        self.label.pack(pady=20)

        self.select_button = tk.Button(
            root, text="フォルダ選択", command=self.select_folder
        )
        self.select_button.pack(pady=10)

        self.progress_label = tk.Label(root, text="")
        self.progress_label.pack(pady=10)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        if folder_path:
            self.process_files(folder_path)

    def process_files(self, folder_path):
        files = os.listdir(folder_path)
        total_files = len(files)
        processed_files = 0

        for file in files:
            processed_files += 1
            self.progress_label.config(
                text=f"進行状況: {processed_files}/{total_files}"
            )
            self.root.update_idletasks()

            if not file.endswith(".pdf"):
                continue

            file_path = os.path.join(folder_path, file)
            temp_file = os.path.join(folder_path, file + "temp")
            remove_copy_protections(file_path, temp_file)
            text_and_coordinates = extract_text_with_positions(temp_file)
            os.remove(temp_file)
            formatted_data = format_data(text_and_coordinates)

            output_file = os.path.join(
                folder_path,
                text_and_coordinates[45][0].rstrip().replace("(cid:8443)", "崎")
                + ".csv",
            )

            with open(
                output_file,
                mode="w",
                errors="ignore",
                newline="",
            ) as f:
                csv.writer(f).writerows(formatted_data)
        messagebox.showinfo("完了", f"{total_files}個のファイルの処理が完了しました")
        self.progress_label.config(text="")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
