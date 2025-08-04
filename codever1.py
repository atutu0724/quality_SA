import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

result_df = None  # グローバル変数で抽出結果を保持

def load_csv_file():
    global result_df

    file_path = filedialog.askopenfilename(
        title="CSVファイルを選択",
        filetypes=[("CSV ファイル", "*.csv"), ("すべてのファイル", "*.*")]
    )

    if not file_path:
        return

    try:
        # エンコーディング自動判定（UTF-8 → cp932）
        try:
            df = pd.read_csv(file_path, encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(file_path, encoding="cp932")

        # 抽出設定
        start_row = 70
        condition_col = "Unnamed: 9"
        extract_cols = (
            [9, 35,37,130,134,104,105,80,81,82,83 ] +
            list(range(156, 161)) +
            [150,151] +
            list(range(115, 118)))
   

        if condition_col not in df.columns:
            messagebox.showerror("エラー", f"列名 '{condition_col}' が見つかりません。")
            return

        # フィルタ処理
        target_df = df.loc[start_row:]
        filtered_rows = target_df[target_df[condition_col].notna()]
        filtered_result = filtered_rows.iloc[:, extract_cols]

        # 強制的に含める行
        fixed_row_indices = [1, 2, 3]
        valid_fixed_rows = df.loc[df.index.intersection(fixed_row_indices)]
        fixed_result = valid_fixed_rows.iloc[:, extract_cols]

        # 結合：強制行を先頭に & 重複削除
        result_df = pd.concat([fixed_result, filtered_result], ignore_index=True).drop_duplicates()

        show_table(result_df)

    except Exception as e:
        import traceback
        print(traceback.format_exc())
        messagebox.showerror("読み込みエラー", f"CSVの読み込みに失敗しました：\n{str(e)}")

def show_table(result):
    for widget in frame.winfo_children():
        widget.destroy()

    tree = ttk.Treeview(frame, columns=list(result.columns), show="headings")
    scroll_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
    scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

    scroll_y.pack(side="right", fill="y")
    scroll_x.pack(side="bottom", fill="x")
    tree.pack(expand=True, fill="both")

    for col in result.columns:
        tree.heading(col, text=str(col))
        max_len = max(result[col].astype(str).str.len().max(), len(str(col)))
        width = max_len * 10 + 30
        tree.column(col, width=width, anchor="center", stretch=False)

    tree.tag_configure("evenrow", background="#e6f2ff")

    for idx, (_, row) in enumerate(result.iterrows()):
        values = ["" if pd.isna(v) else v for v in row]  # NaNを空白に置換
        tag = "evenrow" if idx % 2 == 0 else ""
        tree.insert("", "end", values=values, tags=(tag,))

def save_result():
    global result_df
    if result_df is None:
        messagebox.showwarning("警告", "まだデータがありません。先にCSVを読み込んでください。")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        title="保存先を指定してください"
    )

    if not save_path:
        return

    try:
        result_df.to_excel(save_path, index=False)
        messagebox.showinfo("成功", f"保存しました：\n{save_path}")
    except Exception as e:
        messagebox.showerror("保存エラー", f"保存に失敗しました：\n{str(e)}")

# --- GUI構築 ---
root = tk.Tk()
root.title("📂 CSV 抽出ビューア")
root.geometry("1400x720")

style = ttk.Style()
style.theme_use("clam")
style.configure("Treeview", font=('Helvetica', 11), rowheight=28, borderwidth=1, relief="solid")
style.configure("Treeview.Heading", font=('Helvetica', 12, 'bold'), anchor="center")

btn_frame = tk.Frame(root)
btn_frame.pack(pady=10)

tk.Button(btn_frame, text="📂 CSVファイルを読み込む", font=("Helvetica", 12), command=load_csv_file).pack(side="left", padx=10)
tk.Button(btn_frame, text="💾 抽出結果を保存", font=("Helvetica", 12), command=save_result).pack(side="left", padx=10)

frame = tk.Frame(root, padx=10, pady=10)
frame.pack(expand=True, fill="both")

root.mainloop()

