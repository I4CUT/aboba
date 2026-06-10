import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
from pathlib import Path
import json

SETTINGS_FILE = "settings.json"

def load_settings():
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except:
            pass
    return {}

def save_settings(settings):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f)

settings = load_settings()

def read_file(file_path):
    if file_path.endswith(".csv"):
        with open(file_path, encoding='utf-8') as f:
            lines = f.readlines()
        return pd.DataFrame(lines)
    else:
        return pd.read_excel(file_path, header=None)

def select_files():
    file_paths = filedialog.askopenfilenames(
        title="Выберите Excel или CSV файлы",
        filetypes=[("Excel и CSV файлы", "*.xlsx *.csv")],
        initialdir=settings.get("last_open_dir", os.getcwd())
    )
    if file_paths:
        settings["last_open_dir"] = str(Path(file_paths[0]).parent)
        save_settings(settings)
        output_dir = filedialog.askdirectory(title="Выберите папку для сохранения результатов",
                                             initialdir=settings.get("last_save_dir", os.getcwd()))
        if output_dir:
            settings["last_save_dir"] = output_dir
            save_settings(settings)
            process_files(list(file_paths), Path(output_dir))

def process_files(file_paths, output_dir):
    first_data_lines = []

    for file_path in file_paths:
        df_raw = read_file(file_path)
        for i, row in df_raw.iterrows():
            val = row[0] if not isinstance(row, str) else row
            if isinstance(val, str) and val.count(',') == 4:
                try:
                    numbers = [float(x.strip()) for x in val.split(',')]
                    first_data_lines.append((file_path, i, numbers))
                    break
                except ValueError:
                    continue

    if not first_data_lines:
        messagebox.showerror("Ошибка", "Не удалось найти корректные данные ни в одном файле.")
        return

    preview_window = tk.Toplevel()
    preview_window.title("Предпросмотр и названия столбцов")

    tk.Label(preview_window, text="Введите названия столбцов:").pack(pady=5)

    default_headers = ["freq", "e_r", "e_im", "m_r", "m_im"]
    entries = []
    frame = tk.Frame(preview_window)
    frame.pack()

    for i in range(5):
        entry = tk.Entry(frame, width=15)
        entry.insert(0, default_headers[i])
        entry.grid(row=0, column=i, padx=5)
        entries.append(entry)

    tk.Label(preview_window, text="Пример строки данных:").pack(pady=5)
    sample = first_data_lines[0][2]
    sample_frame = tk.Frame(preview_window)
    sample_frame.pack()

    for val in sample:
        tk.Label(sample_frame, text=str(val), width=15, relief="ridge").pack(side=tk.LEFT, padx=5)

    tan_frame = tk.Frame(preview_window)
    tan_frame.pack(pady=5)
    tk.Label(tan_frame, text="Название колонки tan δ (e_im/e_r):").pack(side=tk.LEFT, padx=5)
    tan_entry = tk.Entry(tan_frame, width=20)
    tan_entry.insert(0, "tan_delta")
    tan_entry.pack(side=tk.LEFT, padx=5)

    def on_confirm():
        headers = [e.get().strip() for e in entries]
        if not all(headers):
            messagebox.showwarning("Предупреждение", "Пожалуйста, заполните все названия столбцов.")
            return
        tan_col_name = tan_entry.get().strip() or "tan_delta"
        preview_window.destroy()
        format_and_save(file_paths, headers, output_dir, tan_col_name)

    tk.Button(preview_window, text="Подтвердить", command=on_confirm).pack(pady=10)

def format_and_save(file_paths, headers, output_dir, tan_col_name="tan_delta"):
    output_dir.mkdir(exist_ok=True)
    saved_paths = []

    e_r_col  = headers[1]
    e_im_col = headers[2]

    for file_path in file_paths:
        df_raw = read_file(file_path)

        data_lines = []
        for i, row in df_raw.iterrows():
            val = row[0] if not isinstance(row, str) else row
            if isinstance(val, str) and val.count(',') == 4:
                try:
                    numbers = [float(x.strip()) for x in val.split(',')]
                    data_lines.append(numbers)
                except ValueError:
                    continue

        df_clean = pd.DataFrame(data_lines, columns=headers)

        df_clean[tan_col_name] = df_clean.apply(
            lambda row: row[e_im_col] / row[e_r_col] if row[e_r_col] != 0 else float('nan'),
            axis=1
        )

        filename = Path(file_path).stem + ".xlsx"
        save_path = output_dir / filename
        df_clean.to_excel(save_path, index=False)
        saved_paths.append(save_path)

    messagebox.showinfo("Готово", f"Файлы сохранены в папке: {output_dir.resolve()}")


def generate_report_txt(df_avg, value_cols, freq_col, file_names, save_dir, report_stem="averaging_report"):
    """Генерирует текстовый отчёт со статистикой по усреднённым данным."""
    from datetime import datetime

    freq_min = df_avg[freq_col].min()
    freq_max = df_avg[freq_col].max()
    n_points = len(df_avg)

    stats = {}
    for col in value_cols:
        s = df_avg[col].dropna()
        stats[col] = {
            "mean": s.mean(),
            "var":  s.var(),
            "std":  s.std(),
            "min":  s.min(),
            "max":  s.max(),
        }

    def fmt(v):
        try:
            return f"{v:.6g}"
        except Exception:
            return "N/A"

    SEP  = "=" * 72
    SEP2 = "-" * 72
    lines = []
    lines.append(SEP)
    lines.append("  ОТЧЁТ ПО УСРЕДНЕНИЮ ИЗМЕРЕНИЙ")
    lines.append(f"  Дата формирования: {datetime.now().strftime('%d.%m.%Y  %H:%M:%S')}")
    lines.append(SEP)
    lines.append("")

    lines.append("1. ПАРАМЕТРЫ УСРЕДНЕНИЯ")
    lines.append(SEP2)
    lines.append(f"  Количество файлов  : {len(file_names)}")
    lines.append(f"  Диапазон частот    : {fmt(freq_min)} — {fmt(freq_max)}")
    lines.append(f"  Количество точек   : {n_points}")
    lines.append("")

    lines.append("2. СВОДНАЯ СТАТИСТИКА ПО ВСЕМУ ДИАПАЗОНУ ЧАСТОТ")
    lines.append("   (рассчитана по усреднённым значениям для каждой частотной точки)")
    lines.append(SEP2)

    col_w = max(len(c) for c in stats) + 2
    lines.append(f"  {'Параметр':<{col_w}}  {'Среднее':>14}  {'Дисперсия':>14}  {'Ст. откл.':>14}  {'Мин.':>14}  {'Макс.':>14}")
    lines.append("  " + "-" * (col_w + 14 * 5 + 10))
    for col, s in stats.items():
        lines.append(
            f"  {col:<{col_w}}  {fmt(s['mean']):>14}  {fmt(s['var']):>14}"
            f"  {fmt(s['std']):>14}  {fmt(s['min']):>14}  {fmt(s['max']):>14}"
        )
    lines.append("")

    lines.append("3. ИСХОДНЫЕ ФАЙЛЫ")
    lines.append(SEP2)
    for i, name in enumerate(file_names, 1):
        lines.append(f"  {i:>3}.  {name}")
    lines.append("")
    lines.append(SEP)

    report_path = str(Path(save_dir) / f"{report_stem}.txt")
    with open(report_path, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
    return report_path


def average_files_window():
    initial_dir = settings.get("last_open_dir", os.getcwd())
    selected_files = filedialog.askopenfilenames(
        title="Выберите отформатированные Excel-файлы для усреднения",
        filetypes=[("Excel", "*.xlsx")],
        initialdir=initial_dir
    )
    if not selected_files:
        return

    settings["last_open_dir"] = str(Path(selected_files[0]).parent)
    save_settings(settings)

    dfs = []
    for f in selected_files:
        try:
            df = pd.read_excel(f)
            dfs.append(df)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось прочитать файл {Path(f).name}:\n{e}")
            return

    first_cols = list(dfs[0].columns)
    for i, df in enumerate(dfs[1:], 2):
        if list(df.columns) != first_cols:
            messagebox.showerror(
                "Ошибка",
                f"Файл №{i} имеет отличающиеся столбцы.\nОжидалось: {first_cols}\nПолучено: {list(df.columns)}"
            )
            return

    freq_col   = first_cols[0]
    value_cols = first_cols[1:]

    combined = pd.concat(dfs, ignore_index=True)
    df_avg = combined.groupby(freq_col, sort=False)[value_cols].mean().reset_index()
    df_avg = df_avg.sort_values(by=freq_col).reset_index(drop=True)

    e_r_col  = first_cols[1]
    e_im_col = first_cols[2]
    tan_candidates = [c for c in value_cols if "tan" in c.lower()]
    if tan_candidates and e_r_col in df_avg.columns and e_im_col in df_avg.columns:
        tan_col = tan_candidates[0]
        df_avg[tan_col] = df_avg.apply(
            lambda row: row[e_im_col] / row[e_r_col] if row[e_r_col] != 0 else float('nan'),
            axis=1
        )

    prev_win = tk.Toplevel()
    prev_win.title("Усреднение: предпросмотр результата")
    prev_win.geometry("750x470")

    tk.Label(prev_win, text=f"Усреднено файлов: {len(dfs)}   |   Точек по частоте: {len(df_avg)}",
             font=("Arial", 10, "bold")).pack(pady=5)

    frame_table = tk.Frame(prev_win)
    frame_table.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

    tree = ttk.Treeview(frame_table, columns=first_cols, show="headings", height=15)
    for col in first_cols:
        tree.heading(col, text=col)
        tree.column(col, width=max(80, len(col) * 10), anchor="center")

    for _, row in df_avg.head(20).iterrows():
        values = [f"{row[c]:.6g}" if pd.notna(row[c]) else "NaN" for c in first_cols]
        tree.insert("", tk.END, values=values)

    scrollbar = ttk.Scrollbar(frame_table, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

    if len(df_avg) > 20:
        tk.Label(prev_win, text=f"(показаны первые 20 из {len(df_avg)} строк)",
                 font=("Arial", 8), fg="gray").pack()

    report_var = tk.BooleanVar(value=True)
    tk.Checkbutton(prev_win, text="Сформировать текстовый отчёт (.txt) со статистикой",
                   variable=report_var).pack(pady=4)

    def save_result():
        save_path = filedialog.asksaveasfilename(
            title="Сохранить усреднённый файл",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=settings.get("last_save_dir", os.getcwd()),
            initialfile="averaged.xlsx"
        )
        if not save_path:
            return
        settings["last_save_dir"] = str(Path(save_path).parent)
        save_settings(settings)
        df_avg.to_excel(save_path, index=False)

        msg = f"Файл сохранён:\n{save_path}"

        if report_var.get():
            try:
                file_names_list = [Path(f).name for f in selected_files]
                report_path = generate_report_txt(
                    df_avg, value_cols, freq_col,
                    file_names_list, Path(save_path).parent,
                    report_stem=Path(save_path).stem
                )
                msg += f"\n\nОтчёт сохранён:\n{report_path}"
            except Exception as e:
                msg += f"\n\nОшибка при генерации отчёта:\n{e}"

        messagebox.showinfo("Готово", msg)
        prev_win.destroy()

    tk.Button(prev_win, text="Сохранить результат", command=save_result, width=30).pack(pady=10)


def main():
    root = tk.Tk()
    root.title("Форматирование и усреднение диэлектрических данных")
    root.geometry("400x200")

    tk.Label(root, text="Выберите действие:", font=("Arial", 12)).pack(pady=20)

    tk.Button(root, text="Форматировать файлы", command=select_files, width=30).pack(pady=10)
    tk.Button(root, text="Усреднение файлов",   command=average_files_window, width=30).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
