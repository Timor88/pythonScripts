import fitz
import sys
from collections import defaultdict
# 新增Tkinter相關導入
import tkinter as tk
from tkinter import filedialog, scrolledtext, messagebox

# 標準紙張尺寸（單位：pt，1pt=1/72英吋）
STANDARD_SIZES = {
    'A0': (2384, 3370),
    'A1': (1684, 2384),
    'A2': (1191, 1684),
    'A3': (842, 1191),
    'A4': (595, 842),
    'A5': (420, 595),
}

# 尺寸容差（允許誤差範圍，單位：pt）
TOLERANCE = 10

def get_size_name(width, height):
    for name, (w, h) in STANDARD_SIZES.items():
        if abs(width - w) <= TOLERANCE and abs(height - h) <= TOLERANCE:
            return name
        if abs(width - h) <= TOLERANCE and abs(height - w) <= TOLERANCE:
            return name
    return f"特殊尺寸({int(width)}x{int(height)})"

def group_pages_by_size(pdf_path):
    doc = fitz.open(pdf_path)
    size_pages = defaultdict(list)
    for i, page in enumerate(doc, 1):
        rect = page.rect
        width, height = rect.width, rect.height
        size_name = get_size_name(width, height)
        size_pages[size_name].append(i)
    return size_pages

def format_page_ranges(pages):
    if not pages:
        return ''
    ranges = []
    start = prev = pages[0]
    for num in pages[1:]:
        if num == prev + 1:
            prev = num
        else:
            if start == prev:
                ranges.append(f"{start}")
            else:
                ranges.append(f"{start}-{prev}")
            start = prev = num
    if start == prev:
        ranges.append(f"{start}")
    else:
        ranges.append(f"{start}-{prev}")
    return ','.join(ranges)

def analyze_pdf(pdf_path):
    size_pages = group_pages_by_size(pdf_path)
    result_lines = []
    for size, pages in sorted(size_pages.items()):
        page_str = format_page_ranges(pages)
        result_lines.append(f"{size}共{len(pages)}頁：{page_str}")
    return '\n'.join(result_lines)

# ====== Tkinter GUI 部分 ======
def run_gui():
    def select_file():
        file_path = filedialog.askopenfilename(
            filetypes=[('PDF Files', '*.pdf')],
            title='選擇PDF文件'
        )
        if file_path:
            entry_file.delete(0, tk.END)
            entry_file.insert(0, file_path)

    def analyze_action():
        pdf_path = entry_file.get()
        if not pdf_path:
            messagebox.showwarning('提示', '請先選擇PDF文件！')
            return
        try:
            result = analyze_pdf(pdf_path)
            text_result.config(state='normal')
            text_result.delete(1.0, tk.END)
            text_result.insert(tk.END, result)
            text_result.config(state='disabled')
        except Exception as e:
            messagebox.showerror('錯誤', f'分析失敗：{e}')

    root = tk.Tk()
    root.title('PDF頁面尺寸分類工具')
    root.geometry('600x400')

    frm = tk.Frame(root)
    frm.pack(pady=10)

    tk.Label(frm, text='PDF文件:').pack(side=tk.LEFT)
    entry_file = tk.Entry(frm, width=50)
    entry_file.pack(side=tk.LEFT, padx=5)
    tk.Button(frm, text='選擇文件', command=select_file).pack(side=tk.LEFT)
    tk.Button(root, text='開始分析', command=analyze_action).pack(pady=5)

    text_result = scrolledtext.ScrolledText(root, width=70, height=18, state='disabled')
    text_result.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    root.mainloop()

# ====== 主程序入口 ======
if __name__ == "__main__":
    if len(sys.argv) == 2 and sys.argv[1].lower().endswith('.pdf'):
        # 命令行模式
        print(analyze_pdf(sys.argv[1]))
    else:
        # GUI模式
        run_gui()
