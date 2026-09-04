"""
Excel 多列数据比对工具 V2.0
功能：图形界面比较两个 Excel 文件中指定列的数据差异。
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import sys
import os
import threading
from datetime import datetime


class ExcelComparator:
    """Excel比对核心类"""
    
    def __init__(self):
        self.df1 = None
        self.df2 = None
        self.result_data = None
        
    def load_file(self, file_path):
        """加载Excel文件"""
        try:
            if file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, engine='openpyxl')
            elif file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='xlrd')
            else:
                raise ValueError("不支持的文件格式")
            
            df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)
            return df
        except Exception as e:
            raise Exception(f"加载文件失败: {str(e)}")
    
    def compare(self, df1, df2, columns, ignore_case=False, ignore_whitespace=True):
        """比对两个DataFrame"""
        self.df1 = df1.copy()
        self.df2 = df2.copy()
        
        missing_cols = []
        for col in columns:
            if col not in self.df1.columns:
                missing_cols.append(f"'{col}' (文件1)")
            if col not in self.df2.columns:
                missing_cols.append(f"'{col}' (文件2)")
        
        if missing_cols:
            raise Exception(f"以下列不存在:\n" + "\n".join(missing_cols))
        
        def preprocess_data(df):
            df_processed = df[columns].copy()
            for col in columns:
                df_processed[col] = df_processed[col].astype(str)
                if ignore_whitespace:
                    df_processed[col] = df_processed[col].str.strip()
                if ignore_case:
                    df_processed[col] = df_processed[col].str.lower()
            return df_processed
        
        df1_processed = preprocess_data(self.df1)
        df2_processed = preprocess_data(self.df2)
        
        df1_processed['_source'] = '文件1'
        df2_processed['_source'] = '文件2'
        
        combined = pd.concat([df1_processed, df2_processed], ignore_index=True)
        duplicate_mask = combined.duplicated(subset=columns, keep=False)
        unique_rows = combined[~duplicate_mask].copy()
        
        self.result_data = {
            'only_in_file1': [],
            'only_in_file2': [],
            'total_diff': len(unique_rows)
        }
        
        if unique_rows.empty:
            return "✅ 两个文件的数据完全一致！"
        
        only_file1 = unique_rows[unique_rows['_source'] == '文件1'].drop(columns=['_source'])
        only_file2 = unique_rows[unique_rows['_source'] == '文件2'].drop(columns=['_source'])
        
        if not only_file1.empty:
            idx1 = only_file1.index
            self.result_data['only_in_file1'] = self.df1.loc[idx1].copy()
        
        if not only_file2.empty:
            idx2 = only_file2.index
            self.result_data['only_in_file2'] = self.df2.loc[idx2].copy()
        
        report_parts = []
        report_parts.append(f"📊 比对完成！共发现 {len(unique_rows)} 条差异记录\n")
        report_parts.append("=" * 60 + "\n")
        
        if not only_file1.empty:
            report_parts.append(f"🔴 仅在文件1中存在的行: {len(only_file1)} 条\n")
            report_parts.append("-" * 40 + "\n")
            report_parts.append(only_file1.to_string(index=True))
            report_parts.append("\n\n")
        
        if not only_file2.empty:
            report_parts.append(f"🟢 仅在文件2中存在的行: {len(only_file2)} 条\n")
            report_parts.append("-" * 40 + "\n")
            report_parts.append(only_file2.to_string(index=True))
            report_parts.append("\n")
        
        return "".join(report_parts)
    
    def export_result(self, file_path, format='excel'):
        """导出比对结果"""
        if not self.result_data or self.result_data['total_diff'] == 0:
            raise Exception("没有差异数据可导出")
        
        if format == 'excel':
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                if self.result_data['only_in_file1']:
                    self.result_data['only_in_file1'].to_excel(
                        writer, sheet_name='仅在文件1中', index=False
                    )
                if self.result_data['only_in_file2']:
                    self.result_data['only_in_file2'].to_excel(
                        writer, sheet_name='仅在文件2中', index=False
                    )
                
                summary = pd.DataFrame({
                    '信息': ['比对时间', '总差异数', '文件1独有', '文件2独有'],
                    '值': [
                        datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                        self.result_data['total_diff'],
                        len(self.result_data['only_in_file1']),
                        len(self.result_data['only_in_file2'])
                    ]
                })
                summary.to_excel(writer, sheet_name='汇总', index=False)
        elif format == 'csv':
            combined = pd.concat([
                self.result_data['only_in_file1'].assign(来源='文件1'),
                self.result_data['only_in_file2'].assign(来源='文件2')
            ], ignore_index=True)
            combined.to_csv(file_path, index=False, encoding='utf-8-sig')


class ExcelCompareGUI:
    """图形界面主类"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel 多列数据比对工具 V2.0")
        self.root.geometry("750x650")
        self.root.minsize(700, 600)
        
        self.comparator = ExcelComparator()
        self._create_widgets()
        self._center_window()
    
    def _center_window(self):
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def _create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        title_label = tk.Label(
            main_frame, 
            text="📊 Excel 多列数据比对工具", 
            font=("Microsoft YaHei", 16, "bold"),
            fg="#2c3e50"
        )
        title_label.pack(pady=(0, 15))
        
        file_frame = ttk.LabelFrame(main_frame, text="📁 文件选择", padding="10")
        file_frame.pack(fill='x', pady=(0, 10))
        
        row1 = ttk.Frame(file_frame)
        row1.pack(fill='x', pady=2)
        ttk.Label(row1, text="文件1:", width=8).pack(side='left')
        self.entry_file1 = ttk.Entry(row1)
        self.entry_file1.pack(side='left', fill='x', expand=True, padx=(5, 5))
        ttk.Button(row1, text="浏览...", command=lambda: self._browse_file(self.entry_file1)).pack(side='right')
        
        row2 = ttk.Frame(file_frame)
        row2.pack(fill='x', pady=2)
        ttk.Label(row2, text="文件2:", width=8).pack(side='left')
        self.entry_file2 = ttk.Entry(row2)
        self.entry_file2.pack(side='left', fill='x', expand=True, padx=(5, 5))
        ttk.Button(row2, text="浏览...", command=lambda: self._browse_file(self.entry_file2)).pack(side='right')
        
        col_frame = ttk.LabelFrame(main_frame, text="🔍 比对设置", padding="10")
        col_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(col_frame, text="要比对的列名（用英文逗号分隔）:").pack(anchor='w')
        self.entry_cols = ttk.Entry(col_frame, font=("Consolas", 10))
        self.entry_cols.pack(fill='x', pady=(5, 5))
        
        tip_frame = ttk.Frame(col_frame)
        tip_frame.pack(fill='x')
        ttk.Label(tip_frame, text="💡 示例: 姓名,身份证号,手机号", font=("Microsoft YaHei", 9), foreground="#7f8c8d").pack(side='left')
        
        options_frame = ttk.Frame(col_frame)
        options_frame.pack(fill='x', pady=(10, 0))
        
        self.ignore_case = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            options_frame, 
            text="忽略大小写", 
            variable=self.ignore_case
        ).pack(side='left', padx=(0, 15))
        
        self.ignore_whitespace = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame, 
            text="忽略前后空格", 
            variable=self.ignore_whitespace
        ).pack(side='left')
        
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(0, 10))
        
        self.btn_compare = ttk.Button(
            btn_frame, 
            text="🚀 开始比对", 
            command=self._run_comparison
        )
        self.btn_compare.pack(side='left', padx=(0, 10))
        
        self.btn_export = ttk.Button(
            btn_frame, 
            text="💾 导出结果", 
            command=self._export_result,
            state='disabled'
        )
        self.btn_export.pack(side='left', padx=(0, 10))
        
        self.btn_clear = ttk.Button(
            btn_frame, 
            text="🗑️ 清空", 
            command=self._clear_all
        )
        self.btn_clear.pack(side='left')
        
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.pack(fill='x', pady=(0, 10))
        
        result_frame = ttk.LabelFrame(main_frame, text="📋 比对结果", padding="10")
        result_frame.pack(fill='both', expand=True)
        
        self.status_label = ttk.Label(result_frame, text="就绪", foreground="#7f8c8d")
        self.status_label.pack(anchor='w', pady=(0, 5))
        
        self.text_result = scrolledtext.ScrolledText(
            result_frame, 
            height=18, 
            font=("Consolas", 10),
            wrap=tk.NONE
        )
        self.text_result.pack(fill='both', expand=True)
        
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill='x', pady=(10, 0))
        ttk.Label(
            footer_frame, 
            text="© 2026 ExcelCompare Tool | 支持 .xlsx / .xls 格式",
            font=("Microsoft YaHei", 8),
            foreground="#95a5a6"
        ).pack(side='right')
    
    def _browse_file(self, entry):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[
                ("Excel文件", "*.xlsx *.xls"),
                ("Excel 2007+", "*.xlsx"),
                ("Excel 97-2003", "*.xls"),
                ("所有文件", "*.*")
            ]
        )
        if file_path:
            entry.delete(0, tk.END)
            entry.insert(0, file_path)
    
    def _run_comparison(self):
        file1 = self.entry_file1.get().strip()
        file2 = self.entry_file2.get().strip()
        cols_text = self.entry_cols.get().strip()
        
        if not file1 or not file2:
            messagebox.showwarning("提示", "请选择两个Excel文件！")
            return
        
        if not os.path.exists(file1):
            messagebox.showerror("错误", f"文件不存在:\n{file1}")
            return
        
        if not os.path.exists(file2):
            messagebox.showerror("错误", f"文件不存在:\n{file2}")
            return
        
        if not cols_text:
            messagebox.showwarning("提示", "请输入要比对的列名！")
            return
        
        columns = [col.strip() for col in cols_text.split(',') if col.strip()]
        if not columns:
            messagebox.showwarning("提示", "请输入有效的列名！")
            return
        
        self.btn_compare.config(state='disabled')
        self.btn_export.config(state='disabled')
        self.progress.start()
        self.text_result.delete(1.0, tk.END)
        self.text_result.insert(tk.END, "⏳ 正在加载和比对数据，请稍候...\n")
        self.status_label.config(text="正在处理...", foreground="#f39c12")
        
        def compare_thread():
            try:
                df1 = self.comparator.load_file(file1)
                df2 = self.comparator.load_file(file2)
                
                self.root.after(0, lambda: self.text_result.insert(tk.END, 
                    f"✅ 文件1加载成功: {len(df1)} 行\n"
                    f"✅ 文件2加载成功: {len(df2)} 行\n\n"
                ))
                
                result = self.comparator.compare(
                    df1, df2, columns,
                    ignore_case=self.ignore_case.get(),
                    ignore_whitespace=self.ignore_whitespace.get()
                )
                
                self.root.after(0, lambda: self._on_compare_complete(result))
                
            except Exception as e:
                self.root.after(0, lambda: self._on_compare_error(str(e)))
        
        threading.Thread(target=compare_thread, daemon=True).start()
    
    def _on_compare_complete(self, result):
        self.progress.stop()
        self.btn_compare.config(state='normal')
        self.status_label.config(text="比对完成", foreground="#27ae60")
        
        self.text_result.delete(1.0, tk.END)
        self.text_result.insert(tk.END, result)
        
        if self.comparator.result_data and self.comparator.result_data['total_diff'] > 0:
            self.btn_export.config(state='normal')
        else:
            self.btn_export.config(state='disabled')
    
    def _on_compare_error(self, error_msg):
        self.progress.stop()
        self.btn_compare.config(state='normal')
        self.btn_export.config(state='disabled')
        self.status_label.config(text="发生错误", foreground="#e74c3c")
        
        self.text_result.delete(1.0, tk.END)
        self.text_result.insert(tk.END, f"❌ 错误:\n{error_msg}")
        messagebox.showerror("比对错误", error_msg)
    
    def _export_result(self):
        if not self.comparator.result_data:
            messagebox.showinfo("提示", "没有数据可导出")
            return
        
        file_path = filedialog.asksaveasfilename(
            title="保存比对结果",
            defaultextension=".xlsx",
            filetypes=[
                ("Excel文件", "*.xlsx"),
                ("CSV文件", "*.csv"),
                ("所有文件", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith('.csv'):
                self.comparator.export_result(file_path, format='csv')
            else:
                if not file_path.endswith('.xlsx'):
                    file_path += '.xlsx'
                self.comparator.export_result(file_path, format='excel')
            
            messagebox.showinfo("成功", f"结果已导出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
    
    def _clear_all(self):
        self.entry_file1.delete(0, tk.END)
        self.entry_file2.delete(0, tk.END)
        self.entry_cols.delete(0, tk.END)
        self.text_result.delete(1.0, tk.END)
        self.btn_export.config(state='disabled')
        self.status_label.config(text="就绪", foreground="#7f8c8d")
        self.comparator.result_data = None
    
    def run(self):
        self.root.mainloop()


def main():
    try:
        app = ExcelCompareGUI()
        app.run()
    except Exception as e:
        messagebox.showerror("严重错误", f"程序启动失败:\n{e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
