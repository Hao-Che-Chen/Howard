import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import filedialog
import pandas as pd
import os

class TimeCalculatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("配件工時試算表")
        self.root.geometry("1000x750")

        # 設定檔案路徑
        desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
        file_path = os.path.join(desktop_path, "高雄物料清單-配件工時表.xlsx")

        try:
            if os.path.exists(file_path):
                self.df = pd.read_excel(file_path, header=None)
                self.setup_data_structures()
                self.create_widgets()
            else:
                messagebox.showerror("錯誤", f"找不到檔案: {file_path}")
                 # 讓使用者重新選擇檔案
                file1_path = filedialog.askopenfilename(title="請選擇 Excel 檔案", filetypes=[("Excel files", "*.xlsx;*.xls")])
                if os.path.basename(file1_path) == "高雄物料清單-配件工時表.xlsx" :
                    self.df = pd.read_excel(file1_path, header=None)
                    self.setup_data_structures()
                    self.create_widgets()

                if not file1_path:  # 使用者取消選擇
                    self.root.destroy()
                    return
                

        except Exception as e:
            messagebox.showerror("檔案錯誤", f"讀取檔案失敗: {str(e)}")
            self.root.destroy()

    def setup_data_structures(self):
        # 初始化數據結構
        self.accessory_times = dict(zip(self.df.iloc[:, 0], self.df.iloc[:, 1]))
        self.furniture_times = dict(zip(self.df.iloc[:, 2].dropna(), self.df.iloc[:, 3].dropna()))
        self.additional_information_on_times = dict(zip(self.df.iloc[:, 11], self.df.iloc[:, 12].dropna()))
        
        packing_option_titles = self.df.iloc[0, 4:11].values  # 提取第1行的第5至第11列（即配件選項名稱）

        self.packing_options = {
            packing_option_titles[0]: int(self.df.iloc[1, 4]),         #'桌板 拆+封箱'
            packing_option_titles[1]: int(self.df.iloc[1, 5]),         #'桌腳 拆+封箱'
            packing_option_titles[2]: int(self.df.iloc[1, 6]),         #'桌腳 EPE裁切'
            packing_option_titles[3]: int(self.df.iloc[1, 7]),         #'次級品桌腳功能檢查'
            packing_option_titles[4]: int(self.df.iloc[1, 8]),         #'纏包膜'
            packing_option_titles[5]: int(self.df.iloc[1, 9]),         #'打包帶'
            packing_option_titles[6]: int(self.df.iloc[1, 10])         #'配件使用另外箱子寄送'
        }

        internal_process_times_titles = self.df.iloc[0, 13:21].values  # 提取第1行的第14至第21列（即流程選項名稱）

        self.internal_process_times = {
            internal_process_times_titles[0]: int(self.df.iloc[1, 13]),         #倉庫雜收客供料
            internal_process_times_titles[1]: int(self.df.iloc[1, 14]),         #生管開重工工單
            internal_process_times_titles[2]: int(self.df.iloc[1, 15]),         #跑簽呈(廠長、經管、倉庫、生管)
            internal_process_times_titles[3]: int(self.df.iloc[1, 16]),         #通知製造單位提供人力
            internal_process_times_titles[4]: int(self.df.iloc[1, 17]),         #倉庫說明訂單內容
            internal_process_times_titles[5]: int(self.df.iloc[1, 18]),         #出貨倉庫拍照回郵件、通知交管
            internal_process_times_titles[6]: int(self.df.iloc[1, 19]),         #產品拉到碼頭
            internal_process_times_titles[7]: int(self.df.iloc[1, 20])          #請款倉庫整理該筆備貨工時
        }

    def create_widgets(self):
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # 左側框架
        left_frame = ttk.Frame(main_frame)
        left_frame.grid(row=0, column=0, sticky='nsew')

        # 配件表格
        self.create_accessory_table(left_frame)
        # 家具表格
        self.create_furniture_table(left_frame)
        # 額外工時
        self.create_additional_information_on_section(left_frame)
        # 勾選項
        self.create_checkboxes(left_frame)
        # 結果區域
        self.create_results_section(left_frame)
        #額外增加工時
        self.create_additional_information_on_section(left_frame)

        # 右側流程顯示
        right_frame = ttk.LabelFrame(main_frame, text="廠內流程")
        right_frame.grid(row=0, column=1, sticky='nsew')
        self.internal_process_text = tk.Text(right_frame, height=5, width=55, font=('Arial', 10))
        self.internal_process_text.pack(fill=tk.BOTH, expand=True)
        self.display_internal_process_times()

    def create_accessory_table(self, parent):
        frame = ttk.LabelFrame(parent, text=self.df.iloc[0, 0])  #配件選擇
        frame.grid(row=0, column=0, sticky='nsew', padx=5, pady=5)
        
        headers = ["No.", "品項", "數量/PCs"]
        for col, text in enumerate(headers):
            ttk.Label(frame, text=text).grid(row=0, column=col, padx=5, pady=2)

        self.accessory_combos = []
        self.accessory_qty = []
        for i in range(16):
            ttk.Label(frame, text=str(i+1)).grid(row=i+1, column=0)
            combo = ttk.Combobox(frame, values=self.df.iloc[1:, 0].dropna().tolist(), width=20)
            combo.grid(row=i+1, column=1, padx=5)
            qty = ttk.Entry(frame, width=8)
            qty.grid(row=i+1, column=2, padx=5)
            self.accessory_combos.append(combo)
            self.accessory_qty.append(qty)

    def create_furniture_table(self, parent):
        frame = ttk.LabelFrame(parent, text=self.df.iloc[0, 2])
        frame.grid(row=0, column=1, sticky='nsew', padx=5, pady=5)
                
        headers = ["No.", "品項", "數量/PCs"]
        for col, text in enumerate(headers):
            ttk.Label(frame, text=text).grid(row=0, column=col, padx=5, pady=2)

        self.furniture_combos = []
        self.furniture_qty = []
        for i in range(6):
            ttk.Label(frame, text=str(i+1)).grid(row=i+1, column=0)
            combo = ttk.Combobox(frame, values=self.df.iloc[1:, 2].dropna().tolist(), width=20)
            combo.grid(row=i+1, column=1, padx=5)
            qty = ttk.Entry(frame, width=8)
            qty.grid(row=i+1, column=2, padx=5)
            self.furniture_combos.append(combo)
            self.furniture_qty.append(qty)


    def create_additional_information_on_section(self, parent):
        frame = ttk.LabelFrame(parent, text=self.df.iloc[0, 11])
        frame.grid(row=1, column=1, sticky='nsew', padx=5, pady=5)

        headers = ["No.", "品項", "數量/PCs"]
        for col, text in enumerate(headers):
            ttk.Label(frame, text=text).grid(row=0, column=col, padx=5, pady=2)

        self.additional_information_on_combos = []
        self.additional_information_on_qty = []
        for i in range(4):
            ttk.Label(frame, text=str(i+1)).grid(row=i+1, column=0)
            combo = ttk.Combobox(frame, values=self.df.iloc[1:, 11].dropna().tolist(), width=20)
            combo.grid(row=i+1, column=1, padx=5)
            qty = ttk.Entry(frame, width=8)
            qty.grid(row=i+1, column=2, padx=5)
            self.additional_information_on_combos.append(combo)
            self.additional_information_on_qty.append(qty)

    def create_checkboxes(self, parent):
        frame = ttk.LabelFrame(parent, text="備貨及包裝方式")
        frame.grid(row=1, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        self.check_vars = {}

        # 備貨方式選項（例如：桌板拆封、桌腳拆封等）
        ttk.Label(frame, text="備貨方式選擇").grid(row=0, column=0, sticky='w')

        # 動態從 self.packing_options 中提取備貨方式選項
        stock_options = list(self.packing_options.keys())[:4]  # 假設備貨方式為前4項
        for i, opt in enumerate(stock_options):
            var = tk.BooleanVar()
            ttk.Checkbutton(frame, text=opt, variable=var).grid(row=i+1, column=0, sticky='w')
            self.check_vars[opt] = var
        
        # 包裝方式選項（例如：纏包膜、打包帶等）
        ttk.Label(frame, text="包裝方式選擇").grid(row=0, column=1, sticky='w')

        # 動態從 self.packing_options 中提取包裝方式選項
        pack_options = list(self.packing_options.keys())[4:]  # 假設包裝方式為後3項
        for i, opt in enumerate(pack_options):
            var = tk.BooleanVar()
            ttk.Checkbutton(frame, text=opt, variable=var).grid(row=i+1, column=1, sticky='w')
            self.check_vars[opt] = var

    def create_results_section(self, parent):
        frame = ttk.Frame(parent)
        frame.grid(row=2, column=0, columnspan=2, sticky='ew', pady=10)

        ttk.Button(frame, text="工時試算", command=self.calculate_time).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame, text="清空選項", command=self.clear_all_inputs).pack(side=tk.LEFT, padx=5)

        self.total_time_var = tk.StringVar(value="總花費工時: 00小時 : 00分鐘 : 00秒")
        ttk.Label(frame, textvariable=self.total_time_var, foreground="blue", 
                 font=('微軟正黑體', 14, 'bold')).pack(side=tk.LEFT, padx=20)

        console_frame = ttk.LabelFrame(parent, text="試算清單")
        console_frame.grid(row=3, column=0, columnspan=2, sticky='nsew', padx=5, pady=5)
        self.console_text = tk.Text(console_frame, height=16, width=55, font=('Arial', 12))
        self.console_text.pack(fill=tk.BOTH, expand=True)

    def clear_all_inputs(self):
        # 清空所有輸入組件
        for combo in self.accessory_combos:
            combo.set('')
        for entry in self.accessory_qty:
            entry.delete(0, tk.END)

        for combo in self.furniture_combos:
            combo.set('')
        for entry in self.furniture_qty:
            entry.delete(0, tk.END)

        for combo in self.additional_information_on_combos:
            combo.set('')
        for entry in self.additional_information_on_qty:
            entry.delete(0, tk.END)

        for var in self.check_vars.values():
            var.set(False)

        self.console_text.delete(1.0, tk.END)
        self.total_time_var.set("總花費工時: 00小時 : 00分鐘 : 00秒")

    def seconds_to_hms(self, total_seconds):
        hours = int(total_seconds // 3600)
        minutes = int((total_seconds % 3600) // 60)
        seconds = int(total_seconds % 60)
        return f"{hours:02d}小時 : {minutes:02d}分鐘 : {seconds:02d}秒"

    def display_internal_process_times(self):
        self.internal_process_text.delete(1.0, tk.END)
        for process, time in self.internal_process_times.items():
            self.internal_process_text.insert(tk.END, f"{process}: {self.seconds_to_hms(time)}\n")
        total = sum(self.internal_process_times.values())
        self.internal_process_text.insert(tk.END, f"\n廠內流程工時: {self.seconds_to_hms(total)}")

    def calculate_time(self):
        total_seconds = 0
        console_lines = []

        # 配件計算
        for i in range(16):
            accessory = self.accessory_combos[i].get()
            qty = self.accessory_qty[i].get()
            if accessory and qty:
                try:
                    qty = int(qty)
                    time = int(self.accessory_times.get(accessory, 0))
                    total_seconds += time * qty
                    console_lines.append(f"配件 {accessory} x{qty} => {self.seconds_to_hms(time * qty)}")
                except:
                    console_lines.append(f"錯誤: {accessory} 數量無效")

        # 家具計算
        for i in range(6):
            furniture = self.furniture_combos[i].get()
            qty = self.furniture_qty[i].get()
            if furniture and qty:
                try:
                    qty = int(qty)
                    base_time = int(self.furniture_times.get(furniture, 0))

                    # 如果選擇的是 "次級品桌腳 4F"，則進行額外的功能檢查處理
                    if furniture == '次級品桌腳 4F':
                        packing_check = self.check_vars.get('次級品桌腳功能檢查', False)
                        if packing_check.get():  # 如果勾選了「次級品桌腳功能檢查」
                            base_time = self.packing_options.get('次級品桌腳功能檢查', 1)  # 檢查工時

                    if qty > 1:
                        total_seconds += base_time + (qty - 1) * 300  # 如果數量大於1，每多一個家具加300秒
                    else:
                        total_seconds += base_time
                    console_lines.append(f"{furniture} x{qty} => {self.seconds_to_hms(base_time + (qty - 1) * 300)}")
                except:
                    console_lines.append(f"錯誤: {furniture} 數量無效")

        # 額外工時
        for i in range(4):
            item = self.additional_information_on_combos[i].get()
            qty = self.additional_information_on_qty[i].get()
            if item and qty:
                try:
                    qty = int(qty)
                    time = int(self.additional_information_on_times.get(item, 0))
                    total_seconds += time * qty
                    console_lines.append(f"額外工時 {item} x{qty} => {self.seconds_to_hms(time * qty)}")
                except:
                    console_lines.append(f"錯誤: {item} 數量無效")

        # 勾選項 - 計算勾選的備貨選項工時
        for opt, var in self.check_vars.items():
            if var.get():
                time = self.packing_options.get(opt, 0)
                
                # 如果勾選的是 "次級品桌腳功能檢查"，需要乘以對應的 "次級品桌腳 4F" 數量
                if opt == "次級品桌腳功能檢查":
                    # 查找對應的家具項目 "次級品桌腳 4F"
                    for i in range(6):
                        furniture = self.furniture_combos[i].get()
                        if furniture == "次級品桌腳 4F":
                            qty = self.furniture_qty[i].get()
                            if qty:
                                try:
                                    qty = int(qty)
                                    time *= qty  # 工時乘以數量
                                    console_lines.append(f"[★] {opt} => {self.seconds_to_hms(time)}")
                                except:
                                    console_lines.append(f"錯誤: {opt} 數量無效")
                            break
                    else:
                        console_lines.append(f"[★] {opt} => {self.seconds_to_hms(time)}")
                total_seconds += time


        # 廠內流程
        total_seconds += sum(self.internal_process_times.values())

        # 更新顯示
        self.total_time_var.set(f"總花費工時: {self.seconds_to_hms(total_seconds)}")
        self.console_text.delete(1.0, tk.END)
        self.console_text.insert(tk.END, "\n".join(console_lines))

if __name__ == "__main__":
    root = tk.Tk()
    app = TimeCalculatorApp(root)
    root.mainloop()