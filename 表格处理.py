import pandas as pd
import openpyxl
import tkinter as tk
from tkinter import filedialog, messagebox, Frame, Label, Button, StringVar, ttk
import os
import traceback
import threading
import re


class ExcelProcessor:
    def __init__(self):
        self.monthly_account = None
        self.billing_period = None
        self.order_count = None
        self.file_path = None
        self.total_fee = None
        self.total_discount = None
        self.total_payable = None
        self.total_claims = None

    def process_excel(self, file_path):
        """处理Excel文件，提取所需信息"""
        self.file_path = file_path

        try:
            # 使用openpyxl加载工作簿以处理合并单元格
            wb = openpyxl.load_workbook(file_path, data_only=True)

            # 1. 获取月结账号 (从账单总览sheet的J6:L6合并单元格)
            overview_sheet = wb["账单总览"]

            # 查找合并单元格
            for merged_range in overview_sheet.merged_cells.ranges:
                min_col, min_row, max_col, max_row = merged_range.min_col, merged_range.min_row, merged_range.max_col, merged_range.max_row

                # 检查J6:L6合并单元格 (J=10, L=12, 行号为6)
                if min_col == 10 and max_col == 12 and min_row == 6 and max_row == 6:
                    # 从合并单元格的左上角获取值
                    self.monthly_account = overview_sheet.cell(row=6, column=10).value
                    break

            # 2. 获取账单周期 (从账单总览sheet的D7:G7合并单元格)
            for merged_range in overview_sheet.merged_cells.ranges:
                min_col, min_row, max_col, max_row = merged_range.min_col, merged_range.min_row, merged_range.max_col, merged_range.max_row

                # 检查D7:G7合并单元格 (D=4, G=7, 行号为7)
                if min_col == 4 and max_col == 7 and min_row == 7 and max_row == 7:
                    # 从合并单元格的左上角获取值
                    self.billing_period = overview_sheet.cell(row=7, column=4).value
                    break

            # 3. 获取当月单量 - 修改为获取账单明细表的A列最后一个数值
            detail_sheet = wb["账单明细"]

            # 从A列提取数字值
            numeric_values = []

            for row in range(2, detail_sheet.max_row + 1):  # 从第2行开始，跳过表头
                cell_value = detail_sheet.cell(row=row, column=1).value

                if cell_value is not None:
                    # 处理不同类型的值
                    try:
                        # 如果是字符串，提取数字部分
                        if isinstance(cell_value, str):
                            numbers = re.findall(r'\d+', cell_value)
                            if numbers:
                                numeric_values.append(int(numbers[0]))
                        # 如果是数字，直接使用
                        elif isinstance(cell_value, (int, float)):
                            numeric_values.append(int(cell_value))
                    except (ValueError, TypeError):
                        continue

            # 如果找到了数值，取最大值作为单量
            if numeric_values:
                self.order_count = max(numeric_values)
            else:
                # 如果无法从A列获取有效数值，计算非空行数
                non_empty_rows = 0
                for row in range(2, detail_sheet.max_row + 1):
                    if detail_sheet.cell(row=row, column=1).value is not None:
                        non_empty_rows += 1
                self.order_count = non_empty_rows

            # 4. 获取账单明细中的汇总数据
            # 动态查找汇总数据所在行
            self._find_summary_values(detail_sheet)

            return {
                "月结账号": self.monthly_account,
                "账单周期": self.billing_period,
                "当月单量": self.order_count,
                "费用(元)": self.total_fee,
                "折扣/促销": self.total_discount,
                "应付金额": self.total_payable,
                "理赔费用合计": self.total_claims,
                "文件名": os.path.basename(file_path)
            }

        except Exception as e:
            error_message = f"处理文件时出错: {str(e)}\n{traceback.format_exc()}"
            raise Exception(error_message)

    def _is_valid_number(self, value):
        """检查值是否为有效数字"""
        if value is None:
            return False

        try:
            # 如果是字符串，尝试转换为浮点数
            if isinstance(value, str):
                # 去除可能的货币符号、括号等
                cleaned_value = re.sub(r'[^\d.-]', '', value)
                if cleaned_value:  # 确保不是空字符串
                    float(cleaned_value)
                    return True
                return False
            # 直接检查是否为数字类型
            return isinstance(value, (int, float))
        except (ValueError, TypeError):
            return False

    def _find_summary_values(self, detail_sheet):
        """查找汇总数据 - 通过搜索文本而不依赖于固定位置"""
        # 初始化默认值为None，表示未找到
        self.total_fee = None
        self.total_discount = None
        self.total_payable = None
        self.total_claims = None

        max_row = detail_sheet.max_row

        try:
            # 从表格的最后30行开始向上查找汇总数据
            search_start = max(2, max_row - 50)

            # 先查找包含"合计"或者"总计"的行
            total_rows = []
            for row in range(max_row, search_start, -1):
                found_total = False
                for col in range(1, 20):  # 扫描更多列以确保能找到关键词
                    cell_value = detail_sheet.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and (
                            "合计" in cell_value or "合 计" in cell_value or "总计" in cell_value):
                        total_rows.append(row)
                        found_total = True
                        break
                if found_total:
                    continue  # 继续搜索其他可能的汇总行

            if total_rows:
                # 找到了可能的汇总行，开始在这些行周围寻找关键字段
                for row in total_rows:
                    # 查找标题行 - 通常在数据的前几行
                    header_row = None
                    fee_col = None  # 初始化fee_col变量
                    discount_col = None
                    payable_col = None

                    for r in range(1, 30):  # 查找更多行以找到标题
                        for c in range(1, 20):
                            header_value = detail_sheet.cell(row=r, column=c).value
                            if header_value and isinstance(header_value, str):
                                # 查找与费用(元)、折扣/促销、应付金额相关的标题
                                if "费用" in header_value and ("元" in header_value or "¥" in header_value):
                                    fee_col = c
                                    header_row = r
                                if "折扣" in header_value or "促销" in header_value:
                                    discount_col = c
                                if "应付" in header_value and "金额" in header_value:
                                    payable_col = c

                    # 如果找到了标题行和至少一个相关列，根据标题位置查找对应的汇总值
                    if header_row and (fee_col or discount_col or payable_col):
                        if fee_col:
                            fee_value = detail_sheet.cell(row=row, column=fee_col).value
                            if self._is_valid_number(fee_value) and self.total_fee is None:
                                self.total_fee = fee_value
                        if discount_col:
                            discount_value = detail_sheet.cell(row=row, column=discount_col).value
                            if self._is_valid_number(discount_value) and self.total_discount is None:
                                self.total_discount = discount_value
                        if payable_col:
                            payable_value = detail_sheet.cell(row=row, column=payable_col).value
                            if self._is_valid_number(payable_value) and self.total_payable is None:
                                self.total_payable = payable_value
                    else:
                        # 如果没有找到标题行，则尝试在合计行查找
                        # 寻找合计行的前一行，检查是否有标题文本
                        prev_row = row - 1
                        if prev_row > 0:
                            for c in range(1, 20):
                                col_header = detail_sheet.cell(row=prev_row, column=c).value
                                if col_header and isinstance(col_header, str):
                                    if "费用" in col_header and ("元" in col_header or "¥" in col_header):
                                        fee_col = c
                                    if "折扣" in col_header or "促销" in col_header:
                                        discount_col = c
                                    if "应付" in col_header and "金额" in col_header:
                                        payable_col = c

                            # 根据找到的列获取对应的汇总值
                            if fee_col:
                                fee_value = detail_sheet.cell(row=row, column=fee_col).value
                                if self._is_valid_number(fee_value) and self.total_fee is None:
                                    self.total_fee = fee_value
                            if discount_col:
                                discount_value = detail_sheet.cell(row=row, column=discount_col).value
                                if self._is_valid_number(discount_value) and self.total_discount is None:
                                    self.total_discount = discount_value
                            if payable_col:
                                payable_value = detail_sheet.cell(row=row, column=payable_col).value
                                if self._is_valid_number(payable_value) and self.total_payable is None:
                                    self.total_payable = payable_value

                        # 如果仍然没有找到主要字段，则尝试从合计行向右查找数值
                        if not all([self.total_fee, self.total_payable]):
                            for right_col in range(1, 20):
                                right_value = detail_sheet.cell(row=row, column=right_col).value
                                if self._is_valid_number(right_value):
                                    # 找到了数值，根据位置依次分配
                                    if self.total_fee is None:
                                        self.total_fee = right_value
                                        continue
                                    if self.total_discount is None:
                                        self.total_discount = right_value
                                        continue
                                    if self.total_payable is None:
                                        self.total_payable = right_value
                                        break

            # 查找理赔费用合计 - 全新的方法，基于"理赔费用"标题和H列的合计值
            # 1. 查找"理赔费用"标题行
            claims_section_row = None
            for row in range(1, max_row + 1):
                for col in range(1, 20):
                    cell_value = detail_sheet.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and "理赔费用" in cell_value:
                        claims_section_row = row
                        break
                if claims_section_row:
                    break

            # 2. 如果找到了理赔费用区域标题行，查找其后的合计行和H列的值
            if claims_section_row:
                claims_total_row = None
                # 从理赔费用标题行开始向下搜索合计行
                for row in range(claims_section_row, min(claims_section_row + 30, max_row + 1)):
                    for col in range(1, 20):
                        cell_value = detail_sheet.cell(row=row, column=col).value
                        if cell_value and isinstance(cell_value, str) and (
                                "合计" in cell_value or "合 计" in cell_value):
                            claims_total_row = row
                            break
                    if claims_total_row:
                        break

                # 3. 如果找到了理赔费用区域的合计行，提取H列(=8)的值
                if claims_total_row:
                    claims_value = detail_sheet.cell(row=claims_total_row, column=8).value  # H列=8
                    if self._is_valid_number(claims_value):
                        self.total_claims = claims_value
                    else:
                        # 如果H列没有有效数值，尝试查找该行其他列的数值
                        for col in range(1, 20):
                            value = detail_sheet.cell(row=claims_total_row, column=col).value
                            if self._is_valid_number(value):
                                self.total_claims = value
                                break
            # 如果未找到理赔费用区域，理赔费用合计将保持为None

        except Exception as e:
            # 记录错误但不中断程序
            print(f"查找汇总数据时出错: {str(e)}")
            traceback.print_exc()


class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("账单数据提取工具")
        self.geometry("1100x600")  # 增加窗口宽度以适应更多列
        self.configure(bg="#f5f5f5")
        self.processor = ExcelProcessor()
        self.results = []
        self.is_processing = False

        self.create_widgets()

    def create_widgets(self):
        # 创建标题栏
        title_frame = Frame(self, bg="#4a7abc", height=60)
        title_frame.pack(fill=tk.X, pady=(0, 20))

        title_label = Label(title_frame, text="供销云仓账单数据提取工具",
                            font=("Microsoft YaHei UI", 16, "bold"), bg="#4a7abc", fg="white")
        title_label.pack(pady=15)

        # 创建按钮框架
        button_frame = Frame(self, bg="#f5f5f5")
        button_frame.pack(pady=10)

        # 设置按钮样式
        button_style = {"font": ("Microsoft YaHei UI", 10),
                        "bg": "#4a7abc", "fg": "white",
                        "activebackground": "#3a5a8c", "activeforeground": "white",
                        "width": 15, "height": 2, "bd": 0}

        # 导入按钮
        self.import_button = Button(button_frame, text="导入账单文件", command=self.import_excel, **button_style)
        self.import_button.pack(side=tk.LEFT, padx=10)

        # 导出按钮
        self.export_button = Button(button_frame, text="导出处理结果", command=self.export_results, **button_style)
        self.export_button.pack(side=tk.LEFT, padx=10)

        # 清除按钮
        self.clear_button = Button(button_frame, text="清除结果", command=self.clear_results, **button_style)
        self.clear_button.pack(side=tk.LEFT, padx=10)

        # 创建一个框架来容纳表格
        table_frame = Frame(self, bg="#f5f5f5")
        table_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # 创建表格 - 添加新的列
        columns = ("文件名", "月结账号", "账单周期", "当月单量", "费用(元)", "折扣/促销", "应付金额", "理赔费用合计")
        self.result_tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=15)

        # 设置列宽
        self.result_tree.column("文件名", width=180)
        self.result_tree.column("月结账号", width=120)
        self.result_tree.column("账单周期", width=180)
        self.result_tree.column("当月单量", width=80)
        self.result_tree.column("费用(元)", width=100)
        self.result_tree.column("折扣/促销", width=100)
        self.result_tree.column("应付金额", width=100)
        self.result_tree.column("理赔费用合计", width=100)

        # 设置列标题
        for col in columns:
            self.result_tree.heading(col, text=col)

        # 添加垂直滚动条
        scrollbar = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.result_tree.yview)
        self.result_tree.configure(yscroll=scrollbar.set)

        # 添加水平滚动条
        h_scrollbar = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.result_tree.xview)
        self.result_tree.configure(xscroll=h_scrollbar.set)

        # 放置表格和滚动条
        self.result_tree.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 进度条
        self.progress_frame = Frame(self, bg="#f5f5f5")
        self.progress_frame.pack(fill=tk.X, padx=20, pady=5)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.progress_frame, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill=tk.X)

        # 状态栏
        self.status_var = StringVar()
        self.status_var.set("准备就绪")
        self.status_bar = Label(self, textvariable=self.status_var, bd=1, relief=tk.SUNKEN,
                                anchor=tk.W, font=("Microsoft YaHei UI", 9), bg="#f0f0f0")
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 设置表格样式
        style = ttk.Style()
        style.configure("Treeview", font=("Microsoft YaHei UI", 9), rowheight=25)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"))

    def process_files_thread(self, file_paths):
        """在单独的线程中处理文件"""
        processed_count = 0
        failed_files = []

        total_files = len(file_paths)
        for i, file_path in enumerate(file_paths):
            try:
                # 更新进度
                progress = (i + 1) / total_files * 100
                self.progress_var.set(progress)

                # 更新状态栏
                file_name = os.path.basename(file_path)
                self.status_var.set(f"正在处理: {file_name} ({i + 1}/{total_files})")
                self.update_idletasks()  # 确保UI更新

                # 处理Excel文件
                result = self.processor.process_excel(file_path)

                # 使用after方法来安全地更新UI
                self.after(10, self.add_result_to_tree, result)
                self.update_idletasks()  # 确保UI更新

                processed_count += 1

            except Exception as e:
                failed_files.append((file_path, str(e)))
                print(f"处理文件 {os.path.basename(file_path)} 失败: {str(e)}")

        # 完成处理后的操作
        self.after(10, self.processing_complete, processed_count, failed_files)

    def add_result_to_tree(self, result):
        """将结果添加到表格中"""
        try:
            self.results.append(result)
            self.result_tree.insert("", tk.END, values=(
                result["文件名"],
                result["月结账号"],
                result["账单周期"],
                result["当月单量"],
                result["费用(元)"],
                result["折扣/促销"],
                result["应付金额"],
                result["理赔费用合计"]
            ))
        except Exception as e:
            print(f"添加结果到表格失败: {str(e)}")

    def processing_complete(self, processed_count, failed_files):
        """处理完成后的操作"""
        self.progress_var.set(0)
        self.is_processing = False

        # 重新启用按钮
        self.import_button.config(state=tk.NORMAL)
        self.export_button.config(state=tk.NORMAL)
        self.clear_button.config(state=tk.NORMAL)

        # 更新状态栏
        if failed_files:
            self.status_var.set(f"已完成处理 {processed_count} 个文件，{len(failed_files)} 个文件失败")
            failed_message = "以下文件处理失败:\n\n"
            # 仅显示前10个失败的文件，以避免消息框过大
            for i, (file_path, error) in enumerate(failed_files[:10]):
                short_error = error.split('\n')[0]  # 只显示错误的第一行
                failed_message += f"{i + 1}. {os.path.basename(file_path)}: {short_error}\n"

            if len(failed_files) > 10:
                failed_message += f"\n...等共 {len(failed_files)} 个文件失败"

            messagebox.showwarning("处理警告", failed_message)
        else:
            self.status_var.set(f"已完成处理 {processed_count} 个文件")

    def import_excel(self):
        """导入Excel文件并处理"""
        if self.is_processing:
            messagebox.showinfo("提示", "正在处理文件，请稍候...")
            return

        file_paths = filedialog.askopenfilenames(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )

        if not file_paths:
            return

        # 设置处理状态
        self.is_processing = True

        # 禁用按钮
        self.import_button.config(state=tk.DISABLED)
        self.export_button.config(state=tk.DISABLED)
        self.clear_button.config(state=tk.DISABLED)

        # 启动处理线程
        threading.Thread(target=self.process_files_thread, args=(file_paths,), daemon=True).start()

    def export_results(self):
        """导出处理结果到Excel文件"""
        if self.is_processing:
            messagebox.showinfo("提示", "正在处理文件，请稍候...")
            return

        if not self.results:
            messagebox.showinfo("提示", "没有可导出的结果")
            return

        file_path = filedialog.asksaveasfilename(
            title="保存结果",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not file_path:
            return

        try:
            df = pd.DataFrame(self.results)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("成功", f"结果已成功导出到:\n{file_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出结果时出错: {str(e)}")

    def clear_results(self):
        """清除结果"""
        if self.is_processing:
            messagebox.showinfo("提示", "正在处理文件，请稍候...")
            return

        if not self.results:
            return

        # 清空表格
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

        # 清空结果列表
        self.results = []

        # 更新状态栏
        self.status_var.set("已清除所有结果")


if __name__ == "__main__":
    app = Application()
    app.mainloop()