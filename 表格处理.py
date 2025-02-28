import streamlit as st
import pandas as pd
import openpyxl
import os
import traceback
import re
import io
import tempfile
import base64


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

    def process_excel(self, file_bytes, file_name):
        """处理Excel文件，提取所需信息"""
        self.file_path = file_name  # 使用文件名代替完整路径

        try:
            # 创建临时文件以使openpyxl能够处理
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
                temp_file.write(file_bytes)
                temp_path = temp_file.name

            # 使用openpyxl加载工作簿以处理合并单元格
            wb = openpyxl.load_workbook(temp_path, data_only=True)

            # 1. 获取月结账号 (从账单总览sheet的J6:L6合并单元格)
            try:
                overview_sheet = wb["账单总览"]
            except KeyError:
                # 如果没有找到账单总览，尝试找其他可能的sheet名
                sheet_names = wb.sheetnames
                overview_sheet = None
                for name in sheet_names:
                    if "总览" in name or "概览" in name:
                        overview_sheet = wb[name]
                        break
                
                if overview_sheet is None:
                    # 如果仍然找不到，使用第一个sheet
                    overview_sheet = wb[sheet_names[0]]

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

            # 3. 获取当月单量 - 修改为统计服务字段为运费的行数
            try:
                detail_sheet = wb["账单明细"]
            except KeyError:
                # 如果没有找到账单明细，尝试找其他可能的sheet名
                sheet_names = wb.sheetnames
                detail_sheet = None
                for name in sheet_names:
                    if "明细" in name or "详情" in name:
                        detail_sheet = wb[name]
                        break
                
                if detail_sheet is None:
                    # 如果仍然找不到，使用第一个不是总览的sheet
                    for name in sheet_names:
                        if name != overview_sheet.title:
                            detail_sheet = wb[name]
                            break

            # 查找标题行
            header_row = None
            for row in range(1, 10):  # 在前10行中查找标题
                for col in range(1, 15):  # 在前15列中查找
                    cell_value = detail_sheet.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and "服务" in cell_value:
                        header_row = row
                        break
                if header_row:
                    break
            
            # 直接使用N列（索引为14）查找运费
            if header_row:
                freight_count = 0
                for row in range(header_row + 1, detail_sheet.max_row + 1):
                    # N列的索引为14（因为列索引从1开始，对应到代码是14）
                    service_value = detail_sheet.cell(row=row, column=14).value
                    if service_value and isinstance(service_value, str) and "运费" in service_value:
                        freight_count += 1
                
                if freight_count > 0:
                    self.order_count = freight_count
                else:
                    # 如果没有找到运费相关行，回退到计算非空行数
                    non_empty_rows = 0
                    for row in range(header_row + 1, detail_sheet.max_row + 1):
                        if detail_sheet.cell(row=row, column=1).value is not None:
                            non_empty_rows += 1
                    self.order_count = non_empty_rows
            else:
                # 如果没有找到服务列，回退到原来的方法
                numeric_values = []
                for row in range(2, detail_sheet.max_row + 1):
                    cell_value = detail_sheet.cell(row=row, column=1).value
                    if cell_value is not None:
                        try:
                            if isinstance(cell_value, str):
                                numbers = re.findall(r'\d+', cell_value)
                                if numbers:
                                    numeric_values.append(int(numbers[0]))
                            elif isinstance(cell_value, (int, float)):
                                numeric_values.append(int(cell_value))
                        except (ValueError, TypeError):
                            continue

                if numeric_values:
                    self.order_count = max(numeric_values)
                else:
                    non_empty_rows = 0
                    for row in range(2, detail_sheet.max_row + 1):
                        if detail_sheet.cell(row=row, column=1).value is not None:
                            non_empty_rows += 1
                    self.order_count = non_empty_rows

            # 4. 获取账单明细中的汇总数据
            # 动态查找汇总数据所在行
            self._find_summary_values(detail_sheet)

            # 删除临时文件
            try:
                os.unlink(temp_path)
            except:
                pass

            return {
                "月结账号": self.monthly_account,
                "账单周期": self.billing_period,
                "当月单量": self.order_count,
                "费用(元)": self.total_fee,
                "折扣/促销": self.total_discount,
                "应付金额": self.total_payable,
                "理赔费用合计": self.total_claims,
                "文件名": file_name
            }

        except Exception as e:
            # 确保清理临时文件
            try:
                if 'temp_path' in locals():
                    os.unlink(temp_path)
            except:
                pass
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

            # 查找理赔费用合计 - 首先检查是否存在理赔费用相关单元格
            has_claims = False
            for row in range(1, max_row + 1):
                for col in range(1, 20):
                    cell_value = detail_sheet.cell(row=row, column=col).value
                    if cell_value and isinstance(cell_value, str) and "理赔" in cell_value:
                        has_claims = True
                        break
                if has_claims:
                    break
            
            # 如果存在理赔费用相关单元格，则在H列中查找最小的负值
            if has_claims:
                h_column_values = []
                for row in range(2, max_row + 1):  # 从第2行开始，跳过表头
                    cell_value = detail_sheet.cell(row=row, column=8).value  # H列 = 8
                    if self._is_valid_number(cell_value):
                        # 如果是字符串，尝试转换为浮点数
                        if isinstance(cell_value, str):
                            try:
                                cleaned_value = re.sub(r'[^\d.-]', '', cell_value)
                                if cleaned_value:  # 确保不是空字符串
                                    h_column_values.append(float(cleaned_value))
                            except ValueError:
                                pass
                        else:
                            h_column_values.append(float(cell_value))

                # 从H列值中找到最小的负值作为理赔费用
                negative_values = [val for val in h_column_values if val < 0]
                if negative_values:
                    self.total_claims = min(negative_values)  # 取最小的负值
                else:
                    self.total_claims = None  # 没有找到负值，理赔费用为空
            else:
                # 如果不存在理赔费用相关单元格，则设置为None
                self.total_claims = None

        except Exception as e:
            # 记录错误但不中断程序
            print(f"查找汇总数据时出错: {str(e)}")
            traceback.print_exc()


def get_table_download_link(df):
    """生成一个下载链接，允许下载DataFrame作为Excel文件"""
    # 将DataFrame转换为Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_data = output.getvalue()
    
    # 使用base64编码
    b64 = base64.b64encode(excel_data).decode()
    
    # 创建下载链接
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="账单数据提取结果.xlsx" class="download-btn">📥 下载Excel文件</a>'
    return href


def main():
    st.set_page_config(
        page_title="供销云仓账单数据提取工具",
        page_icon="📊",
        layout="wide"
    )

    # 自定义CSS样式
    st.markdown("""
    <style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1E88E5;
        margin-bottom: 1rem;
        text-align: center;
    }
    .sub-header {
        font-size: 1.5rem;
        color: #424242;
        margin-bottom: 1rem;
    }
    .success-message {
        padding: 1rem;
        background-color: #E8F5E9;
        border-left: 5px solid #4CAF50;
        margin-bottom: 1rem;
    }
    .info-box {
        padding: 1rem;
        background-color: #E3F2FD;
        border-radius: 5px;
        margin-bottom: 1rem;
    }
    .download-btn {
        text-decoration: none;
        padding: 0.5rem 1rem;
        background-color: #1E88E5;
        color: white;
        border-radius: 4px;
        display: inline-block;
        margin: 1rem 0;
        font-weight: bold;
    }
    .download-btn:hover {
        background-color: #1565C0;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown("<div class='main-header'>供销云仓账单数据提取工具</div>", unsafe_allow_html=True)
    st.markdown("---")

    # 侧边栏 - 用于上传文件和显示操作状态
    with st.sidebar:
        st.markdown("<div class='sub-header'>操作面板</div>", unsafe_allow_html=True)
        
        st.markdown("<div class='info-box'>上传您的账单Excel文件，系统将自动提取关键数据并生成汇总报表。</div>", unsafe_allow_html=True)
        
        uploaded_files = st.file_uploader(
            "上传账单Excel文件",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="支持上传多个.xlsx或.xls格式的文件"
        )
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("清除结果", key="clear_button", use_container_width=True):
                # 清除结果
                if "results" in st.session_state:
                    st.session_state.results = []
                    st.markdown("<div class='success-message'>已清除所有结果！</div>", unsafe_allow_html=True)
        
        with col2:
            if st.button("刷新", key="refresh_button", use_container_width=True):
                st.rerun()
                
        # 添加版本信息
        st.markdown("---")
        st.markdown("<div style='text-align: center; color: #9E9E9E; font-size: 0.8rem;'>版本 1.2.0</div>", unsafe_allow_html=True)
    
    # 主界面 - 显示结果表格
    if "results" not in st.session_state:
        st.session_state.results = []
    
    # 处理上传的文件
    if uploaded_files:
        processor = ExcelProcessor()
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total_files = len(uploaded_files)
        processed_count = 0
        failed_files = []
        
        for i, file in enumerate(uploaded_files):
            try:
                # 更新进度
                progress = (i + 1) / total_files
                progress_bar.progress(progress)
                
                # 更新状态栏
                status_text.text(f"正在处理: {file.name} ({i + 1}/{total_files})")
                
                # 处理Excel文件
                file_bytes = file.read()
                result = processor.process_excel(file_bytes, file.name)
                
                # 检查结果是否已存在（按文件名）
                exists = False
                for r in st.session_state.results:
                    if r["文件名"] == result["文件名"]:
                        exists = True
                        break
                
                # 只有不存在时才添加
                if not exists:
                    st.session_state.results.append(result)
                
                processed_count += 1
                
            except Exception as e:
                failed_files.append((file.name, str(e)))
                st.error(f"处理文件 {file.name} 失败")
                print(f"处理文件 {file.name} 失败: {str(e)}")
        
        # 完成处理后的操作
        progress_bar.empty()
        
        if failed_files:
            status_text.markdown(f"<div style='color: #FF5722; font-weight: bold;'>已完成处理 {processed_count} 个文件，{len(failed_files)} 个文件失败</div>", unsafe_allow_html=True)
            
            with st.expander("查看失败文件详情", expanded=True):
                for i, (file_name, error) in enumerate(failed_files):
                    st.markdown(f"**{i + 1}. {file_name}**")
                    st.markdown(f"<div style='color: #D32F2F; background-color: #FFEBEE; padding: 10px; border-radius: 5px;'>错误: {error.split('Traceback')[0]}</div>", unsafe_allow_html=True)  # 只显示错误的第一部分
        else:
            status_text.markdown(f"<div style='color: #2E7D32; font-weight: bold;'>✅ 已成功处理 {processed_count} 个文件</div>", unsafe_allow_html=True)
    
    # 显示结果表格
    if st.session_state.results:
        st.markdown("<div class='sub-header'>📊 处理结果</div>", unsafe_allow_html=True)
        
        # 转换为DataFrame
        results_df = pd.DataFrame(st.session_state.results)
        
        # 统计信息
        total_files = len(results_df)
        total_orders = results_df['当月单量'].sum() if '当月单量' in results_df.columns else 0
        total_payable = results_df['应付金额'].sum() if '应付金额' in results_df.columns else 0
        
        # 显示统计信息
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("处理文件总数", f"{total_files}个")
        with col2:
            st.metric("总单量", f"{int(total_orders)}单")
        with col3:
            st.metric("总应付金额", f"¥{total_payable:.2f}")
        
        # 显示表格
        st.dataframe(
            results_df,
            hide_index=True,
            use_container_width=True
        )
        
        # 下载按钮
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            st.markdown(get_table_download_link(results_df), unsafe_allow_html=True)
    else:
        st.markdown("<div class='info-box'>请上传账单Excel文件以开始处理</div>", unsafe_allow_html=True)
    
    # 显示使用说明
    with st.expander("查看使用说明", expanded=False):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("""
            ### 基本操作
            
            1. **上传文件** - 在左侧操作面板上传一个或多个账单Excel文件
            2. **处理数据** - 系统自动提取关键数据并汇总
            3. **查看结果** - 所有提取的数据将显示在表格中
            4. **导出数据** - 点击"下载Excel文件"保存结果
            5. **清除结果** - 使用"清除结果"按钮重新开始
            
            ### 数据提取说明
            
            系统会自动识别并提取以下字段：
            - **月结账号** - 从账单总览sheet的合并单元格中提取
            - **账单周期** - 从账单总览sheet中提取
            - **当月单量** - 统计账单明细sheet中N列为"运费"的行数
            - **费用(元)** - 从账单明细sheet的汇总行提取
            - **折扣/促销** - 从账单明细sheet的汇总行提取
            - **应付金额** - 从账单明细sheet的汇总行提取
            - **理赔费用合计** - 如存在理赔相关单元格，取H列中最小负值
            """)
            
        with col2:
            st.markdown("""
            ### 常见问题解答
            
            **Q: 为什么某些字段显示为空？**  
            A: 可能是因为账单格式与系统预期不符，或该字段在原始账单中不存在。
            
            **Q: 数据提取有误怎么办？**  
            A: 请检查原始Excel文件格式是否符合标准，或联系技术支持。
            
            **Q: 是否支持批量处理？**  
            A: 是的，您可以一次上传多个文件进行批量处理。
            
            **Q: 数据会被上传到服务器吗？**  
            A: 不会，所有处理都在您的浏览器中完成，数据不会离开您的设备。
            
            ### 技术说明
            
            - 支持的文件格式：**.xlsx**, **.xls**
            - 最佳分辨率：1920×1080
            - 推荐浏览器：Chrome, Edge, Firefox
            """)
    
    # 页脚
    st.markdown("---")
    footer_cols = st.columns([2, 1, 2])
    with footer_cols[1]:
        st.markdown("<div style='text-align: center;'>供销云仓账单数据提取工具 © 2025</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
