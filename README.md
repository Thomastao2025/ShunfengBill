# 供销云仓账单数据提取工具

这是一个基于Streamlit开发的网页应用，用于提取和处理供销云仓的账单Excel文件，并提供数据导出功能。

## 功能

- 上传多个Excel账单文件进行批量处理
- 自动提取关键信息：月结账号、账单周期、当月单量、费用、折扣、应付金额和理赔费用
- 显示处理结果并支持导出为Excel格式
- 简洁清晰的用户界面，操作简单直观

## 部署说明

### 在Streamlit Cloud上部署

1. Fork这个仓库到自己的GitHub账号
2. 访问 [Streamlit Sharing](https://share.streamlit.io/) 并登录
3. 点击"New app"按钮
4. 选择你Fork的仓库和主分支
5. 设置主文件为`app.py`
6. 点击"Deploy"

### 本地运行

1. 克隆仓库到本地
```bash
git clone https://github.com/你的用户名/供销云仓账单数据提取工具.git
cd 供销云仓账单数据提取工具
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 运行应用
```bash
streamlit run app.py
```

## 使用指南

1. 在应用左侧的操作面板点击"上传账单Excel文件"上传一个或多个账单文件
2. 系统会自动处理上传的文件并提取关键数据
3. 处理结果将显示在表格中
4. 点击"下载Excel文件"链接将结果下载为Excel文件
5. 使用"清除结果"按钮可以清空当前结果

## 文件结构

- `app.py`: 主应用程序
- `requirements.txt`: 依赖库列表
- `README.md`: 项目说明文档

## 依赖库

- streamlit
- pandas
- openpyxl

## 注意事项

- 程序会尝试从不同的Excel表格位置提取数据，但如果账单格式与预期差异较大，可能无法正确提取所有信息
- 所有数据处理都在用户浏览器中完成，数据不会被上传到服务器
