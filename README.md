# Excel Sheet 合并工具 📊

将多个 Excel 文件的所有 Sheet 合并到一个文件中。

## 功能

- ✅ 支持多文件合并
- ✅ 自动合并所有 Sheet
- ✅ 保留来源信息（文件名、Sheet名）
- ✅ 命令行 + Web 两种使用方式

## 安装

```bash
pip install flask pandas openpyxl
```

## 使用方式

### Web 版（推荐）

```bash
python excel_merge_web.py
```

打开浏览器访问 http://localhost:5050，拖拽上传文件即可。

### 命令行版

```bash
# 合并单个文件的所有 sheet
python excel_merge.py data.xlsx

# 合并多个文件
python excel_merge.py file1.xlsx file2.xlsx -o combined.xlsx

# 不添加来源信息列
python excel_merge.py data.xlsx --no-source
```

## 输出说明

合并后的文件会自动添加两列（可选）：
- `_来源文件` - 原始文件名
- `_来源Sheet` - 原始 Sheet 名

方便追溯数据来源。

## License

MIT
