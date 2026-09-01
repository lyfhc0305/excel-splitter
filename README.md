# Excel 拆表工具

## 功能

支持两种拆分方式，可在界面顶部或命令行中切换：

### 方式 A：按关键列拆分（拆数据行）

- 用户可自行选择表头占用前几行
- 用户可自行选择表尾占用后几行（如"合计"行），会固定附加到每个拆分文件末尾；表尾里的合计公式会按各文件自身数据自动重算
- 用户可自行选择关键单元格所在列
- 关键列中每个不同的关键字生成一个文件，包含表头 + 该关键字对应的所有数据行 + 表尾

### 方式 B：按关键行拆分（拆数据列）

- 用户可自行选择左侧固定列数（如"序号""项目名称"等列，会保留在每个拆分文件左侧）
- 用户可自行选择右侧固定列数（如"合计"列，会保留在每个拆分文件右侧）；固定列里的公式会按各文件自身保留的数据列自动重算
- 用户可自行选择关键行序号（该行中的关键字决定数据列的分组，支持横向合并单元格的关键字）
- 关键行中每个不同的关键字生成一个文件，包含左侧固定列 + 该关键字对应的所有数据列 + 右侧固定列；所有行都会完整保留
- 跨越被剔除数据列的合并单元格（如覆盖整行的大标题）会自动收缩到保留的列，标题文字不丢失

### 通用

- 拆分后的文件保留原表的样式、单元格格式、列宽、行高、合并单元格和打印设置
- 支持 `xlsx`、`xlsm`、`xltx`、`xltm`、部分 `et` 文件
- `xls` 和非 xlsx 内核的 `et` 文件会先尝试调用 LibreOffice、Excel 或 WPS 转换为临时 `xlsx` 后再拆分

## 格式说明

- 新版 Office/WPS 表格通常可以直接处理
- 老式 `xls` 文件不是 openpyxl 原生支持格式，程序会自动寻找可用转换器
- 如果电脑没有安装 LibreOffice、Excel 或 WPS，老式 `xls` 和部分 `et` 可能无法自动转换
- 拆分后的输出文件统一为 `xlsx`

## 目录结构

```text
project_拆表/
├─ src/        # 源代码（excel_splitter.py）
├─ data/       # 样例输入文件
├─ output/     # 拆分输出结果
├─ scripts/    # 打包脚本（build_windows.bat / build_linux.sh）
├─ dist/       # 已打包好的可执行程序
├─ ExcelSplitter-Windows.spec
├─ requirements.txt
└─ README.md
```

## 运行方式

### 方式一：单窗口图形界面

```powershell
python .\src\excel_splitter.py
```

运行后会打开一个主界面，在同一个窗口里完成：

1. 选择待拆分的 Excel 文件
2. 选择工作表
3. 选择拆分方式（按关键列 / 按关键行）
4. 按关键列拆分时：输入表头行数、表尾行数、关键列序号（可点选预览表格的列标题快速选择）
5. 按关键行拆分时：输入左侧固定列数、右侧固定列数、关键行序号（可直接点击预览表格中的某一行快速选择）
6. 选择输出目录
7. 点击“开始拆分”

界面会自动加载工作表列表，并实时预览拆分对象和分组数量。

### 方式二：命令行直接指定参数

按关键列拆分（默认模式）：

```powershell
python .\src\excel_splitter.py --input ".\data\示例分配表.xlsx" --sheet 示例工作表 --header-rows 3 --footer-rows 1 --key-column 7
```

按关键行拆分：

```powershell
python .\src\excel_splitter.py --input ".\data\某分季度统计表.xlsx" --mode row --header-cols 1 --footer-cols 1 --key-row 2
```

参数说明：

- `--input`：输入 Excel 文件
- `--sheet`：要拆分的工作表名称，不传时默认第一个工作表
- `--mode`：拆分方式，`column`=按关键列（默认），`row`=按关键行
- `--header-rows`：表头占用行数（column 模式）
- `--footer-rows`：表尾占用行数（固定在每个文件末尾，column 模式），默认 0
- `--key-column`：关键列序号，从 1 开始（column 模式）
- `--header-cols`：左侧固定列数（row 模式）
- `--footer-cols`：右侧固定列数（固定在每个文件右侧，row 模式），默认 0
- `--key-row`：关键行序号，从 1 开始（row 模式）
- `--output-dir`：输出目录，不传时默认生成到 `split_output`

## 测试结果

已使用 `data/` 中的通用样例文件 `示例分配表.xlsx` 测试通过（工作表 `示例工作表`、表头 `3` 行、关键列 `7`），拆分结果见 `output/` 目录。

## 免环境打包

可以将程序打包成无需预装 Python 的可执行版本：

### Windows

运行：

```powershell
.\scripts\build_windows.bat
```

生成目录：

```text
dist\ExcelSplitter-Windows
```

可将整个目录拷贝到其他 Windows 电脑直接使用，启动文件在该目录内。

### Linux

在 Linux 环境中运行：

```bash
chmod +x scripts/build_linux.sh
./scripts/build_linux.sh
```

生成目录：

```text
dist/ExcelSplitter-Linux
```

可将整个目录拷贝到其他 Linux 电脑直接使用。

### 注意

- Windows 可执行文件需要在 Windows 上构建
- Linux 可执行文件需要在 Linux 上构建
- 这类桌面程序通常不能直接跨平台打包出另一种系统的可执行文件
