# CSV / Excel 文件合并工具（Windows GUI）

这是一个按列名对齐数据的 CSV / Excel 批量合并工具。界面支持选择输入文件夹、列名模板和输出文件夹，并自动记住上一次的选择路径。

## 界面功能

- 输入文件夹：合并该文件夹第一层中的 `.csv`、`.xlsx`、`.xlsm` 文件。Excel 输入读取第一个工作表，第一行作为表头。
- 列名模板（可选）：支持 `.csv`、`.xlsx`、`.xlsm`，读取第一个工作表的第一行作为目标列名及顺序。
- 输出文件夹：输出 `merged_output.csv`，文件名可在界面修改。
- 来源列：默认添加 `Source_File`，用于追溯每一行来自哪个文件。
- 路径记忆：设置保存在当前 Windows 用户的 `%APPDATA%\CSV文件合并工具\settings.json`。
- 编码：自动兼容 UTF-8、UTF-8 BOM、GB18030 和 GBK，输出统一为 Excel 可直接打开的 UTF-8 BOM CSV。

不选择模板时，程序会按照各输入文件的出现顺序生成全部列名的并集。选择模板时，不在模板中的源列不会写入，缺少的模板列会补空值；异常情况记录在输出目录的 `csv_column_warning_report.csv`。程序始终输出 CSV 文件。

## 使用 GitHub Actions 生成 Windows EXE（推荐）

项目已包含 `.github/workflows/build-windows.yml`。将项目推送到 GitHub 后：

1. 打开 GitHub 仓库的 **Actions** 页面。
2. 在左侧选择 **Windows EXE 打包**。
3. 点击 **Run workflow**，再次点击绿色的 **Run workflow**。
4. 等待构建任务显示绿色对勾，然后打开该次运行记录。
5. 在页面底部的 **Artifacts** 中下载 `CSV文件合并工具-Windows`。
6. 解压下载的 ZIP，即可得到 `CSV文件合并工具.exe`。

向 `main` 或 `master` 分支推送代码、推送以 `v` 开头的版本标签时，也会自动触发打包。构建产物默认保留 30 天。

`.gitignore` 已排除 `.venv`、`input` 和 `sum` 等本地环境及数据文件，避免将测试数据上传到 GitHub。

## 在 Windows 本地生成 EXE

将整个项目复制到 Windows 电脑，安装 Python 3.9 或更高版本，然后双击：

```text
build_windows.bat
```

打包完成后，可独立分发的程序位于：

```text
dist\CSV文件合并工具.exe
```

目标电脑不需要安装 Python。由于 PyInstaller 不能跨系统生成 Windows 程序，请在 Windows 上执行打包脚本。

## 从源码运行

```bat
py -3 -m venv .venv-win
.venv-win\Scripts\python -m pip install -r requirements.txt
.venv-win\Scripts\python gui_app.py
```

命令行方式仍然可用：

```bat
.venv-win\Scripts\python run.py "D:\input" -o "D:\output" -t "D:\template.xlsx"
```
