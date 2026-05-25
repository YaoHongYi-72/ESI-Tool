# ESI-Tool

`ESI-Tool` 是一套面向非计算机专业人员的 Windows 桌面统计工具，用来把 `6` 期 ESI 附表自动整理成年度统计工作簿和配套报告，并支持生成可分发的 Windows 安装包。

当前版本聚焦 `2025` 年统计口径，能够自动生成：

- `表12` 入围期数分布
- `表13` 院系及学科分布
- `表14` 院系及作者分布
- 去重总表 Excel
- 年度正式报告 DOCX
- 统计口径说明 DOCX
- 作者统计流程详解 DOCX

## 仓库结构

```text
analyze_2025_esi.py
generate_2025_esi_doc.py
generate_2025_esi_process_doc.py
generate_author_flow_doc.py
generate_esi_stats_workbook.py
generate_windows_tool_tech_doc.py
windows_tool/
  esi_gui_app.py
  esi_pipeline.py
  build_windows_exe.bat
  build_windows_installer.bat
  build_iexpress_installer.ps1
  install.cmd
  run_esi_tool.bat
  esi_tool.spec
  requirements-windows.txt
.github/workflows/
  build-windows-exe.yml
```

## 给老师直接使用

### 方法一：下载 GitHub Actions 自动打包结果

1. 打开仓库的 `Actions`
2. 进入最近一次 `Build Windows Installer Package`
3. 下载构建产物：
   - `ESI统计工具-installer`
   - `ESI统计工具-portable`
4. 安装版直接运行 `ESI-Tool-Setup.exe`
5. 便携版解压后运行 `ESI统计工具.exe`

### 方法二：在 Windows 本机从源码启动

1. 安装 `Python 3.11` 或 `3.12`
2. 下载或克隆本仓库
3. 双击运行 `windows_tool/run_esi_tool.bat`
4. 在界面中选择：
   - `附表文件夹`
   - `2024模板报告`
   - `输出工作目录`
5. 点击 `开始生成`

## 在 Windows 上本地打包

### 打包安装版

直接双击：

```text
windows_tool/build_windows_installer.bat
```

打包完成后，安装包位于：

```text
release/ESI-Tool-Setup.exe
```

### 仅打包单文件 exe

直接双击：

```text
windows_tool/build_windows_exe.bat
```

打包完成后，程序目录位于：

```text
dist/ESI统计工具.exe
```

## GitHub 自动打包

仓库内置了 GitHub Actions 工作流：

- 推送到 `main` 或 `master` 时自动构建 Windows 安装包和便携版
- 手动触发 `workflow_dispatch` 时也会构建
- 如果推送的是 `v*` 标签，会额外把安装包和便携版上传到 GitHub Release

构建产物为：

```text
ESI-Tool-Setup.exe
ESI-Tool-portable.zip
```

## 输入要求

当前版本按 `2025` 年度统计流程设计，默认识别以下附表命名模式：

- `ESI*第1期附表*.xlsx`
- `ESI*第2期附表*.xlsx`
- `ESI*第3期附表*.xlsx`
- `ESI*第4-5期附表*.xlsx`
- `ESI*第6期附表*.xlsx`

其中第 `4-5` 期默认从同一个工作簿中读取两个 sheet。

## 模板说明

年度正式报告依赖 `2024` 年正式报告模板。出于使用场景考虑，模板文件默认不随仓库提交；使用时请在图形界面中手工选择对应的 `.docx` 模板。

如果希望把模板固定在仓库内，可以放到：

```text
templates/
```

## 开发说明

创建虚拟环境并安装依赖：

```bash
python -m venv .venv
```

Windows:

```powershell
.venv\Scripts\python -m pip install -r windows_tool/requirements-windows.txt
```

macOS / Linux:

```bash
.venv/bin/python -m pip install -r windows_tool/requirements-windows.txt
```

本地启动图形界面：

```bash
python -m windows_tool.esi_gui_app
```

## 当前边界

- 当前版本主要针对 `2025` 年口径
- `表14` 采用的是“每篇论文、每个学院只记 `1` 位代表作者”的口径，不是全作者展开表
- 个别仅保留拼音的作者仍需要人工后续核实
