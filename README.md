# Luna 预翻译工具

![license](https://img.shields.io/badge/license-MIT-blue.svg)
![python](https://img.shields.io/badge/python-3.8%2B-blue)
![platform](https://img.shields.io/badge/platform-Windows-lightgrey)

面向 **Galgame / 视觉小说** 翻译的文本预处理小工具。它把两步重复劳动自动化：

1. **译前处理**：把原文本 `.txt` 清洗后导出为待翻译的 `.xlsx`（原文统一放在 A 列），直接喂给 Linguagacha / AiNiee 等翻译工具。
2. **译后处理**：把译文 `.xlsx`（A 列原文 / B 列译文）按一套**替换规则**处理后，写入内置 sqlite 模板的 `machineTrans` 列，导出为最终的 `.sqlite` 预翻译文件。

**完全离线**：不联网、不上传任何文本、不写注册表、不读取个人目录。

---

## 目录

- [功能特性](#功能特性)
- [界面截图](#界面截图)
- [快速开始](#快速开始)
- [从源码运行](#从源码运行)
- [打包成 exe](#打包成-exe)
- [使用说明](#使用说明)
- [内置模板](#内置模板)
- [项目结构](#项目结构)
- [常见问题](#常见问题)
- [许可](#许可)

---

## 功能特性

- **一键译前清洗**：去除全角空格（保留半角）、去除空行/纯空白行、去除重复行（保留首次出现），三条规则可单独开关；支持「包含子文件夹」按源目录结构镜像输出。
- **一键译后处理**：套用替换规则 → 用原文 A 列 + 内置模板建底库 → 写入 `machineTrans` → 空译文自动用原文回填，全程无需手动操作 sqlite。
- **可视化替换规则表**：预置一套默认规则（引号归一、`...`→`…`、`俺/老子`→我、`妳`→你、`。」`→`」`、首尾 `^`/`$` 包裹等），支持添加 / 编辑 / 删除 / 上移下移 / 恢复默认 / 导入导出，自动保存在 exe 旁的 `替换规则.json`。
- **原生 ttk 界面**：平角卡片、原生控件，清晰、性能好；支持浅色 / 深色主题并自动记忆；高 DPI 感知（125% / 150% 缩放下不发虚）。
- **实时日志与进度**：处理进度、行数、回填数量、异常信息集中显示，内容超出窗口时可滚动。
- **命令行自检模式**：`--selftest` / `--selftest-trans`，便于批量验证与 CI。

## 界面截图

> 截图待补充。请把界面截图放到 `docs/screenshots/` 下，并在此处引用，例如：
>
> ```markdown
> ![译前处理](docs/screenshots/light-1.png)
> ![译后处理](docs/screenshots/dark-1.png)
> ```

## 快速开始

1. 打开仓库的 [**Releases**](../../releases) 页面，下载最新的 `Luna预翻译工具.exe`。
2. 双击运行（免安装，支持 Windows 10 / 11 x64）。
3. 按界面提示选择文件夹并点「开始处理」。

## 从源码运行

需要 **Python 3.8+**（Windows，且带 tkinter）。

```bash
pip install -r requirements.txt
python luna_tool.py
```

## 打包成 exe

```bash
pip install pyinstaller
python build_exe.py
# 产物：dist/Luna预翻译工具.exe
```

## 使用说明

三个页面（顶部按钮切换）：**① 译前处理 ② 译后处理 ③ 日志**。

### ① 译前处理（txt → xlsx）

1. 「原文 txt 文件夹」选择未处理的 `.txt` 所在文件夹。
2. 「输出文件夹」选一个位置。
3. 点「开始处理」：每个 `.txt` 生成一个 `.xlsx`，清洗后的原文写在 A 列。
   - 清洗规则可勾选：去全角空格（保留半角）/ 去空行纯空白行 / 去重（保留首次出现）。
   - 可勾选「包含子文件夹」，输出按源目录结构镜像。

### ② 译后处理（xlsx → sqlite）

翻译完成后，把译文 xlsx（A=原文，B=译文）所在文件夹选进来，一键完成：

1. B 列逐行套用「替换规则」；
2. 用 A 列 + 内置模板建底库（无需另外准备 sqlite）；
3. 写入 `machineTrans` 列；
4. 空译文 `{"rengong": ""}` 自动用该行原文回填。

输出文件夹里每个 xlsx 对应一个 `.sqlite` 终稿。日志会汇报行数、回填数量、A 列与原文不一致数量。

### ③ 日志

所有处理日志集中显示，出错原因也在此页；可一键清空。

## 内置模板

- **sqlite 模板**：50 万行（id 1..500000 连续），表结构
  `artificialtrans(id INTEGER PRIMARY KEY AUTOINCREMENT, source TEXT, machineTrans TEXT, origin TEXT)`。
- **xlsx 模板**：空白单表 `Sheet1`。
- 超过 50 万行会在日志报错。

## 项目结构

```
luna-pretrans-tool/
├── luna_tool.py            # 主程序（单文件，含 UI + 业务逻辑）
├── build_exe.py            # PyInstaller 打包脚本
├── requirements.txt        # 运行依赖
├── resources/
│   ├── default.sqlite      # 内置 sqlite 模板（50 万行）
│   ├── default.xlsx        # 内置 xlsx 模板
│   └── feather.ico         # 软件图标
├── docs/screenshots/       # 界面截图
├── 使用说明.txt            # 随 exe 分发的说明
├── LICENSE
└── README.md
```

## 常见问题

**Q：处理超慢 / 卡住？**
A：大文件（几万行）需要一些时间，进度条与日志会实时更新；可点「停止」中断。

**Q：替换规则改乱了？**
A：点规则表右侧的「恢复默认」即可还原预置规则。

**Q：exe 换了图标但资源管理器里没变？**
A：那是 Windows 图标缓存。重启资源管理器、把 exe 复制到新路径，或运行 `ie4uinit.exe -show` 刷新即可。

**Q：支持 Mac / Linux 吗？**
A：目前界面基于 Windows 的 tkinter + 高 DPI API，未在其它平台测试。

## 许可

本项目基于 [MIT License](LICENSE) 开源。
