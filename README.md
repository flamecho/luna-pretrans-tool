# Luna 预翻译工具

给 [LunaTranslator](https://github.com/HIllya51/LunaTranslator) 生成 sqlite 预翻译文件 的小工具。

## 流程

```
原文 txt ─【译前处理】→ 待翻译 xlsx / txt ─【翻译器】→ 译文 ─【译后处理】→ sqlite
```

译前负责清洗和拆分，译后负责套替换规则、写进 sqlite 模板。两种链路并存：

| 链路 | 译前产出 | 译后输入 |
| --- | --- | --- |
| 表格 | `.xlsx`，原文在 A 列 | `.xlsx`，A 列原文 / B 列译文 |
| 纯文本 | `.txt`，一行一条 | 原文 `.txt` + 译文 `.txt`，两个文件夹 |

两条链路用的替换规则和最终结果完全一致，可以按需混用。

## 用法

**译前处理**：选原文 txt 文件夹 → 选输出文件夹 → 勾 `xlsx` / `txt`（可同时勾）→ 开始处理。每个 txt 生成一个同名待翻译文件。

**译后处理**：先切「输入格式」，再选对应文件夹 → 开始处理。每个文件生成一个同名 sqlite。

几点约定：

- 纯文本链路按相对路径同名配对，子文件夹结构要一致
- 译文里的空行占一行，视为该行未翻译，`machineTrans` 回填原文
- 输出文件夹不能和原文文件夹相同，否则导出 txt 会覆盖原文（程序会拦下）

## 下载

到 [Releases](../../releases) 下载 `LunaPretransTool.zip`（免安装，Windows 10 / 11 x64）。

## 从源码运行

```bash
pip install -r requirements.txt
python luna_tool.py        # 直接运行

pip install pyinstaller
python build_exe.py        # 打包成 dist/Luna预翻译工具.exe
```

批量核对时可以用命令行自检，不开界面：

```bash
python luna_tool.py --selftest 原文目录 输出目录 both
python luna_tool.py --selftest-trans 译文xlsx目录 输出目录
python luna_tool.py --selftest-trans-txt 原文目录 译文目录 输出目录
```

## 许可

[MIT License](LICENSE)
