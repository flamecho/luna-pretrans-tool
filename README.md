# Luna 预翻译工具

为 [LunaTranslator](https://github.com/HIllya51/LunaTranslator) 制作预翻译文件的工具。

## 简易流程

```
原文 txt ──【译前处理】──> 待翻译 xlsx / txt
                                 │
                         【翻译器机翻】
                                 │
           译文 xlsx 或 译文 txt（配原文 txt）
                                 │
                         【译后处理】──> sqlite 预翻译文件
```

两条链路并存，按手上翻译器支持的格式选：

| 链路 | 译前产出 | 译后输入 |
| --- | --- | --- |
| 表格 | 待翻译 `.xlsx`，原文在 A 列 | 译文 `.xlsx`，A 列原文 / B 列译文 |
| 纯文本 | 待翻译 `.txt`，一行一条 | 原文 `.txt` 与译文 `.txt` 各放一个文件夹 |

纯文本链路不依赖表格，适合只认 txt 的翻译器；老版本翻译器可以继续走表格链路。

## 使用说明

### ① 译前处理（原文 txt → 待翻译 xlsx / txt）

1. 选择「原文 txt 文件夹」→ 选择「输出文件夹」
2. 勾选清洗规则（默认全开）：
   - 去除全角空格（保留半角）
   - 去除空行 / 纯空白行
   - 去除重复行（保留首次出现）
3. 在「包含子文件夹」右侧勾选输出格式：`xlsx` / `txt`，两个都勾就是同时导出
4. 点「开始处理」
5. 每个 `.txt` 生成一个同名文件：
   - `xlsx`：原文写在 A 列
   - `txt`：一行一条，UTF-8（带 BOM）+ LF
   - 需要时勾「包含子文件夹」，输出按源目录结构镜像

> 输出文件夹不能和原文文件夹相同 —— 导出 txt 会覆盖原文，程序会直接拦下。

### ② 译后处理（译文 xlsx → sqlite 预翻译文件）

1. 「输入格式」选「译文 xlsx」
2. 选择「译文 xlsx 文件夹」→ 选择「输出文件夹」
3. 点「开始处理」
4. 每个 `.xlsx` 生成一个同名 `.sqlite`

### ③ 译后处理（原文 txt + 译文 txt → sqlite 预翻译文件）

1. 「输入格式」选「原文 + 译文 txt」
2. 选择「原文 txt 文件夹」（译前导出的 txt，或任何一行一条的原文）
3. 选择「译文 txt 文件夹」
4. 选择「输出文件夹」→ 点「开始处理」

配对规则：

- 按相对路径同名配对，`a/b.txt` 对 `a/b.txt`，子文件夹结构要一致
- 输出文件名取原文 txt 的名字
- 译文里的空行照样占一行，视为该行未翻译，`machineTrans` 回填原文
- 有原文、缺译文的文件会被跳过，并在日志里列出

两种链路都套用同一套「替换规则」。

## 下载

到 [Releases](../../releases) 下载最新的 `Luna预翻译工具.exe`（免安装，Windows 10 / 11 x64）。

## 从源码运行 / 打包

```bash
pip install -r requirements.txt
python luna_tool.py            # 直接运行

pip install pyinstaller
python build_exe.py            # 打包成 dist/Luna预翻译工具.exe
```

命令行自检（不开界面，便于批量核对）：

```bash
python luna_tool.py --selftest 原文目录 输出目录 both            # 译前：xlsx / txt / both
python luna_tool.py --selftest-trans 译文xlsx目录 输出目录        # 译后：xlsx → sqlite
python luna_tool.py --selftest-trans-txt 原文目录 译文目录 输出目录  # 译后：txt → sqlite
```

## 许可

本项目基于 [MIT License](LICENSE) 开源。
