# Luna 预翻译工具

为 [LunaTranslator](https://github.com/HIllya51/LunaTranslator) 制作预翻译文件的工具。

## 简易流程

```
原文 txt → 【译前处理】 → 待翻译 xlsx → 【翻译器机翻】 → 译文 xlsx → 【译后处理】 → sqlite 预翻译文件
```

## 使用说明

### ① 译前处理（原文 txt → 待翻译 xlsx）

1. 选择「原文 txt 文件夹」→ 选择「输出文件夹」
2. 勾选清洗规则（默认全开）：
   - 去除全角空格（保留半角）
   - 去除空行 / 纯空白行
   - 去除重复行（保留首次出现）
3. 点「开始处理」
4. 每个 `.txt` 生成一个同名 `.xlsx`，原文写在 A 列
   - 需要时勾「包含子文件夹」，输出按源目录结构镜像

### ② 译后处理（译文 xlsx → sqlite 预翻译文件）

1. 选择「译文 xlsx 文件夹」→ 选择「输出文件夹」
2. 点「开始处理」
3. 每个 `.xlsx` 生成一个同名 `.sqlite`

## 下载

到 [Releases](../../releases) 下载最新的 `Luna预翻译工具.exe`（免安装，Windows 10 / 11 x64）。

## 从源码运行 / 打包

```bash
pip install -r requirements.txt
python luna_tool.py            # 直接运行

pip install pyinstaller
python build_exe.py            # 打包成 dist/Luna预翻译工具.exe
```

## 许可

本项目基于 [MIT License](LICENSE) 开源。
