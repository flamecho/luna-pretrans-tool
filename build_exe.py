# -*- coding: utf-8 -*-
"""把 luna_tool.py 用 PyInstaller 打包成单文件 exe（Windows）。

用法：
    pip install pyinstaller
    python build_exe.py
产物：dist/Luna预翻译工具.exe
"""
import os
import sys
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
NAME = 'Luna预翻译工具'
ICON = os.path.join(HERE, 'resources', 'feather.ico')
RES = os.path.join(HERE, 'resources')

cmd = [
    sys.executable, '-m', 'PyInstaller',
    '--noconfirm', '--onefile', '--windowed',
    '--name', NAME,
    '--icon', ICON,
    '--add-data', RES + os.pathsep + 'resources',
    '--exclude-module', 'numpy',
    '--exclude-module', 'pandas',
    '--exclude-module', 'matplotlib',
    '--exclude-module', 'scipy',
    '--exclude-module', 'PIL',
    '--exclude-module', 'pytest',
    '--exclude-module', 'IPython',
    '--exclude-module', 'notebook',
    os.path.join(HERE, 'luna_tool.py'),
]

print(' '.join(f'"{c}"' if ' ' in c else c for c in cmd))
subprocess.run(cmd, cwd=HERE, check=True)
print('\n完成：', os.path.join(HERE, 'dist', NAME + '.exe'))
