# ResumeAssistant 简历助手

解析个人简历 Word 文档（`.docx`），以「左侧大纲树 + 右侧内容面板」的形式展示，支持点击复制与窗口置顶。

## 功能

- 打开 exe 后引导选择 `.docx` 简历文档
- 左侧大纲：一级标题可折叠/展开，二级子节可点击
- 右侧内容：按「名称 / 内容 / 小节」分块展示，选中一级显示全部子节，选中子节仅显示该节
- 内容只读，单击任意文本块自动复制到剪贴板
- 支持最小化、关闭、置顶（置顶后始终显示在其他窗口之上）

## 运行（源码）

```bash
pip install -r requirements.txt
python main.py
```

## 打包 exe

双击运行 `build.bat`，或在独立虚拟环境中执行：

```bash
python -m venv .venv
.venv/Scripts/python -m pip install python-docx pyinstaller
.venv/Scripts/python -m PyInstaller --noconsole --onefile --name ResumeAssistant main.py
```

生成文件位于 `dist/ResumeAssistant.exe`。建议在独立 venv 中打包（Anaconda 自带的过时 `pathlib` 反向移植包会与 PyInstaller 冲突）。
