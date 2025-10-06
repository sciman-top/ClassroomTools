# 打包指令建议

以下命令在 Windows PowerShell 或 CMD 中均可执行，默认假设当前目录位于 `ClassroomTools.py` 所在文件夹，并且已经为 Nuitka 配置好所需的 C/C++ 编译环境与依赖缓存。

## 单文件（Onefile）

```powershell
python -m nuitka ClassroomTools.py --onefile --remove-output --assume-yes-for-downloads --jobs=%NUMBER_OF_PROCESSORS% --enable-plugin=pyqt6 --include-qt-plugins=sensible --include-package=pyttsx3.drivers --include-module=comtypes.gen --include-package-data=openpyxl --noinclude-data-files=openpyxl/tests/* --include-data-file=students.xlsx=students.xlsx --include-data-file=settings.ini=settings.ini --windows-console-mode=disable --windows-icon-from-ico=icon.ico --windows-file-version=2.2.0.0 --windows-product-version=2.2.0.0 --windows-company-name="sciman逸居" --windows-product-name="课堂工具" --windows-file-description="课堂教学辅助工具"
```

## 多文件（Standalone 目录）

```powershell
python -m nuitka ClassroomTools.py --standalone --remove-output --output-dir=dist --assume-yes-for-downloads --jobs=%NUMBER_OF_PROCESSORS% --enable-plugin=pyqt6 --include-qt-plugins=sensible --enable-plugin=numpy --include-package=pyttsx3.drivers --include-module=comtypes.gen --include-package-data=openpyxl --noinclude-data-files=openpyxl/tests/* --include-data-file=students.xlsx=students.xlsx --include-data-file=settings.ini=settings.ini --windows-console-mode=disable --windows-icon-from-ico=icon.ico --windows-file-version=2.2.0.0 --windows-product-version=2.2.0.0 --windows-company-name="sciman逸居" --windows-product-name="课堂工具" --windows-file-description="课堂教学辅助工具"
```

> `--jobs=%NUMBER_OF_PROCESSORS%` 会根据机器可用 CPU 内核数自动并行编译，加快构建速度；同时移除了冗余的 `--include-module=pyttsx3.drivers.sapi5`，并通过 `--noinclude-data-files=openpyxl/tests/*` 排除体积较大的测试数据，以缩减产物体积。
