# ClassroomTools 打包说明

本文档说明如何在 Windows 平台使用 [Nuitka](https://nuitka.net/) 将 ClassroomTools 打包为可分发的单文件或独立目录。请在执行命令前逐项确认环境、依赖与运行时资源，确保生成物与源代码保持一致。

## 1. 环境与准备

1. **Python 与依赖**
   - 使用与生产环境一致的 64 位 Python 3.10 及以上版本。
   - 运行 `pip install -r requirements.txt`（如未提供该文件，可参照 `ClassroomTools.py` 顶部的导入列表手动安装）。
2. **Nuitka 与编译工具**
   - 安装 Nuitka 推荐依赖：`pip install "nuitka[recommended]"`。
   - 准备 Visual Studio Build Tools（含 MSVC v142+ 与 Windows 10/11 SDK）或与 Python 架构匹配的 MinGW64。
3. **运行时资源**
   - `settings.ini`：使用仓库内默认文件，或根据需要调整后同步至构建目录。
   - `students.xlsx`：仓库不直接保存真实学生名单，可通过以下脚本快速生成模板：
     ```powershell
     python - <<'PY'
     import pandas as pd
     df = pd.DataFrame({
         "学号": [101, 102, 103],
         "姓名": ["张三", "李四", "王五"],
         "分组": ["A", "B", "A"],
         "成绩": [0, 0, 0],
     })
     with pd.ExcelWriter("students.xlsx", engine="openpyxl") as writer:
         df.to_excel(writer, sheet_name="班级1", index=False)
     PY
     ```
    若旧版曾生成 `students.xlsx.enc`，请先手动解密为 `students.xlsx` 再执行打包；当前版本不再内置加密/解密流程。
   - `icon.ico`：请从品牌素材库中获取 256×256 ICO 文件并放置于项目根目录。

> 所有资源需与 `ClassroomTools.py` 位于同一目录，以便 Nuitka `--include-data-file` 参数引用。

> **注意**：仓库未提供 `.gitignore`。执行打包命令会生成 `build/`、`dist/`、`__pycache__/` 等目录，请在提交前删除这些构建产物，避免污染版本历史。

## 2. 单文件（Onefile）构建

```powershell
python -m nuitka ClassroomTools.py ^
  --onefile ^
  --remove-output ^
  --assume-yes-for-downloads ^
  --jobs=%NUMBER_OF_PROCESSORS% ^
  --lto=no ^
  --enable-plugin=pyqt6 ^
  --include-qt-plugins=sensible ^
  --enable-plugin=numpy ^
  --include-package=pyttsx3.drivers ^
  --include-module=pyttsx3,pyttsx3.drivers.sapi5,pythoncom,win32api,win32con,win32gui,win32clipboard,win32com.client,win32com.server ^
  --include-package=comtypes ^
  --include-package=comtypes.gen ^
  --include-package=win32com ^
  --include-package-data=openpyxl ^
  --noinclude-data-files=openpyxl/tests/* ^
  --nofollow-import-to=numpy.tests ^
  --include-data-file=students.xlsx=students.xlsx ^
  --include-data-file=settings.ini=settings.ini ^
  --include-data-file=icon.ico=icon.ico ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=icon.ico ^
  --windows-file-version=4.0.0.0 ^
  --windows-product-version=4.0.0.0 ^
  --windows-company-name="sciman逸居" ^
  --windows-product-name="课堂工具" ^
  --windows-file-description="课堂教学辅助工具"
```

## 3. 独立目录（Standalone）构建

```powershell
python -m nuitka ClassroomTools.py ^
  --standalone ^
  --output-dir=dist ^
  --remove-output ^
  --assume-yes-for-downloads ^
  --jobs=%NUMBER_OF_PROCESSORS% ^
  --lto=no ^
  --enable-plugin=pyqt6 ^
  --include-qt-plugins=sensible ^
  --enable-plugin=numpy ^
  --include-package=pyttsx3.drivers ^
  --include-module=pyttsx3,pyttsx3.drivers.sapi5,pythoncom,win32api,win32con,win32gui,win32clipboard,win32com.client,win32com.server ^
  --include-package=comtypes ^
  --include-package=comtypes.gen ^
  --include-package=win32com ^
  --include-package-data=openpyxl ^
  --noinclude-data-files=openpyxl/tests/* ^
  --nofollow-import-to=numpy.tests ^
  --include-data-file=students.xlsx=students.xlsx ^
  --include-data-file=settings.ini=settings.ini ^
  --include-data-file=icon.ico=icon.ico ^
  --windows-console-mode=disable ^
  --windows-icon-from-ico=icon.ico ^
  --windows-file-version=4.0.0.0 ^
  --windows-product-version=4.0.0.0 ^
  --windows-company-name="sciman逸居" ^
  --windows-product-name="课堂工具" ^
  --windows-file-description="课堂教学辅助工具"
```

## 4. 命令审核与补充建议

- 上述指令已覆盖 PyQt6、pyttsx3、pywin32、comtypes、openpyxl 等依赖，显式捆绑配置与数据文件，足以满足当前功能。
- `--jobs=%NUMBER_OF_PROCESSORS%` 允许并行编译，建议保留；若在老旧机器上遇到过热或内存不足，可改为具体数字限制并发度。
- `--nofollow-import-to=numpy.tests` 与 `--noinclude-data-files=openpyxl/tests/*` 可显著减小体积；如需调试这些模块可暂时移除。
- 对需要控制包体大小的发布，可再加上 `--nofollow-import-to=matplotlib.tests` 等补充排除项（视实际依赖情况调整）。
- 调试失败时可临时追加 `--show-scons` 或将 `--windows-console-mode` 调为 `force` 以观察标准输出。

## 5. 常见问题排查

| 现象 | 排查建议 |
| --- | --- |
| 构建时提示缺少 `students.xlsx` | 按上文脚本生成模板，或确认资源文件与命令中的路径一致。 |
| 运行后语音合成报错 | 确认系统已安装 SAPI5 中文语音包，可在控制面板语音设置里检查。 |
| 打包产物被杀毒软件误报 | Onefile 会自解压，必要时使用 Standalone 构建或在安全软件中添加信任。 |
| 程序无法读取设置 | 确保 `settings.ini` 与可执行文件同级；若需要自定义路径，可在应用启动后通过设置界面保存一次。 |

发布前请在干净环境中运行打包产物，完成冒烟测试与主要功能回归。
