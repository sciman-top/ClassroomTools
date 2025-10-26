# ClassroomTools Repository Guide

## 项目概览
- 仓库采用**单一的 `ClassroomTools.py`**，其中按照逻辑区域分块（通用工具 → 系统集成 → 启动器/叠加层 → 点名/画板/计时器 → 应用入口）。
- 如需新增功能，请先查找是否已有相近的 mixin 或 helper，再决定扩展或复用，避免在窗口类里堆叠长函数。
- PyQt 相关改动务必测试在 Windows 桌面环境上的行为，并保持 UI 字符串为中文。

## 贡献准则
- 仅在明确需要时才引入第三方依赖，所有平台特定逻辑仍需通过 `if sys.platform == "win32"` 保护。
- 日志请通过共享的 `logger` 或 `_resolve_debug_logger()` 等工具记录，禁止使用 `print()`。
- 修改或新增 helper 时，同时更新 `tests/test_helpers.py` 中的对应单元测试；新增行为若无法测试，请在 PR 中解释原因。
- 运行 `pytest` 是提交前的最低要求，如涉及 UI 行为可额外附加截图或录屏（仓库不保存二进制资产）。

## 配置与资源
- `settings.ini` 中的默认值应与 `SettingsManager.defaults` 保持同步；若新增字段请同时更新两处并说明迁移策略。
- 运行时生成的 `students.xlsx`（以及旧版本遗留的 `students.xlsx.enc`）等用户数据，请勿提交到版本库，可在发布前从受控位置拷贝进入构建目录。
- 仓库不直接托管大体积或版权受限的二进制文件（如真实学生名单、正式图标）；请在 `PACKAGING.md` 说明如何准备这些资源。

## 文档与打包
- 调整构建流程时更新 `PACKAGING.md`，并保持文件中 Onefile/Standalone 命令与 `ClassroomTools.py` 底部注释一致。
- 仓库不再提供 `.gitignore`；构建或测试后请手动清理 `build/`、`dist/`、`__pycache__/` 等临时目录，确保提交前工作区干净。

## 测试与质量
- 运行 `pytest` 并在 PR 描述中附上命令与结果。
- 对性能或鲁棒性有影响的改动，应配套添加回归测试或 profiling 说明。

