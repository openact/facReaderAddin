# facReaderAddin

`facReaderAddin` 是面向 Excel 的 `.fac` 结果查询工具链。

它包含：

- Go DLL（`facReaderAddin.dll`）：执行 `.fac` 单值和批量查询
- VBA 模块（`facReaderAddin.bas`）：暴露 Excel UDF，并把 F9 绑定为“只失效变化的 facReaderAddin 缓存后重算”
- XLL wrapper（`build/excel_xll/`）：暴露更快的原生 Excel XLL 单值函数
- XLL controller XLAM（`facReaderAddinXControl.xlam`）：XLL 模式下只绑定 F9/Shift+F9，不暴露 UDF
- 部署脚本：用于生成和打包 Excel 加载项（`.xlam` / `.xll`）

## 目录结构

- `build/dll/`：Go c-shared DLL 入口
- `build/excel_xlam/`：VBA 模块和 XLAM 打包脚本；XLAM 会把 F9 绑定为“只失效变化的 facReaderAddin 缓存后重算”
- `build/excel_xll/`：XLL wrapper 源码、XLL controller XLAM 和构建脚本；复用 `facReaderAddin.dll`，不合并二进制
- `internal/`：共享查询逻辑
- `dist/`：平铺发布产物，包含 XLAM、XLL、DLL 和说明文件
- `docs/`：使用说明
- `input/`：本地样例和测试输入数据
- `main.go`：命令行测试入口

## 构建与打包

推荐使用统一部署脚本：

```powershell
Set-Location "als/app/facReaderAddin"
.\deploy.ps1 -Target all
```

`Target` 可选：`dll`、`xlam`、`xll`、`all`。其中 `xlam` 和 `xll` 都依赖已存在的 `dist/facReaderAddin.dll`。

也可以分别运行子脚本。DLL：

```powershell
Set-Location "als/app/facReaderAddin/build/dll"
.\build.ps1
```

XLAM：

```powershell
Set-Location "als/app/facReaderAddin/build/excel_xlam"
.\deploy.ps1
```

发布产物输出到 `dist/`：

- `facReaderAddin.dll`
- `facReaderAddin.xlam`（自动打包；若环境受限则需手工生成）
- `facReaderAddinX.xll`
- `facReaderAddinXControl.xlam`（XLL 模式的 F9/Shift+F9 controller）
- `facReaderAddin_README.md`（用户版说明）

XLL：

```powershell
Set-Location "als/app/facReaderAddin/build/excel_xll"
.\build.ps1
```

XLL 产物也输出到 `dist/`：

- `facReaderAddinX.xll`
- `facReaderAddinXControl.xlam`
- `facReaderAddin.dll`
- `facReaderAddin_README.md`

## Excel 函数

单值预测查询：

`EProj_Result(resVault, resID, product, spCode, resultType, variable, timePeriod, [simID])`

通用单值查询：

`ERead_Result(resVault, resID, filename, key1, key2, ..., colKey)`

大范围批量查询：

`ERead_Results(resVault, resID, filename, rowKeysRange, colKeysRange)`

`EProj_Results(resVault, resID, product, projKeysRange, timePeriodsRange, [simID])`

说明：

- `resultType` 是兼容保留参数，当前查询逻辑不使用。
- `simID` 省略时默认为 `0`。
- `ERead_Result` 要求 `filename` 后的坐标数量等于 `.fac` 维度数：先传所有行维度，最后传列维度；数量不匹配时返回 `#ERROR: ...`。
- FAC 复制或解析失败会显示具体错误（包括解析失败的行号）；只有坐标实际为 `*` 时才提示不支持通配符。空坐标可用于匹配 FAC 中的空行键。
- `ERead_Results` 返回二维动态数组；`rowKeysRange` 必须包含 `NumDims - 1` 列，`colKeysRange` 提供最后一维列 key。
- `EProj_Results` 返回二维动态数组；`projKeysRange` 必须有两列：`SP_CODE` 和 `VAR_NAME`，`timePeriodsRange` 提供期间列。
- XLAM 加载后，F9 会先只失效发生变化的 facReaderAddin DLL 内存缓存，再调用 Excel workbook 重算；Shift+F9 会先只失效发生变化的缓存，再只重算当前 sheet；Ctrl+Alt+F9 / Ctrl+Shift+Alt+F9 会先做同样的变化检查，再调用对应的 Excel 全量重算。`AppReaderClearCache` 仍可用于手工全清内存缓存。

## 迁移 / 分发

详细流程见：

- `docs/README.md`
- `docs/README_CN.md`
- `dist/facReaderAddin_README.md`

## 远程 `.fac` 缓存

当 `.fac` 位于 UNC 路径或映射网络盘时，`facReaderAddin.dll` 会先复制到本地缓存再解析，避免在网络盘上随机读取导致 Excel 重算很慢。

缓存目录：

```text
%LOCALAPPDATA%\facReaderAddin\fac-cache
```

缓存文件名由远程完整路径 hash、文件大小和修改时间组成；同目录下会生成同名 `.meta.json`，记录原始路径、本地缓存路径、文件大小、修改时间、缓存时间和最后访问时间。每次 F9 / Shift+F9 触发的重算会先对已加载的 `.fac` 各检查一次大小和修改时间；没变的继续复用内存表，变化的才失效并在重算首次使用时重新复制/加载。同一个远程源文件只保留最新版本的本地缓存；缓存目录默认最多保留 100GB，超过后按 `lastAccessAt` 删除最久未使用的旧缓存。
