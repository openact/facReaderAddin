# facReaderAddin Excel 查询说明

## 目的

`facReaderAddin` 提供面向 Excel 的 DLL 函数，用于从 `.fac` 结果文件中进行单值和大范围批量查询。

## 函数签名

在 VBA（`facReaderAddin.bas`）中：

`EProj_Result(resVault, resID, product, spCode, resultType, variable, timePeriod, [simID])`

通用按表维度顺序查询：

`ERead_Result(resVault, resID, filename, key1, key2, ..., colKey)`

大范围批量查询：

`ERead_Results(resVault, resID, filename, rowKeysRange, colKeysRange)`

`EProj_Results(resVault, resID, product, projKeysRange, timePeriodsRange, [simID])`

- `resVault`：vault 完整路径（例如 `...\\rbc2512_4`）
- `resID`：子目录（例如 `E_1`）
- `product`：文件名，可带或不带 `.fac`
- `spCode`：SP_CODE 值
- `resultType`：兼容保留参数（当前查询逻辑不使用）
- `variable`：VAR_NAME 值
- `timePeriod`：用于匹配 `.fac` 列名的期间键
- `simID`：可选 SIM_ID（默认 `0`）
- `filename`：文件名，可带或不带 `.fac`
- `key1, key2, ..., colKey`：按 `.fac` 维度顺序传入的精确匹配坐标；先传所有行维度，最后传列维度

## timePeriod 匹配规则

`timePeriod` 在查询前会先经过 `facReaderAddin.bas` 的归一化处理。

1. 日期类输入（Excel 日期序列值，或 `2025/12/31` 这类日期文本）-> `yyyymm`（例如 `202512`）
2. 数字/文本 `YYYYMM` -> 保持 `YYYYMM`
3. 数字/文本 `YYYY` -> 保持 `YYYY`
4. `*` 或空值 -> 视为无效（返回 `#VALUE!`）

### 说明

- 如果 `.fac` 的列名是 `2025`，Excel 输入 `2025` 可以直接命中。
- 如果 `.fac` 的列名是 `202512`，输入 `2025/12/31` 会在归一化后命中。
- 行维度参数（如 `resID`、`product`、`variable`）与 `timePeriod` 列匹配均不区分大小写。

## 查询行为

- 按表结构中的行维度做精确匹配：
  - `NAME/PROD_NAME/PRODUCT/PRODUCT_NAME` <- `product`
  - `SP_CODE` <- `spCode`
  - `VAR_NAME/VARIABLE/VARIABLE_NAME` <- `variable`
  - `SIM_ID/SIMULATION` <- `simID`（空值时默认 `0`）
- `RESULT_TYPE` 在当前查询逻辑中被有意忽略。
- `ERead_Result` 严格按表维度顺序读取。`filename` 后面的坐标参数个数必须等于 `.fac` 维度个数；否则公式所在单元格返回 `#ERROR: ...` 文本。
- FAC 文件复制或解析失败时，单值和批量查询返回具体加载错误（解析错误含行号）；只有坐标确实为 `*` 时才显示通配符错误，空坐标可正常匹配。
- `ERead_Result` 已改为一次 DLL 查询，不再先 `Found` 再读取。
- `ERead_Results` 和 `EProj_Results` 用于大块结果读取，返回二维动态数组；单个坐标不存在时返回 `#N/A`。
- XLAM 加载后，F9 会先只失效发生变化的 facReaderAddin DLL 内存缓存，再调用 Excel workbook 重算；Shift+F9 会先只失效发生变化的缓存，再只重算当前 sheet；Ctrl+Alt+F9 / Ctrl+Shift+Alt+F9 会先做同样的变化检查，再调用对应的 Excel 全量重算。`AppReaderClearCache` 仍可用于手工全清内存缓存。

## 远程 `.fac` 缓存

当 `.fac` 位于 UNC 路径或映射网络盘时，`facReaderAddin.dll` 会先复制到本地缓存再解析，避免在网络盘上反复随机读取。

缓存目录：

`%LOCALAPPDATA%\facReaderAddin\fac-cache`

缓存文件名由远程完整路径 hash、文件大小和修改时间组成；同目录下会生成同名 `.meta.json`，记录原始路径、本地缓存路径、文件大小、修改时间、缓存时间和最后访问时间。每次 F9 / Shift+F9 触发的重算会先对已加载的 `.fac` 各检查一次大小和修改时间；没变的继续复用内存表，变化的才失效并在重算首次使用时重新复制/加载。同一个远程源文件只保留最新版本的本地缓存；缓存目录默认最多保留 100GB，超过后按 `lastAccessAt` 删除最久未使用的旧缓存。

如果 workbook 位于 OneDrive/SharePoint 且 Excel AutoSave 为开启状态，刷新 volatile UDF 后 Excel 可能自动保存 workbook；这是 Excel AutoSave 行为，不是 facReaderAddin 主动保存。

## 批量函数用法

### ERead_Results

`rowKeysRange` 每一行是一组行维度坐标，列数必须等于 `.fac` 的 `NumDims - 1`；`colKeysRange` 是最后一维列 key，可以是一行或一列。

示例：

`=ERead_Results($A$1, $B$1, "TEST_ACCUM", A5:D100004, F4:DA4)`

### EProj_Results

`projKeysRange` 必须有两列：`SP_CODE` 和 `VAR_NAME`；`timePeriodsRange` 是期间列，可以是一行或一列。

示例：

`=EProj_Results($A$1, $B$1, "PHKL_0", A5:B100004, C4:CX4, "0")`

## 部署

推荐使用：

`deploy.ps1 -Target all`

也可以拆分执行：

1. `build/dll/build.ps1`
2. `build/excel_xlam/deploy.ps1`
3. `build/excel_xll/build.ps1`

发布产物输出到 `facReaderAddin/dist`：

- `facReaderAddin.dll`
- `facReaderAddin.xlam`（自动打包；若环境受限则需手工生成）
- `facReaderAddinX.xll`
- `facReaderAddinXControl.xlam`（XLL 模式下只负责 F9/Shift+F9 刷新绑定，不暴露 UDF）
- `facReaderAddin_README.md`（用户版说明）

## 三阶段流程（开发者 / 分发 / 用户）

### 阶段 1：开发者

1. 运行 `deploy.ps1 -Target all`，生成 `facReaderAddin/dist` 产物。
2. 将 `dist` 目录中的文件复制到 `C:\Tools\OpenAct\Addins`（不要复制 `dist` 文件夹本身）。
3. 若自动打包失败，使用源码目录中的 `build/excel_xlam/facReaderAddin.bas` 手工生成 `facReaderAddin.xlam`。

### 阶段 2：分发

对外分发最小包：

- `facReaderAddin.dll`
- `facReaderAddin.xlam`
- `facReaderAddinX.xll`
- `facReaderAddinXControl.xlam`
- `facReaderAddin_README.md`

### 阶段 3：用户

1. 用户将 `facReaderAddin.dll` 放到：
   `C:\Tools\OpenAct\Addins\facReaderAddin.dll`
2. XLAM 模式：Excel -> 选项 -> 加载项 -> 管理“Excel 加载项”-> 浏览并勾选 `facReaderAddin.xlam`。
3. XLL 模式：浏览并勾选 `facReaderAddinX.xll` 和 `facReaderAddinXControl.xlam`；不要同时加载完整 `facReaderAddin.xlam`。
3. 在任意工作簿调用 `EProj_Result(...)`、`ERead_Result(...)`、`ERead_Results(...)` 或 `EProj_Results(...)`。

说明：

- `.bas` 仅用于开发者手工生成 xlam，不进入最终用户分发包。
- 运行时依赖 `Lib` 声明中的 DLL 路径。
- 自动打包 xlam 依赖本机安装 Excel，且可能需要在 Excel 信任中心启用“信任对 VBA 工程对象模型的访问”。
