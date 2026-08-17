# facReaderAddin Excel 加载项使用说明

## 安装

1. 将发布包中的文件复制到：
   `C:\Tools\OpenAct\Addins`
2. 发布包包含：
   - `facReaderAddin.dll`
   - `facReaderAddin.xlam`
   - `facReaderAddinX.xll`
   - `facReaderAddinXControl.xlam`
   - `facReaderAddin_README.md`
3. XLAM 方式：打开 Excel -> 文件 -> 选项 -> 加载项。
4. 底部“管理：Excel 加载项” -> 转到...
5. 浏览并勾选：
   `C:\Tools\OpenAct\Addins\facReaderAddin.xlam`
6. XLL 方式：在同一界面浏览并勾选：
   `C:\Tools\OpenAct\Addins\facReaderAddinX.xll`
   并同时勾选：
   `C:\Tools\OpenAct\Addins\facReaderAddinXControl.xlam`

不要在同一个 Excel 进程里同时加载完整 `facReaderAddin.xlam` 和 `facReaderAddinX.xll`，因为两者暴露相同函数名。`facReaderAddinXControl.xlam` 不暴露 UDF，只负责 XLL 模式下的 F9/Shift+F9 刷新绑定，可以和 `facReaderAddinX.xll` 同时加载。

## 函数用法

### EProj_Result：按预测常用字段读取单个值

`=EProj_Result(resVault, resID, product, spCode, resultType, variable, timePeriod, [simID])`

示例：

`=EProj_Result($A$1,$B$1,"PHKL_0",0,"","DISC_A_PC",202512,0)`

说明：

- `resVault`：结果 vault 完整路径，例如 `D:\result\rbc2512_4`
- `resID`：结果子目录，例如 `E_1`
- `product`：`.fac` 文件名，可带或不带 `.fac`
- `spCode`：`SP_CODE`
- `resultType`：兼容保留参数，当前不参与匹配
- `variable`：`VAR_NAME`
- `timePeriod`：期间列，例如 `202512`、`2025` 或 Excel 日期
- `simID`：可选，默认 `0`
- XLAM 和 XLL 都支持该函数。

### ERead_Result：按 fac 表维度顺序读取单个值

`=ERead_Result(resVault, resID, filename, key1, key2, ..., colKey)`

示例：

`=ERead_Result($A$1,$B$1,"TEST_ACCUM",0,"TEST_ACCUM",0,"DISC_SH_NET_BT_A",202512)`

说明：

- `filename`：`.fac` 文件名，可带或不带 `.fac`
- `key1...colKey`：必须严格按 `.fac` 维度顺序传入；先传所有行维度，最后传列维度
- 坐标数量不等于 `.fac` 维度数时返回 `#ERROR: ...`
- XLAM 和 XLL 都支持该函数。

### ERead_Results：按 fac 表维度顺序批量读取

`=ERead_Results(resVault, resID, filename, rowKeysRange, colKeysRange)`

示例：

`=ERead_Results($A$1,$B$1,"TEST_ACCUM",A5:D100004,F4:DA4)`

说明：

- `rowKeysRange`：每一行是一组行维度坐标，列数必须等于 `.fac` 的 `NumDims - 1`
- `colKeysRange`：最后一维列 key，可以是一行或一列
- 返回二维动态数组；单个坐标不存在时返回 `#N/A`

Excel 填法：

- 如果单值函数原来是 `=ERead_Result(..., key1, key2, key3, key4, colKey)`，那么 `ERead_Results` 里 `rowKeysRange` 就要横向放 `key1:key4` 这 4 个行 key。
- 每一行代表一次查询的行坐标；有多少组 row key，就填多少行。
- `colKeysRange` 放要展开读取的列 key，例如期间列、金额列等。

示例布局：

| A | B | C | D | F | G |
|---|---|---|---|---|---|
| SIM_ID | NAME | SP_CODE | VAR_NAME | 202412 | 2025 |
| 0 | TEST_ACCUM | 0 | DISC_SH_NET_BT_A |  |  |
| 0 | TEST_ACCUM | 0 | SEG_CFL(1) |  |  |

公式填在 `F2`：

`=ERead_Results($A$1,$B$1,"TEST_ACCUM",A2:D3,F1:G1)`

Office 365 会自动 spill 到 `F2:G3`。

Office 2016 不支持动态数组 spill，需要这样输入：

1. 先选中完整结果区域 `F2:G3`
2. 输入公式 `=ERead_Results($A$1,$B$1,"TEST_ACCUM",A2:D3,F1:G1)`
3. 按 `Ctrl + Shift + Enter` 确认为数组公式

### EProj_Results：按 SP_CODE / VAR_NAME 批量读取

`=EProj_Results(resVault, resID, product, projKeysRange, timePeriodsRange, [simID])`

示例：

`=EProj_Results($A$1,$B$1,"PHKL_0",A5:B100004,C4:CX4,0)`

说明：

- `projKeysRange`：必须有两列，依次为 `SP_CODE`、`VAR_NAME`
- `timePeriodsRange`：期间列，可以是一行或一列
- 返回二维动态数组；单个坐标不存在时返回 `#N/A`

Excel 填法：

- `EProj_Results` 是 `EProj_Result` 的批量版，row key 固定简化为两列：`SP_CODE`、`VAR_NAME`。
- 每一行是一组 `SP_CODE + VAR_NAME`。
- `timePeriodsRange` 放要横向展开的期间。

示例布局：

| A | B | C | D |
|---|---|---|---|
| SP_CODE | VAR_NAME | 202412 | 2025 |
| 0 | DISC_A_PC |  |  |
| 0 | DISC_B_PC |  |  |

公式填在 `C2`：

`=EProj_Results($A$1,$B$1,"PHKL_0",A2:B3,C1:D1,0)`

Office 365 会自动 spill 到 `C2:D3`。

Office 2016 不支持动态数组 spill，需要这样输入：

1. 先选中完整结果区域 `C2:D3`
2. 输入公式 `=EProj_Results($A$1,$B$1,"PHKL_0",A2:B3,C1:D1,0)`
3. 按 `Ctrl + Shift + Enter` 确认为数组公式

## 常见问题

- 如果 Excel 提示 DLL 或加载项错误，请确认 `facReaderAddin.dll` 位于 `C:\Tools\OpenAct\Addins\facReaderAddin.dll`。
- 如果函数不存在，请确认已启用 `facReaderAddin.xlam`。
- 大范围读取优先使用 `ERead_Results` 或 `EProj_Results`，不要铺大量单格公式。
- Office 2016 不支持动态数组 spill，`_Results` 函数必须先选中完整输出区域，再用 `Ctrl + Shift + Enter` 输入数组公式。
- XLL 当前只提供单值函数 `ERead_Result` 和 `EProj_Result`；批量 `_Results` 函数请使用完整 XLAM。
- XLL 需要 64 位 Excel，并要求 `facReaderAddinX.xll` 与 `facReaderAddin.dll` 位于同一目录。
- XLL 模式如需 F9 / Shift+F9 刷新语义，请同时加载 `facReaderAddinXControl.xlam`；它只绑定快捷键，不注册同名工作表函数。
- 如果文件来自网络下载，先在文件属性中点击“解除锁定”；如被信任中心阻止，请把 `C:\Tools\OpenAct\Addins` 加入 Excel 受信任位置。
- XLAM 加载后，F9 会先只失效发生变化的 facReaderAddin DLL 内存缓存，再调用 Excel workbook 重算；Shift+F9 会先只失效发生变化的缓存，再只重算当前 sheet；Ctrl+Alt+F9 / Ctrl+Shift+Alt+F9 会先做同样的变化检查，再调用对应的 Excel 全量重算。`AppReaderClearCache` 仍可用于手工全清内存缓存。
- 远程 `.fac`（UNC 路径或映射网络盘）会先复制到 `%LOCALAPPDATA%\facReaderAddin\fac-cache` 再解析；同目录 `.meta.json` 记录原始路径、大小、修改时间和最后访问时间。每次 F9 / Shift+F9 触发的重算会先对已加载的 `.fac` 各检查一次大小和修改时间；没变的继续复用内存表，变化的才失效并在重算首次使用时重新复制/加载。同一个远程源文件只保留最新版本的本地缓存；缓存目录默认最多保留 100GB，超过后按 `lastAccessAt` 删除最久未使用的旧缓存。
- 如果 workbook 位于 OneDrive/SharePoint 且 Excel AutoSave 为开启状态，刷新 volatile UDF 后 Excel 可能自动保存 workbook；这是 Excel AutoSave 行为，不是 facReaderAddin 主动保存。
