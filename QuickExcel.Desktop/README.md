# QuickExcel 0.1 原型

Windows x64 桌面表格，基于 WPF、ExcelIO 和实际的 QuickJS 引擎。专属扩展名为 **`.xlsq`**。

## 开始试用

在仓库根目录双击 `Start-QuickExcel.cmd`，或直接运行发布目录中的 `artifacts/QuickExcel/QuickExcel.exe`。当前发布包依赖 .NET 10 Windows Desktop Runtime；本机已经安装。复制给另一台机器时，需要复制整个发布目录并安装相应运行时。

1. 启动后会展示示例工作簿和 JavaScript 源码。
2. 点击右侧 **应用脚本并重算**（Ctrl+Enter）。
3. 双击价格、数量等单元格修改数据，观察自定义公式和合计更新。也可以在公式栏输入内容并按 Enter。
4. 修改脚本，再点击应用；支持同步的数字、字符串、布尔值或空结果。
5. **保存 .xlsq**（Ctrl+S）保存数据、公式、缓存结果和脚本。重新打开时脚本默认停用，确认源码后可再次应用。
6. **导出 .xlsx** 将所有公式转换为当前缓存结果，供其他表格软件查看。原 `.xlsq` 文档不受影响。

## 脚本 API

```javascript
quickExcel.functions.register("DISCOUNT", (price, rate) => price * (1 - rate));
quickExcel.functions.register("DOUBLE", value => value * 2);
```

单元格公式：`=DISCOUNT(B2,D2)`、`=DOUBLE(A1)`、`=SUM(E2:E5)`。

函数名使用英文字母，不区分大小写。首版仅支持同步标量返回；不支持 Promise、对象、动态数组溢出。自定义函数只应使用显式参数；脚本不提供工作簿写入、文件系统、网络或 .NET 反射接口。运行时设有执行时间和内存限制，但此原型没有提供操作系统级进程沙箱。

脚本源代码在文档内保存，打开文件不会自动执行它。不要启用不了解的脚本。编辑脚本后会停用旧版本，点击应用后才执行新版本。

## 编辑操作

- Ctrl+O 打开；Ctrl+N 新建；Ctrl+S 保存；Ctrl+Shift+S 另存为。
- 表格中 Ctrl+C 复制当前值、Ctrl+V 粘贴制表符分隔的矩形数据。
- Delete 清空选中内容；Ctrl+Z / Ctrl+Y 撤销或重做单元格编辑和粘贴。
- 脚本编辑区使用自己的文本撤销历史。
- 底部可切换/新增工作表，以及扩展可见行列。

## 文件格式

`.xlsq` 是一个 ZIP 包：

```text
manifest.json
workbook.xlsx
scripts/functions.js
```

manifest 记录格式标识和版本。只保存 JavaScript 源码，不保存 QuickJS 字节码。脚本启用状态不会写入文档。保存使用同目录临时文件再替换，避免直接截断现有文件。

## 原型边界

- xlsx 导入导出支持现有 ExcelIO 模型覆盖的特性，并不保证任意 Excel 文件无损往返。请使用工作副本试用。
- 当前底层模型以字符串保存单元格值；导出的字面值沿用文本类型，尚未实现数字、日期与文本的完整类型保真。
- 导出 xlsx 是**当前结果快照**：所有公式转为值。脚本停用时可能包含旧的缓存结果，请先应用脚本并检查运行记录。
- 支持同表公式依赖重算。跨表计算、完整 Excel 函数语义、合并单元格/图片在界面中的呈现、高级数字格式、筛选和图表不在本原型验收范围。
- 表格界面最多显示/编辑 100,000 行 × 256 列，单次粘贴不超过 20,000 格。导入模型中界面未展示的受支持内容仍随文档保存。
- 当前撤销仅覆盖单元格内容，不覆盖新增工作表和整个脚本应用操作。
- 脚本批量修改工作簿、异步函数和插件机制待后续版本。

## 构建与验证

```powershell
dotnet build .\QuickExcel.Desktop\QuickExcel.Desktop.csproj -v minimal
dotnet test .\QuickExcel.Core.Test\QuickExcel.Core.Test.csproj -v minimal
.\scripts\build-quickexcel.ps1
# GUI 集成验收（创建测试输出并自动退出）：
.\artifacts\QuickExcel\QuickExcel.exe --smoke-test .\artifacts\quickexcel-smoke
```

首次从源码构建原生引擎所需的工具及版本见 `QuickExcel.Scripting/native` 下的说明。独立的 `QuickExcel.slnx` 不依赖 Web/WASM 工作负载；原有 Web 代码保留。
