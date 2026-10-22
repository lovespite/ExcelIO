# 04 · 脚本与运行时

本文定义 QuickExcel 的 JavaScript 执行边界、值协议、工作进程和操作脚本提交机制。
“当前”指 M0 原型中已经存在的代码；“目标”及所有标注为“拟议”的 API 都需要后续实现。
整体职责见 [系统架构](02-system-architecture.md)，文件持久化见 [文件格式与兼容性](03-file-format-and-compatibility.md)。

## 1. 当前实现与目标边界

当前实现入口是 [QuickJsSession](../../QuickExcel.Scripting/QuickJsSession.cs) 与 [原生桥接](../../QuickExcel.Scripting/native/quickexcel-js.c)。

| 项目 | M0 当前实现 | 目标实现 |
| --- | --- | --- |
| 引擎 | QuickJS-NG 0.17.0，Windows x64 原生 DLL | 继续使用固定版本、可复现构建的 QuickJS 引擎 |
| 生命周期 | 每个启用脚本的文档一个进程内会话 | 每个启用脚本的文档一个专用工作进程 |
| 串行化 | 托管锁串行调用，每次更新 QuickJS 栈顶 | 工作进程内单一执行线程、单一请求队列 |
| 执行上限 | 每次原生调用约 1 秒中断期限 | 保留引擎中断，增加宿主看门狗与进程终止 |
| 引擎内存 | 64 MiB QuickJS 分配上限 | 保留引擎限额，增加独立的进程资源预算 |
| JS 栈 | 256 KiB | 保持可配置、有版本的宿主策略 |
| 输入/输出 | 输入检查序列化字符串长度约 4 MiB；输出 4 MiB UTF-8 字节 | 按 UTF-8 字节严格限制 IPC 帧、嵌套深度与集合数量 |
| 调用结果 | 有限数字、文本、布尔、null；undefined 归一为 null | 带类型的 Number/Text/Boolean/Error/Blank 协议 |
| 区域参数 | 兼容层传递一维原始字符串数组 | v1 保持兼容，v2 显式传递二维区域形状 |
| 运行日志 | 最多保存 100 条，每条截断至 4,000 个 JS 字符 | 按请求关联、限流、可定位源码的结构化诊断 |
| 工作簿访问 | 未暴露读写对象 | UDF 只接收参数；操作脚本接收快照并生成变更集 |
| 脚本授权 | 打开不执行；用户应用后运行；源码变化停用旧会话 | 本地显式授权绑定文档会话、源码摘要与运行策略 |
| 操作脚本 | 尚未实现 | 变更集经文档命令队列原子提交，支持一次撤销 |

当前 `QuickJsSession` 已提供构造、`FunctionNames`、`Invoke`、`GetLogs`、`Dispose`。
当前 Core 将 JS 的 null 结果转为空字符串，将执行失败显示为 `#SCRIPT!`；这还不是完整的类型化错误模型。
当前没有 `CancellationToken`、版本化 IPC、工作进程恢复、脚本断点或 OS 访问隔离。
已有锁和 `JS_UpdateStackTop` 允许顺序切换托管线程，但不允许同一运行时并发执行。
核心不链接 `quickjs-libc`，未注册 `std`、`os`、文件、网络、模块加载器或 CLR 桥接。

## 2. 组件职责与依赖

```text
Desktop 脚本编辑器 / 运行记录
        │ 用户意图、取消请求
        ▼
DocumentSession / 文档命令队列  ← 唯一可提交文档变更的入口
        │ 不可变参数快照、revision、calculationGeneration、scriptGeneration
        ▼
CalculationCoordinator / ScriptCoordinator
        │ IScriptRuntime
        ├── InProcessQuickJsRuntime（M1 过渡适配器）
        └── WorkerScriptRuntime（M4 默认）
                   │ 有界、版本化 IPC
                   ▼
          QuickExcel.ScriptWorker
                   │ 串行请求、资源预算
                   ▼
            QuickJS 原生桥接
```

文档领域层持有工作簿、类型、命令历史和修订号；运行时不能拥有这些状态的权威副本。
计算协调器决定需要哪些输入、何时重算以及结果是否仍然有效；运行时只执行明确的请求。
脚本协调器管理源码、授权、注册表、工作进程和诊断，不直接写入 `XlWorkbook`。
工作进程仅接收必要数据的复制快照；禁止跨 IPC 传递托管对象引用、文件句柄或活动工作簿代理。
UI 不持有原生指针，也不在 UI 线程执行脚本、等待工作进程退出或序列化大块区域。

## 3. 脚本 API：先保持现有函数可用

当前 v1 接口继续受支持：

```javascript
quickExcel.functions.register("DISCOUNT", (price, rate) => price * (1 - rate));
quickExcel.functions.register("TOTAL", values =>
    values.reduce((sum, value) => sum + Number(value), 0));
```

对应公式为 `=DISCOUNT(B2,D2)`、`=TOTAL(B2:B4)`。
这里的 v1/v2 指 `scriptApiVersion`，与 `.xlsq` 的 `formatVersion` 独立；v2 文件可以继续声明脚本 API v1。
包升级为 `scripts/main.js` 与模块映射，不会自动改变函数注册写法或区域参数语义。
当前函数名只允许英文字母，不区分大小写，注册后转为大写；重复名称在同一脚本中报错。
不能在 M1 或 M2 提前宣传 `JS.DISCOUNT`：当前公式解析器尚未支持这种函数命名。
后续注册表必须明确定义与内置名称冲突的错误；不允许无提示覆盖 `SUM` 等内置函数。

以下描述符接口是拟议的 v2 API，必须等到 M3 支持命名空间解析后再开放：

```javascript
quickExcel.functions.register({
    name: "JS.DISCOUNT",
    apiVersion: 2,
    description: "计算折扣价",
    parameters: [{ name: "price", type: "number" }, { name: "rate", type: "number" }],
    returns: "number",
    profile: "strict"
}, (price, rate) => price * (1 - rate));
```

描述符用于校验参数个数、生成帮助与补全，并声明版本和执行策略；声明本身不能证明函数纯净。
v1 函数名保持可解析，不自动改写用户现有公式；命名空间迁移必须是可预览、可撤销的文档操作。
注册只在加载阶段发生，成功后注册表冻结；单元格求值期间不得添加、删除或替换函数。
第一阶段 UDF 只支持同步标量返回，不运行 Promise 任务循环，不支持异步 UDF 或动态数组溢出。
返回 Promise、对象、函数、Symbol、BigInt、NaN 或 Infinity 均应产生明确的类型错误。

## 4. 类型桥接与区域版本

目标 IPC 使用判别联合表示值，避免把 `"00123"`、`123`、空值和错误都压成字符串。

| 类型 | 拟议载荷 | JS v2 侧语义 |
| --- | --- | --- |
| Number | `{ "kind": "number", "value": 123.5 }` | 有限 IEEE 754 number |
| Text | `{ "kind": "text", "value": "00123" }` | string，禁止自动数值化 |
| Boolean | `{ "kind": "boolean", "value": true }` | boolean |
| Blank | `{ "kind": "blank" }` | null；与空字符串不同 |
| Error | `{ "kind": "error", "code": "DIV0" }` | 默认由宿主传播，不作为普通字符串调用 UDF |

错误显示文本由领域层根据错误码生成；用户输入的文本 `"#DIV/0!"` 不自动升级为错误值。
默认错误传播规则在调用 JS 前执行；若未来允许函数处理 Error，必须通过版本化描述符显式声明。
不得根据 JSON 对象具有某个字段就把任意 UDF 对象结果识别为宿主 Error。
日期在领域层采用数值序列、格式和工作簿 1900/1904 日期系统标志；此阶段不会隐式转换为 JS `Date`。
需要日期语义的函数显式接收序列和约定，后续辅助 API 再定义时区、闰年兼容与舍入规则。

v1 兼容模式保持当前区域行为：按既定顺序展开为一维原始字符串数组，由脚本自行转换。
此模式不能区分区域中的空单元格与空字符串，也不承诺恢复二维坐标，迁移时必须保留该限制。
v2 区域显式携带 `rows`、`columns` 与行主序 `values`，并可带只读的来源坐标描述。
例如 2×2 区域的拟议载荷为：

```json
{
  "kind": "range", "rows": 2, "columns": 2,
  "values": [
    {"kind":"number","value":1}, {"kind":"text","value":"02"},
    {"kind":"blank"}, {"kind":"boolean","value":true}
  ]
}
```

JS v2 适配器可提供只读 `range.get(row, column)` 与迭代方法；这些是快照视图，不是文档回调。
宿主必须校验行列乘积、载荷长度、最大单元格数、深度和编码字节数，避免在解码前分配超大数组。
区域 API 版本按函数描述符/脚本 API 版本显式选择，不根据参数形状猜测，不静默升级已保存的 v1 脚本。

## 5. 纯函数约定、确定性与缓存

UDF 的契约是仅依赖显式参数、返回一个值；没有修改文档、读取隐藏单元格或发起外部 I/O 的能力。
全局变量、闭包计数器、当前时间和随机数仍会使 JavaScript 结果依赖调用顺序或运行时状态。
不暴露工作簿 API 并不能从语言层面证明纯函数；冻结注册对象也不能证明整个 JavaScript 环境纯净。

目标严格执行策略必须明确：禁用或替换非确定时间/随机源，限制可变共享状态，约束动态代码入口。
若提供时间或随机数，使用宿主显式注入的计算批次时间/种子，并把它们纳入输入和缓存键。
严格策略需要专门的实现与反例测试；完成之前不宣称“任意 JS 函数天然确定”。
M3 必须先定义保守调度契约，再实现增量计算；该契约是 `CALC-003` 的前置条件，不能等到 M4 才决定。
在尚未证明 UDF 纯度时，含 JS 的工作簿每轮重建运行上下文，重新加载当前已授权源码，按稳定顺序完整重算。
即使某个 UDF 的显式输入没有变化，也不复用该 UDF 或其下游在上一轮的结果；不以局部脏标记跳过它们。
超时或能力不支持导致未完成时，保存结果仅供过期预览，不能算作该轮完整重算成功。
无 JS 的确定性公式子集可以使用增量调度；全量/增量等价证明仅覆盖受支持的确定性公式子集。
每轮重建消除跨轮闭包残留，但无法证明任意 JS 的数学纯度；同轮全局状态、时间和随机仍按上述策略处理。
M4 worker 实现继承此契约；严格策略的实现与反例测试未通过前，不开放 JS UDF 结果复用。
缓存新鲜度只表示给定源码、绑定、输入与执行策略下该轮计算已完成，不保证函数满足数学纯度。
后续缓存键至少包含函数身份、源码摘要、`scriptGeneration`、引擎版本、API 版本、策略及类型化输入。
未定义语义的非纯函数不参与公共子表达式合并、参数相同复用或跨文档缓存。
同一次重算中的执行顺序必须固定并可追踪；如果脚本依赖闭包状态，诊断应说明它不满足纯函数契约。

## 6. 运行时门面与生命周期

M1 引入拟议的 `IScriptRuntime`，先用适配器封装现有 `QuickJsSession`，避免计算层依赖 P/Invoke。
接口形状如下；具体类型与命名在 M1 契约测试中冻结，目前均未实现：

```csharp
public interface IScriptRuntime : IAsyncDisposable
{
    Task<ScriptLoadResult> LoadAsync(ScriptLoadRequest request, CancellationToken cancellationToken);
    Task<ScriptBatchResult> InvokeBatchAsync(ScriptBatchRequest request, CancellationToken cancellationToken);
    Task<ScriptActionResult> RunActionAsync(ScriptActionRequest request, CancellationToken cancellationToken);
}
```

取消令牌是门面契约；进程内适配器必须如实报告它只能在请求边界取消或依靠现有原生期限结束。
不得通过 `Task.Run` 外层取消后返回来假装原生执行已经停止，更不能让仍在执行的会话参与下一次计算。
工作进程实现通过中断信号和宿主终止能力完成真正的执行取消。

目标状态包括 `Disabled → Starting → Loading → Ready → Executing`，以及 `Faulted/Stopping/Disposed`。
源码编辑、函数绑定变化、启用状态变化、关闭文档均使旧注册表失效，取消排队任务，并提升 `scriptGeneration`。
加载成功后一次性发布新注册表；加载失败不保留部分注册的函数，也不默默回退执行旧源码。
脚本禁用时保留可展示的保存结果，但将其标记为未重算/过期；这不允许继续执行旧运行时。
每次重新打开文档生成新的文档会话 ID；进程重启生成新的 worker 实例 ID，防止旧回复混入新实例。

持久化与会话计数使用不同语义，字段不能因为数值相近而互换：

| 标识 | 生命周期与用途 |
| --- | --- |
| `scriptRevision` / `bindingRevision` | 文件中的源码集合/入口与绑定修订；内容命令推进，撤销也是新修订；见格式文档 |
| 文档 `Revision` / IPC `revision` | 当前会话唯一写入队列的输入修订，用于拒绝过期计算和动作；保存快照同时捕获各类文件修订 |
| `ScriptGeneration` / IPC `scriptGeneration` | 当前会话的执行状态代次；启停、加载失败等即使没有改变保存源码，也会使旧代次失效 |
| `CalculationGeneration` / IPC `calculationGeneration` | 当前会话的重算请求代次；同一文档修订再次重算也获得新代次 |

文件的 `contentRevision`、`scriptRevision`、`bindingRevision` 用于持久化内容和缓存来源核验，不能恢复运行授权。
重新打开时建立新的运行代次；不得直接把保存的 `scriptRevision` 当作 `ScriptGeneration`，或接受旧会话的消息。

## 7. 工作进程协议与资源预算

M4 默认采用每文档专用 `QuickExcel.ScriptWorker`，按需启动，关闭文档或停用脚本后释放。
工作进程内部只有一个 JS 执行队列；控制通道可请求中断，但不能从另一个线程直接操作 JS 对象。
初期采用专用继承管道或受当前用户 ACL 保护的本地管道，禁止在控制台文本中混合协议与日志。
协议采用定长长度前缀加 UTF-8 JSON；读取长度后先校验上限，再分配缓冲区。

所有消息使用如下拟议信封：

```json
{
  "protocolVersion": 1,
  "sessionId": "document-open-instance",
  "workerInstanceId": "worker-process-instance",
  "requestId": "unique-request",
  "revision": 42,
  "calculationGeneration": 17,
  "scriptGeneration": 3,
  "kind": "invokeBatch",
  "timeoutMs": 1000,
  "payload": {}
}
```

握手核验协议主版本、API 版本、引擎版本与策略；不支持的主版本立即失败，不能降级为无约束执行。
请求类型包含握手、加载、调用批次、运行操作、取消和关闭；回复包含成功、结构化失败和受限日志。
完整身份元组为 `(sessionId, workerInstanceId, requestId, revision, calculationGeneration, scriptGeneration)`。
所有回复原样关联该元组；宿主同时核对会话、worker、请求、输入修订、计算代次和脚本代次后才使用结果。
UDF 批次的 `calculationGeneration` 必填；动作、加载及非计算控制请求可为 null，并与对应请求严格匹配。
动作的 null 不免除 `requestId`、`revision`、`scriptGeneration` 与 worker/会话身份核验，不能作为通配值。
未知请求、重复响应、非法枚举、异常长度、截断数据或超量日志均是协议错误，关闭该 worker 并记录原因。
文档不接受 worker 主动发起的任意文件读取或命令执行请求；调用方向和消息种类均固定。

| 预算 | M4 初始目标 | 处理方式 |
| --- | --- | --- |
| 单 IPC 帧 | 4 MiB UTF-8 | 编码端与解码端都拒绝超限 |
| 排队请求 | 每文档最多 32 个 | 反压；合并尚未启动的重算，操作请求不静默丢失 |
| 单函数/加载 | 1 秒引擎预算 | 原生中断；宿主留出有界通信宽限 |
| 整批计算/操作 | 5 秒初始预算 | 耗尽即取消整批；不按函数个数无限放大 |
| 取消宽限 | 250 ms 初始目标 | 超期终止 worker；以压力测试校准最终值 |
| 引擎分配/栈 | 64 MiB / 256 KiB | 保留原生限制 |
| 进程内存 | 256 MiB 初始测量预算 | 包含 .NET 与原生开销；在 M4 实测后冻结发布策略 |

`timeoutMs` 是宿主许可内的预算，不接受工作簿任意提高上限；启动/握手另有明确超时。
大区域按计算计划分块；整块语义要求完整输入的 UDF 超限时应报错，不能暗中拆开求值改变结果。
批次未启动时以新修订替代旧重算；正在执行的过期批次请求取消，其返回值无论成功与否均不提交。

## 8. UDF 结果提交与崩溃恢复

计算开始时捕获文档 `revision`、`calculationGeneration`、`scriptGeneration`、函数注册表快照与显式输入。
工作进程返回结果包后，计算协调器核对完整身份元组，领域层在命令队列中再次核对全部代次后发布。
同一 revision 下后发重算也会使前一次计算代次过期；不能只比较输入修订和脚本代次。
过期结果被丢弃并安排新的计算；不能凭单元格坐标相同就覆盖用户后来的编辑。
计算结果更新与用户编辑的修订规则由领域层统一定义，避免计算自身无限触发重算。

进程退出、管道断开或协议失效时，待处理请求全部结束为受控错误，UI 和文档会话保持可用。
保存和编辑不依赖 worker 存活；未提交的计算结果和操作变更集全部丢弃。
恢复时重建空运行时，仅在当前源码的本地授权仍有效时加载相同 generation 的源码。
原请求不自动重放；用户发起重新计算后产生新的请求，操作脚本必须重新显式运行。
记录连续失败次数并限制重启，避免恶意初始化脚本形成自动崩溃循环。
引擎异常或宿主强制终止后不复用原生句柄，不把“超时后曾能调用”当作所有故障均可安全恢复的证明。

## 9. 操作脚本：先生成变更集，再统一提交

操作脚本与 UDF 使用不同入口；UDF 永远不能获取操作写入能力。
拟议接口是 `quickExcel.actions.register(name, action)`，由用户点击运行，不作为单元格公式调用。
UDF 和操作使用分离的 context/会话装载源码，避免操作的全局状态污染公式计算环境。
每种上下文只保留当前 `scriptBindings.kind` 对应的注册；另一种声明被注册门面忽略，宿主永不在该上下文调用其函数体。
同一模块可声明两种注册，但宿主门面按上下文分配能力；UDF 门面不能取得动作调用入口或 patch builder。
入口接收带 `baseRevision` 的只读工作簿快照和有界 patch builder；两者均不连接活动工作簿。

```javascript
// 拟议 API，当前原型尚不支持。
quickExcel.actions.register("规范商品名", ({ snapshot, patch }) => {
    const values = snapshot.getRange("sheet-stable-id", "A2:A10");
    patch.setValues("sheet-stable-id", "A2:A10",
        values.map(row => row.map(value => String(value ?? "").trim())));
});
```

第一阶段动作仅支持受限的单元格批量赋值/公式写入，不开放任意对象路径或任意反射式命令。
返回结果使用完整 IPC 身份元组，并含 `baseRevision`、操作 ID 和变更列表；工作进程只能提出变更，不能提交。
领域层校验稳定 sheet ID、区域边界、值类型、公式长度、总格数、重复/重叠写入及授权来源。
文档命令队列取出动作时再次检查 `currentRevision == baseRevision`；不匹配则拒绝整个变更集。
初期不自动合并冲突，也不在后台重复运行脚本；UI 提供重新运行，以新快照重新计算建议变更。
校验通过后以一个事务提交，先构造反向命令，再一次性更新文档、修订号和依赖失效集合。
整个动作对应一次撤销；Undo/Redo 重放已记录的正反向数据，不重新执行 JavaScript。
中途异常、取消、超时、worker 崩溃或验证失败均不产生部分文档写入。
首版操作仍为同步脚本；异步操作、网络请求、外部服务凭据与权限代理留到单独版本设计。
M4 动作持久化使用 `formatVersion: 2`、`scriptApiVersion: 2`，并在 `requiredFeatures` 加入 `script-actions`。
动作绑定采用 `{ "module": "main", "scope": "workbook", "kind": "actions" }`，其模块必须已在内容映射和入口列表声明。
`formula-functions` 仍是公式绑定 kind；一个模块需要两种能力时分别声明绑定，由宿主分别装载到相应上下文。
只有公式的 v2 文件继续声明脚本 API v1 是合法情况；不能因支持动作而强制改写所有原有文档的 API 版本。
不支持 `script-actions` 的读取器按必需能力规则拒绝完整文档模式，不能把动作误当公式入口或静默执行。
保存后重开必须恢复动作源码、绑定和帮助信息，但保持脚本禁用；用户再次启用且显式运行后才生成 patch。

## 10. 信任与安全边界

`.xlsq` 中保存源码和兼容性元数据，不保存可使宿主自动授权的有效许可。
manifest 中的作者、来源、能力声明或 `trusted` 字段不能赋予文件、网络或脚本执行权限。
打开、预览、导入、保存、导出、索引和生成缩略图均不得隐式执行嵌入脚本。
启用操作由本地用户明确发起，绑定源码摘要；源码变化撤销已有授权并停用旧代码。
在禁用状态展示已有缓存结果与明确的过期状态，不通过显示单元格间接启动脚本。

进程隔离提供崩溃隔离和可终止执行；同权限进程并不因此失去用户文件或网络访问权限。
Job Object 用于进程树、资源限额和关闭时清理；它本身不是阻止文件或网络 I/O 的访问策略。
Beta 威胁模型必须包括恶意源码、畸形消息、原生引擎漏洞、超量资源以及被攻陷 worker 的系统调用。
M4 实现受限令牌、最小继承句柄、Job Object 与文件系统/网络 OS 访问策略的隔离原型，并完成开发级探测。
M6 对该策略执行威胁模型复核、故障与越界测试、净环境安装验证，按质量门槛决定 Beta 放行与安全声明。
文件策略需要显式的可读运行库与最小临时目录范围；网络策略需要独立验证回环、DNS 和外网访问。
可评估 AppContainer/能力模型或其他 Windows 隔离机制；选择必须经过部署权限与兼容性验证。
不得把降低完整性级别、删去 JS API、设置 Job Object 或运行普通子进程称为完整沙箱。
OS 隔离验收未通过时，只能继续作为“可信脚本试用”候选，不能算作完成 M6 Beta 门禁。
不得通过弱化 Beta 标记绕过门禁，也不得承诺可安全运行任意第三方工作簿脚本。
安全控制不能依赖脚本自报能力；所有预算、授权和最终提交判定都由宿主决定。

## 11. 诊断、调试与源码保存

目标诊断模型区分 Syntax、Registration、ArgumentType、ResultType、Execution、Timeout、Cancelled、
MemoryLimit、WorkerExit、Protocol、StaleResult 与 ActionConflict，并保留稳定错误码。
UDF 失败映射为领域层 Error 值，显示 `#SCRIPT!` 或已有兼容错误；详细信息在运行记录中展开。
取消和过期结果通常不改写单元格为错误，而是保留标记过期的结果并等待重算。
每条诊断携带完整请求身份元组、函数名、单元格地址或动作 ID、源码文件与行列。
v1 单脚本使用 `scripts/functions.js`；v2 使用 manifest 映射的 `scripts/main.js`、`scripts/lib/*.js` 等稳定包内路径。
宿主包装代码的行偏移必须转换回对应用户模块源码，不能把所有模块都报为 `functions.js`。
首期调试提供编辑器定位、调用参数摘要、耗时、受限 console 日志与可复制诊断；不承诺原生断点调试。
日志按请求限条数、单条字节数和总量；含敏感单元格的数据只保存在本地会话，不默认上传。
异常对象转字符串也属于预算内执行；避免 getter/toString 再次执行导致无限异常或日志攻击。

`.xlsq` 仅保存 JavaScript 源码，不保存 QuickJS 字节码、函数指针、原生对象或运行时堆快照。
引擎来源、MIT 许可证与源码 SHA-256 见 [运行时来源说明](../../QuickExcel.Scripting/native/SOURCES.md)。
版本升级必须更新引擎摘要、许可证清单与回归结果；保存的缓存标注引擎/API/脚本版本，不冒充新引擎结果。
M2 只保存和验证 v2 的 `scripts.entryModules`、`scripts.modules` 内容映射；M4 才实现受控包内模块加载器。
模块 import 只解析 manifest 已声明并校验摘要的包内逻辑 ID，入口按清单顺序加载；注册失败整体回滚。
包内用户源码模块属于目标能力；外部包管理器、自动安装第三方依赖和运行时下载模块不在首期范围。

## 12. 实施任务与依赖

| 任务 | 阶段 | 交付与完成条件 |
| --- | --- | --- |
| JS-001 | M0 当前 | 固定引擎、原生限额、v1 注册/调用及 smoke；保留事实基线 |
| JS-002 | M1 领域/会话 | IScriptRuntime 门面与进程内适配器；领域层不再依赖 QuickJsSession |
| JS-003 | M1 领域/会话 | 会话/请求身份、revision、两类 generation 与文件脚本/绑定修订分离；变更撤销旧请求 |
| JS-004 | M2 文件 | 持久化脚本 API/兼容元数据；打开不运行；迁移不导入授权；只存源码 |
| JS-005 | M3 计算 | 类型化值与 Error 传播、v1 区域兼容、v2 二维区域契约测试 |
| JS-006 | M3 计算 | 命名空间解析完成后开放描述符；注册冲突、元数据和代次失效规则 |
| JS-007 | M4 worker/脚本/动作 | 专用 worker、包内模块加载、握手、帧限额、串行执行、反压与完整身份核验 |
| JS-008 | M4 worker/脚本/动作 | 取消、中断、看门狗、终止、退出处理与有界恢复；桌面不崩溃 |
| JS-009 | M4 worker/脚本/动作 | 快照动作、分离上下文、patch/CAS/撤销；动作绑定保存重开不执行的往返验收 |
| JS-010 | M3 契约 / M4 实现 | M3 在 CALC-003 前定义含 JS 每轮重建与完整重算契约；M4 落地 worker 严格策略及反例测试 |
| JS-011 | M5 桌面体验 | 显式启用、过期状态、运行/停止、源码定位、错误及日志界面 |
| JS-012 | M6 Beta | 复核 M4 隔离策略；执行 OS I/O 越界与净环境质量放行测试，确定发布声明 |
| JS-013 | M6 Beta | 引擎升级流程、许可证/来源清单、崩溃与性能回归、净环境安装验证 |
| JS-014 | M4 worker/脚本/动作 | 受限令牌、句柄、Job Object 及独立文件/网络 OS 策略原型与开发级验证 |

M4 的动作提交依赖 M1 命令队列和撤销模型，typed bridge 依赖 M3 单元格类型；不能提前绕过领域层落地。
JS-010 的 M3 契约先于 CALC-003；动作持久化与文件层的 `script-actions` 必需能力和 API v2 校验共同交付。
任务排期和跨模块依赖以 [实施计划](05-implementation-plan.md) 为准。

## 13. 验收与回归清单

| 场景 | 必须观察到的结果 |
| --- | --- |
| 现有 DISCOUNT/TOTAL 文档 | v1 行为兼容；名称/参数错误能定位；Unicode 正常 |
| 文本 00123、空字符串、Blank、Error、布尔 | 类型不互相猜测；错误按规则传播；v1/v2 差异有固定样例 |
| 2×3 区域与空区域 | v2 保留形状与次序；维度不一致/超限在执行前拒绝 |
| 无限循环、深递归、超量分配、异常 toString | 在预算内中断或终止；UI 可用；无无限等待 |
| 计算期间编辑单元格、源码或同修订再次重算 | 旧 revision/calculationGeneration/scriptGeneration 回复被丢弃，不能覆盖新输入或新计算 |
| 重新打开、启停脚本、修改绑定与恢复源码 | 文件 scriptRevision/bindingRevision 不充当会话代次或授权；旧会话回复失效 |
| 取消后迟到回复、重复回复、worker 重启旧消息 | 全部按身份拒绝；新请求不接收旧日志或旧函数结果 |
| 半帧、超大长度、未知协议版本 | 受控协议失败，不先分配超大内存，不运行载荷 |
| worker 原生崩溃/强杀 | 文档可编辑和保存；未提交 patch 消失；恢复不自动重放动作 |
| 操作执行中用户编辑 | baseRevision 冲突，整批不落地；重新运行使用新快照 |
| 操作成功后撤销/重做 | 一次撤销恢复全部旧值；不重新执行脚本；依赖只统一失效一次 |
| 动作源码/绑定保存后重开 | 格式 v2、API v2、script-actions 与 actions 绑定完整往返；打开和恢复不执行 |
| 同模块包含 UDF 与动作注册 | 分离上下文只保留当前 kind；UDF 无 patch 能力，不调用被忽略的动作体 |
| 打开含自称 trusted 的 xlsq | 不执行源码，不继承权限，不发起外部请求 |
| 时间、随机数、可变闭包及显式输入未变的 UDF | 每轮重建上下文并稳定顺序完整重算；UDF/下游不复用旧结果；等价证明不覆盖任意 JS |
| 文件/网络/子进程/句柄越界探测 | 通过 OS 层测试验证实际策略，不能仅测试 typeof API 不存在 |
| OS 隔离验收失败 | M6 Beta 门禁未完成，只能保留可信脚本试用候选 |
| 多文档与大量调用 | 文档间运行时和日志隔离；队列有界；批量 IPC 性能达到约定预算 |

M0 smoke 已覆盖标量/数组、Unicode、超时、内存/栈限制、中断后调用和顺序切换线程。
这些测试不能替代 worker 协议、事务一致性或 Beta OS 隔离验收。
完整测试分层与性能口径见 [质量、性能与交付](06-quality-performance-and-delivery.md)。
