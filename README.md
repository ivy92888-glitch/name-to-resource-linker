# name-to-resource-linker / 名称匹配资源链接器

> 零依赖的 VBA 宏：对源表里每个自由文本品名，从一张或多张可配置的参考表里找出**最匹配的资源记录**，写回最多 N 条带可点击超链接的排序链接。
> A dependency-free VBA macro that, for each free-text item name, finds the best-matching resource records from reference tables and writes up to N ranked, hyperlinked links back into the sheet.

> **非官方。** 与 Microsoft（Excel）及任何电子表格厂商无隶属、赞助或背书关系；"workbook" 仅作为表格文件的通用指代。

---

## 痛点：它解决什么问题（中文）

另一类高频工作，形状也高度一致：**你有一堆自由文本名称，想给每一个挂上「最合适」的支持资源——文档、供应商页面、知识库文章、规格书——并按价格或相关度排好序**。

典型例子：

- 给每个请购物料贴上「最便宜的供应商链接」；
- 给每张工单推荐「最相关的知识库文章」；
- 给订单里每个产品链上「最匹配的规格书」。

**手工做法的崩溃点**：

- 名称不规范，人肉去资源库里翻、还容易错链（如把 `Desk leg` 错链到 `Desk leg kit`）；
- 同一名字可能有好几条候选，要按价格/相关度挑前 N 条，手工排很累；
- 还要过滤掉「缺货/停售」之类的无效链接；
- 几百个名字 × 几十条资源，纯手工基本不可行。

**这个宏的价值**：内置一套**通用链接引擎**——

1. **名称→资源注册表**：每个参考名对应一组资源记录，每条带多个属性 + 一个 URL；
2. **两级名称匹配**：优先精确匹配，退而求其次用「边界感知」的模糊匹配（所以 `leg` 不会误配 `legacy`）；
3. **链接过滤**：含排除关键词（如"缺货"）的记录在排序前先丢掉；
4. **指标排序**：按数值指标（如最低单价优先）给每条记录排序；
5. **Top-N 混合**：精确命中先填、模糊命中补位，最多 `MAX_LINKS`（默认 3）条；
6. **复合输出 + 超链接**：写回可读的属性串（`名称=编码=比率=重量=链接`）+ 可点击超链接。

全程配置驱动，没有任何特定业务被写死。

## What problem it solves (English)

Another recurring shape: you have free-text names and want each linked to the *best* supporting resource (doc, supplier page, KB article, spec) ranked by price or relevance. By hand this is slow and error-prone — wrong links (e.g. `Desk leg` → `Desk leg kit`), no ranking, no filtering of out-of-stock items. This macro ships a reusable engine: two-tier matching, link filtering, metric ranking, and a Top-N blend with clickable hyperlinks — all configuration-driven.

---

## 快速运行（中文）

**环境要求**

- Excel 或任何兼容的电子表格软件（支持 VBA）；
- 一张源表 + 一张或多张**放在同一工作簿内**的资源表。

**五步跑起来**

1. 打开工作簿，按 `Alt + F11` 导入 `src/NameToResource.bas`；
2. 把资源表放进**同一个工作簿**（每张表有表头、数据从第 2 行起）；
3. 改模块顶部常量，匹配你的表结构（见下表）；
4. 把工作簿另存为 `.xlsm`；
5. 按 `Alt + F8` 运行宏 **`LinkResources`**。链接默认写到 AG 起的输出列。

**默认列布局**

| 列（默认） | 含义 | 必填 |
|-----------|------|------|
| C（NAME_COL） | 品名（自由文本） | 是 |
| C3（表头） | 期望表头文本 `Item name`（`""` 可跳过） | 否 |

资源表列（A..J）：`A` 资源名（可用 `/` 别名、`-` 去后缀）、`B` 显示名、`C` 编码、`D` 比率、`E` 价格、`F` 重量、`G` 链接（按关键词过滤）、`H` 类别、`I` 组、`J` 排序指标。

**想立刻看到效果**：打开 `examples/sample-data.xlsx` 直接运行（示例：`Desk leg` 精确命中两条并按指标排序取最便宜，且不会误配 `Desk leg kit`）。

## Quick start (English)

1. Open your workbook, `Alt+F11` → import `src/NameToResource.bas`.
2. Put resource tables in the **same workbook** (header + data from row 2).
3. Edit the config constants at the top.
4. Save as `.xlsm`, press `Alt+F8`, run **`LinkResources`** (output starts at column AG).
5. Try `examples/sample-data.xlsx` first.

---

## 工作原理 / How it works

1. 读取源表品名；
2. 扫描配置的资源表，建立「名称→记录」索引，应用链接过滤、按指标排序；
3. 对每个名字收集精确与模糊命中，混合（精确优先、上限 `MAX_LINKS`），写回复合串 + 超链接；
4. 无匹配则清空输出单元格。

**匹配规则一览**：精确（名字相等或 `/` 别名之一相等）；智能（名字包含该项且位置检查通过：在结尾 OK、在开头加 `-` 后缀 OK、在开头加额外词 NO、在中间 NO）；再加**后缀黑名单**防止 `Desk leg` 误配 `Desk leg kit`。

**关键常量（唯一真源 `src/NameToResource.bas`）**

| 常量 | 默认 | 含义 |
|------|------|------|
| `NAME_COL` | `"C"` | 品名列 |
| `RESOURCE_SHEETS` | `"Resources,ResourcesExtra"` | 资源表（逗号分隔） |
| `OUTPUT_LINK_COL_START` | `"AG"` | 首个输出列 |
| `MAX_LINKS` | `3` | 每个名字最多链接数 |
| `HYPERLINK_ON` | `True` | 附加可点击链接 |
| `EXCLUDE_LINK_KEYWORDS` | `out of stock,…` | 含这些词的链接跳过 |
| `EXCLUDE_SUFFIX_WORDS` | `accessory,kit,…` | 过宽的名字后缀黑名单 |

## 局限 / Limitations

- 资源表必须和宏在同一工作簿（`ThisWorkbook`）；
- 仅做整名相等或边界感知子串匹配，不做模糊/拼音匹配；
- 候选列表先填精确再填模糊（按索引序），不对不同名字做全局重排；
- 作者环境未实际执行该宏；正式使用前请在数据副本上测试。

## 适用场景 / Use cases

给请购物料贴最便宜供应商链接、给工单推最相关 KB 文章、给订单产品链最匹配规格书——任何「自由文本名称 → Top-N 资源链接」的映射。

## 法律声明 / Legal & Non-affiliation

本仓库是通用匹配模式的**独立、clean-room 重写**，不含任何第三方的专有数据、内部命名或业务特定逻辑，作者与任何可能启发类似流程的公司无隶属关系、也不代表其立场。MIT 许可证仅覆盖作者本人贡献；使用者需确保用途符合所在地法律与单位政策。

## License

[MIT License](LICENSE) · Copyright (c) 2026 ivy92888-glitch.

---

<details><summary>维护说明（中英双语单文件，怎么避免漂移）</summary>

本文档是**单文件中英双语**：中文为权威详细描述，英文为简洁摘要（不强求逐句对等）。所有参数名（如 `MAX_LINKS`、`RESOURCE_SHEETS`）直接取自源码常量，唯一真源是 `src/NameToResource.bas`。修改代码逻辑时，只要同步更新源码常量名与本文档常量表中的同名项即可，叙述性文字按需增删，不影响一致性。

</details>
