# CampusSync 模拟面试深挖 Q&A

> **适用场景**：国际学校 IT 部门转正面试  
> **准备建议**：通读代码后，用自己的话复述答案，不要死记硬背。面试官追问的方向通常是"为什么这样设计"和"如果遇到 XX 情况怎么办"。

---

## 一、项目背景与业务动机

### Q1: 为什么要做这个项目？学校现有的工单系统有什么问题？

**参考答案**：
校内的 IT 报修目前可能还依赖邮件或微信群，信息散落各处，难以追踪进度。这个项目是为了把工单流程标准化：
1. **集中化管理** — 所有报修进 SharePoint List，IT 团队统一视图。
2. **状态可追溯** — 从 `open` → `in_progress` → `resolved` → `closed` 有明确生命周期。
3. **报表归档** — 定期导出 CSV 到 SharePoint Document Library，方便学期末做设备故障分析。
4. **轻量可落地** — 作为外包/实习期间的独立交付物，证明我能独立完成从需求到代码的闭环。

**💡 追问防御**：
- "SharePoint List 和 Excel 比优势在哪？" → List 有 API，可以程序读写；Excel 在线编辑容易冲突。
- "这个系统现在在学校跑了吗？" → **诚实回答**：目前是在 Mock 模式下调通的完整原型，已有真实的 SharePoint 集成代码，待拿到正式 Azure 凭证后即可切换 Live 模式。

---

## 二、技术选型

### Q2: 为什么选 FastAPI？不用 Flask 或 Django？

**参考答案**：
1. **原生异步** (`async`/`await`)：IT 工单系统虽然并发不高，但调用 Microsoft Graph API 是网络 I/O 密集型，异步可以避免线程阻塞。
2. **自动文档**：FastAPI 基于 Pydantic 自动生成 OpenAPI/Swagger，IT 部门的同事可以直接在 `/docs` 页面试用接口，降低沟通成本。
3. **类型安全**：请求参数、响应模型都有类型提示，配合 Pydantic 校验，减少运行时错误。
4. **轻量**：Django 太重，Flask 需要额外配 `flask-pydantic` 或 `marshmallow`，FastAPI 开箱即用。

**💡 追问防御**：
- "FastAPI 的异步是真的异步吗？" → 是的，基于 `asyncio` 和 `starlette`，但前提是下游操作也要异步（比如用 `httpx` 而不是 `requests`）。我目前用的是 `requests` 调 Graph API，严格来说这部分是同步阻塞的，后续可以优化为 `httpx.AsyncClient`。

---

### Q3: 为什么用 Pydantic？`TicketCreate`、`TicketUpdate`、`TicketOut` 为什么要分开定义？

**参考答案**：
这是**分层校验**的思想：
- `TicketCreate`：必填字段有 `min_length=1`，`title` 必须存在。`reporter_email` 是可选的。
- `TicketUpdate`：全部字段 `Optional`，允许 PATCH 只更新部分字段。我代码里用了 `model_dump(exclude_unset=True)`，确保不会把未传的字段覆盖成 `null`。
- `TicketOut`：包含服务端生成的字段（`id`, `created_at`, `updated_at`），这些不该由客户端传入。

分开定义可以避免"前端传了什么不该传的字段"的安全隐患，也让 API 契约更清晰。

**💡 追问防御**：
- "`exclude_unset=True` 的作用是什么？" → Pydantic v2 的方法，只序列化用户显式传入的字段。比如 PATCH 只传了 `status`，就不会把 `title` 覆盖掉。
- "如果前端传了一个模型里没有的字段怎么办？" → FastAPI 默认会忽略，如果想严格拒绝，可以在 Pydantic Config 里设 `extra='forbid'`。

---

## 三、OAuth2 与 Microsoft Graph API

### Q4: 你们学校的 SharePoint 认证是怎么做的？OAuth2 的哪种授权模式？

**参考答案**：
用的是 **OAuth2 Client Credentials**（客户端凭证模式）。

原因：这是一个后台服务（Daemon App），没有用户交互界面，需要以"应用自身"的身份去读写 SharePoint，而不是"代表某个用户"。

流程：
1. 在 Azure Portal 注册应用，获取 `client_id`、`client_secret`、`tenant_id`。
2. 后端用 MSAL 库的 `ConfidentialClientApplication`。
3. 调用 `acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])` 获取 Access Token。
4. 后续所有 Graph API 请求在 Header 里带 `Authorization: Bearer {token}`。

**代码位置**：`sharepoint_client.py:86-96`

**💡 追问防御**：
- "Client Credentials 和 Delegated 权限的区别？" → Client Credentials 是应用身份，适合无人值守的后台服务；Delegated 是"代表用户"，需要用户登录授权，适合前端应用。
- "Token 过期了怎么办？" → 我目前做了简单的实例变量缓存 (`self._access_token`)，但没有做过期检查。MSAL 的 `acquire_token_for_client` 内部有缓存机制，但在生产环境中应该检查 `result.get("expires_in")` 并定时刷新。
- "如果 `client_secret` 泄露了怎么办？" → 这是学校 IT 管理的范畴，技术上应该：1) 用 Azure Key Vault 存储 secret；2) 定期轮换；3) 最小权限原则，只申请 `Sites.Read.All` 和必要的写权限。

---

### Q5: SharePoint List 的 CRUD 具体怎么实现的？创建一条工单时，Graph API 的 Payload 长什么样？

**参考答案**：
SharePoint List 通过 Microsoft Graph API 的 `/sites/{site-id}/lists/{list-id}/items` 端点操作。

**创建工单**（`sharepoint_client.py:143-154`）：
```python
endpoint = f"/sites/{self.site_id}/lists/{self.list_id}/items"
graph_payload = {"fields": payload}  # 关键：必须用 fields 包裹
resp = self._graph_request("POST", endpoint, json_payload=graph_payload)
```

Graph API 要求 List Item 的字段必须放在 `fields` 对象里，不能直接传平铺字段。

**查询工单**（`sharepoint_client.py:156-195`）：
用 OData 过滤：
```python
filters.append(f"fields/status eq '{status}'")
endpoint += f"&$filter={filter_str}"
```
同时加 `?expand=fields`，否则返回的 item 不包含自定义列的值。

**💡 追问防御**：
- "如果工单量很大，分页怎么做？" → 我目前用了 `$top` (limit) 和 `$skip` (offset)，但 Graph API 的推荐做法是用 `@odata.nextLink` 做游标分页，而不是 `skip`，因为 SharePoint List 的 skip 性能在大数据量时会下降。
- "List 里的列名和代码里的字段名必须一致吗？" → 是的，Graph API 的 `fields/{fieldName}` 是区分大小写的，需要在 SharePoint 里提前建好对应的列（如 `status`, `category`, `reporterEmail` 等）。

---

### Q6: 文件上传到 SharePoint Document Library 是怎么实现的？

**参考答案**：
通过 Graph API 的 Drive 端点：
```python
endpoint = f"/drives/{target_drive}/root:/{file_name}:/content"
resp = self._graph_request("PUT", endpoint, data=file_bytes, headers=headers)
```

这是**简单上传**（Simple Upload），适合 < 4MB 的文件。对于 CSV 报表这种小文件完全够用。如果是大文件，需要分块上传（Resumable Upload）。

**代码位置**：`sharepoint_client.py:235-259`

**💡 追问防御**：
- "上传的 MIME Type 有什么用？" → SharePoint 会根据 MIME Type 决定文件的预览方式，比如 `text/csv` 可以在线预览，而 `application/octet-stream` 只能下载。
- "如果文件名冲突怎么办？" → Graph API 的 PUT 会**覆盖同名文件**。如果需要保留历史版本，可以在文件名里加时间戳，比如 `weekly_report_20240816.csv`。

---

## 四、Mock 架构设计

### Q7: 为什么要做 Mock 模式？怎么设计的 Fallback？

**参考答案**：
原因很实际：作为外包/实习生，我可能没有权限直接拿到学校的 Azure App Registration 凭证，也不能随意读写生产环境的 SharePoint List。Mock 模式让我能在**零外部依赖**的情况下完成开发和演示。

**Fallback 逻辑**（`sharepoint_client.py:58-63`）：
```python
if force_mock or not self._has_real_credentials():
    self._mock = MockSharePointClient()
```

只要 `TENANT_ID`、`CLIENT_ID`、`CLIENT_SECRET`、`SITE_ID` 任意一个缺失，就自动进 Mock。

**Mock 的实现**（`mock_sharepoint.py`）：
- **Tickets**：SQLite 持久化，表结构模拟 SharePoint List 的字段（`id`, `title`, `status`, `priority` 等）。
- **Files**：本地文件系统 + SQLite 记录元数据。
- **API 响应结构**：和真实 Graph API 返回的 JSON 结构保持一致（如 `id`, `name`, `webUrl`, `createdDateTime`），确保前端无感切换。

**💡 追问防御**：
- "Mock 和真实的 SharePoint 数据怎么迁移？" → 目前不支持自动迁移。如果后续要上线，可以写一条迁移脚本，把 SQLite 的数据通过 `sharepoint_client.py` 的 Live 接口批量导入 SharePoint List。
- "Mock 模式支持并发吗？" → SQLite 默认支持单文件多连接读取，但写入有锁。校园 IT 工单的并发极低，所以够用。如果真的要支撑高并发，可以把 SQLite 换成 PostgreSQL。

---

### Q8: SQLite 的表结构是怎么设计的？metadata 字段是做什么的？

**参考答案**：
```sql
CREATE TABLE tickets (
    id TEXT PRIMARY KEY,
    title TEXT NOT NULL,
    description TEXT,
    category TEXT,
    reporter_email TEXT,
    status TEXT DEFAULT 'open',
    priority TEXT DEFAULT 'medium',
    assigned_to TEXT,
    sharepoint_item_id TEXT,
    created_at TEXT,
    updated_at TEXT,
    metadata TEXT
);
```

`metadata` 是一个 JSON 文本列。Pydantic 模型可能包含未来扩展的字段，我不想每加一个字段就改表结构，所以把"非核心字段"序列化成 JSON 存进去。读取时用 `json.loads` 反序列化并合并到返回字典里。

**代码位置**：`mock_sharepoint.py:32-58`, `73-81`

**💡 追问防御**：
- "为什么不用 ALTER TABLE 加列？" → 为了演示方便，减少 schema 变更。生产环境应该用迁移工具（如 Alembic）。
- "UUID 做主键有什么优缺点？" → 优点：本地生成，不依赖数据库自增，分布式友好。缺点：无序、占用空间大、查询效率略低于自增 ID。对于校园工单这种数据量，完全没问题。

---

## 五、CSV 报告导出

### Q9: 导出 CSV 的功能是怎么实现的？为什么用 `utf-8-sig`？

**参考答案**：
1. 先调用 `list_tickets` 拿到符合条件的工单列表。
2. 用 Python 的 `io.StringIO` 在**内存中**构建 CSV，而不是写临时文件。
3. 用 `csv.DictWriter` 自动根据字典 keys 生成表头。
4. 编码为 `utf-8-sig`：这是为了兼容 Windows 上的 Excel。Excel 打开 UTF-8 CSV 时，如果没有 BOM（Byte Order Mark），中文会乱码。`utf-8-sig` 会在文件头自动加上 BOM。
5. 最后通过 `upload_file` 上传到 SharePoint Document Library（或 Mock 的本地目录）。

**代码位置**：`main.py:211-260`

**💡 追问防御**：
- "如果工单量很大（比如 10 万条），内存会不会爆？" → 会。目前的实现是把全部数据加载到内存再生成 CSV。优化方案是用生成器逐条写入 `StreamingResponse`，或者分批次查询、追加写入临时文件。
- "CSV 的字段顺序能控制吗？" → 目前用 `list(tickets[0].keys())`，顺序取决于字典的插入顺序（Python 3.7+ 保证有序）。如果需要固定顺序，应该显式定义 `fieldnames`。

---

## 六、前端 Dashboard

### Q10: 为什么把前端页面直接内嵌在 Python 文件里？没有分离前后端吗？

**参考答案**：
这是一个**有意为之的简化设计**：
1. **部署简单**：学校 IT 环境可能只需要跑一个 `python main.py`，不需要额外配 Nginx 或部署前端静态文件。
2. **内网场景**：校园 IT 系统通常在内网使用，不需要考虑 CDN、SEO 等复杂前端工程化问题。
3. **开发效率**：作为独立原型，把 HTML/CSS/JS 内嵌在 `main.py` 末尾，可以减少文件数量和构建步骤。

前端用的是**原生 JS + Tailwind CDN**，没有引入 React/Vue，因为页面逻辑很简单（表单提交、列表渲染、筛选、统计卡片），不需要状态管理库。

**代码位置**：`main.py:271-587`

**💡 追问防御**：
- "如果后续要扩展功能，这种内嵌方式会不会成为瓶颈？" → 会的。如果未来需要多人协作开发复杂交互，应该把前端拆出去，用 Vue/React + Vite 单独构建，FastAPI 只提供 API。
- "XSS 怎么防？" → 我在前端用了 `esc()` 函数（`textContent` 转义），把用户输入的标题、描述等字段在插入 DOM 前做 HTML Entity 编码，防止反射型 XSS。

---

## 七、错误处理与日志

### Q11: 错误处理是怎么分层的？为什么 Graph API 出错返回 502？

**参考答案**：
分三层：
1. **自定义异常**：`SharePointClientError`，封装所有 SharePoint/Graph 相关的错误。
2. **HTTP 状态码映射**：
   - `SharePointClientError` → `502 Bad Gateway`（表示下游服务出问题）
   - 参数校验失败 → `400 Bad Request`
   - 找不到资源 → `404 Not Found`
   - 未知内部错误 → `500 Internal Server Error`
3. **日志记录**：用 Python 标准 `logging` 模块，`logger.error` 记业务错误，`logger.exception` 记堆栈（用于 500 未知错误）。

**代码位置**：`main.py:104-208`

**💡 追问防御**：
- "为什么不直接用 `try/except Exception` 兜住所有异常？" → 我在代码里特意区分了 `SharePointClientError` 和裸 `Exception`，因为 502 和 500 的语义不同。502 说明 Graph API 挂了或权限不足，500 说明我自己的代码有 Bug。
- "如果 MSAL 获取 token 失败，错误信息会暴露 client_secret 吗？" → 不会。MSAL 的错误描述通常是"Invalid client secret"这种笼统提示，不会把 secret 打印出来。但为了安全，生产环境应该在日志里脱敏。

---

## 八、工程实践

### Q12: 环境变量怎么管理的？`.env.example` 里有什么？

**参考答案**：
用 `python-dotenv` + Pydantic Settings（或直接用 `os.getenv`）。

`.env.example` 里包含：
- `TENANT_ID`, `CLIENT_ID`, `CLIENT_SECRET`：Azure AD 凭证
- `SITE_ID`, `LIST_ID`, `DRIVE_ID`：SharePoint 资源 ID
- `MOCK_MODE=true/false`：强制 Mock
- `DEBUG=true/false`：FastAPI 热重载
- `PORT`：服务端口

代码里用 `load_dotenv()` 在启动时加载，所有配置集中管理，不硬编码在业务逻辑里。

**💡 追问防御**：
- ".env 文件上传 Git 了怎么办？" → `.env` 已加入 `.gitignore`，只提交 `.env.example` 作为模板。
- "如果学校有多个环境（测试/生产），怎么切换？" → 目前通过不同的 `.env` 文件切换。更成熟的方案是用 Pydantic Settings 的 `env_file` 参数，配合 `--env-file=.env.production` 启动。

---

### Q13: 测试是怎么做的？冒烟测试覆盖了哪些内容？

**参考答案**：
项目里有 `test_smoke.py`，用 pytest 做最基本的**冒烟测试**：
- 验证所有模块能正常导入（`import main`, `import sharepoint_client`, `import mock_sharepoint`）
- 验证 FastAPI 应用实例能创建
- 验证 Mock 模式的 CRUD 基本链路

这是"最小可接受测试"，确保代码没有语法错误和循环导入。因为这是一个原型项目，还没有做完整的单元测试（如 mock MSAL 的 token 获取、模拟 Graph API 的 HTTP 响应等）。

**💡 追问防御**：
- "如果要补全测试，你会怎么写？" → 
  1. 用 `pytest-asyncio` 测试异步 endpoint。
  2. 用 `unittest.mock` mock `ConfidentialClientApplication`，测试 token 获取失败和成功的分支。
  3. 用 `responses` 或 `httpx.MockTransport` 拦截 HTTP 请求，测试 Graph API 的 401、403、500 返回。
  4. 给 `MockSharePointClient` 本身写单元测试，验证 SQLite 的增删改查。

---

## 九、转正面试特供题

### Q14: 你在 IT 部门的日常工作是什么？这个项目和你的日常工作有什么关系？

**参考答案准备方向**：
- **日常**：处理师生硬件/软件/网络报修、维护校内设备、管理 Office 365 / SharePoint 站点。
- **痛点观察**：发现报修信息分散，主管难以统计每个 technician 的工作量，也无法分析哪个教室/哪类设备故障率最高。
- **项目动机**：主动向主管提议做一个轻量级工单系统，把流程线上化。这个项目是**从真实工作需求中长出来的**，不是拍脑袋的练手项目。

---

### Q15: 如果转正后要把这个系统真正部署到学校环境里，你下一步会做什么？

**参考答案**：
1. **权限与安全**：
   - 把 `client_secret` 迁移到 Azure Key Vault 或学校内部的 secret manager。
   - 给 API 加身份验证（OAuth2 Bearer + JWT 或集成学校的 SSO）。
2. **部署与运维**：
   - 用 Docker 容器化，方便在学校的 Windows Server 或 Azure VM 上部署。
   - 加 `docker-compose.yml`，包含应用 + 反向代理（Traefik 或 Nginx）。
3. **数据库升级**：
   - Mock 模式的 SQLite 仅供开发；上线后如果不用 SharePoint List 做主存储，可以切到 PostgreSQL。
4. **功能扩展**：
   - 邮件/Teams 通知：工单状态变更时自动通知报修人。
   - 资产管理对接：把工单和具体的设备资产标签关联。
   - SLA 统计：自动计算"从报修到解决"的平均时长。
5. **监控与日志**：
   - 接入学校的日志平台（如 Azure Application Insights 或 ELK）。

---

### Q16: 你在做这个项目时遇到的最大技术挑战是什么？怎么解决的？

**参考答案准备方向**（选一个你最熟悉的）：

**选项 A：Microsoft Graph API 的调试困难**
- 问题：Graph API 的报错信息有时很模糊，比如 `Access Denied` 不知道是权限不足还是 scope 不对。
- 解决：用 Microsoft Graph Explorer 手动构造请求，确认权限和 payload 格式；阅读 MSAL 文档理解 token 缓存机制。

**选项 B：Mock 与 Live 的行为一致性**
- 问题：Mock 返回的数据结构和 Graph API 不完全一致，导致前端在切换模式时出错。
- 解决：仔细阅读 Graph API 文档，确保 Mock 的 `upload_file` 返回的字段（`id`, `name`, `webUrl`, `createdDateTime`）和真实 API 一致。

**选项 C：OData Filter 的语法**
- 问题：SharePoint List 的字段名在 Graph API 里有特殊规则，比如空格要转义，大小写敏感。
- 解决：在 Graph Explorer 里先手动测试 `$filter=fields/status eq 'open'`，确认语法后再写到代码里。

---

## 十、快速自检清单

在面试前，确保你能流畅回答以下问题（无需背诵，理解即可）：

- [ ] `asynccontextmanager` (`lifespan`) 在 FastAPI 里做什么用？
- [ ] `exclude_unset=True` 在 PATCH 更新里为什么重要？
- [ ] `ConfidentialClientApplication` 和 `PublicClientApplication` 的区别？
- [ ] Graph API 的 `fields` wrapper 是什么意思？
- [ ] 为什么 Mock 模式用 SQLite 而不是直接存内存字典？
- [ ] `utf-8-sig` 和 `utf-8` 的区别？
- [ ] 502 和 500 的语义区别，以及我在代码里怎么区分它们的？
- [ ] 如果学校给你生产环境的 Azure 凭证，你第一件事会做什么？

---

**祝你面试顺利！** 🎯
