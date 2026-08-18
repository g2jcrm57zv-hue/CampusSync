# SAS-CampusSync IT Helpdesk

> **SAS** = **Shanghai American School**（上海美国学校）  
> 本项目为作者在 SAS 校内 IT 部门实习/外包期间独立开发的内部系统原型，用于展示校园 IT 工单管理与 SharePoint 集成能力。

**SAS-CampusSync** (SharePoint-integrated Academic Services Campus Sync) 是一款面向校园 IT 部门的轻量级工单服务。它基于 **Microsoft Graph API** 与校内 **SharePoint Online** 深度集成（List 存储工单、Document Library 归档报表），同时内置 **Mock Mode**，方便在无 Azure 凭证的环境下本地开发与演示。

---

## 目录 / Table of Contents

- [功能特性 / Features](#功能特性--features)
- [系统架构 / Architecture](#系统架构--architecture)
- [项目结构 / Project Structure](#项目结构--project-structure)
- [快速开始 / Quick Start](#快速开始--quick-start)
- [API 文档 / API Documentation](#api-文档--api-documentation)
- [Azure / SharePoint 应用注册指引](#azure--sharepoint-应用注册指引)
- [环境变量说明 / Environment Variables](#环境变量说明--environment-variables)
- [Mock 模式说明 / Mock Mode](#mock-模式说明--mock-mode)
- [运行测试 / Testing](#运行测试--testing)
- [许可证 / License](#许可证--license)

---

## 功能特性 / Features

- **工单全生命周期管理**：提交、查询、状态更新、删除 REST API。
- **SharePoint 深度集成**：通过 Microsoft Graph API 直接读写 SharePoint List 与 Document Library。
- **OAuth2 Client Credentials**：全自动令牌获取与缓存，支持校内 Azure AD / Microsoft Entra ID 环境。
- **Mock fallback**：缺少真实 Azure 凭证时，自动切换为本地 SQLite + 文件系统模拟，保证一键跑通演示。
- **报告导出**：一键将工单数据导出为 CSV 并上传至 SharePoint 文档库（或本地模拟目录）。
- **Pydantic 数据校验**：请求与响应均具备强类型验证。
- **异步 FastAPI**：高性能、自动生成的 OpenAPI/Swagger 交互文档。

---

## 系统架构 / Architecture

```
┌─────────────────┐      HTTP/REST       ┌─────────────────────────────┐
│   Client / UI   │ ◄──────────────────► │   FastAPI (main.py)         │
└─────────────────┘                      │   - Pydantic validation     │
                                         │   - Business logic          │
                                         └─────────────┬───────────────┘
                                                       │
                              ┌────────────────────────┼────────────────────────┐
                              │                        │                        │
                    LIVE Mode │                MOCK Mode (auto fallback)       │
                              │                        │                        │
                 ┌────────────▼────────────┐  ┌────────▼─────────┐  ┌─────────▼──────────┐
                 │  Microsoft Graph API    │  │  SQLite DB       │  │  Local Filesystem  │
                 │  - OAuth2 token mgmt    │  │  (tickets)       │  │  (mock_files)      │
                 │  - SharePoint List CRUD │  │                  │  │                    │
                 │  - Drive file up/down   │  │                  │  │                    │
                 └─────────────────────────┘  └──────────────────┘  └────────────────────┘
```

---

## 项目结构 / Project Structure

```
.
├── .env.example          # Azure / SharePoint configuration template
├── main.py               # FastAPI application entry point
├── sharepoint_client.py  # Microsoft Graph API wrapper (OAuth2 + List/Drive ops)
├── mock_sharepoint.py    # Local mock implementation (SQLite + filesystem)
├── requirements.txt      # Python dependencies
├── README.md             # This file
└── data/                 # Auto-created in Mock Mode
    ├── campus_sync.db    # SQLite database
    └── mock_files/       # Simulated Document Library files
```

---

## 快速开始 / Quick Start

### 1. 克隆并进入项目目录

```bash
cd SAS-CampusSync
```

### 2. 创建虚拟环境（推荐）

```bash
python -m venv venv

# Windows
venv\Scripts\activate

# macOS / Linux
source venv/bin/activate
```

### 3. 安装依赖

```bash
pip install -r requirements.txt
```

### 4. 配置环境变量

复制示例文件并按需填写：

```bash
cp .env.example .env
```

> **提示**：若您暂时没有 Azure 凭证，可将 `.env` 中的 `MOCK_MODE=true`，或直接留空 Azure 字段，系统会在启动时自动进入 Mock 模式。

### 5. 启动服务

```bash
python main.py
```

或者使用 uvicorn 直接启动：

```bash
uvicorn main:app --reload --host 0.0.0.0 --port 8001
```

### 6. 访问交互式文档

打开浏览器访问：

- Swagger UI: [http://localhost:8001/docs](http://localhost:8001/docs)
- ReDoc: [http://localhost:8001/redoc](http://localhost:8001/redoc)

---

## API 文档 / API Documentation

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/tickets` | Submit a new helpdesk ticket |
| `GET`  | `/tickets` | Query tickets (filters + pagination) |
| `GET`  | `/tickets/{id}` | Retrieve a single ticket |
| `PATCH`| `/tickets/{id}` | Update ticket fields (e.g., status) |
| `DELETE`| `/tickets/{id}` | Delete a ticket |
| `POST` | `/reports/export` | Export CSV report to SharePoint Document Library |
| `GET`  | `/health` | Health check |

### 示例请求 / Example Requests

**Create ticket:**
```bash
curl -X POST "http://localhost:8001/tickets" \
  -H "Content-Type: application/json" \
  -d '{
    "title": " projector in Room 301 not working",
    "description": "No signal from HDMI, lamp indicator red.",
    "category": "hardware",
    "reporter_email": "student@campus.edu",
    "priority": "high"
  }'
```

**List tickets (filtered):**
```bash
curl "http://localhost:8001/tickets?status=open&limit=10"
```

**Update status:**
```bash
curl -X PATCH "http://localhost:8001/tickets/{ticket_id}" \
  -H "Content-Type: application/json" \
  -d '{"status": "resolved", "assigned_to": "tech@campus.edu"}'
```

**Export report:**
```bash
curl -X POST "http://localhost:8001/reports/export" \
  -H "Content-Type: application/json" \
  -d '{"file_name": "weekly_report.csv", "status_filter": "open"}'
```

---

## Azure / SharePoint 应用注册指引

### 步骤 1：在 Azure 门户注册应用

1. 登录 [Azure Portal](https://portal.azure.com/).
2. 导航至 **Microsoft Entra ID** > **App registrations** > **New registration**.
3. 填写应用名称（如 `SAS-CampusSync`），选择 **Accounts in this organizational directory only**。
4. 点击 **Register**。

### 步骤 2：记录凭证

在应用概览页复制：

- **Application (client) ID** → `CLIENT_ID`
- **Directory (tenant) ID** → `TENANT_ID`

### 步骤 3：创建客户端密钥

1. 进入 **Certificates & secrets** > **Client secrets** > **New client secret**.
2. 添加描述并选择过期时间，点击 **Add**。
3. 立即复制生成的 **Value**（只显示一次）→ `CLIENT_SECRET`。

### 步骤 4：配置 API 权限

1. 进入 **API permissions** > **Add a permission** > **Microsoft Graph** > **Application permissions**。
2. 添加以下权限：
   - `Sites.Read.All` （读取站点）
   - `Sites.FullControl.All` （若需写操作，部分租户需管理员同意）
3. 点击 **Grant admin consent for [tenant]**。

### 步骤 5：获取 SharePoint Site / List / Drive ID

使用 [Microsoft Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) 或以下请求查询：

```http
GET https://graph.microsoft.com/v1.0/sites?search=campus
GET https://graph.microsoft.com/v1.0/sites/{site-id}/lists
GET https://graph.microsoft.com/v1.0/sites/{site-id}/drives
```

将对应 ID 填入 `.env`：

```env
SITE_ID=your-site-id
LIST_ID=your-list-id
DRIVE_ID=your-drive-id
```

---

## 环境变量说明 / Environment Variables

| Variable | Required (Live) | Description |
|----------|-----------------|-------------|
| `TENANT_ID` | Yes | Azure AD / Microsoft Entra tenant ID |
| `CLIENT_ID` | Yes | Azure App Registration client ID |
| `CLIENT_SECRET` | Yes | Azure App client secret |
| `SITE_ID` | Yes | SharePoint Site ID |
| `LIST_ID` | Yes* | SharePoint List ID for tickets (*optional in Mock) |
| `DRIVE_ID` | Yes* | SharePoint Drive ID for document library (*optional in Mock) |
| `MOCK_MODE` | No | Set `true` to force Mock mode |
| `DEBUG` | No | Set `true` to enable FastAPI auto-reload |
| `PORT` | No | Server port (default: `8001`) |

---

## Mock 模式说明 / Mock Mode

当以下任一条件满足时，服务自动进入 Mock 模式：

1. `MOCK_MODE=true` 被显式设置。
2. `TENANT_ID`、`CLIENT_ID`、`CLIENT_SECRET` 中任意一项缺失。

Mock 模式行为：

- **Tickets**：持久化至本地 SQLite (`data/campus_sync.db`)，支持完整的 CRUD 与过滤。
- **Files**：上传的文件保存在 `data/mock_files/`，并在 SQLite 中记录元数据。
- **API 响应**：与真实 SharePoint 返回的结构保持一致，前端/调用方可无感切换。

---

## 运行测试 / Testing

安装依赖后，运行 pytest 确保无语法与导入错误：

```bash
pytest --tb=short -q
```

项目内置了最小化冒烟测试脚本，您也可以快速执行：

```bash
python -c "import main; import sharepoint_client; import mock_sharepoint; print('All modules imported successfully')"
```

---

## 许可证 / License

MIT License. See project repository for details.

---

> **Built with ❤️ for campus IT teams.**  
> For questions or contributions, please open an issue or pull request.
