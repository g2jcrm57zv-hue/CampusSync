"""
main.py
FastAPI application for CampusSync IT Helpdesk.
Provides REST APIs for ticket lifecycle and report export to SharePoint.
"""

import csv
import io
import logging
import os
from contextlib import asynccontextmanager
from datetime import datetime
from typing import List, Optional

from fastapi import FastAPI, HTTPException, Query, UploadFile, File
from fastapi.responses import StreamingResponse, HTMLResponse
from pydantic import BaseModel, Field
from dotenv import load_dotenv

from sharepoint_client import SharePointClient, SharePointClientError

# ---------------------------------------------------------------------------
# Bootstrap
# ---------------------------------------------------------------------------
load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# Pydantic schemas
# ---------------------------------------------------------------------------
class TicketCreate(BaseModel):
    title: str = Field(..., min_length=1, max_length=200, description="Ticket title")
    description: Optional[str] = Field(None, description="Detailed description")
    category: Optional[str] = Field("general", description="Category: general, hardware, software, network")
    reporter_email: Optional[str] = Field(None, description="Email of the reporter")
    priority: Optional[str] = Field("medium", description="Priority: low, medium, high, critical")
    assigned_to: Optional[str] = Field(None, description="Assigned technician email or name")


class TicketUpdate(BaseModel):
    title: Optional[str] = Field(None, max_length=200)
    description: Optional[str] = None
    category: Optional[str] = None
    status: Optional[str] = Field(None, description="Status: open, in_progress, resolved, closed")
    priority: Optional[str] = None
    assigned_to: Optional[str] = None


class TicketOut(BaseModel):
    id: str
    title: str
    description: Optional[str] = None
    category: Optional[str] = None
    reporter_email: Optional[str] = None
    status: Optional[str] = None
    priority: Optional[str] = None
    assigned_to: Optional[str] = None
    created_at: Optional[str] = None
    updated_at: Optional[str] = None


class ExportRequest(BaseModel):
    file_name: Optional[str] = Field("tickets_report.csv", description="Name of the exported file")
    status_filter: Optional[str] = Field(None, description="Filter tickets by status before export")


# ---------------------------------------------------------------------------
# Lifespan & client initialization
# ---------------------------------------------------------------------------
@asynccontextmanager
async def lifespan(app: FastAPI):
    force_mock = os.getenv("MOCK_MODE", "false").lower() in ("true", "1", "yes")
    app.state.sp_client = SharePointClient(force_mock=force_mock)
    logger.info("CampusSync service started. Mock mode: %s", force_mock)
    yield
    logger.info("CampusSync service shutting down.")


app = FastAPI(
    title="CampusSync IT Helpdesk",
    description="Lightweight SharePoint-integrated campus IT helpdesk service.",
    version="1.0.0",
    lifespan=lifespan,
)


# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------
def _get_client() -> SharePointClient:
    return app.state.sp_client


# ---------------------------------------------------------------------------
# API Endpoints
# ---------------------------------------------------------------------------

@app.post("/tickets", response_model=TicketOut, status_code=201)
async def create_ticket(payload: TicketCreate):
    """
    Submit a new IT helpdesk ticket.
    """
    try:
        data = payload.model_dump(exclude_unset=True)
        result = _get_client().create_ticket(data)
        return result
    except SharePointClientError as exc:
        logger.error("create_ticket error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc))
    except Exception as exc:
        logger.exception("Unexpected error creating ticket")
        raise HTTPException(status_code=500, detail="Internal server error")


@app.get("/tickets", response_model=List[TicketOut])
async def list_tickets(
    status: Optional[str] = Query(None, description="Filter by status"),
    category: Optional[str] = Query(None, description="Filter by category"),
    reporter_email: Optional[str] = Query(None, description="Filter by reporter email"),
    limit: int = Query(100, ge=1, le=500),
    offset: int = Query(0, ge=0),
):
    """
    Query tickets with optional filters and pagination.
    """
    try:
        results = _get_client().list_tickets(
            status=status,
            category=category,
            reporter_email=reporter_email,
            limit=limit,
            offset=offset,
        )
        return results
    except SharePointClientError as exc:
        logger.error("list_tickets error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc))
    except Exception as exc:
        logger.exception("Unexpected error listing tickets")
        raise HTTPException(status_code=500, detail="Internal server error")


@app.get("/tickets/{ticket_id}", response_model=TicketOut)
async def get_ticket(ticket_id: str):
    """
    Retrieve a single ticket by ID.
    """
    try:
        result = _get_client().get_ticket(ticket_id)
        if not result:
            raise HTTPException(status_code=404, detail="Ticket not found")
        return result
    except HTTPException:
        raise
    except SharePointClientError as exc:
        logger.error("get_ticket error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc))
    except Exception as exc:
        logger.exception("Unexpected error retrieving ticket")
        raise HTTPException(status_code=500, detail="Internal server error")


@app.patch("/tickets/{ticket_id}", response_model=TicketOut)
async def update_ticket(ticket_id: str, payload: TicketUpdate):
    """
    Update ticket fields (e.g., change status to resolved).
    """
    try:
        data = payload.model_dump(exclude_unset=True)
        if not data:
            raise HTTPException(status_code=400, detail="No fields provided for update")
        result = _get_client().update_ticket(ticket_id, data)
        if not result:
            raise HTTPException(status_code=404, detail="Ticket not found")
        return result
    except HTTPException:
        raise
    except SharePointClientError as exc:
        logger.error("update_ticket error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc))
    except Exception as exc:
        logger.exception("Unexpected error updating ticket")
        raise HTTPException(status_code=500, detail="Internal server error")


@app.delete("/tickets/{ticket_id}", status_code=204)
async def delete_ticket(ticket_id: str):
    """
    Delete a ticket by ID.
    """
    try:
        ok = _get_client().delete_ticket(ticket_id)
        if not ok:
            raise HTTPException(status_code=404, detail="Ticket not found")
    except HTTPException:
        raise
    except SharePointClientError as exc:
        logger.error("delete_ticket error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc))
    except Exception as exc:
        logger.exception("Unexpected error deleting ticket")
        raise HTTPException(status_code=500, detail="Internal server error")


@app.post("/reports/export")
async def export_report(payload: ExportRequest):
    """
    Export tickets to a CSV report and upload it to the SharePoint Document Library.
    Returns the file metadata returned by SharePoint (or mock).
    """
    try:
        client = _get_client()
        # Fetch filtered tickets
        tickets = client.list_tickets(
            status=payload.status_filter,
            limit=10000,
        )

        # Build CSV in memory
        output = io.StringIO()
        if tickets:
            fieldnames = list(tickets[0].keys())
            writer = csv.DictWriter(output, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(tickets)
        else:
            output.write("No tickets found for the given criteria.\n")
        csv_bytes = output.getvalue().encode("utf-8-sig")
        output.close()

        # Upload to SharePoint / mock
        file_name = payload.file_name
        if not file_name.endswith(".csv"):
            file_name += ".csv"

        upload_stream = io.BytesIO(csv_bytes)
        result = client.upload_file(
            file_name=file_name,
            content=upload_stream,
            mime_type="text/csv",
        )

        return {
            "message": "Report exported successfully",
            "file_name": file_name,
            "record_count": len(tickets),
            "sharepoint_file": result,
        }
    except SharePointClientError as exc:
        logger.error("export_report error: %s", exc)
        raise HTTPException(status_code=502, detail=str(exc))
    except Exception as exc:
        logger.exception("Unexpected error exporting report")
        raise HTTPException(status_code=500, detail="Internal server error")


@app.get("/health")
async def health_check():
    """
    Health check endpoint.
    """
    return {"status": "ok", "timestamp": datetime.utcnow().isoformat() + "Z"}


@app.get("/", response_class=HTMLResponse)
async def dashboard():
    """
    Single-page dashboard for ticket management and report export.
    """
    return """<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CampusSync Dashboard</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    @keyframes fadeIn { from { opacity:0; transform: translateY(6px); } to { opacity:1; transform: translateY(0); } }
    .card-anim { animation: fadeIn 0.3s ease-out both; }
    .loader { border: 2px solid #f3f3f3; border-top: 2px solid #3b82f6; border-radius: 50%; width: 16px; height: 16px; animation: spin 1s linear infinite; display: inline-block; }
    @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
  </style>
</head>
<body class="bg-gray-50 min-h-screen font-sans text-gray-800">

  <!-- Header -->
  <header class="bg-white border-b border-gray-200 sticky top-0 z-30">
    <div class="max-w-5xl mx-auto px-4 py-4 flex items-center justify-between">
      <div class="flex items-center gap-3">
        <div class="w-9 h-9 bg-gradient-to-br from-blue-600 to-indigo-700 rounded-lg flex items-center justify-center text-white font-bold text-sm shadow-sm">IT</div>
        <div>
          <h1 class="text-base font-semibold text-gray-900 leading-tight">CampusSync IT Helpdesk</h1>
          <p class="text-[11px] text-gray-500">SharePoint Integrated Ticketing</p>
        </div>
      </div>
      <button onclick="exportReport()" id="export-btn" class="bg-green-600 hover:bg-green-700 text-white text-sm font-medium px-4 py-2 rounded-lg transition-colors flex items-center gap-2 shadow-sm">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 16v-8m0 0l-4 4m4-4l4 4M4 17v2a2 2 0 002 2h12a2 2 0 002-2v-2"></path></svg>
        导出报表
      </button>
    </div>
  </header>

  <main class="max-w-5xl mx-auto px-4 py-6 space-y-6">

    <!-- Stats -->
    <div class="grid grid-cols-2 md:grid-cols-4 gap-4" id="stats-bar">
      <div class="bg-white rounded-xl border border-gray-200 p-4 card-anim" style="animation-delay:0ms">
        <p class="text-xs text-gray-500 font-medium">全部工单</p>
        <p class="text-2xl font-bold text-gray-900 mt-1" id="stat-all">-</p>
      </div>
      <div class="bg-white rounded-xl border border-gray-200 p-4 card-anim" style="animation-delay:50ms">
        <p class="text-xs text-gray-500 font-medium">待处理</p>
        <p class="text-2xl font-bold text-orange-600 mt-1" id="stat-open">-</p>
      </div>
      <div class="bg-white rounded-xl border border-gray-200 p-4 card-anim" style="animation-delay:100ms">
        <p class="text-xs text-gray-500 font-medium">已解决</p>
        <p class="text-2xl font-bold text-green-600 mt-1" id="stat-resolved">-</p>
      </div>
      <div class="bg-white rounded-xl border border-gray-200 p-4 card-anim" style="animation-delay:150ms">
        <p class="text-xs text-gray-500 font-medium">今日新增</p>
        <p class="text-2xl font-bold text-blue-600 mt-1" id="stat-today">-</p>
      </div>
    </div>

    <!-- New Ticket Form -->
    <div class="bg-white rounded-xl border border-gray-200 shadow-sm overflow-hidden card-anim" style="animation-delay:200ms">
      <div class="px-5 py-4 border-b border-gray-100 flex items-center justify-between">
        <h2 class="text-sm font-semibold text-gray-800 flex items-center gap-2">
          <span class="w-5 h-5 bg-blue-100 text-blue-600 rounded flex items-center justify-center text-xs">+</span>
          新建报修工单
        </h2>
        <span class="text-[11px] text-gray-400">数据将同步至 SharePoint List</span>
      </div>
      <form id="ticket-form" class="p-5 grid grid-cols-1 md:grid-cols-2 gap-4">
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1.5">标题 <span class="text-red-500">*</span></label>
          <input type="text" name="title" required placeholder="例如：301教室投影仪故障"
            class="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all">
        </div>
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1.5">分类</label>
          <select name="category"
            class="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-white">
            <option value="general">一般咨询</option>
            <option value="hardware">硬件故障</option>
            <option value="software">软件问题</option>
            <option value="network">网络故障</option>
          </select>
        </div>
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1.5">优先级</label>
          <select name="priority"
            class="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all bg-white">
            <option value="low">低</option>
            <option value="medium" selected>中</option>
            <option value="high">高</option>
            <option value="critical">紧急</option>
          </select>
        </div>
        <div>
          <label class="block text-xs font-medium text-gray-600 mb-1.5">报修人邮箱</label>
          <input type="email" name="reporter_email" placeholder="user@campus.edu"
            class="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all">
        </div>
        <div class="md:col-span-2">
          <label class="block text-xs font-medium text-gray-600 mb-1.5">问题描述</label>
          <textarea name="description" rows="3" placeholder="请详细描述遇到的问题..."
            class="w-full rounded-lg border border-gray-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500 transition-all resize-none"></textarea>
        </div>
        <div class="md:col-span-2 flex justify-end">
          <button type="submit" id="submit-btn"
            class="bg-blue-600 hover:bg-blue-700 text-white text-sm font-medium px-6 py-2 rounded-lg transition-colors shadow-sm flex items-center gap-2">
            <span>提交工单</span>
          </button>
        </div>
      </form>
    </div>

    <!-- Filter Tabs -->
    <div class="flex items-center gap-2 overflow-x-auto pb-1 card-anim" style="animation-delay:250ms">
      <button onclick="setFilter('all')" data-filter="all" class="filter-btn px-4 py-1.5 rounded-full text-sm font-medium transition-all bg-blue-600 text-white shadow-sm">全部</button>
      <button onclick="setFilter('open')" data-filter="open" class="filter-btn px-4 py-1.5 rounded-full text-sm font-medium transition-all bg-white text-gray-600 border border-gray-200 hover:bg-gray-50">待处理</button>
      <button onclick="setFilter('in_progress')" data-filter="in_progress" class="filter-btn px-4 py-1.5 rounded-full text-sm font-medium transition-all bg-white text-gray-600 border border-gray-200 hover:bg-gray-50">处理中</button>
      <button onclick="setFilter('resolved')" data-filter="resolved" class="filter-btn px-4 py-1.5 rounded-full text-sm font-medium transition-all bg-white text-gray-600 border border-gray-200 hover:bg-gray-50">已解决</button>
      <button onclick="setFilter('closed')" data-filter="closed" class="filter-btn px-4 py-1.5 rounded-full text-sm font-medium transition-all bg-white text-gray-600 border border-gray-200 hover:bg-gray-50">已关闭</button>
    </div>

    <!-- Ticket List -->
    <div id="ticket-list" class="space-y-3">
      <!-- JS fills -->
    </div>

    <!-- Empty state -->
    <div id="empty-state" class="hidden text-center py-16">
      <div class="w-16 h-16 bg-gray-100 rounded-full flex items-center justify-center text-3xl mx-auto mb-3">📝</div>
      <p class="text-sm text-gray-500">暂无工单，点击上方表单提交第一个报修</p>
    </div>
  </main>

  <!-- Toast -->
  <div id="toast" class="fixed bottom-6 right-6 transform translate-y-24 opacity-0 transition-all duration-300 z-50">
    <div class="bg-gray-900 text-white px-4 py-3 rounded-lg shadow-xl text-sm flex items-center gap-2 min-w-[200px]">
      <span id="toast-icon">✅</span>
      <span id="toast-msg">操作成功</span>
    </div>
  </div>

  <script>
    let currentFilter = 'all';
    let allTickets = [];

    const statusMap = {
      open: { label: '待处理', color: 'bg-orange-50 text-orange-700 border-orange-200' },
      in_progress: { label: '处理中', color: 'bg-blue-50 text-blue-700 border-blue-200' },
      resolved: { label: '已解决', color: 'bg-green-50 text-green-700 border-green-200' },
      closed: { label: '已关闭', color: 'bg-gray-100 text-gray-600 border-gray-200' }
    };
    const priorityMap = {
      low: { label: '低', color: 'bg-gray-100 text-gray-600' },
      medium: { label: '中', color: 'bg-yellow-50 text-yellow-700' },
      high: { label: '高', color: 'bg-orange-50 text-orange-700' },
      critical: { label: '紧急', color: 'bg-red-50 text-red-700' }
    };
    const categoryMap = {
      general: '一般咨询', hardware: '硬件故障', software: '软件问题', network: '网络故障'
    };

    function showToast(msg, icon='✅') {
      const t = document.getElementById('toast');
      document.getElementById('toast-msg').textContent = msg;
      document.getElementById('toast-icon').textContent = icon;
      t.classList.remove('translate-y-24', 'opacity-0');
      setTimeout(() => t.classList.add('translate-y-24', 'opacity-0'), 2800);
    }

    async function api(url, opts={}) {
      try {
        const res = await fetch(url, opts);
        if (!res.ok) throw new Error(await res.text());
        return await res.json();
      } catch (e) {
        showToast('请求失败: ' + e.message, '❌');
        throw e;
      }
    }

    async function loadTickets() {
      const list = document.getElementById('ticket-list');
      list.innerHTML = '<div class="text-center py-10 text-gray-400 text-sm"><span class="loader"></span> 加载中...</div>';
      try {
        const params = currentFilter === 'all' ? '' : '?status=' + currentFilter;
        allTickets = await api('/tickets' + params);
        renderTickets(allTickets);
        updateStats();
      } catch (e) { /* toast handled */ }
    }

    function renderTickets(tickets) {
      const list = document.getElementById('ticket-list');
      const empty = document.getElementById('empty-state');
      if (!tickets || tickets.length === 0) {
        list.innerHTML = '';
        empty.classList.remove('hidden');
        return;
      }
      empty.classList.add('hidden');
      list.innerHTML = tickets.map((t, i) => {
        const st = statusMap[t.status] || statusMap.open;
        const pr = priorityMap[t.priority] || priorityMap.medium;
        const cat = categoryMap[t.category] || t.category;
        const date = t.created_at ? new Date(t.created_at).toLocaleString('zh-CN') : '';
        return `
          <div class="bg-white rounded-xl border border-gray-200 p-4 md:p-5 card-anim hover:shadow-md transition-shadow" style="animation-delay:${i*40}ms">
            <div class="flex flex-col md:flex-row md:items-start md:justify-between gap-3">
              <div class="flex-1 min-w-0">
                <div class="flex items-center gap-2 flex-wrap mb-1">
                  <h3 class="text-sm font-semibold text-gray-900 truncate">${esc(t.title)}</h3>
                  <span class="text-[10px] px-2 py-0.5 rounded-full border ${st.color} font-medium">${st.label}</span>
                  <span class="text-[10px] px-2 py-0.5 rounded-full ${pr.color} font-medium">${pr.label}</span>
                </div>
                <p class="text-xs text-gray-500 mb-2">${esc(t.description || '无描述')}</p>
                <div class="flex items-center gap-3 text-[11px] text-gray-400 flex-wrap">
                  <span>📁 ${cat}</span>
                  <span>📧 ${esc(t.reporter_email || '-')}</span>
                  <span>🕐 ${date}</span>
                  <span class="font-mono text-gray-300">#${esc((t.id||'').slice(0,8))}</span>
                </div>
              </div>
              <div class="flex items-center gap-2 shrink-0">
                ${t.status !== 'resolved' && t.status !== 'closed'
                  ? `<button onclick="resolveTicket('${t.id}')" class="text-xs bg-green-600 hover:bg-green-700 text-white px-3 py-1.5 rounded-md transition-colors font-medium">解决</button>`
                  : ''}
                <button onclick="deleteTicket('${t.id}')" class="text-xs bg-white hover:bg-red-50 text-red-600 border border-red-200 px-3 py-1.5 rounded-md transition-colors font-medium">删除</button>
              </div>
            </div>
          </div>`;
      }).join('');
    }

    function updateStats() {
      document.getElementById('stat-all').textContent = allTickets.length;
      document.getElementById('stat-open').textContent = allTickets.filter(t=>t.status==='open').length;
      document.getElementById('stat-resolved').textContent = allTickets.filter(t=>t.status==='resolved').length;
      const today = new Date().toISOString().slice(0,10);
      document.getElementById('stat-today').textContent = allTickets.filter(t=>(t.created_at||'').startsWith(today)).length;
    }

    function setFilter(status) {
      currentFilter = status;
      document.querySelectorAll('.filter-btn').forEach(btn => {
        if (btn.dataset.filter === status) {
          btn.className = 'filter-btn px-4 py-1.5 rounded-full text-sm font-medium transition-all bg-blue-600 text-white shadow-sm';
        } else {
          btn.className = 'filter-btn px-4 py-1.5 rounded-full text-sm font-medium transition-all bg-white text-gray-600 border border-gray-200 hover:bg-gray-50';
        }
      });
      loadTickets();
    }

    document.getElementById('ticket-form').addEventListener('submit', async (e) => {
      e.preventDefault();
      const btn = document.getElementById('submit-btn');
      const fd = new FormData(e.target);
      const payload = Object.fromEntries(fd.entries());
      if (!payload.title.trim()) return;
      btn.disabled = true;
      btn.innerHTML = '<span class="loader"></span> 提交中...';
      try {
        await api('/tickets', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify(payload) });
        showToast('工单提交成功');
        e.target.reset();
        setFilter('all');
      } catch (e) {} finally {
        btn.disabled = false;
        btn.innerHTML = '<span>提交工单</span>';
      }
    });

    async function resolveTicket(id) {
      if (!confirm('确认将该工单标记为已解决？')) return;
      try {
        await api(`/tickets/${id}`, { method: 'PATCH', headers: {'Content-Type':'application/json'}, body: JSON.stringify({status:'resolved'}) });
        showToast('工单已解决');
        loadTickets();
      } catch (e) {}
    }

    async function deleteTicket(id) {
      if (!confirm('确认删除该工单？此操作不可恢复。')) return;
      try {
        await api(`/tickets/${id}`, { method: 'DELETE' });
        showToast('工单已删除');
        loadTickets();
      } catch (e) {}
    }

    async function exportReport() {
      const btn = document.getElementById('export-btn');
      const original = btn.innerHTML;
      btn.disabled = true;
      btn.innerHTML = '<span class="loader"></span> 导出中...';
      try {
        const res = await api('/reports/export', { method: 'POST', headers: {'Content-Type':'application/json'}, body: JSON.stringify({file_name:'tickets_report.csv'}) });
        showToast(`报表已导出: ${res.file_name}`, '📊');
      } catch (e) {} finally {
        btn.disabled = false;
        btn.innerHTML = original;
      }
    }

    function esc(s) {
      const d = document.createElement('div');
      d.textContent = s || '';
      return d.innerHTML;
    }

    // Init
    loadTickets();
  </script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.getenv("PORT", "8001")),
        reload=os.getenv("DEBUG", "false").lower() == "true",
    )
