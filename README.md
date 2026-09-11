# 考勤核对 Web 版（可部署 / GitHub 版）

纯 Python 标准库实现的考勤核对系统：上传总表 → 自动按符号统计 → 主管标色+传凭证 → 管理员审批 → 三色状态 → 导出 Excel。
无需 Flask，只额外依赖 `openpyxl`（处理 xlsx）。

## 这个仓库里该放哪些文件
| 文件 | 作用 | 是否必须 |
|------|------|----------|
| `app.py` | 后端 HTTP 服务（绑定 0.0.0.0，自动建库） | ✅ 必须 |
| `index.html` | 前端页面（运行时从磁盘读取） | ✅ 必须 |
| `requirements.txt` | 依赖：`openpyxl` | ✅ 必须 |
| `Procfile` | 启动命令 `web: python app.py`（Heroku/Render/Railway 通用） | ✅ 建议 |
| `render.yaml` | Render 一键部署配置 | 可选 |
| `.gitignore` | 排除运行时文件 | 建议 |
| `att.db` | 数据库 | ❌ 不进库（首次启动自动建空库，管理员账号自动生成） |

> 不要放进库：`cloudflared`、`*.bat`、`uploads/`、`tests/`、`test_*.py`、`*.log`、`__pycache__/`、`版式预览.html`、各测试 xlsx。

## 本地跑（先验证）
```bash
pip install -r requirements.txt
python app.py
# 浏览器打开 http://localhost:8000
```
默认管理员：`admin` / `admin123`（部署时务必通过环境变量 `ADMIN_PASS` 改成强密码）。

## 部署到公网（生成别人能访问的网址）
GitHub 只负责存代码 + 自动部署；真正给网址的是一台能跑 Python 的免费托管平台。

### 推荐：Render（和 GitHub 联动最顺）
1. GitHub 新建仓库（如 `kaoqin-hedui`），把本文件夹内容 push 上去。
2. 打开 render.com → 用 GitHub 登录 → New → Web Service → 关联该仓库。
3. 配置：Runtime = Python 3；Build Command = `pip install -r requirements.txt`；Start Command = `python app.py`；Plan = Free。
4. 在 Environment 里加变量 `ADMIN_PASS=你的强密码`。
5. 点 Deploy，稍等得到 `https://kaoqin-hedui.onrender.com`，把这个链接发主管/管理员即可。
6. 用 `admin` / 你设的密码登录，首次在后台上传总表、建主管账号。

### 备选：Railway / Fly.io
同样连 GitHub 仓库，启动命令 `python app.py`，并注入 `PORT` 与 `ADMIN_PASS` 环境变量即可。

## ⚠️ 数据持久化（很重要）
免费托管平台的磁盘通常是**临时**的：每次重新部署 / 重启会清空 `att.db`（含上传的考勤、账号、凭证照片）。
- 想长期留存数据：在平台上挂一个**持久磁盘**（Render/Railway/Fly 都支持，费用极低，约 ¥1–2/月级别）。
- 或接受"重新部署=回到空库"，每次部署后重新上传总表。
- 若想保留本机已有的 8 月考勤数据：把本机的 `web版/att.db` 复制进仓库、并从 `.gitignore` 移除 `att.db` 再提交（仅首次需要）。

## 默认账号
- 管理员：`admin` / 见上（环境变量 `ADMIN_PASS`，未设则 `admin123`）
- 主管：由管理员在后台按部门创建
