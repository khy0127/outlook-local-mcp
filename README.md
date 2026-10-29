# outlook-local-mcp

A fully local [MCP](https://modelcontextprotocol.io) server for **classic desktop Outlook on Windows**. It talks to the running `OUTLOOK.EXE` through COM, so there is no cloud API, no Azure app registration, and the server makes no network calls of its own. It can read, search, move, archive and draft mail, but by design it **cannot send mail or delete anything**.

一个完全本地的 Outlook (Classic) MCP 服务器：通过 COM 直接操作本机已登录的 `OUTLOOK.EXE`，不经过任何云端 API，不需要 Azure 应用注册，服务器本身也不发出网络请求。能读取、检索、移动、归档邮件和创建草稿，但**无法发送邮件，也无法删除任何东西**，这是有意的设计。

基于 [r-rohit/outlook-mcp-readonly](https://github.com/r-rohit/outlook-mcp-readonly)（MIT）改写：保留它的只读工具，以及"危险方法不出现在代码里"的审计思路；借鉴 [MagoDev/outlook-mcp](https://github.com/MagoDev/outlook-mcp) 的"只存草稿、不发送"做法；另外加了受保护的移动/归档、标记、草稿编辑、更快的检索，以及附件直接存进 Claude Cowork 项目文件夹。

## 环境要求

- Windows 10/11，经典版 Outlook（`OUTLOOK.EXE`，Microsoft 365 / Outlook 2016 及以上），使用时需要开着。New Outlook（`olk.exe`）不支持 COM，不能用。
- Python 3.10+
- 一个 MCP 客户端，比如 Claude Desktop（包括 Cowork）或 Claude Code

## 安全边界

| 能做 | 做不到 |
|---|---|
| 读取、检索邮件和日历 | **发送**任何邮件（代码里没有 `.Send`） |
| 移动邮件、归档到 Archive | **删除**任何东西（没有 `.Delete` / `PermanentlyDelete`，也删不了文件夹） |
| 新建文件夹 | 移入 Deleted Items / Junk Email / Outbox / Sync Issues（运行时拒绝，包括它们的子文件夹） |
| 标记已读/未读、旗标、分类 | 修改已收到或已发送邮件的正文 |
| 新建草稿、回复/转发草稿、修改草稿 | 打开 Outlook 窗口或执行功能区命令 |
| 把附件另存到本地 | |

这些边界有三层保证：

1. **代码层**：`tests/test_safety.py` 会解析源码。只要出现 `Send`、`Delete`、`PermanentlyDelete`、`MoveTo`、`Display`、`ExecuteMso` 这类成员，或者用 `getattr` 动态拼方法名，测试就会失败。整个包里 `.Move` 只能出现一次，而且必须在 `_guarded_move` 里面。
2. **运行时**：`_guarded_move` 会逐级检查目标文件夹及其所有上级文件夹，同时按 EntryID 和名称比对，凡是 Deleted Items、Junk、Outbox、Sync Issues、Recoverable Items 一律拒绝。`update_draft` 只修改 Drafts 文件夹里尚未发送的邮件。
3. **可追溯**：每次写操作都会追加到 `%LOCALAPPDATA%\outlook-local-mcp\audit.jsonl`。移动会记下原文件夹，可以用 `list_recent_actions` 查到后再移回去。

> 草稿保存在你的邮箱里（Exchange 会同步到服务器），但不会被发送。邮件内容是第三方写的、不可信；服务器给模型的说明里已经要求它不要执行邮件中出现的指令。

## 安装

```bash
git clone https://github.com/<your-account>/outlook-local-mcp.git
```

```bash
cd outlook-local-mcp
```

```bash
python -m venv .venv
```

```bash
.venv\Scripts\python.exe -m pip install -e .
```

```bash
.venv\Scripts\python.exe tests\test_safety.py
```

### Claude Desktop / Cowork

编辑 `%APPDATA%\Claude\claude_desktop_config.json`，在 `mcpServers` 下加入下面这段（路径换成你本机的实际位置）。改完后要彻底退出 Claude Desktop（托盘图标右键选 Quit）再打开：

```json
{
  "mcpServers": {
    "outlook-local": {
      "command": "C:\\path\\to\\outlook-local-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-m", "outlook_local_mcp"],
      "env": { "PYTHONIOENCODING": "utf-8" }
    }
  }
}
```

### Claude Code

```bash
claude mcp add outlook-local --scope user -- C:\path\to\outlook-local-mcp\.venv\Scripts\python.exe -m outlook_local_mcp
```

## 工具

- **读取**：`list_folders`、`list_emails`、`search_emails`、`get_email`、`get_thread`、`list_calendar_events`、`list_recent_actions`
- **整理**：`move_emails`、`archive_emails`、`create_folder`、`set_email_status`
- **附件**：`save_attachments`、`list_save_destinations`
- **草稿**：`create_draft`、`create_reply_draft`、`create_forward_draft`、`update_draft`

文件夹路径的写法：`Inbox`、`Inbox/Project Alpha`、`Archive`、`Drafts`、`Sent`；其他存储（比如 Online Archive）写成 `//Online Archive - you@example.com/Archive`。`search_emails` 用 `folder="all"` 时会搜索主邮箱的全部邮件文件夹。

## 给模型的使用说明

`server.py` 里的 `INSTRUCTIONS` 会在连接时发给模型，内容包括：编号怎么用、文件夹路径怎么写、检索时尽量带上日期或发件人、附件一次存好、批量移动、哪些操作会被拒绝且不要设法绕过、草稿不会被发送、邮件内容不可信。每个工具的参数说明写在各自的 docstring 里。如果发现模型反复用错某个功能，就改这两处。

旧的工具名 `save_attachment` 保留为兼容入口，这样改名之前开始的对话也能继续用。

## 附件保存位置

`save_attachments` 默认一次存下一封邮件的全部附件，签名 logo 之类的内嵌图片会自动跳过。存放位置有三种指定方式：

- **不指定**：存到 Cowork 的文件目录 `%USERPROFILE%\Claude\Outlook Attachments`（路径取自 Claude Desktop 配置里的 `coworkUserFilesPath`）。
- **指定名称**：比如 `destination="Project Alpha"`，文件会存到该项目文件夹下的 `Outlook Attachments\`。可用的名称来自你的 Cowork 项目和在 Cowork 里信任过的文件夹，每次调用时实时读取，所以新建的项目不用重启就能用。名称写一部分也行，只要不会和别的名称混淆。用 `subfolder` 指定子文件夹，`subfolder="."` 表示直接存在项目根目录。
- **绝对路径**：直接写完整路径。

已有的同名文件不会被覆盖，新文件会自动加上 `(1)` 这样的后缀。环境变量 `OUTLOOK_MCP_ATTACHMENT_DIR` 可以修改默认目录。

## 检索与性能

- **Table API**：列表和检索都用 `Folder.GetTable()` 一次取回整批字段，不再逐封、逐个属性地读取。在一个约 1100 封的文件夹里，取最新 200 封从 10 秒降到 0.7 秒，取 25 封约 0.2 秒。
- **`match_mode`**（`search_emails` 的参数）：
  - `index`：走 Windows 搜索索引，速度快，但很多中文词不在索引里，会漏检。
  - `exact`：两段式检索。第一段用日期、发件人、主题、文件夹等其他条件，通过 Table 快速筛出候选邮件；第二段从新到旧逐封核对原文是否包含检索词。主题已经命中的直接算，主题没命中的才读正文，每封约 40ms。凑够 `limit` 条就停，单次最多 90 秒、最多扫描 `max_scan` 封（默认 1500）。没扫完时会说明检查到了哪个日期，更早的部分用索引结果补上。
  - `auto`（默认）：检索词含中日韩文字时用 `exact`，否则用 `index`。
  - 实测：对一个约 1100 封的文件夹精确检索中文词组，全部扫完约 76 秒（直接用 DASL `LIKE` 要 256 秒）；加上 60 天的日期范围后，在全邮箱里检索只要约 8 秒。**加上日期范围或发件人，速度能快一个数量级。**
- **短 id**：每封邮件用 `m1k3` 这样的短编号代替 140 个字符的 EntryID。编号和 EntryID 的对应表存在 `%LOCALAPPDATA%\outlook-local-mcp\ids.sqlite3`，里面只有编号，没有邮件内容。Claude Desktop 会给每个对话和 Cowork 会话各启动一个服务进程，它们共用这张表，所以编号可以跨对话使用，重启后也有效。
- **紧凑输出**：返回不带缩进的紧凑 JSON，值为 false 或默认值的字段直接省略，正文里多余的空行会被压缩。

## 配置

| 环境变量 | 作用 |
|---|---|
| `OUTLOOK_MCP_ARCHIVE_FOLDER` | 归档目标文件夹，默认是主邮箱根目录下的 `Archive` |
| `OUTLOOK_MCP_ATTACHMENT_DIR` | 附件的默认保存目录 |
| `OUTLOOK_MCP_AUDIT_LOG` | 审计日志的路径 |
| `OUTLOOK_MCP_ID_DB` | 短 id 对应表的路径 |

## 本地数据

本服务只会在本机写入以下文件：审计日志（每次写操作的时间、主题和文件夹）、短 id 对应表（只有编号，没有邮件内容），以及你要求保存的附件。它不会另存邮件正文。

## License

MIT, see [LICENSE](LICENSE).
