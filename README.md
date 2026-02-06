# TCMSP Related Targets Export (Playwright / Python)

给定一个中药名（中文），脚本会在 TCMSP 网站完成检索与多层跳转，进入药材详情页后抓取 **Related Targets → Targets Infomation**（站点原文拼写如此）的全部靶点数据，并导出到桌面 `Success*.xlsx`。

这个仓库适合作为课程/科研助教示例：包含“稳定抓取策略 + 可复现调试证据链”（URL/DOM 断言、截图/HTML dump、XHR 日志、Playwright trace、自检模式）。

## 你会得到什么
- 输入任意药材中文名（例如：`杜仲`、`三七`、`陈皮`）
- 自动搜索并进入详情页（支持多层 drill-down）
- 抓取 `Related Targets` 下 `Targets Infomation` 的全部条目（例如：杜仲约 100 页，共 1487 条）
- 导出 Excel 到桌面：`Success.xlsx` 或 `Success{序号}.xlsx`
- 自动生成 debug 目录：每一步 URL、页面 HTML、截图、可选 trace

## 运行环境
- macOS（已在 Apple Silicon 上稳定验证）
- Python 3.10+（Python 3.14 亦可）

> 其他系统（Windows/Linux）理论可用，但默认输出路径是 `~/Desktop`；若你的系统没有该目录，请自行创建或改代码中的输出路径。

## macOS 一次性准备（给学生）
1. 安装 Xcode 命令行工具（只需一次）：
```bash
xcode-select --install
```

2. 安装 Python（任选其一）：
- 方式 A：用系统自带 Python（版本可能偏旧，不推荐）
- 方式 B：用 Homebrew 安装（推荐）
```bash
brew install python
```

> 如果你的机器上 `python` 指向的是 Python 2 或版本不对，请用 `python3` 替代 README 里的 `python`。

## 安装（推荐 venv）
```bash
git clone <your-repo-url>
cd <repo>

python -m venv .venv
source .venv/bin/activate

python -m pip install -U pip
python -m pip install -r requirements.txt
python -m playwright install chromium
```

## 最快上手（推荐学生第一次照做）
```bash
HEADLESS=0 SLOW_MO=150 TRACE=1 python tcmsp_related_targets_export_v2.py 杜仲
```

你将看到：
- 输出文件：`~/Desktop/Success.xlsx`
- 调试目录：`./debug/<timestamp>_杜仲/`（包含截图、HTML、URL 记录、trace 等）

## 常用运行方式

### 1) Headless 批量跑（默认）
```bash
HEADLESS=1 python tcmsp_related_targets_export_v2.py 杜仲
```

### 2) 连续多次运行：导出到 `Success1.xlsx`、`Success2.xlsx`…
用环境变量指定序号：
```bash
SUCCESS_INDEX=1 HEADLESS=1 python tcmsp_related_targets_export_v2.py 三七
SUCCESS_INDEX=2 HEADLESS=1 python tcmsp_related_targets_export_v2.py 陈皮
```

或用参数指定：
```bash
HEADLESS=1 python tcmsp_related_targets_export_v2.py 三七 --success-index 1
```

### 3) 自检模式（跑 3 次，覆盖 headless + headed）
```bash
SELF_CHECK=1 python tcmsp_related_targets_export_v2.py 杜仲
```

## 输出说明（Excel）
输出为一个 Sheet（列以站点返回为准），通常包含：
- `molecule_ID`、`MOL_ID`、`molecule_name`
- `target_name`、`target_ID`、`drugbank_ID`
- `validated`、`SVM_score`、`RF_score`

文件位置：
- `~/Desktop/Success.xlsx`
- 或 `~/Desktop/Success{序号}.xlsx`

### 可选：和“参考答案”对比
如果你的桌面存在 `~/Desktop/<药名>_RelatedTargets.xlsx`，脚本会自动做一致性比对：
- 一致：控制台打印 `[COMPARE] ... matched=True`
- 不一致：直接报错，并在 debug 目录落下 `failure.html/.png`、`error.txt` 等

## Debug 目录（如何拿到“可复现证据”）
每次运行都会生成：
`./debug/<timestamp>_<herb>/`

常见文件：
- `entry_*.html/.png`：入口页快照
- `after_search.html/.png`：搜索结果页快照
- `drill_*.html/.png`：每层 drill-down 快照
- `failure.html/.png`：失败现场
- `error.txt`：异常信息 + Python traceback
- `xhr_log.json`：XHR/FETCH 响应日志（有些内容并非 XHR）
- `step_urls.txt`：每一步 URL/状态记录
- `trace.zip`：Playwright trace（当 `TRACE=1`）

打开 trace：
```bash
python -m playwright show-trace debug/<timestamp>_<herb>/trace.zip
```

## 稳定性策略（为什么“能跑通”）
TCMSP 详情页的 `Related Targets → Targets Infomation` 使用 Kendo Grid（`#grid2`）。

直接“点击 tab / 点击分页”在真实网页中非常脆弱（每日弹窗、遮罩层、焦点丢失、UI 重绘等）。因此脚本采用“数据优先”的策略：

1. 尽力切换到 `Related Targets`（但不依赖点击一定成功）
2. **首选**：从浏览器上下文直接读取 Kendo Grid dataSource（值最权威、通常包含全量）
   - `$('#grid2').data('kendoGrid').dataSource.data().toJSON()`
3. **降级**：从 HTML 内嵌脚本解析 `#grid2` 的 `dataSource.data=[...]`
4. **兜底**：逐页采集 DOM 可见表格（兼容极端情况）

## 环境变量（可复现实验参数）
- `HEADLESS=1|0`：默认 `1`
- `SLOW_MO=150`：有界面运行时减速（便于观察）
- `TRACE=1`：生成 `trace.zip`
- `SUCCESS_INDEX=1`：输出到 `Success1.xlsx`
- `SELF_CHECK=1`：连续跑 3 次自检

## 常见问题（Troubleshooting）

### 1) 搜索结果为 0（No items to display）
通常是输入词不在站点里（错别字/别名/繁简体差异）。
- 建议：先在网页上手工搜索，确认“能出结果的写法”
- 脚本行为：导出一个空的 `Success*.xlsx`（列齐全但无数据），并在控制台提示 `No search results`

### 2) `playwright` 浏览器未安装
```bash
python -m playwright install chromium
```

### 3) `ModuleNotFoundError`（缺依赖，例如 `bs4`/`pandas`）
```bash
python -m pip install -r requirements.txt
```

### 4) 站点弹窗遮挡/点击无响应
站点存在每日弹窗 `#dp/#dc`，脚本会自动关闭并移除遮罩。若想“肉眼确认每一步”，用：
```bash
HEADLESS=0 SLOW_MO=150 TRACE=1 python tcmsp_related_targets_export_v2.py 杜仲
```

### 5) 站点偶发超时/抽风
建议重试（脚本已内置 retry）并保留 debug 目录作为证据；必要时提高超时（修改脚本常量 `DEFAULT_TIMEOUT_MS`）。

## 教学/科研使用声明
本脚本仅用于科研/教学演示自动化流程。请遵守 TCMSP 网站的使用条款与合理访问频率，避免对公共服务造成压力。

## 代码入口
- 导出脚本：`tcmsp_related_targets_export_v2.py`
- 依赖文件：`requirements.txt`
