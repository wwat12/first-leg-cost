# FBA 头程运费分摊工具（MVP）

上传美通风格物流账单 → 只取「运费」→ 按计费重分摊 →（可选）匹配赛狐 FBA 货件导出 → 得到 SKU 单件头程成本。中文界面，无需登录。

## 永久公网地址（推荐）

**主分享链接以云托管为准**（固定 HTTPS，不依赖本机开机）。部署完成后地址写在下方；若尚未填入，说明正在等待 GitHub / 托管账号连通后自动部署。

| 项 | 值 |
| --- | --- |
| 永久 URL | _（部署完成后填写，形如 `https://first-leg-cost.onrender.com`）_ |
| 托管方 | Render（免费 Web Service）或同等 Python 主机 |
| 休眠说明 | Render 免费实例约 **15 分钟无访问会休眠**；下次打开约需 **30–60 秒** 冷启动。仍比 Cloudflare 临时隧道稳定（域名不变）。 |

### 重新部署

代码改完后：

```bash
# 已连接 GitHub + Render 时
git add -A && git commit -m "update" && git push
# Render 会自动从仓库重新构建；或在 Render Dashboard → Manual Deploy
```

本地 Docker 自测：

```bash
docker build -t first-leg-cost .
docker run --rm -p 8765:8765 -e PORT=8765 first-leg-cost
```

### 不再依赖 Windows Cloudflare 隧道

以前用本机 `cloudflared` Quick Tunnel（`*.trycloudflare.com`）给同事用：每次重启域名会变，且电脑关机即失效。该方式仅作备用，**请改用上表永久 URL**。

备用（本机 Windows）脚本仍在 `C:\Users\WIN10\first-leg-cost-deploy\_win_start.ps1`，仅在云服务不可用时临时使用。

## 本地快速启动

```bash
cd /workspace/first-leg-cost   # 或你的部署目录
python3 -m venv .venv          # 首次
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

浏览器打开：<http://127.0.0.1:8765>

生产启动（云主机）：`gunicorn app:app --bind 0.0.0.0:$PORT`

## 上传什么

1. **物流账单（必需）**  
   美通等对账单 `.xlsx`，含「应收汇总」+ 明细表即可。  
   程序会：
   - 跳过名称含「汇总」的表
   - 自动找表头含「运单号」「费用类型」的行
   - **只保留费用类型 = 运费**（排除清关费等）
   - 金额优先用「折合人民币」，否则用「金额」
   - 计费重用「收费重(KG/CBM)」

2. **赛狐货件导出（可选）**  
   用于把运费分到 SKU。列名不固定时，上传后在页面上映射：
   - 货件号 (FBA)
   - SKU
   - 数量
   - 单件重量 **或** 总重量

3. **店铺**  
   文本框，写入导出 Excel，方便多店铺。

## 分摊公式

1. 货件运费 = 整票总运费 × (该货件计费重 / 整票总计费重)
2. SKU分摊费用 = 货件运费 × (SKU分摊重量 / 货件分摊总重量)
3. 单件头程 = 该SKU所有货件分摊费用之和 / 该SKU总发货数量

说明：MVP 中，账单每一行「运费」先视为货件级成本；若同一运单号对应多个 FBA 且赛狐有各货件重量，则按赛狐重量把整票运费切开再进步骤 2。

**包含全部 SKU**（含名称含「LR」的 SKU），**不按店铺拆分**；店铺字段可选且默认留空。

## 样例账单校验

```bash
source .venv/bin/activate
python verify_sample.py
```

期望：样例 `sample_bill.xlsx` 运费合计 ≈ **36095.4**。

## 目录

```
first-leg-cost/
  app.py              # Flask 入口（读 PORT，绑 0.0.0.0）
  lib/bill_parser.py  # 账单解析
  lib/allocate.py     # 三步分摊
  lib/saihu.py        # 赛狐灵活列映射
  templates/          # 中文 UI
  static/
  uploads/            # 会话上传（云上多为临时盘）
  exports/            # 导出文件
  sample_bill.xlsx    # 样例账单
  Procfile / Dockerfile / render.yaml  # 云部署
```

## 尚未支持（后续）

- 赛狐 API / 固定模板一键导入
- 报价卡自动核价（`rate card` 已留文件，未接 UI）
- 多店铺账套与历史账单库
- 登录与权限
- 持久化磁盘（免费层重启后 uploads 可能清空）

## 无登录提示

任何拿到链接的人都能上传/计算；不要把含敏感账单的链接公开发到外网。
