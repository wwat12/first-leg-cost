# FBA 头程运费分摊工具

每月仅上传两份文件：美通风格**物流账单** + 赛狐**装箱/货件导出** → 按账单实付运费与装箱**实际毛重**三步分摊 → SKU 单件头程（含 LR，按店铺分别计算）。**不用物流报价表**。中文界面，无需登录。

## 永久公网地址

| 项 | 值 |
| --- | --- |
| 永久 URL | **https://first-leg-cost.onrender.com** |
| 托管方 | Render Web Service |
| 休眠说明 | 免费实例约 15 分钟无访问会休眠；下次打开约需 30–60 秒冷启动。 |

### 重新部署

```bash
git add -A && git commit -m "update" && git push
# Render 自动从 main 构建；或：
# render deploys create srv-dbi33ih42hec73cai940 --confirm
```

## 分摊公式

1. **货件运费**＝整票实际运费 × 该货件实际毛重 ÷ 本票所有货件实际毛重之和
2. **SKU 分摊**：SKU总重量＝发货数量×单件含包装重量；SKU分摊＝货件运费×SKU总重÷货件内SKU总重之和。**单 SKU 货件：运费全归该 SKU**（无需单件重量）
3. **单件头程**＝SKU分摊÷数量；同 SKU 多货件：合计分摊÷合计数量

## 匹配规则

- 每月两份输入：物流账单 + 赛狐装箱信息（**无物流报价表**）
- 账单：合并全部明细表（跳过汇总）；**不按费用类型筛选**（请上传前自行删除税金/清关费行）；按运单号成票；**货件列表取自 G 列「扩展单号」**中的 FBA… 号（优先于客户单号）
- 赛狐 **装箱信息**：货件编号 ↔ 扩展单号；实际毛重＝总重量或箱子重量(kg)；数量＝单箱数量×箱数；SKU 列；店铺可从同文件「货件信息」带出
- 多 SKU 同货件需要单件含包装重量表；缺则中文报错列出货件
- 分摊合计应等于已匹配票的账单运费（舍入范围内）

## 本地启动

```bash
cd /workspace/first-leg-cost
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
python app.py   # http://127.0.0.1:8765
```

生产：`gunicorn app:app --bind 0.0.0.0:$PORT`

## 回归样例

运单 `NJHL2609107951`（运费 ¥562.50；货件 FBA15MB02L89 / FBA15MB030DD / FBA15MB0BFP3）→ LR-BB223 ≈ ¥33.24，LR-BB232 ≈ ¥20.86。

```bash
python verify_sample.py
```

样例文件：`sample_bill_sept.xlsx` + `sample_saihu_sept.xlsx`。
