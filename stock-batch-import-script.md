# 库存批次数据导入 — 开发说明

**文档用途**：根据本文档编写脚本，向数据库导入库存批次数据。脚本需复现线上「创建批次」「调整库存」两个能力对数据库的写入逻辑，保证导入后小程序端展示与业务规则一致。

**数据库**：`guangheshop`（以下 SQL 均使用该库名）

**对应线上接口**（仅供对照，脚本可直接写库，不必调接口）：

| 能力 | 方法 | 路径 |
|------|------|------|
| 创建批次 | POST | `/wx/stock/batch/create` |
| 调整库存 | POST | `/wx/stock/batch/adjust` |

---

## 一、业务背景（必读）

### 1.1 库存怎么算

系统**不**在商品 SKU 表记库存，而是用「批次表」记账：

```text
剩余库存 remaining = total_quantity + adjust_quantity - sold_quantity
```

- `total_quantity`：创建时一次性入库数量，之后一般不改  
- `adjust_quantity`：人工调增/调减的累计量（可正可负）  
- `sold_quantity`：订单卖出扣减的数量（导入历史数据时才可能需要非 0）

### 1.2 创建 / 调整会动哪些表

| 操作 | 写入的表 |
|------|----------|
| **创建批次** | `litemall_stock_batch`、`litemall_stock_batch_region`、`litemall_stock_log` |
| **调整库存** | `litemall_stock_batch`（UPDATE）、`litemall_stock_log`（INSERT） |

**不会写入**：`litemall_goods`、`litemall_goods_product` 的库存字段，活动商品表，订单表。

### 1.3 导入前必须准备好的数据

| 数据 | 来源 | 说明 |
|------|------|------|
| `product_id` | `litemall_goods_product.id` | SKU 必须已存在 |
| `goods_id` | `SELECT goods_id FROM litemall_goods_product WHERE id = ?` | 必须与 SKU 一致，禁止乱填 |
| `creator_id` | 业务提供 | 团长用户 ID，写入批次与日志 |
| `region_code` / `region_name` | 业务提供 | 创建时**至少一个城市**，市级编码如 `320400` |

---

## 二、表结构（完整）

### 2.1 litemall_stock_batch（库存批次主表）

```sql
CREATE TABLE `litemall_stock_batch` (
  `id`               INT NOT NULL AUTO_INCREMENT COMMENT '批次ID',
  `product_id`       INT NOT NULL COMMENT 'SKU规格ID',
  `goods_id`         INT NOT NULL COMMENT '商品ID(冗余)',
  `creator_id`       INT NOT NULL COMMENT '创建者用户ID',
  `scope`            VARCHAR(10) NOT NULL DEFAULT 'PRIVATE' COMMENT 'PRIVATE/PUBLIC',
  `cost_price`       DECIMAL(10,2) DEFAULT NULL COMMENT '采购单价',
  `cost_total`       DECIMAL(12,2) DEFAULT NULL COMMENT '采购总价',
  `wholesale_price`  DECIMAL(10,2) DEFAULT NULL COMMENT '批发单价',
  `wholesale_total`  DECIMAL(12,2) DEFAULT NULL COMMENT '批发总价',
  `total_quantity`   INT NOT NULL COMMENT '入库总量',
  `sold_quantity`    INT NOT NULL DEFAULT 0 COMMENT '已售量',
  `adjust_quantity`  INT NOT NULL DEFAULT 0 COMMENT '调整量',
  `batch_no`         VARCHAR(50) DEFAULT NULL,
  `remark`           VARCHAR(200) DEFAULT NULL,
  `status`           VARCHAR(10) NOT NULL DEFAULT 'AVAILABLE' COMMENT 'AVAILABLE/EXHAUSTED',
  `deleted`          TINYINT(1) NOT NULL DEFAULT 0,
  `add_time`         DATETIME DEFAULT CURRENT_TIMESTAMP,
  `update_time`      DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
);
```

### 2.2 litemall_stock_batch_region（批次可用城市）

```sql
CREATE TABLE `litemall_stock_batch_region` (
  `id`           INT NOT NULL AUTO_INCREMENT,
  `batch_id`     INT NOT NULL COMMENT '库存批次ID',
  `region_code`  INT NOT NULL COMMENT '市级行政区编码，如320400',
  `region_name`  VARCHAR(50) NOT NULL COMMENT '城市名称',
  `add_time`     DATETIME DEFAULT CURRENT_TIMESTAMP,
  `deleted`      TINYINT(1) DEFAULT 0,
  PRIMARY KEY (`id`)
);
```

### 2.3 litemall_stock_log（库存操作日志）

```sql
CREATE TABLE `litemall_stock_log` (
  `id`                INT NOT NULL AUTO_INCREMENT,
  `batch_id`          INT NOT NULL,
  `type`              VARCHAR(20) NOT NULL,
  `quantity`          INT NOT NULL COMMENT '变动数量，正数',
  `before_remaining`  INT NOT NULL,
  `after_remaining`   INT NOT NULL,
  `order_id`          INT DEFAULT NULL,
  `activity_id`       INT DEFAULT NULL,
  `reason`            VARCHAR(200) DEFAULT NULL,
  `operator_id`       INT NOT NULL,
  `add_time`          DATETIME DEFAULT CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
);
```

**日志类型 `type` 取值**（本脚本只需用到前 3 种）：

| type | 含义 | 何时写入 |
|------|------|----------|
| `INBOUND` | 创建入库 | 创建批次时 |
| `ADJUST_ADD` | 调增 | 调整接口 ADD 时 |
| `ADJUST_SUB` | 调减 | 调整接口 SUB 时 |

---

## 三、创建批次 — 脚本要做的事

等价于调用 `POST /wx/stock/batch/create`。一次创建 = **1 条批次 + N 条区域 + 1 条日志**，放在**同一事务**里。

### 3.1 步骤总览

```
步骤1  INSERT  litemall_stock_batch
步骤2  INSERT  litemall_stock_batch_region  （每个城市一行，batch_id = 步骤1 的自增ID）
步骤3  INSERT  litemall_stock_log           （type = INBOUND）
```

### 3.2 步骤1：INSERT litemall_stock_batch

| 字段 | 值 |
|------|-----|
| product_id | 源数据 SKU ID |
| goods_id | 由 SKU 查出的 goods_id |
| creator_id | 团长用户 ID |
| total_quantity | 入库数量，**必须 > 0** |
| sold_quantity | **0** |
| adjust_quantity | **0** |
| scope | `PRIVATE` 或 `PUBLIC`，默认 `PRIVATE` |
| cost_price / cost_total | 采购价，至少填一种，见下方换算 |
| wholesale_price / wholesale_total | 批发价，至少填一种，见下方换算 |
| remark | 可选，NULL 或字符串 |
| batch_no | 可选，NULL |
| status | **`AVAILABLE`**（有库存时） |
| deleted | **0** |

**价格换算规则**（`qty = total_quantity`）：

1. 只给单价 → `总价 = 单价 × qty`（金额保留 2 位小数）  
2. 只给总价 → `单价 = 总价 ÷ qty`（四舍五入到 2 位）  
3. 单价、总价都给 → **以总价为准**，重算单价  
4. 采购价、批发价各自独立换算  

### 3.3 步骤2：INSERT litemall_stock_batch_region

对每个允许使用的城市执行一次 INSERT：

| 字段 | 值 |
|------|-----|
| batch_id | 步骤1 的 `LAST_INSERT_ID()` |
| region_code | 如 `320400` |
| region_name | 如 `常州市`（可为空字符串） |
| deleted | **0** |

**注意**：线上创建接口要求至少 1 个城市；不写区域会导致按城市筛选时批次不可用。

### 3.4 步骤3：INSERT litemall_stock_log

| 字段 | 值 |
|------|-----|
| batch_id | 步骤1 的批次 ID |
| type | **`INBOUND`** |
| quantity | = total_quantity |
| before_remaining | **0** |
| after_remaining | = total_quantity |
| order_id | **NULL** |
| activity_id | **NULL** |
| reason | 固定写 **`创建批次入库`** |
| operator_id | = creator_id |

### 3.5 创建 — 完整 SQL 示例

```sql
START TRANSACTION;

INSERT INTO guangheshop.litemall_stock_batch (
  product_id, goods_id, creator_id, scope,
  cost_price, cost_total, wholesale_price, wholesale_total,
  total_quantity, sold_quantity, adjust_quantity,
  remark, status, deleted
) VALUES (
  101, 50, 266, 'PRIVATE',
  12.00, 1200.00, 15.00, 1500.00,
  100, 0, 0,
  '3月发货', 'AVAILABLE', 0
);
SET @batch_id = LAST_INSERT_ID();

INSERT INTO guangheshop.litemall_stock_batch_region (batch_id, region_code, region_name, deleted)
VALUES (@batch_id, 320400, '常州市', 0);

INSERT INTO guangheshop.litemall_stock_batch_region (batch_id, region_code, region_name, deleted)
VALUES (@batch_id, 320100, '南京市', 0);

INSERT INTO guangheshop.litemall_stock_log (
  batch_id, type, quantity, before_remaining, after_remaining,
  order_id, activity_id, reason, operator_id
) VALUES (
  @batch_id, 'INBOUND', 100, 0, 100,
  NULL, NULL, '创建批次入库', 266
);

COMMIT;
```

### 3.6 创建 — 接口请求体参考（若改调 HTTP）

```json
{
  "productId": 101,
  "totalQuantity": 100,
  "costPrice": "12.00",
  "wholesalePrice": "15.00",
  "scope": "PRIVATE",
  "remark": "3月发货",
  "regionCodes": [320400, 320100],
  "regionNames": ["常州市", "南京市"]
}
```

采购价、批发价可传单价或总价，规则同 3.2。

---

## 四、调整库存 — 脚本要做的事

等价于调用 `POST /wx/stock/batch/adjust`。一次调整 = **UPDATE 批次 + INSERT 日志**，建议同一事务。

### 4.1 步骤总览

```
步骤1  读取当前批次，计算 before_remaining / after_remaining
步骤2  UPDATE litemall_stock_batch  （改 adjust_quantity，可能改 status）
步骤3  INSERT litemall_stock_log    （type = ADJUST_ADD 或 ADJUST_SUB）
```

**不修改**：total_quantity、sold_quantity、价格、区域表。

### 4.2 输入参数

| 参数 | 说明 |
|------|------|
| batch_id | 已存在的批次 ID |
| adjust_type | `ADD` 增加 / `SUB` 减少 |
| quantity | 调整数量，**必须 > 0** |
| reason | 原因说明，建议填写 |
| operator_id | 操作人用户 ID |

### 4.3 计算规则

先查批次当前值：

```text
before_remaining = total_quantity + adjust_quantity - sold_quantity
```

| adjust_type | adjust_quantity 新值 | after_remaining |
|-------------|----------------------|-----------------|
| ADD | 原 adjust_quantity + quantity | before_remaining + quantity |
| SUB | 原 adjust_quantity - quantity | before_remaining - quantity |

**SUB 必须校验**：若 `quantity > before_remaining`，不允许执行（线上返回错误「调减数量超过可用库存」）。

**status 更新规则**：

| 条件 | status 设为 |
|------|-------------|
| after_remaining = 0 | `EXHAUSTED` |
| after_remaining > 0 且当前 status = `EXHAUSTED` | `AVAILABLE` |
| 其他情况 | 保持原 status 不变 |

### 4.4 步骤3：INSERT litemall_stock_log

| 字段 | ADD | SUB |
|------|-----|-----|
| type | `ADJUST_ADD` | `ADJUST_SUB` |
| quantity | 调整数量 | 同左 |
| before_remaining | 见 4.3 | 同左 |
| after_remaining | 见 4.3 | 同左 |
| order_id | NULL | NULL |
| activity_id | NULL | NULL |
| reason | 传入的 reason | 同左 |
| operator_id | 操作人 ID | 同左 |

### 4.5 调整 — 完整 SQL 示例（ADD +10）

```sql
START TRANSACTION;

SELECT total_quantity, adjust_quantity, sold_quantity, status
INTO @tq, @aq, @sq, @old_status
FROM guangheshop.litemall_stock_batch
WHERE id = 291 AND deleted = 0
FOR UPDATE;

SET @qty = 10;
SET @before = @tq + @aq - @sq;
SET @after = @before + @qty;

UPDATE guangheshop.litemall_stock_batch
SET adjust_quantity = @aq + @qty,
    status = CASE
      WHEN @after = 0 THEN 'EXHAUSTED'
      WHEN @after > 0 AND @old_status = 'EXHAUSTED' THEN 'AVAILABLE'
      ELSE status
    END
WHERE id = 291 AND deleted = 0;

INSERT INTO guangheshop.litemall_stock_log (
  batch_id, type, quantity, before_remaining, after_remaining,
  order_id, activity_id, reason, operator_id
) VALUES (
  291, 'ADJUST_ADD', @qty, @before, @after,
  NULL, NULL, '盘点补录', 266
);

COMMIT;
```

**SUB 示例**：`@after = @before - @qty`，须先判断 `@qty <= @before`；`adjust_quantity = @aq - @qty`；日志 `type = 'ADJUST_SUB'`。

### 4.6 调整 — 接口请求体参考（若改调 HTTP）

```json
{
  "batchId": 291,
  "adjustType": "ADD",
  "quantity": 10,
  "reason": "盘点补录"
}
```

`adjustType` 只能是 `ADD` 或 `SUB`。

---

## 五、导入方案（三选一）

### 方案 A：按接口逻辑逐步导入（推荐）

1. 每条新库存先走 **第三章（创建）** 三步  
2. 若源数据还有多次调增/调减，按时间顺序对同一 `batch_id` 重复 **第四章（调整）**  

优点：与线上一致，日志完整。

### 方案 B：直接写入最终状态（快速）

只 INSERT 一条批次，字段设为：

```text
total_quantity  = 初始入库量
adjust_quantity = 历史净调整量（可为负）
sold_quantity   = 历史已售量（无则 0）
status          = (total_quantity + adjust_quantity - sold_quantity) > 0 ? 'AVAILABLE' : 'EXHAUSTED'
```

仍须 INSERT 区域表；建议至少补 1 条 `INBOUND` 日志。

### 方案 C：调用 HTTP 接口

对少量数据 POST 上述两个接口，需登录态 token，由服务端写库。

---

## 六、源数据 CSV 建议格式

### 6.1 创建类数据（create.csv）

| 列名 | 必填 | 说明 |
|------|------|------|
| product_id | 是 | SKU ID |
| creator_id | 是 | 团长 ID |
| total_quantity | 是 | 入库数量 |
| cost_price | 与 cost_total 二选一 | 采购单价 |
| cost_total | 同上 | 采购总价 |
| wholesale_price | 与 wholesale_total 二选一 | 批发单价 |
| wholesale_total | 同上 | 批发总价 |
| scope | 否 | PRIVATE / PUBLIC，默认 PRIVATE |
| remark | 否 | 备注 |
| region_codes | 是 | 逗号分隔，如 `320400,320100` |
| region_names | 否 | 逗号分隔，与 codes 一一对应 |

示例一行：

```text
101,266,100,12.00,,15.00,,PRIVATE,3月发货,320400|320100,常州市|南京市
```

脚本处理：查 `goods_id` → 执行创建三步 → 记录返回的 `batch_id` 供后续调整使用。

### 6.2 调整类数据（adjust.csv）

| 列名 | 必填 | 说明 |
|------|------|------|
| batch_id | 是 | 批次 ID（可由创建步骤生成） |
| adjust_type | 是 | ADD / SUB |
| quantity | 是 | 正整数 |
| reason | 否 | 原因 |
| operator_id | 是 | 操作人 |

示例：

```text
291,ADD,10,盘点补录,266
291,SUB,5,损耗,266
```

---

## 七、速查对照表

| 操作 | 表 | SQL | 要点 |
|------|-----|-----|------|
| 创建 | litemall_stock_batch | INSERT | sold=0, adjust=0, status=AVAILABLE |
| 创建 | litemall_stock_batch_region | INSERT×N | 至少 1 个城市 |
| 创建 | litemall_stock_log | INSERT | type=INBOUND, before=0, after=total_quantity |
| 调整 | litemall_stock_batch | UPDATE | 只改 adjust_quantity、可能 status |
| 调整 | litemall_stock_log | INSERT | type=ADJUST_ADD 或 ADJUST_SUB |

---

## 八、常见错误与校验清单

1. `goods_id` 未从 SKU 查询，与 `product_id` 不匹配 → 数据错乱  
2. 创建未写 `litemall_stock_batch_region` → 按城市选品看不到库存  
3. 调整时修改了 `total_quantity` → 与线上逻辑不一致  
4. SUB 时 `quantity` 大于当前 remaining → 应拒绝执行  
5. 只插批次不插日志 → 小程序详情「操作记录」为空  
6. `remaining` 无数据库字段，不要 INSERT 该列；由三字段计算  
7. 含历史销量时设置 `sold_quantity` 属于方案 B；对应订单扣减日志类型为 `ORDER_DEDUCT`，不在本文档 create/adjust 范围内  

---

## 九、交付物建议

开发完成后脚本应至少支持：

- [ ] 读取 CSV（或 Excel 导出 CSV）  
- [ ] 创建：事务内写 3 张表，输出 `batch_id`  
- [ ] 调整：事务内 UPDATE + INSERT 日志  
- [ ] 失败行记录日志（SKU 不存在、SUB 超库存等）  
- [ ] 可选：dry-run 只打印将执行的 SQL 不写库  

---

文档版本：2026-06-02
