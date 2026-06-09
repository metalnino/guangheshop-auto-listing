# -*- coding: utf-8 -*-
"""
自动商品上架程序 — 配置文件
"""
from decimal import Decimal
from pymysql.cursors import DictCursor

# ============================================================
# 数据库配置 (线上库)
# ============================================================
DB_CONFIG = {
    'host': 'rm-uf6c5b0j4975frod0wo.mysql.rds.aliyuncs.com',
    'port': 33306,
    'user': 'guanghe',
    'password': 'guanghe@Z100H.2025',
    'database': 'guangheshop',
    'charset': 'utf8mb4',
    'cursorclass': DictCursor
}

# ============================================================
# 腾讯云 COS 配置
# ============================================================
COS_CONFIG = {
    'secret_id': 'AKIDd18zauJm8D1e4FqqGPGq9CXbNt85r9KE',
    'secret_key': 'mJ6croUu3vWIpQkndEYaMs88quiFWYIp',
    'region': 'ap-shanghai',
    'bucket': 'mpfamily-1301068541',
    'upload_prefix': 'activity/',  # 上传路径前缀
    'upload_to_cos': True,         # 是否上传到 COS (测试环境可置为 False)
}

# ============================================================
# AI 生图管线配置
# ============================================================
AI_OUTPUT_DIR = r"D:\work\AI\AI生图流程化项目_v2\output"

# ============================================================
# 商品基础默认值配置（仅保留不在 Excel 表格中体现的内部字段）
# 用户可见的默认值（单位、公共/私有）已直接预填在 Excel 模板中
# ============================================================
DEFAULTS = {
    'category_id': 0,
    'brand_id': 0,
    'is_on_sale': 1,
    'sort_order': 100,
    'counter_price': Decimal('0.00'),
    'retail_price': Decimal('0.00'),
}
