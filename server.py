# -*- coding: utf-8 -*-
"""词库词根管理系统 - Python后端服务（逐条存储+分页）"""
import http.server
import json
import os
import sys
import urllib.parse
import datetime
import hashlib
import secrets

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PORT = int(os.environ.get('PORT', 8080))
DATABASE_URL = os.environ.get('DATABASE_URL', '')
LOG_FILE = os.path.join(SCRIPT_DIR, 'server.log')

# 数据库适配层：支持 SQLite 和 PostgreSQL
USE_PG = DATABASE_URL.startswith('postgresql')

if USE_PG:
    import psycopg2
    import psycopg2.extras
else:
    import sqlite3
    DATA_DIR = '/data' if os.path.isdir('/data') else SCRIPT_DIR
    DB_FILE = os.path.join(DATA_DIR, 'wordlib.db')

def _pg_connect():
    """创建 PostgreSQL 连接"""
    conn = psycopg2.connect(DATABASE_URL)
    conn.autocommit = False
    return conn

class PgCursorWrapper:
    """PostgreSQL 游标包装器，将 ? 占位符转为 %s，并模拟 sqlite3.Row 行为"""
    def __init__(self, conn):
        self._conn = conn
        self._cursor = conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        self._lastrowid = None

    def execute(self, sql, params=None):
        sql = sql.replace('?', '%s')
        # SQLite 函数适配
        sql = sql.replace("datetime('now','localtime')", "NOW()")
        sql = sql.replace("date('now','localtime')", "CURRENT_DATE")
        sql = sql.replace("date('now','start of month')", "DATE_TRUNC('month', CURRENT_DATE)::date")
        # PostgreSQL 大小写敏感列名处理
        import re as _re
        # 在非建表语句中，给 cnDesc/enDesc/dataType/dataLen/enumValues 加引号
        if 'CREATE TABLE' not in sql.upper():
            sql = _re.sub(r'\bcnDesc\b', '"cnDesc"', sql)
            sql = _re.sub(r'\benDesc\b', '"enDesc"', sql)
            sql = _re.sub(r'\bdataType\b', '"dataType"', sql)
            sql = _re.sub(r'\bdataLen\b', '"dataLen"', sql)
            sql = _re.sub(r'\benumValues\b', '"enumValues"', sql)
        self._cursor.execute(sql, params or ())
        # 获取 lastrowid（INSERT 时）
        if sql.strip().upper().startswith('INSERT') and 'RETURNING' not in sql.upper():
            try:
                # 尝试获取序列值
                self._cursor.execute("SELECT lastval()")
                row = self._cursor.fetchone()
                self._lastrowid = row['lastval'] if row else None
            except:
                self._lastrowid = None
        return self

    def fetchone(self):
        row = self._cursor.fetchone()
        if row is None:
            return None
        return PgRow(row)

    def fetchall(self):
        rows = self._cursor.fetchall()
        return [PgRow(r) for r in rows]

    @property
    def lastrowid(self):
        return self._lastrowid

    def close(self):
        self._cursor.close()

class PgRow:
    """模拟 sqlite3.Row，支持 row['col'] 和 row[0] 访问"""
    def __init__(self, data):
        self._data = data
        self._keys = list(data.keys())

    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self._data.values())[key]
        return self._data.get(key)

    def __contains__(self, key):
        return key in self._data

    def keys(self):
        return self._keys

    def get(self, key, default=None):
        return self._data.get(key, default)

class PgConnWrapper:
    """PostgreSQL 连接包装器，模拟 sqlite3 连接接口"""
    def __init__(self, conn):
        self._conn = conn

    def execute(self, sql, params=None):
        cur = PgCursorWrapper(self._conn)
        cur.execute(sql, params)
        return cur

    def commit(self):
        self._conn.commit()

    def rollback(self):
        self._conn.rollback()

    def close(self):
        self._conn.close()

    @property
    def row_factory(self):
        return None

    @row_factory.setter
    def row_factory(self, val):
        pass

def write_log(msg):
    """写入日志文件并输出到控制台"""
    ts = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    line = f"[{ts}] {msg}"
    # 写日志文件
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(line + '\n')
    except Exception as e:
        print(f"日志写入失败: {e}", flush=True)
    # 输出到控制台（stdout，cmd窗口可见）
    print(line, flush=True)

def _hash_password(password):
    """SHA-256 哈希密码"""
    return hashlib.sha256(password.encode('utf-8')).hexdigest()

def _init_admin_account(conn):
    """初始化 admin 账号（如不存在则创建，已存在则确保密码正确）"""
    existing = conn.execute("SELECT id FROM users WHERE username='admin'").fetchone()
    if not existing:
        conn.execute(
            "INSERT INTO users(username, password_hash, role) VALUES(?,?,?)",
            ('admin', _hash_password('admin123'), 'admin')
        )
    else:
        # 确保 admin 密码哈希是最新的
        conn.execute(
            "UPDATE users SET password_hash=? WHERE username='admin'",
            (_hash_password('admin123'),)
        )

def _verify_token(token):
    """验证 token 有效性，检查是否过期（24小时）
    返回: user_dict 或 None
    """
    conn = get_db()
    row = conn.execute(
        "SELECT s.user_id, s.last_active, u.username, u.role FROM sessions s JOIN users u ON s.user_id=u.id WHERE s.token=?",
        (token,)
    ).fetchone()
    if not row:
        conn.close()
        return None
    # 检查是否超过24小时未活动
    last_active = datetime.datetime.strptime(row['last_active'], '%Y-%m-%d %H:%M:%S')
    if (datetime.datetime.now() - last_active).total_seconds() > 86400:
        conn.execute("DELETE FROM sessions WHERE token=?", (token,))
        conn.commit()
        conn.close()
        return None
    # 更新最后活动时间
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn.execute("UPDATE sessions SET last_active=? WHERE token=?", (now, token))
    conn.commit()
    result = {'user_id': row['user_id'], 'username': row['username'], 'role': row['role']}
    conn.close()
    return result

def get_db():
    if USE_PG:
        raw_conn = _pg_connect()
        conn = PgConnWrapper(raw_conn)
        # PostgreSQL 建表
        conn.execute("""
            CREATE TABLE IF NOT EXISTS words (
                id SERIAL PRIMARY KEY,
                cn TEXT NOT NULL DEFAULT '',
                en TEXT NOT NULL DEFAULT '',
                cat TEXT DEFAULT '',
                roots TEXT DEFAULT '',
                score REAL DEFAULT 0,
                abbr TEXT DEFAULT '',
                "cnDesc" TEXT DEFAULT '',
                "enDesc" TEXT DEFAULT '',
                ref TEXT DEFAULT '',
                "dataType" TEXT DEFAULT '',
                "dataLen" TEXT DEFAULT '',
                "enumValues" TEXT DEFAULT '',
                status TEXT DEFAULT 'draft',
                time TEXT DEFAULT CURRENT_DATE,
                deleted INTEGER DEFAULT 0,
                deleted_time TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS roots (
                id SERIAL PRIMARY KEY,
                name TEXT NOT NULL DEFAULT '',
                en TEXT DEFAULT '',
                mean TEXT DEFAULT '',
                src TEXT DEFAULT '',
                cat TEXT DEFAULT '',
                status TEXT DEFAULT 'draft',
                examples TEXT DEFAULT '[]',
                deleted INTEGER DEFAULT 0,
                deleted_time TEXT
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS asset_history (
                id SERIAL PRIMARY KEY,
                filename TEXT NOT NULL,
                l4_count INTEGER DEFAULT 0,
                l5_count INTEGER DEFAULT 0,
                issue_count INTEGER DEFAULT 0,
                change_count INTEGER DEFAULT 0,
                result_json TEXT DEFAULT '',
                time TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS')
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS word_versions (
                id SERIAL PRIMARY KEY,
                word_id INTEGER NOT NULL,
                version INTEGER NOT NULL,
                snapshot TEXT NOT NULL,
                op_type TEXT NOT NULL,
                operator TEXT DEFAULT 'system',
                time TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS')
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS import_logs (
                id SERIAL PRIMARY KEY,
                batch_id TEXT NOT NULL,
                row_num INTEGER,
                reason TEXT,
                raw_data TEXT,
                time TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS')
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT DEFAULT 'user',
                time TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS')
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                id SERIAL PRIMARY KEY,
                user_id INTEGER NOT NULL,
                token TEXT UNIQUE NOT NULL,
                last_active TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS'),
                time TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS')
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS synonyms (
                id SERIAL PRIMARY KEY,
                word TEXT NOT NULL,
                standard TEXT NOT NULL,
                time TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS')
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS extract_history (
                id SERIAL PRIMARY KEY,
                filename TEXT NOT NULL,
                root_count INTEGER DEFAULT 0,
                field_count INTEGER DEFAULT 0,
                result_json TEXT DEFAULT '',
                time TEXT DEFAULT TO_CHAR(NOW(), 'YYYY-MM-DD HH24:MI:SS')
            )
        """)
        _init_admin_account(conn)
        conn.commit()
        return conn
    else:
        conn = sqlite3.connect(DB_FILE)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        # 词条表
        conn.execute("""
            CREATE TABLE IF NOT EXISTS words (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                cn TEXT NOT NULL DEFAULT '',
                en TEXT NOT NULL DEFAULT '',
                cat TEXT DEFAULT '',
                roots TEXT DEFAULT '',
                score REAL DEFAULT 0,
                abbr TEXT DEFAULT '',
                cnDesc TEXT DEFAULT '',
                enDesc TEXT DEFAULT '',
                ref TEXT DEFAULT '',
                dataType TEXT DEFAULT '',
                dataLen TEXT DEFAULT '',
                enumValues TEXT DEFAULT '',
                status TEXT DEFAULT 'draft',
                time TEXT DEFAULT (date('now','localtime'))
            )
        """)
        # 词根表
        conn.execute("""
            CREATE TABLE IF NOT EXISTS roots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL DEFAULT '',
                en TEXT DEFAULT '',
                mean TEXT DEFAULT '',
                src TEXT DEFAULT '',
                cat TEXT DEFAULT '',
                status TEXT DEFAULT 'draft',
                examples TEXT DEFAULT '[]'
            )
        """)
        # 资产解析历史表
        conn.execute("""
            CREATE TABLE IF NOT EXISTS asset_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                l4_count INTEGER DEFAULT 0,
                l5_count INTEGER DEFAULT 0,
                issue_count INTEGER DEFAULT 0,
                change_count INTEGER DEFAULT 0,
                result_json TEXT DEFAULT '',
                time TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        try:
            conn.execute("SELECT result_json FROM asset_history LIMIT 1")
        except:
            try:
                conn.execute("ALTER TABLE asset_history ADD COLUMN result_json TEXT DEFAULT ''")
            except: pass
        conn.execute("""
            CREATE TABLE IF NOT EXISTS word_versions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                word_id INTEGER NOT NULL,
                version INTEGER NOT NULL,
                snapshot TEXT NOT NULL,
                op_type TEXT NOT NULL,
                operator TEXT DEFAULT 'system',
                time TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS import_logs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                batch_id TEXT NOT NULL,
                row_num INTEGER,
                reason TEXT,
                raw_data TEXT,
                time TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                role TEXT DEFAULT 'user',
                time TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL,
                token TEXT UNIQUE NOT NULL,
                last_active TEXT DEFAULT (datetime('now','localtime')),
                time TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS synonyms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                word TEXT NOT NULL,
                standard TEXT NOT NULL,
                time TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS extract_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT NOT NULL,
                root_count INTEGER DEFAULT 0,
                field_count INTEGER DEFAULT 0,
                result_json TEXT DEFAULT '',
                time TEXT DEFAULT (datetime('now','localtime'))
            )
        """)
        for table in ('words', 'roots'):
            try:
                conn.execute(f"SELECT deleted FROM {table} LIMIT 1")
            except:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN deleted INTEGER DEFAULT 0")
                conn.execute(f"ALTER TABLE {table} ADD COLUMN deleted_time TEXT")
        # 自动迁移：words 表新增技术属性字段
        for col in ('dataType', 'dataLen', 'enumValues'):
            try:
                conn.execute(f"SELECT {col} FROM words LIMIT 1")
            except:
                try:
                    conn.execute(f"ALTER TABLE words ADD COLUMN {col} TEXT DEFAULT ''")
                except: pass
        _init_admin_account(conn)
        conn.commit()
        return conn

def parse_query(qs):
    """解析查询参数"""
    params = {}
    if qs:
        for p in qs.split('&'):
            kv = p.split('=', 1)
            if len(kv) == 2:
                params[kv[0]] = urllib.parse.unquote(kv[1])
    return params

def row_to_dict(row):
    """sqlite3.Row 或 PgRow 转字典"""
    if row is None:
        return None
    if isinstance(row, dict):
        return row
    if hasattr(row, '_data'):
        return dict(row._data)
    return dict(row)

def _soft_delete(conn, table, record_id):
    """软删除：标记 deleted=1 并记录删除时间"""
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn.execute(
        f"UPDATE {table} SET deleted=1, deleted_time=? WHERE id=?",
        (now, record_id)
    )

def _restore_record(conn, table, record_id):
    """恢复：重置 deleted=0 并清除删除时间"""
    conn.execute(
        f"UPDATE {table} SET deleted=0, deleted_time=NULL WHERE id=?",
        (record_id,)
    )

# ========== 版本记录辅助函数（任务 5.1） ==========

def _create_word_version(conn, word_id, old_data, op_type, operator='system'):
    """创建词条版本记录"""
    max_ver = conn.execute(
        "SELECT MAX(version) FROM word_versions WHERE word_id=?", (word_id,)
    ).fetchone()[0] or 0
    now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    conn.execute(
        """INSERT INTO word_versions(word_id, version, snapshot, op_type, operator, time)
           VALUES(?,?,?,?,?,?)""",
        (word_id, max_ver + 1, json.dumps(old_data, ensure_ascii=False), op_type, operator, now)
    )

def _diff_versions(old_snapshot, new_snapshot):
    """对比两个版本快照，返回差异字段列表"""
    diffs = []
    all_keys = set(list(old_snapshot.keys()) + list(new_snapshot.keys()))
    for key in all_keys:
        ov = old_snapshot.get(key, '')
        nv = new_snapshot.get(key, '')
        if str(ov) != str(nv):
            diffs.append({'field': key, 'old_value': str(ov), 'new_value': str(nv)})
    return diffs

# ========== 钢铁行业词根字典（任务 7.1） ==========

KNOWN_ROOTS = {
    # 时间类
    'time': '时间', 'date': '日期', 'datetime': '日期时间', 'year': '年', 'month': '月', 'day': '日',
    'hour': '小时', 'minute': '分钟', 'second': '秒', 'period': '期间', 'start': '开始', 'end': '结束',
    'create': '创建', 'update': '更新', 'delete': '删除', 'modify': '修改',
    # 标识类
    'id': '标识', 'no': '编号', 'code': '代码', 'name': '名称', 'key': '键', 'pk': '主键',
    'seq': '序号', 'num': '数量', 'count': '计数', 'index': '索引', 'order': '排序',
    'type': '类型', 'kind': '种类', 'class': '分类', 'category': '类别', 'group': '分组',
    'status': '状态', 'state': '状态', 'flag': '标志', 'mark': '标记', 'tag': '标签',
    # 操作类
    'add': '新增', 'insert': '插入', 'remove': '移除', 'cancel': '取消',
    'submit': '提交', 'approve': '审批', 'reject': '驳回', 'confirm': '确认',
    'check': '检查', 'verify': '验证', 'audit': '审核', 'review': '评审',
    'import': '导入', 'export': '导出', 'upload': '上传', 'download': '下载',
    # 属性类
    'desc': '描述', 'description': '描述', 'remark': '备注', 'remarks': '备注', 'memo': '备忘',
    'note': '注释', 'comment': '评论', 'content': '内容', 'text': '文本',
    'value': '值', 'amount': '金额', 'price': '价格', 'cost': '成本', 'fee': '费用',
    'rate': '比率', 'ratio': '比例', 'percent': '百分比', 'weight': '重量', 'qty': '数量',
    'size': '尺寸', 'length': '长度', 'width': '宽度', 'height': '高度', 'area': '面积',
    'min': '最小', 'max': '最大', 'avg': '平均', 'sum': '合计', 'total': '总计',
    # 位置类
    'org': '组织', 'dept': '部门', 'company': '公司', 'factory': '工厂', 'plant': '工厂',
    'warehouse': '仓库', 'location': '位置', 'address': '地址', 'area': '区域', 'zone': '区域',
    'line': '产线', 'station': '工位', 'position': '岗位', 'site': '站点',
    # 人员类
    'user': '用户', 'operator': '操作员', 'creator': '创建人', 'updater': '更新人',
    'handler': '处理人', 'inspector': '检验员', 'submitter': '提交人', 'approver': '审批人',
    'person': '人员', 'employee': '员工', 'staff': '员工', 'customer': '客户', 'cust': '客户',
    'supplier': '供应商', 'vendor': '供应商',
    # 钢铁行业
    'heat': '炉次', 'furnace': '炉', 'steel': '钢', 'iron': '铁', 'slab': '铸坯',
    'coil': '卷', 'plate': '板', 'bar': '棒', 'wire': '线材', 'rod': '棒材',
    'rolling': '轧制', 'casting': '铸造', 'smelting': '冶炼', 'refining': '精炼',
    'inspection': '检验', 'test': '试验', 'sample': '样品', 'specimen': '试样',
    'defect': '缺陷', 'quality': '质量', 'grade': '等级', 'standard': '标准',
    'temperature': '温度', 'temp': '温度', 'pressure': '压力', 'speed': '速度',
    'thickness': '厚度', 'diameter': '直径', 'dia': '直径',
    'carbon': '碳', 'silicon': '硅', 'manganese': '锰', 'sulfur': '硫', 'phosphorus': '磷',
    'tensile': '抗拉', 'yield': '屈服', 'elongation': '延伸率', 'hardness': '硬度',
    'batch': '批次', 'lot': '批号', 'charge': '炉号',
    'material': '物料', 'mat': '物料', 'product': '产品', 'prod': '产品', 'item': '项目',
    'equip': '设备', 'equipment': '设备', 'device': '设备', 'machine': '机器',
    'process': '工序', 'procedure': '流程', 'step': '步骤', 'phase': '阶段',
    'plan': '计划', 'schedule': '排程', 'task': '任务', 'job': '作业',
    'result': '结果', 'rslt': '结果', 'record': '记录', 'log': '日志', 'history': '历史',
    'config': '配置', 'setting': '设置', 'param': '参数', 'parameter': '参数',
    'rule': '规则', 'formula': '公式', 'method': '方法', 'mode': '模式',
    'level': '级别', 'tier': '层级', 'version': '版本', 'revision': '修订',
    'source': '来源', 'target': '目标', 'origin': '原始', 'dest': '目的',
    'upper': '上限', 'lower': '下限', 'limit': '限制', 'threshold': '阈值',
    'actual': '实际', 'std': '标准', 'ref': '参考', 'base': '基础',
    'is': '是否', 'has': '是否有', 'can': '是否可', 'enable': '启用', 'disable': '禁用',
    'valid': '有效', 'invalid': '无效', 'active': '活跃', 'inactive': '非活跃',
    'old': '旧', 'new': '新', 'prev': '前', 'next': '后', 'current': '当前', 'cur': '当前',
    'parent': '父级', 'child': '子级', 'master': '主', 'detail': '明细', 'dtl': '明细',
    'main': '主', 'sub': '子', 'header': '表头', 'body': '表体',
    'file': '文件', 'path': '路径', 'url': '链接', 'image': '图片', 'img': '图片',
    'sign': '符号', 'symbol': '符号', 'unit': '单位', 'uom': '单位',
    'sort': '排序', 'filter': '筛选', 'search': '搜索', 'query': '查询',
    'send': '发送', 'receive': '接收', 'transfer': '转移', 'move': '移动',
    'open': '打开', 'close': '关闭', 'lock': '锁定', 'unlock': '解锁',
    'online': '线上', 'offline': '线下', 'announcement': '公告',
    'org_id': '组织ID', 'workstation': '作业站',
    # 华新数据标准规范补充 - 业务域缩写
    'prod': '生产', 'qual': '质量', 'inv': '库存', 'sale': '销售',
    'purch': '采购', 'logi': '物流', 'fin': '财务', 'energy': '能源',
    # 华新数据标准规范补充 - 产品类型
    'hot': '热', 'cold': '冷', 'rolled': '轧', 'galvanized': '镀锌',
    'stainless': '不锈钢', 'strip': '带钢', 'sheet': '薄板', 'billet': '钢坯',
    'bloom': '方坯', 'ingot': '钢锭', 'scrap': '废钢', 'alloy': '合金',
    'pig': '生铁', 'coke': '焦炭', 'ore': '矿石', 'slag': '炉渣',
    # 华新数据标准规范补充 - 工艺参数
    'converter': '转炉', 'ladle': '钢包', 'tundish': '中间包', 'mold': '结晶器',
    'strand': '流', 'pass': '道次', 'reduction': '压下量', 'draft': '压下率',
    'torque': '扭矩', 'tension': '张力', 'gap': '辊缝', 'crown': '凸度',
    'flatness': '平坦度', 'roughness': '粗糙度', 'camber': '镰刀弯',
    'cooling': '冷却', 'heating': '加热', 'annealing': '退火', 'quenching': '淬火',
    'tempering': '回火', 'normalizing': '正火', 'pickling': '酸洗',
    # 华新数据标准规范补充 - 检验相关
    'qualified': '合格', 'unqualified': '不合格', 'acceptance': '验收',
    'certificate': '证书', 'report': '报告', 'specification': '规格',
    'dimension': '尺寸', 'tolerance': '公差', 'deviation': '偏差',
    'surface': '表面', 'internal': '内部', 'external': '外部',
    'crack': '裂纹', 'inclusion': '夹杂', 'porosity': '气孔',
    'impact': '冲击', 'bending': '弯曲', 'flattening': '压扁',
    # 华新数据标准规范补充 - 订单/业务
    'order': '订单', 'contract': '合同', 'delivery': '交货', 'shipment': '发货',
    'invoice': '发票', 'payment': '付款', 'receipt': '收据',
    'inventory': '库存', 'stock': '库存', 'warehouse': '仓库',
    'inbound': '入库', 'outbound': '出库', 'dispatch': '调度',
    'demand': '需求', 'forecast': '预测', 'budget': '预算',
    'project': '项目', 'department': '部门', 'organization': '组织',
    # 华新数据标准规范补充 - 通用后缀
    'remark': '备注', 'memo': '备忘', 'ext': '扩展',
    'seq': '序号', 'idx': '索引', 'cnt': '计数', 'avg': '平均',
    'min': '最小', 'max': '最大', 'sum': '合计', 'total': '总计',
    'begin': '开始', 'finish': '完成', 'duration': '持续时间',
    'frequency': '频率', 'interval': '间隔', 'cycle': '周期',
    'input': '输入', 'output': '输出', 'consumption': '消耗',
    'efficiency': '效率', 'utilization': '利用率', 'capacity': '产能',
    'target': '目标', 'actual': '实际', 'standard': '标准',
    'upper': '上限', 'lower': '下限', 'limit': '限制',
    'previous': '上一个', 'next': '下一个', 'last': '最后',
    'first': '第一', 'second': '第二', 'third': '第三',
    # 补充常见操作/属性词根
    'direct': '直接', 'indirect': '间接', 'set': '设定', 'get': '获取',
    'list': '列表', 'detail': '明细', 'dtl': '明细', 'info': '信息',
    'apply': '申请', 'request': '请求', 'response': '响应',
    'start': '开始', 'stop': '停止', 'begin': '开始', 'finish': '完成',
    'success': '成功', 'fail': '失败', 'error': '错误', 'warn': '警告',
    'local': '本地', 'remote': '远程', 'global': '全局',
    'rebate': '返利', 'discount': '折扣', 'deduct': '扣减',
    'dimension': '维度', 'factor': '因子', 'version': '版本',
    'week': '周', 'daily': '日', 'monthly': '月', 'yearly': '年',
    'price': '价格', 'tax': '税', 'profit': '利润', 'margin': '毛利',
    'serial': '序列', 'ser': '序列', 'no': '编号', 'number': '编号',
    'desc': '描述', 'mm': '毫米', 'kg': '千克', 'pcs': '件',
    'avg': '平均', 'min': '最小', 'max': '最大', 'sum': '合计',
    'plan': '计划', 'real': '实际', 'diff': '差异', 'compare': '比较',
    'assign': '分配', 'alloc': '分配', 'split': '拆分', 'merge': '合并',
    'approve': '审批', 'reject': '驳回', 'confirm': '确认',
    'print': '打印', 'scan': '扫描', 'copy': '复制',
    'attach': '附件', 'doc': '文档', 'template': '模板',
    'notify': '通知', 'message': '消息', 'alert': '告警',
    'sp': '特殊', 'pk': '主键', 'fk': '外键',
}

def _classify_root(root):
    """词根分类"""
    steel_roots = {'heat','furnace','steel','iron','slab','coil','plate','bar','wire','rod','rolling','casting','smelting','refining','inspection','test','sample','specimen','defect','quality','grade','standard','temperature','temp','pressure','speed','thickness','diameter','dia','carbon','silicon','manganese','sulfur','phosphorus','tensile','yield','elongation','hardness','batch','lot','charge','hot','cold','rolled','galvanized','stainless','strip','sheet','billet','bloom','ingot','scrap','alloy','pig','coke','ore','slag','converter','ladle','tundish','mold','strand','pass','reduction','draft','torque','tension','gap','crown','flatness','roughness','camber','cooling','heating','annealing','quenching','tempering','normalizing','pickling','qualified','unqualified','acceptance','certificate','specification','dimension','tolerance','deviation','surface','internal','external','crack','inclusion','porosity','impact','bending','flattening'}
    time_roots = {'time','date','datetime','year','month','day','hour','minute','second','period','start','end','create','update','delete','modify','begin','finish','duration','frequency','interval','cycle','previous','last','first'}
    id_roots = {'id','no','code','name','key','pk','seq','num','count','index','order','idx','cnt'}
    person_roots = {'user','operator','creator','updater','handler','inspector','submitter','approver','person','employee','staff','customer','cust','supplier','vendor'}
    equip_roots = {'equip','equipment','device','machine'}
    qty_roots = {'value','amount','price','cost','fee','rate','ratio','percent','weight','qty','avg','min','max','sum','total','budget'}
    loc_roots = {'org','dept','company','factory','plant','warehouse','location','address','area','zone','line','station','position','site','department','organization'}
    biz_roots = {'prod','qual','inv','sale','purch','logi','fin','energy','order','contract','delivery','shipment','invoice','payment','receipt','inventory','stock','inbound','outbound','dispatch','demand','forecast','project'}
    if root in steel_roots: return '钢铁'
    if root in time_roots: return '时间'
    if root in id_roots: return '标识'
    if root in person_roots: return '人员'
    if root in equip_roots: return '设备'
    if root in qty_roots: return '数量'
    if root in loc_roots: return '位置'
    if root in biz_roots: return '业务'
    return '通用'

# ========== 词根剥离与导入函数（任务 7.1） ==========

def _extract_roots_from_xlsx(filepath):
    """从 xlsx/docx 文件动态剥离词根
    智能识别所有包含英文字段名的列，按 _ 拆分词根，反向查找中文名
    """
    from collections import Counter
    fname = os.path.basename(filepath).lower()

    # 收集所有英文-中文字段对
    field_pairs = set()  # (en, cn)

    if fname.endswith('.docx'):
        # docx 格式：从表格中提取字段
        field_pairs = _extract_fields_from_docx(filepath)
    else:
        # xlsx 格式：智能识别列
        field_pairs = _extract_fields_from_xlsx(filepath)

    if not field_pairs:
        return [], []

    # 构建词条列表（每个字段对就是一个词条）
    word_list = []
    seen_words = set()
    for item in field_pairs:
        # 兼容 (en, cn) 和 (en, cn, tp, ln) 两种格式
        if len(item) >= 4:
            en, cn, tp, ln = item[0], item[1], item[2], item[3]
        else:
            en, cn = item[0], item[1]
            tp, ln = '', ''
        if not en or not cn: continue
        key = en.lower() + '|' + cn
        if key in seen_words: continue
        seen_words.add(key)
        # 从英文名推断分类
        parts = en.lower().split('_')
        cat = '通用类'
        for p in parts:
            rc = _classify_root(p)
            if rc != '通用':
                cat = rc + '类'
                break
        word_list.append({
            'cn': cn, 'en': en, 'cat': cat,
            'roots': json.dumps([cn + '-' + cn], ensure_ascii=False),
            'score': 0, 'abbr': '', 'cnDesc': '', 'enDesc': '',
            'ref': '', 'dataType': tp, 'dataLen': ln, 'enumValues': '', 'status': 'draft',
            'time': datetime.datetime.now().strftime('%Y-%m-%d')
        })

    # ========== 词性判断与词根剥离 ==========
    # 动词→名词映射（常见动词转为对应名词形式）
    _VERB_TO_NOUN = {
        'calculate': 'calculation', 'approve': 'approval', 'delete': 'deletion',
        'create': 'creation', 'update': 'update', 'modify': 'modification',
        'submit': 'submission', 'reject': 'rejection', 'confirm': 'confirmation',
        'check': 'check', 'verify': 'verification', 'audit': 'audit',
        'import': 'import', 'export': 'export', 'upload': 'upload', 'download': 'download',
        'send': 'dispatch', 'receive': 'receipt', 'transfer': 'transfer',
        'open': 'opening', 'close': 'closure', 'lock': 'lock', 'unlock': 'unlock',
        'add': 'addition', 'insert': 'insertion', 'remove': 'removal', 'cancel': 'cancellation',
        'apply': 'application', 'assign': 'assignment', 'alloc': 'allocation',
        'split': 'split', 'merge': 'merge', 'sort': 'sort', 'filter': 'filter',
        'search': 'search', 'query': 'query', 'print': 'print', 'scan': 'scan',
        'copy': 'copy', 'move': 'movement', 'notify': 'notification',
        'inspect': 'inspection', 'test': 'test', 'sample': 'sample',
        'roll': 'rolling', 'cast': 'casting', 'smelt': 'smelting', 'refine': 'refining',
        'cool': 'cooling', 'heat': 'heat', 'anneal': 'annealing', 'quench': 'quenching',
        'temper': 'tempering', 'pickle': 'pickling', 'normalize': 'normalizing',
        'dispatch': 'dispatch', 'settle': 'settlement', 'deduct': 'deduction',
        'evaluate': 'evaluation', 'compare': 'comparison',
    }
    # 修饰词（形容词/副词/过去分词作修饰，不作为独立词根）
    _MODIFIERS = {
        'deleted', 'locked', 'active', 'inactive', 'valid', 'invalid',
        'enabled', 'disabled', 'hidden', 'visible', 'required', 'optional',
        'primary', 'secondary', 'main', 'sub', 'old', 'new', 'prev', 'next',
        'current', 'cur', 'last', 'first', 'upper', 'lower', 'max', 'min',
        'total', 'actual', 'real', 'std', 'avg', 'sum', 'net', 'gross',
        'hot', 'cold', 'raw', 'final', 'initial', 'temp', 'tmp',
        'synced', 'pending', 'approved', 'rejected', 'completed', 'cancelled',
        'updated', 'created', 'modified', 'submitted', 'confirmed',
        'direct', 'indirect', 'local', 'remote', 'global', 'internal', 'external',
        'annual', 'monthly', 'daily', 'weekly', 'yearly',
        'blind', 'mixed', 'manual', 'auto', 'default', 'custom',
    }
    # 停用词
    _STOP_WORDS = {
        'a', 'an', 'the', 'and', 'or', 'not', 'is', 'are', 'was', 'were',
        'be', 'been', 'being', 'have', 'has', 'had', 'do', 'does', 'did',
        'will', 'would', 'shall', 'should', 'may', 'might', 'can', 'could',
        'of', 'in', 'on', 'at', 'to', 'for', 'with', 'by', 'from', 'as',
        'into', 'about', 'between', 'through', 'after', 'before', 'above',
        'below', 'up', 'down', 'out', 'off', 'over', 'under', 'again',
        'then', 'than', 'so', 'if', 'but', 'no', 'yes', 'all', 'each',
        'every', 'both', 'few', 'more', 'most', 'other', 'some', 'such',
        'only', 'own', 'same', 'too', 'very', 'just', 'because', 'also',
        'it', 'its', 'this', 'that', 'these', 'those', 'my', 'your', 'his',
        'her', 'we', 'they', 'me', 'him', 'us', 'them', 'who', 'which',
        'what', 'where', 'when', 'how', 'why', 'there', 'here',
        'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm',
        'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z',
    }

    def _extract_core_root(parts):
        """从单词列表中提取核心名词词根（取最后一个名词）"""
        nouns = []
        for part in parts:
            if not part or len(part) <= 1: continue
            if part.isdigit(): continue
            if part in _STOP_WORDS: continue
            if part in _MODIFIERS: continue
            # 动词转名词（动词不作为独立词根，只用转换后的名词形式）
            if part in _VERB_TO_NOUN:
                noun_form = _VERB_TO_NOUN[part]
                if noun_form in KNOWN_ROOTS:
                    nouns.append(noun_form)
                # 动词本身不保留，跳过
                continue
            # 名词（在 KNOWN_ROOTS 中的）
            if part in KNOWN_ROOTS:
                nouns.append(part)
        # 返回最后一个名词作为核心词根（复合词中核心实体在尾部）
        return nouns[-1] if nouns else None

    # 词根剥离：按 _ 拆分，提取核心名词词根
    root_counter = Counter()
    root_examples = defaultdict(set)

    for item in field_pairs:
        en, cn = item[0], item[1]
        parts = en.lower().strip().split('_')
        # 提取核心词根
        core = _extract_core_root(parts)
        if core:
            root_counter[core] += 1
            root_examples[core].add(en)

    results = []
    for root, count in root_counter.most_common():
        if count < 2: continue
        cn_name = KNOWN_ROOTS.get(root, '')
        if not cn_name: continue
        cat = _classify_root(root)
        results.append({
            'name': cn_name,
            'en': root,
            'mean': cn_name,
            'src': '数据资产',
            'cat': cat,
            'status': 'approved',
            'examples': list(root_examples[root])[:5],
            'count': count
        })
    # 标记中文名重复的词根（不去重，留给用户抉择）
    cn_count = {}
    for r in results:
        cn_count[r['name']] = cn_count.get(r['name'], 0) + 1
    for r in results:
        r['duplicate'] = cn_count[r['name']] > 1
    return sorted(results, key=lambda x: x['count'], reverse=True), word_list


def _extract_fields_from_xlsx(filepath):
    """从 xlsx 文件智能提取英文-中文字段对"""
    import zipfile
    import xml.etree.ElementTree as ET
    z = zipfile.ZipFile(filepath)
    ns_s = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    shared = []
    try:
        ss_raw = z.read('xl/sharedStrings.xml')
        ss_root = ET.fromstring(ss_raw)
        for si in ss_root.findall(f'.//{{{ns_s}}}si'):
            shared.append(''.join(t.text or '' for t in si.findall(f'.//{{{ns_s}}}t')))
    except: pass
    wb_raw = z.read('xl/workbook.xml')
    wb_root = ET.fromstring(wb_raw)
    rels_raw = z.read('xl/_rels/workbook.xml.rels')
    rels_root = ET.fromstring(rels_raw)

    # 遍历所有 sheet（不再只找特定 sheet）
    field_pairs = set()
    rel_map = {r.attrib.get('Id', ''): r.attrib.get('Target', '') for r in rels_root}

    for s in wb_root.findall(f'.//{{{ns_s}}}sheet'):
        rid = s.attrib.get(f'{{{ns_r}}}id', '')
        target = rel_map.get(rid, '')
        if not target: continue
        sheet_path = 'xl/' + target if not target.startswith('/') else target.lstrip('/')
        try:
            sheet_raw = z.read(sheet_path)
        except: continue
        sheet_root = ET.fromstring(sheet_raw)

        def col_idx(col_str):
            r = 0
            for c in col_str:
                r = r * 26 + (ord(c) - ord('A') + 1)
            return r - 1

        rows_data = {}
        for row in sheet_root.findall(f'.//{{{ns_s}}}sheetData/{{{ns_s}}}row'):
            rn = int(row.attrib['r'])
            for cell in row.findall(f'{{{ns_s}}}c'):
                ref = cell.attrib.get('r', '')
                m = re.match(r'([A-Z]+)(\d+)', ref)
                if not m: continue
                ci = col_idx(m.group(1))
                v_el = cell.find(f'{{{ns_s}}}v')
                val = v_el.text if v_el is not None else ''
                if cell.attrib.get('t') == 's' and val:
                    idx = int(val)
                    val = shared[idx] if idx < len(shared) else val
                if rn not in rows_data: rows_data[rn] = {}
                rows_data[rn][ci] = (val or '').strip()

        if not rows_data: continue

        # 智能识别英文列和中文列
        # 策略：找表头中包含"英文"/"字段名"的列作为英文列，相邻的"中文"列作为中文列
        hdr_row = rows_data.get(1, {})
        en_cols = []  # [(en_col_idx, cn_col_idx, tp_col_idx, len_col_idx)]

        # 先找类型列和长度列
        tp_col_global = None
        len_col_global = None
        for ci, val in hdr_row.items():
            if ('类型' in val or 'type' in val.lower()) and tp_col_global is None:
                tp_col_global = ci
            if ('长度' in val or 'length' in val.lower() or '精度' in val) and len_col_global is None:
                len_col_global = ci

        for ci, val in hdr_row.items():
            val_lower = val.lower().replace(' ', '')
            # 识别英文字段列
            if any(kw in val for kw in ['英文', '字段英文', '属性英文', 'English', 'Field']) and '中文' not in val:
                # 找相邻的中文列（通常在英文列的下一列）
                cn_ci = None
                for offset in [1, -1, 2, -2]:
                    neighbor = hdr_row.get(ci + offset, '')
                    if '中文' in neighbor:
                        cn_ci = ci + offset
                        break
                en_cols.append((ci, cn_ci, tp_col_global, len_col_global))

        # 如果没找到明确的表头，尝试自动检测：找包含下划线英文的列
        if not en_cols:
            # 扫描前 20 行数据，找哪些列主要包含 xxx_yyy 格式的英文
            col_scores = defaultdict(int)
            for rn in sorted(rows_data.keys())[:20]:
                row = rows_data[rn]
                for ci, val in row.items():
                    if re.match(r'^[a-zA-Z][a-zA-Z0-9_]*_[a-zA-Z0-9_]+$', val):
                        col_scores[ci] += 1
            # 得分最高的列作为英文列
            if col_scores:
                best_en_col = max(col_scores, key=col_scores.get)
                if col_scores[best_en_col] >= 3:
                    # 找相邻的中文列
                    cn_col = None
                    for offset in [1, -1, 2, -2]:
                        test_col = best_en_col + offset
                        cn_count = 0
                        for rn in sorted(rows_data.keys())[:20]:
                            val = rows_data.get(rn, {}).get(test_col, '')
                            if val and re.search(r'[\u4e00-\u9fff]', val):
                                cn_count += 1
                        if cn_count >= 3:
                            cn_col = test_col
                            break
                    en_cols.append((best_en_col, cn_col, tp_col_global, len_col_global))

        # 提取字段对（含类型和长度）
        for en_ci, cn_ci, tp_ci, len_ci in en_cols:
            for rn in sorted(rows_data.keys()):
                if rn <= 1: continue
                row = rows_data[rn]
                en = row.get(en_ci, '').strip()
                cn = row.get(cn_ci, '').strip() if cn_ci is not None else ''
                tp = row.get(tp_ci, '').strip() if tp_ci is not None else ''
                ln = row.get(len_ci, '').strip() if len_ci is not None else ''
                if en and re.match(r'^[a-zA-Z][a-zA-Z0-9_]*$', en):
                    field_pairs.add((en, cn, tp, ln))

    z.close()
    return field_pairs


def _extract_fields_from_docx(filepath):
    """从 docx 文件提取英文-中文字段对（含字段类型和长度）"""
    import zipfile
    import xml.etree.ElementTree as ET
    z = zipfile.ZipFile(filepath)
    try:
        doc_xml = z.read('word/document.xml')
    except:
        z.close()
        return set()
    root = ET.fromstring(doc_xml)
    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    body = root.find(f'{{{WNS}}}body')
    z.close()
    if body is None:
        return set()

    def get_accepted_text(node):
        """提取接受修订后的文本（忽略 w:del）"""
        text = ''
        for child in node:
            tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
            if tag == 'del': continue
            if tag == 't': text += (child.text or '')
            elif tag == 'delText': continue
            else: text += get_accepted_text(child)
        return text

    field_pairs = set()
    # 遍历所有表格
    for tbl in body.findall(f'.//{{{WNS}}}tbl'):
        rows = []
        for tr in tbl.findall(f'{{{WNS}}}tr'):
            cells = []
            for tc in tr.findall(f'{{{WNS}}}tc'):
                cells.append(get_accepted_text(tc).strip())
            rows.append(cells)
        if len(rows) < 2: continue
        # 找英文列、中文列、类型列、长度列
        hdr = rows[0]
        en_ci, cn_ci, tp_ci, len_ci = None, None, None, None
        for ci, val in enumerate(hdr):
            if '英文' in val and '中文' not in val and en_ci is None:
                en_ci = ci
            elif '中文' in val and cn_ci is None:
                cn_ci = ci
            elif ('类型' in val or 'type' in val.lower()) and tp_ci is None:
                tp_ci = ci
            elif ('长度' in val or 'length' in val.lower() or '精度' in val) and len_ci is None:
                len_ci = ci
        # 如果没找到表头，尝试自动检测
        if en_ci is None:
            for ci, val in enumerate(hdr):
                if re.match(r'^[a-zA-Z][a-zA-Z0-9_]*_', val):
                    en_ci = ci
                    break
        if en_ci is None: continue
        # 提取字段对（含类型和长度）
        for row in rows[1:]:
            if en_ci >= len(row): continue
            en = row[en_ci].strip()
            cn = row[cn_ci].strip() if cn_ci is not None and cn_ci < len(row) else ''
            tp = row[tp_ci].strip() if tp_ci is not None and tp_ci < len(row) else ''
            ln = row[len_ci].strip() if len_ci is not None and len_ci < len(row) else ''
            if en and re.match(r'^[a-zA-Z][a-zA-Z0-9_]*$', en):
                field_pairs.add((en, cn, tp, ln))

    return field_pairs

def _import_extracted_roots(roots_list, mode='skip'):
    """将剥离结果导入 roots 表"""
    conn = get_db()
    imported, skipped = 0, 0
    for r in roots_list:
        existing = conn.execute("SELECT id FROM roots WHERE en=? AND deleted=0", (r['en'],)).fetchone()
        if existing:
            if mode == 'merge':
                conn.execute("UPDATE roots SET name=?,mean=?,src=?,cat=?,status=?,examples=? WHERE id=?",
                    (r['name'], r['mean'], r['src'], r['cat'], r['status'],
                     json.dumps(r.get('examples', []), ensure_ascii=False), existing['id']))
                imported += 1
            else:
                skipped += 1
        else:
            conn.execute("INSERT INTO roots(name,en,mean,src,cat,status,examples) VALUES(?,?,?,?,?,?,?)",
                (r['name'], r['en'], r['mean'], r['src'], r['cat'], r['status'],
                 json.dumps(r.get('examples', []), ensure_ascii=False)))
            imported += 1
    conn.commit()
    conn.close()
    return {'imported': imported, 'skipped': skipped}

# ========== 批量操作辅助函数（任务 8.1） ==========

def _batch_approve_words(word_ids, target_status, operator='system'):
    """批量审核词条，最多 500 条"""
    if len(word_ids) > 500:
        return {'success': 0, 'failed': [{'id': 0, 'reason': '单次最多处理500条'}]}
    conn = get_db()
    success, failed = 0, []
    for wid in word_ids:
        try:
            old_row = conn.execute("SELECT * FROM words WHERE id=? AND deleted=0", (wid,)).fetchone()
            if not old_row:
                failed.append({'id': wid, 'reason': '词条不存在'})
                continue
            old_data = dict(old_row)
            _create_word_version(conn, wid, old_data, '审核变更', operator)
            conn.execute("UPDATE words SET status=?, time=? WHERE id=?", (target_status, datetime.datetime.now().strftime('%Y-%m-%d'), wid))
            success += 1
        except Exception as e:
            failed.append({'id': wid, 'reason': str(e)})
    conn.commit()
    conn.close()
    return {'success': success, 'failed': failed}

def _log_import_error(conn, batch_id, row_num, reason, raw_data):
    """记录导入错误到 import_logs 表"""
    conn.execute("INSERT INTO import_logs(batch_id, row_num, reason, raw_data) VALUES(?,?,?,?)",
        (batch_id, row_num, reason, json.dumps(raw_data, ensure_ascii=False) if isinstance(raw_data, dict) else str(raw_data)))

# ========== 相似度计算辅助函数（任务 9.1） ==========

def _levenshtein_distance(s1, s2):
    """计算 Levenshtein 编辑距离"""
    if len(s1) < len(s2):
        return _levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    prev_row = range(len(s2) + 1)
    for i, c1 in enumerate(s1):
        curr_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = prev_row[j + 1] + 1
            deletions = curr_row[j] + 1
            substitutions = prev_row[j] + (c1 != c2)
            curr_row.append(min(insertions, deletions, substitutions))
        prev_row = curr_row
    return prev_row[-1]

def _normalized_similarity(s1, s2):
    """归一化相似度: 1 - (distance / max(len(s1), len(s2)))"""
    if not s1 and not s2: return 1.0
    if not s1 or not s2: return 0.0
    s1, s2 = s1.lower(), s2.lower()
    if s1 == s2: return 1.0
    dist = _levenshtein_distance(s1, s2)
    max_len = max(len(s1), len(s2))
    return 1 - dist / max_len if max_len > 0 else 0.0

# 简易拼音映射
_PINYIN_MAP = {}
try:
    from pypinyin import lazy_pinyin
    def _to_pinyin(text):
        return ''.join(lazy_pinyin(text))
except ImportError:
    def _to_pinyin(text):
        return text  # 无 pypinyin 时返回原文

def _find_similar_words(en, cn='', top_n=5):
    """查找相似词条"""
    conn = get_db()
    rows = conn.execute("SELECT id, cn, en, cat FROM words WHERE deleted=0 AND status='approved'").fetchall()
    conn.close()
    # 加载同义词映射
    syn_conn = get_db()
    syn_rows = syn_conn.execute("SELECT word, standard FROM synonyms").fetchall()
    syn_conn.close()
    syn_map = {r['word']: r['standard'] for r in syn_rows}
    # 同义词替换
    en_std = en
    for w, s in syn_map.items():
        en_std = en_std.replace(w, s)
    results = []
    for r in rows:
        r_en = r['en'] or ''
        r_cn = r['cn'] or ''
        # 英文相似度
        en_sim = _normalized_similarity(en_std, r_en)
        # 同义词替换后再算
        r_en_std = r_en
        for w, s in syn_map.items():
            r_en_std = r_en_std.replace(w, s)
        en_syn_sim = _normalized_similarity(en_std, r_en_std)
        en_score = max(en_sim, en_syn_sim)
        # 拼音相似度
        py_score = 0.0
        if cn and r_cn:
            py1 = _to_pinyin(cn)
            py2 = _to_pinyin(r_cn)
            py_score = _normalized_similarity(py1, py2)
        # 加权合并
        if cn:
            score = en_score * 0.6 + py_score * 0.4
        else:
            score = en_score
        match_type = []
        if en_sim >= 0.6: match_type.append('编辑距离')
        if en_syn_sim > en_sim: match_type.append('同义词')
        if py_score >= 0.6: match_type.append('拼音')
        if score >= 0.3:
            results.append({
                'id': r['id'], 'cn': r_cn, 'en': r_en, 'cat': r['cat'],
                'score': round(score, 4),
                'match_type': ','.join(match_type) if match_type else '编辑距离'
            })
    results.sort(key=lambda x: x['score'], reverse=True)
    return results[:top_n]

# ========== 统计报表辅助函数（任务 11.1） ==========

def _get_report_overview():
    conn = get_db()
    word_total = conn.execute("SELECT COUNT(*) FROM words WHERE deleted=0").fetchone()[0]
    root_total = conn.execute("SELECT COUNT(*) FROM roots WHERE deleted=0").fetchone()[0]
    # 本月新增：用 Python 计算月份字符串
    month_start = datetime.datetime.now().strftime('%Y-%m-01')
    word_month = conn.execute("SELECT COUNT(*) FROM words WHERE deleted=0 AND time >= ?", (month_start,)).fetchone()[0]
    root_month = conn.execute("SELECT COUNT(*) FROM roots WHERE deleted=0").fetchone()[0]
    conn.close()
    return {'word_total': word_total, 'root_total': root_total, 'word_month': word_month, 'root_month': root_month}

def _get_trend(months=12):
    conn = get_db()
    results = []
    now = datetime.datetime.now()
    for i in range(months - 1, -1, -1):
        # 用 Python 计算月份
        d = now - datetime.timedelta(days=i * 30)
        month_str = d.strftime('%Y-%m')
        count = conn.execute(
            "SELECT COUNT(*) FROM words WHERE deleted=0 AND time LIKE ?", (month_str + '%',)
        ).fetchone()[0]
        results.append({'month': month_str, 'count': count})
    conn.close()
    return results

def _get_approval_rate():
    conn = get_db()
    total = conn.execute("SELECT COUNT(*) FROM words WHERE deleted=0").fetchone()[0]
    statuses = ['draft', 'pending', 'approved', 'rejected', 'offline']
    status_names = {'draft': '草稿', 'pending': '待审核', 'approved': '已通过', 'rejected': '已驳回', 'offline': '已下线'}
    results = []
    for s in statuses:
        count = conn.execute("SELECT COUNT(*) FROM words WHERE deleted=0 AND status=?", (s,)).fetchone()[0]
        results.append({
            'status': s,
            'name': status_names.get(s, s),
            'count': count,
            'percent': round(count / total * 100, 1) if total > 0 else 0
        })
    conn.close()
    return results

def _get_hot_roots(limit=20):
    conn = get_db()
    all_roots = conn.execute("SELECT id, name, en, cat FROM roots WHERE deleted=0").fetchall()
    words_rows = conn.execute("SELECT roots, cn FROM words WHERE deleted=0 AND status='approved'").fetchall()
    conn.close()
    ref_counts = {}
    for r in all_roots:
        count = 0
        for w in words_rows:
            if r['name'] in (w['roots'] or '') or r['name'] in (w['cn'] or ''):
                count += 1
        ref_counts[r['id']] = {'name': r['name'], 'en': r['en'], 'cat': r['cat'], 'ref_count': count}
    sorted_roots = sorted(ref_counts.values(), key=lambda x: x['ref_count'], reverse=True)
    return sorted_roots[:limit]

def _get_category_dist():
    conn = get_db()
    word_cats = conn.execute("SELECT cat, COUNT(*) as cnt FROM words WHERE deleted=0 GROUP BY cat").fetchall()
    root_cats = conn.execute("SELECT cat, COUNT(*) as cnt FROM roots WHERE deleted=0 GROUP BY cat").fetchall()
    conn.close()
    result = {}
    for r in word_cats:
        cat = r['cat'] or '未分类'
        result[cat] = result.get(cat, {'word_count': 0, 'root_count': 0})
        result[cat]['word_count'] = r['cnt']
    for r in root_cats:
        cat = r['cat'] or '未分类'
        result[cat] = result.get(cat, {'word_count': 0, 'root_count': 0})
        result[cat]['root_count'] = r['cnt']
    return [{'cat': k, **v} for k, v in result.items()]

class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        write_log(args[0])

    def _log_op(self, op, detail=''):
        """打印业务操作日志"""
        msg = f"  ✅ {op}"
        if detail:
            msg += f" - {detail}"
        write_log(msg)

    def _cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET,POST,PUT,DELETE,OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')

    def _send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self._cors()
        self.send_header('Content-Length', len(body))
        self.end_headers()
        self.wfile.write(body)

    def _read_body(self):
        length = int(self.headers.get('Content-Length', 0))
        if length == 0:
            return {}
        raw = self.rfile.read(length)
        ct = self.headers.get('Content-Type', '')
        if 'multipart' in ct:
            return raw  # multipart 返回原始字节
        return json.loads(raw.decode('utf-8'))

    def _check_auth(self):
        """认证中间件：验证 API 请求的 token
        返回: user_dict 或 True（无需认证的路径） 或 None（已发送 401 响应）
        """
        path = urllib.parse.urlparse(self.path).path
        # 非 API 路径或登录接口不需要认证
        if not path.startswith('/api/') or path == '/api/login':
            return True
        auth = self.headers.get('Authorization', '')
        token = auth.replace('Bearer ', '') if auth.startswith('Bearer ') else ''
        if not token:
            self._send_json(401, {'error': '未登录，请先登录'})
            return None
        user = _verify_token(token)
        if not user:
            self._send_json(401, {'error': '会话已过期，请重新登录'})
            return None
        return user

    def do_OPTIONS(self):
        self.send_response(204)
        self._cors()
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        params = parse_query(parsed.query)

        # 认证中间件
        if not self._check_auth():
            return

        # === 会话验证API ===
        if path == '/api/session':
            auth = self.headers.get('Authorization', '')
            token = auth.replace('Bearer ', '') if auth.startswith('Bearer ') else ''
            user = _verify_token(token) if token else None
            if user:
                self._send_json(200, {'valid': True, 'username': user['username'], 'role': user['role']})
            else:
                self._send_json(401, {'valid': False, 'error': '未登录或会话已过期'})
            return

        # === 版本历史API（必须在 /api/words 之前） ===
        # /api/words/:id/versions
        if path.startswith('/api/words/') and path.endswith('/versions'):
            wid = int(path.split('/')[3])
            conn = get_db()
            rows = conn.execute("SELECT * FROM word_versions WHERE word_id=? ORDER BY version DESC", (wid,)).fetchall()
            conn.close()
            result = []
            for r in rows:
                d = row_to_dict(r)
                d['snapshot'] = json.loads(d['snapshot']) if d.get('snapshot') else {}
                result.append(d)
            self._send_json(200, result)
            return

        # /api/word_versions/:vid
        if path.startswith('/api/word_versions/'):
            vid = int(path.split('/')[-1])
            conn = get_db()
            row = conn.execute("SELECT * FROM word_versions WHERE id=?", (vid,)).fetchone()
            if not row:
                conn.close()
                self._send_json(404, {'error': '版本不存在'})
                return
            d = row_to_dict(row)
            d['snapshot'] = json.loads(d['snapshot']) if d.get('snapshot') else {}
            # 获取前一版本做差异对比
            prev = conn.execute("SELECT snapshot FROM word_versions WHERE word_id=? AND version=?",
                                (d['word_id'], d['version'] - 1)).fetchone()
            conn.close()
            if prev:
                prev_snap = json.loads(prev['snapshot']) if prev['snapshot'] else {}
                d['diff'] = _diff_versions(prev_snap, d['snapshot'])
            else:
                d['diff'] = []
            self._send_json(200, d)
            return

        # === 相似度API（必须在 /api/words 之前） ===
        if path == '/api/words/similar':
            en = params.get('en', '')
            cn = params.get('cn', '')
            if not en and not cn:
                self._send_json(400, {'error': '请提供 en 或 cn 参数'})
                return
            results = _find_similar_words(en, cn)
            self._send_json(200, results)
            return

        # === 批量操作相关API（必须在 /api/words 之前） ===
        if path == '/api/import_logs':
            conn = get_db()
            rows = conn.execute("SELECT * FROM import_logs ORDER BY time DESC LIMIT 100").fetchall()
            conn.close()
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        # === 词根剥离历史API ===
        if path == '/api/extract_history':
            conn = get_db()
            rows = conn.execute("SELECT id, filename, root_count, field_count, time FROM extract_history ORDER BY id DESC LIMIT 20").fetchall()
            conn.close()
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        if path.startswith('/api/extract_history/'):
            hid = int(path.split('/')[-1])
            conn = get_db()
            row = conn.execute("SELECT * FROM extract_history WHERE id=?", (hid,)).fetchone()
            conn.close()
            if row:
                d = row_to_dict(row)
                if d.get('result_json'):
                    try:
                        parsed = json.loads(d['result_json'])
                        if isinstance(parsed, dict):
                            d['roots'] = parsed.get('roots', [])
                            d['words'] = parsed.get('words', [])
                        elif isinstance(parsed, list):
                            d['roots'] = parsed
                            d['words'] = []
                    except: d['roots'] = []; d['words'] = []
                del d['result_json']
                self._send_json(200, d)
            else:
                self._send_json(404, {'error': '记录不存在'})
            return

        # === 同义词API ===
        if path == '/api/synonyms':
            conn = get_db()
            rows = conn.execute("SELECT * FROM synonyms ORDER BY id DESC").fetchall()
            conn.close()
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        # === 用户管理API（仅admin） ===
        if path == '/api/users':
            auth_header = self.headers.get('Authorization', '')
            tk = auth_header.replace('Bearer ', '') if auth_header.startswith('Bearer ') else ''
            cur_user = _verify_token(tk) if tk else None
            if not cur_user or cur_user.get('role') != 'admin':
                self._send_json(403, {'error': '仅管理员可查看用户列表'})
                return
            conn = get_db()
            rows = conn.execute("SELECT id, username, role, time FROM users ORDER BY id").fetchall()
            conn.close()
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        # === 统计报表API ===
        if path == '/api/report/overview':
            self._send_json(200, _get_report_overview())
            return
        if path == '/api/report/trend':
            months = int(params.get('months', 12))
            self._send_json(200, _get_trend(months))
            return
        if path == '/api/report/approval_rate':
            self._send_json(200, _get_approval_rate())
            return
        if path == '/api/report/hot_roots':
            limit = int(params.get('limit', 20))
            self._send_json(200, _get_hot_roots(limit))
            return
        if path == '/api/report/category_dist':
            self._send_json(200, _get_category_dist())
            return

        # === 词条API ===
        if path == '/api/words':
            conn = get_db()
            search = params.get('search', '')
            cat = params.get('cat', '')
            status = params.get('status', '')
            page = int(params.get('page', 1))
            size = int(params.get('size', 50))

            where = ["deleted=0"]
            args = []
            if search:
                where.append("(cn LIKE ? OR en LIKE ? OR roots LIKE ? OR cnDesc LIKE ?)")
                args.extend([f'%{search}%'] * 4)
            if cat:
                where.append("cat=?")
                args.append(cat)
            if status:
                where.append("status=?")
                args.append(status)

            w = (" WHERE " + " AND ".join(where)) if where else ""
            total = conn.execute(f"SELECT COUNT(*) FROM words{w}", args).fetchone()[0]
            rows = conn.execute(
                f"SELECT * FROM words{w} ORDER BY id DESC LIMIT ? OFFSET ?",
                args + [size, (page - 1) * size]
            ).fetchall()
            conn.close()
            self._log_op('查询词条', f'共{total}条, 第{page}页, 每页{size}条' + (f', 搜索:{search}' if search else '') + (f', 分类:{cat}' if cat else ''))
            self._send_json(200, {
                "total": total, "page": page, "size": size,
                "data": [row_to_dict(r) for r in rows]
            })
            return

        # === 词根API ===
        if path == '/api/roots':
            conn = get_db()
            search = params.get('search', '')
            cat = params.get('cat', '')

            where = ["deleted=0"]
            args = []
            if search:
                where.append("(name LIKE ? OR en LIKE ? OR mean LIKE ?)")
                args.extend([f'%{search}%'] * 3)
            if cat:
                where.append("cat=?")
                args.append(cat)

            w = (" WHERE " + " AND ".join(where)) if where else ""
            rows = conn.execute(f"SELECT * FROM roots{w} ORDER BY id DESC", args).fetchall()
            conn.close()
            self._log_op('查询词根', f'共{len(rows)}条' + (f', 搜索:{search}' if search else '') + (f', 分类:{cat}' if cat else ''))
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        # === 回收站查询API ===
        if path == '/api/recycle_bin':
            rtype = params.get('type', 'words')
            if rtype not in ('words', 'roots'):
                self._send_json(400, {'error': '类型参数无效，请使用 words 或 roots'})
                return
            conn = get_db()
            rows = conn.execute(f"SELECT * FROM {rtype} WHERE deleted=1 ORDER BY deleted_time DESC").fetchall()
            conn.close()
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        # === 统计API ===
        if path == '/api/stats':
            conn = get_db()
            wc = conn.execute("SELECT COUNT(*) FROM words").fetchone()[0]
            rc = conn.execute("SELECT COUNT(*) FROM roots").fetchone()[0]
            conn.close()
            self._log_op('统计', f'词条:{wc}, 词根:{rc}')
            self._send_json(200, {"wordCount": wc, "rootCount": rc})
            return

        # === 初始化数据检查 ===
        if path == '/api/check_init':
            conn = get_db()
            wc = conn.execute("SELECT COUNT(*) FROM words").fetchone()[0]
            rc = conn.execute("SELECT COUNT(*) FROM roots").fetchone()[0]
            conn.close()
            self._send_json(200, {"hasData": wc > 0 or rc > 0})
            return

        # === 资产解析历史 ===
        if path == '/api/asset_history':
            conn = get_db()
            rows = conn.execute("SELECT id,filename,l4_count,l5_count,issue_count,change_count,time FROM asset_history ORDER BY id DESC LIMIT 20").fetchall()
            conn.close()
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        # === 资产历史详情（含完整数据） ===
        if path.startswith('/api/asset_history/'):
            hid = path.split('/')[-1]
            conn = get_db()
            row = conn.execute("SELECT * FROM asset_history WHERE id=?", (hid,)).fetchone()
            conn.close()
            if row:
                d = row_to_dict(row)
                if d.get('result_json'):
                    try: d['result'] = json.loads(d['result_json'])
                    except: d['result'] = None
                del d['result_json']
                self._send_json(200, d)
            else:
                self._send_json(404, {'error': '记录不存在'})
            return

        # === 静态文件 ===
        if path == '/': path = '/index.html'
        filepath = '.' + path
        if os.path.isfile(filepath):
            ext = os.path.splitext(filepath)[1].lower()
            ct_map = {
                '.html': 'text/html; charset=utf-8',
                '.js': 'application/javascript; charset=utf-8',
                '.css': 'text/css; charset=utf-8',
                '.json': 'application/json; charset=utf-8',
                '.png': 'image/png', '.jpg': 'image/jpeg',
                '.svg': 'image/svg+xml', '.ico': 'image/x-icon',
            }
            with open(filepath, 'rb') as f:
                content = f.read()
            self.send_response(200)
            self.send_header('Content-Type', ct_map.get(ext, 'application/octet-stream'))
            self.send_header('Content-Length', len(content))
            self.end_headers()
            self.wfile.write(content)
        else:
            self._send_json(404, {"error": "not found"})

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        # 认证中间件
        if not self._check_auth():
            return

        # asset_parse 需要自己处理 multipart body，必须在 _read_body 之前拦截
        if path == '/api/asset_parse':
            self._handle_asset_parse()
            return

        # extract_roots 也需要处理 multipart 文件上传，必须在 _read_body 之前拦截
        if path == '/api/extract_roots':
            self._handle_extract_roots()
            return

        data = self._read_body()

        # === 登录API ===
        if path == '/api/login':
            username = data.get('username', '')
            password = data.get('password', '')
            conn = get_db()
            user = conn.execute("SELECT id, username, role, password_hash FROM users WHERE username=?", (username,)).fetchone()
            if not user or user['password_hash'] != _hash_password(password):
                conn.close()
                self._send_json(401, {'error': '用户名或密码错误'})
                return
            token = secrets.token_hex(32)
            conn.execute("INSERT INTO sessions(user_id, token) VALUES(?,?)", (user['id'], token))
            conn.commit()
            conn.close()
            self._log_op('用户登录', f'{username}')
            self._send_json(200, {'token': token, 'username': user['username'], 'role': user['role']})
            return

        # === 登出API ===
        if path == '/api/logout':
            auth = self.headers.get('Authorization', '')
            token = auth.replace('Bearer ', '') if auth.startswith('Bearer ') else ''
            if token:
                conn = get_db()
                conn.execute("DELETE FROM sessions WHERE token=?", (token,))
                conn.commit()
                conn.close()
            self._log_op('用户登出')
            self._send_json(200, {'msg': 'ok'})
            return

        # === 回收站恢复API ===
        if path == '/api/recycle_bin/restore':
            rtype = data.get('type', '')
            rid = data.get('id', 0)
            if rtype not in ('words', 'roots') or not rid:
                self._send_json(400, {'error': '参数无效'})
                return
            conn = get_db()
            _restore_record(conn, rtype, rid)
            conn.commit()
            conn.close()
            self._log_op('恢复记录', f'{rtype} ID:{rid}')
            self._send_json(200, {'msg': 'ok'})
            return

        # === 回收站清理过期记录API ===
        if path == '/api/recycle_bin/cleanup':
            conn = get_db()
            for table in ('words', 'roots'):
                cutoff = (datetime.datetime.now() - datetime.timedelta(days=30)).strftime('%Y-%m-%d %H:%M:%S')
                conn.execute(f"DELETE FROM {table} WHERE deleted=1 AND deleted_time < ?", (cutoff,))
            conn.commit()
            conn.close()
            self._log_op('清理过期回收站记录')
            self._send_json(200, {'msg': 'ok'})
            return

        # === 批量审核API（必须在 /api/words 之前） ===
        if path == '/api/words/batch_approve':
            word_ids = data.get('ids', [])
            target_status = data.get('status', 'approved')
            result = _batch_approve_words(word_ids, target_status)
            self._log_op('批量审核', f'成功:{result["success"]}, 失败:{len(result["failed"])}')
            self._send_json(200, result)
            return

        # === 批量删除/下线API ===
        if path == '/api/words/batch_action':
            ids = data.get('ids', [])
            action = data.get('action', '')
            if not ids or action not in ('delete', 'offline'):
                self._send_json(400, {'error': '参数无效'})
                return
            conn = get_db()
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            count = 0
            for wid in ids:
                try:
                    if action == 'delete':
                        conn.execute("UPDATE words SET deleted=1, deleted_time=? WHERE id=?", (now, wid))
                    else:
                        conn.execute("UPDATE words SET status='offline', time=? WHERE id=?", (datetime.datetime.now().strftime('%Y-%m-%d'), wid))
                    count += 1
                except: pass
            conn.commit()
            conn.close()
            self._log_op(f'批量{action}词条', f'{count}/{len(ids)}')
            self._send_json(200, {'success': count, 'total': len(ids)})
            return

        if path == '/api/roots/batch_action':
            ids = data.get('ids', [])
            action = data.get('action', '')
            if not ids or action not in ('delete', 'offline'):
                self._send_json(400, {'error': '参数无效'})
                return
            conn = get_db()
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            count = 0
            for rid in ids:
                try:
                    if action == 'delete':
                        conn.execute("UPDATE roots SET deleted=1, deleted_time=? WHERE id=?", (now, rid))
                    else:
                        conn.execute("UPDATE roots SET status='offline' WHERE id=?", (rid,))
                    count += 1
                except: pass
            conn.commit()
            conn.close()
            self._log_op(f'批量{action}词根', f'{count}/{len(ids)}')
            self._send_json(200, {'success': count, 'total': len(ids)})
            return

        # === 导出API（必须在 /api/words 之前） ===
        if path == '/api/words/export':
            conn = get_db()
            where = ["deleted=0"]
            args = []
            if data.get('cat'):
                where.append("cat=?")
                args.append(data['cat'])
            if data.get('status'):
                where.append("status=?")
                args.append(data['status'])
            if data.get('time_start'):
                where.append("time>=?")
                args.append(data['time_start'])
            if data.get('time_end'):
                where.append("time<=?")
                args.append(data['time_end'])
            w = " WHERE " + " AND ".join(where)
            rows = conn.execute(f"SELECT * FROM words{w} ORDER BY id DESC", args).fetchall()
            conn.close()
            self._send_json(200, [row_to_dict(r) for r in rows])
            return

        # === 导出剥离结果为 xlsx ===
        if path == '/api/extract_roots/export':
            roots_list = data.get('roots', [])
            if not roots_list:
                self._send_json(400, {'error': '没有可导出的数据'})
                return
            try:
                import xlsxwriter
            except ImportError:
                self._send_json(500, {'error': '缺少 xlsxwriter'})
                return
            buf = io.BytesIO()
            wb = xlsxwriter.Workbook(buf)
            hdr_fmt = wb.add_format({'bold': 1, 'bg_color': '#4472C4', 'font_color': '#FFF', 'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 11})
            cell_fmt = wb.add_format({'border': 1, 'text_wrap': 1, 'valign': 'vcenter', 'font_size': 10})
            ws = wb.add_worksheet('词根剥离结果')
            headers = ['序号', '词根名称', '英文', '含义', '分类', '出现次数', '示例']
            for c, h in enumerate(headers):
                ws.write(0, c, h, hdr_fmt)
            for i, r in enumerate(roots_list):
                ex = r.get('examples', [])
                if isinstance(ex, list): ex = '、'.join(ex)
                ws.write(i+1, 0, i+1, cell_fmt)
                ws.write(i+1, 1, r.get('name', ''), cell_fmt)
                ws.write(i+1, 2, r.get('en', ''), cell_fmt)
                ws.write(i+1, 3, r.get('mean', ''), cell_fmt)
                ws.write(i+1, 4, r.get('cat', ''), cell_fmt)
                ws.write(i+1, 5, r.get('count', 0), cell_fmt)
                ws.write(i+1, 6, ex, cell_fmt)
            for c, w in enumerate([6, 14, 14, 20, 10, 10, 40]):
                ws.set_column(c, c, w)
            ws.autofilter(0, 0, len(roots_list), 6)
            ws.freeze_panes(1, 0)
            wb.close()
            excel_bytes = buf.getvalue()
            self.send_response(200)
            self._cors()
            self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition', 'attachment; filename="roots_export.xlsx"')
            self.send_header('Content-Length', str(len(excel_bytes)))
            self.end_headers()
            self.wfile.write(excel_bytes)
            return

        # === 导入词根API ===
        if path == '/api/import_roots':
            roots_list = data.get('roots', [])
            mode = data.get('mode', 'skip')
            result = _import_extracted_roots(roots_list, mode)
            self._send_json(200, result)
            return

        # === 同义词API ===
        if path == '/api/synonyms':
            word = data.get('word', '')
            standard = data.get('standard', '')
            if not word or not standard:
                self._send_json(400, {'error': '请提供 word 和 standard'})
                return
            conn = get_db()
            existing = conn.execute("SELECT id FROM synonyms WHERE word=?", (word,)).fetchone()
            if existing:
                conn.execute("UPDATE synonyms SET standard=? WHERE id=?", (standard, existing['id']))
            else:
                conn.execute("INSERT INTO synonyms(word, standard) VALUES(?,?)", (word, standard))
            conn.commit()
            conn.close()
            self._send_json(200, {'msg': 'ok'})
            return

        # === 用户管理API（仅admin） ===
        if path == '/api/users':
            # 检查当前用户是否为admin
            auth_header = self.headers.get('Authorization', '')
            tk = auth_header.replace('Bearer ', '') if auth_header.startswith('Bearer ') else ''
            cur_user = _verify_token(tk) if tk else None
            if not cur_user or cur_user.get('role') != 'admin':
                self._send_json(403, {'error': '仅管理员可管理用户'})
                return
            username = data.get('username', '').strip()
            password = data.get('password', '').strip()
            role = data.get('role', 'user')
            if not username or not password:
                self._send_json(400, {'error': '请提供用户名和密码'})
                return
            if role not in ('admin', 'user'):
                role = 'user'
            conn = get_db()
            existing = conn.execute("SELECT id FROM users WHERE username=?", (username,)).fetchone()
            if existing:
                conn.close()
                self._send_json(400, {'error': f'用户名 {username} 已存在'})
                return
            conn.execute("INSERT INTO users(username, password_hash, role) VALUES(?,?,?)",
                (username, _hash_password(password), role))
            conn.commit()
            conn.close()
            self._log_op('新增用户', f'{username}, 角色:{role}')
            self._send_json(200, {'msg': 'ok'})
            return

        # === 重置用户密码API（仅admin） ===
        if path == '/api/users/reset_password':
            auth_header = self.headers.get('Authorization', '')
            tk = auth_header.replace('Bearer ', '') if auth_header.startswith('Bearer ') else ''
            cur_user = _verify_token(tk) if tk else None
            if not cur_user or cur_user.get('role') != 'admin':
                self._send_json(403, {'error': '仅管理员可重置密码'})
                return
            uid = data.get('id', 0)
            new_password = data.get('password', '').strip()
            if not uid or not new_password:
                self._send_json(400, {'error': '请提供用户ID和新密码'})
                return
            conn = get_db()
            conn.execute("UPDATE users SET password_hash=? WHERE id=?", (_hash_password(new_password), uid))
            conn.commit()
            conn.close()
            self._log_op('重置密码', f'用户ID:{uid}')
            self._send_json(200, {'msg': 'ok'})
            return

        if path == '/api/words':
            conn = get_db()
            cur = conn.execute(
                """INSERT INTO words(cn,en,cat,roots,score,abbr,cnDesc,enDesc,ref,dataType,dataLen,enumValues,status,time)
                   VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (data.get('cn',''), data.get('en',''), data.get('cat',''),
                 data.get('roots',''), data.get('score',0), data.get('abbr',''),
                 data.get('cnDesc',''), data.get('enDesc',''), data.get('ref',''),
                 data.get('dataType',''), data.get('dataLen',''), data.get('enumValues',''),
                 data.get('status','draft'), data.get('time',''))
            )
            conn.commit()
            new_id = cur.lastrowid
            conn.close()
            self._log_op('新增词条', f'ID:{new_id}, {data.get("cn","")}({data.get("en","")}), 分类:{data.get("cat","")}, 状态:{data.get("status","")}')
            self._send_json(201, {"id": new_id, "msg": "ok"})
            return

        if path == '/api/roots':
            conn = get_db()
            examples = data.get('examples', [])
            if isinstance(examples, list):
                examples = json.dumps(examples, ensure_ascii=False)
            cur = conn.execute(
                """INSERT INTO roots(name,en,mean,src,cat,status,examples)
                   VALUES(?,?,?,?,?,?,?)""",
                (data.get('name',''), data.get('en',''), data.get('mean',''),
                 data.get('src',''), data.get('cat',''),
                 data.get('status','draft'), examples)
            )
            conn.commit()
            new_id = cur.lastrowid
            conn.close()
            self._log_op('新增词根', f'ID:{new_id}, {data.get("name","")}({data.get("en","")}), 分类:{data.get("cat","")}')
            self._send_json(201, {"id": new_id, "msg": "ok"})
            return

        # 批量初始化导入
        if path == '/api/init_words':
            items = data if isinstance(data, list) else data.get('words', [])
            conn = get_db()
            for w in items:
                conn.execute(
                    """INSERT INTO words(cn,en,cat,roots,score,abbr,cnDesc,enDesc,ref,dataType,dataLen,enumValues,status,time)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (w.get('cn',''), w.get('en',''), w.get('cat',''),
                     w.get('roots',''), w.get('score',0), w.get('abbr',''),
                     w.get('cnDesc',''), w.get('enDesc',''), w.get('ref',''),
                     w.get('dataType',''), w.get('dataLen',''), w.get('enumValues',''),
                     w.get('status','approved'), w.get('time',''))
                )
            conn.commit()
            conn.close()
            self._log_op('批量初始化词条', f'共{len(items)}条')
            self._send_json(200, {"msg": "ok", "count": len(items)})
            return

        if path == '/api/init_roots':
            items = data if isinstance(data, list) else data.get('roots', [])
            conn = get_db()
            for r in items:
                examples = r.get('examples', [])
                if isinstance(examples, list):
                    examples = json.dumps(examples, ensure_ascii=False)
                conn.execute(
                    """INSERT INTO roots(name,en,mean,src,cat,status,examples)
                       VALUES(?,?,?,?,?,?,?)""",
                    (r.get('name',''), r.get('en',''), r.get('mean',''),
                     r.get('src',''), r.get('cat',''),
                     r.get('status','approved'), examples)
                )
            conn.commit()
            conn.close()
            self._log_op('批量初始化词根', f'共{len(items)}条')
            self._send_json(200, {"msg": "ok", "count": len(items)})
            return

        # ========== 数据资产整改 API ==========
        if path == '/api/asset_analyze':
            self._handle_asset_analyze(data)
            return
        if path == '/api/asset_export':
            self._handle_asset_export(data)
            return

        self._send_json(404, {"error": "not found"})

    def do_PUT(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path

        # 认证中间件
        if not self._check_auth():
            return

        data = self._read_body()

        # /api/words/123
        if path.startswith('/api/words/'):
            wid = int(path.split('/')[-1])
            conn = get_db()
            # 在 UPDATE 之前创建版本记录
            old_row = conn.execute("SELECT * FROM words WHERE id=?", (wid,)).fetchone()
            if old_row:
                old_data = dict(old_row)
                op_type = '审核变更' if old_data.get('status') != data.get('status', '') else '编辑'
                _create_word_version(conn, wid, old_data, op_type)
            conn.execute(
                """UPDATE words SET cn=?,en=?,cat=?,roots=?,score=?,abbr=?,
                   cnDesc=?,enDesc=?,ref=?,dataType=?,dataLen=?,enumValues=?,status=?,time=? WHERE id=?""",
                (data.get('cn',''), data.get('en',''), data.get('cat',''),
                 data.get('roots',''), data.get('score',0), data.get('abbr',''),
                 data.get('cnDesc',''), data.get('enDesc',''), data.get('ref',''),
                 data.get('dataType',''), data.get('dataLen',''), data.get('enumValues',''),
                 data.get('status',''), data.get('time',''), wid)
            )
            conn.commit()
            conn.close()
            self._log_op('更新词条', f'ID:{wid}, {data.get("cn","")}({data.get("en","")}), 状态:{data.get("status","")}')
            self._send_json(200, {"msg": "ok"})
            return

        if path.startswith('/api/roots/'):
            rid = int(path.split('/')[-1])
            examples = data.get('examples', [])
            if isinstance(examples, list):
                examples = json.dumps(examples, ensure_ascii=False)
            conn = get_db()
            conn.execute(
                """UPDATE roots SET name=?,en=?,mean=?,src=?,cat=?,status=?,examples=?
                   WHERE id=?""",
                (data.get('name',''), data.get('en',''), data.get('mean',''),
                 data.get('src',''), data.get('cat',''),
                 data.get('status',''), examples, rid)
            )
            conn.commit()
            conn.close()
            self._log_op('更新词根', f'ID:{rid}, {data.get("name","")}({data.get("en","")}), 状态:{data.get("status","")}')
            self._send_json(200, {"msg": "ok"})
            return

        self._send_json(404, {"error": "not found"})

    def do_DELETE(self):
        path = urllib.parse.urlparse(self.path).path

        # 认证中间件
        if not self._check_auth():
            return

        # === 回收站永久删除 API ===
        import re as _re_mod
        m = _re_mod.match(r'^/api/recycle_bin/(words|roots)/(\d+)$', path)
        if m:
            rtype, rid = m.group(1), int(m.group(2))
            conn = get_db()
            conn.execute(f"DELETE FROM {rtype} WHERE id=? AND deleted=1", (rid,))
            conn.commit()
            conn.close()
            self._log_op('永久删除', f'{rtype} ID:{rid}')
            self._send_json(200, {'msg': 'ok'})
            return

        if path.startswith('/api/words/'):
            wid = int(path.split('/')[-1])
            conn = get_db()
            # 先查出词条信息用于日志
            row = conn.execute("SELECT cn,en FROM words WHERE id=?", (wid,)).fetchone()
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            conn.execute("UPDATE words SET deleted=1, deleted_time=? WHERE id=?", (now, wid))
            conn.commit()
            conn.close()
            if row:
                self._log_op('删除词条', f'ID:{wid}, {row["cn"]}({row["en"]})')
            else:
                self._log_op('删除词条', f'ID:{wid}')
            self._send_json(200, {"msg": "ok"})
            return

        if path.startswith('/api/roots/'):
            rid = int(path.split('/')[-1])
            conn = get_db()
            # 先查出词根信息用于日志
            row = conn.execute("SELECT name,en FROM roots WHERE id=?", (rid,)).fetchone()
            now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            conn.execute("UPDATE roots SET deleted=1, deleted_time=? WHERE id=?", (now, rid))
            conn.commit()
            conn.close()
            if row:
                self._log_op('删除词根', f'ID:{rid}, {row["name"]}({row["en"]})')
            else:
                self._log_op('删除词根', f'ID:{rid}')
            self._send_json(200, {"msg": "ok"})
            return

        self._send_json(404, {"error": "not found"})

    # ========== 词根剥离：上传 xlsx 并剥离词根 ==========
    def _handle_extract_roots(self):
        ct = self.headers.get('Content-Type', '')
        if 'multipart' not in ct:
            self._send_json(400, {'error': '请上传 multipart/form-data'})
            return

        boundary = ct.split('boundary=')[1].strip()
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length)

        fname = 'unknown'
        file_data = None
        for part in body.split(('--' + boundary).encode()):
            if b'filename=' in part:
                idx = part.find(b'\r\n\r\n')
                if idx >= 0:
                    file_data = part[idx + 4:]
                    if file_data.endswith(b'\r\n'):
                        file_data = file_data[:-2]
                hdr_part = part[:part.find(b'\r\n\r\n')].decode('utf-8', errors='ignore')
                import re as _re
                fn_match = _re.search(r'filename="([^"]+)"', hdr_part)
                if fn_match:
                    fname = fn_match.group(1)

        if not file_data:
            self._send_json(400, {'error': '未找到上传文件'})
            return

        if not (fname.lower().endswith('.xlsx') or fname.lower().endswith('.docx')):
            self._send_json(400, {'error': '仅支持 xlsx 和 docx 格式文件'})
            return

        import tempfile
        suffix = '.docx' if fname.lower().endswith('.docx') else '.xlsx'
        tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        tmp.write(file_data)
        tmp.close()

        try:
            results, word_list = _extract_roots_from_xlsx(tmp.name)
            self._log_op('词根剥离', f'{fname} → 剥离出 {len(results)} 个词根, {len(word_list)} 个词条')
            # 保存剥离历史
            try:
                conn = get_db()
                conn.execute("INSERT INTO extract_history(filename, root_count, field_count, result_json) VALUES(?,?,?,?)",
                    (fname, len(results), len(word_list),
                     json.dumps({'roots': results, 'words': word_list}, ensure_ascii=False)))
                conn.commit()
                conn.close()
            except: pass
            self._send_json(200, {'roots': results, 'words': word_list, 'count': len(results), 'wordCount': len(word_list)})
        except Exception as e:
            self._send_json(500, {'error': str(e)})
        finally:
            os.unlink(tmp.name)

    # ========== 数据资产整改：解析 docx ==========
    def _handle_asset_parse(self):
        ct = self.headers.get('Content-Type', '')
        if 'multipart' not in ct:
            self._send_json(400, {'error': '请上传 multipart/form-data'})
            return

        boundary = ct.split('boundary=')[1].strip()
        length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(length)

        fname = 'unknown'
        file_data = None
        for part in body.split(('--' + boundary).encode()):
            if b'filename=' in part:
                idx = part.find(b'\r\n\r\n')
                if idx >= 0:
                    file_data = part[idx + 4:]
                    if file_data.endswith(b'\r\n'):
                        file_data = file_data[:-2]
                hdr_part = part[:part.find(b'\r\n\r\n')].decode('utf-8', errors='ignore')
                import re as _re
                fn_match = _re.search(r'filename="([^"]+)"', hdr_part)
                if fn_match:
                    fname = fn_match.group(1)
            elif b'name="filename"' in part:
                idx = part.find(b'\r\n\r\n')
                if idx >= 0:
                    fname = part[idx + 4:].strip().decode('utf-8', errors='ignore').strip().rstrip('\r\n')

        if not file_data:
            self._send_json(400, {'error': '未找到上传文件'})
            return

        import tempfile
        suffix = '.xlsx' if fname.lower().endswith('.xlsx') else '.docx'
        tmp = tempfile.NamedTemporaryFile(suffix=suffix, delete=False)
        tmp.write(file_data)
        tmp.close()

        try:
            fname_clean = fname.strip().rstrip('\r\n')
            if fname_clean.lower().endswith('.xlsx'):
                result = _parse_xlsx_file(tmp.name)
            else:
                result = _parse_docx_file(tmp.name)
            # 同时做分析
            write_log(f'  解析完成: L4={len(result["l4"])}, L5={len(result["l5"])}')
            if result['l5']:
                write_log(f'  L5样本: en={result["l5"][0].get("en","")!r}, cn={result["l5"][0].get("cn","")!r}')
                valid_count = sum(1 for a in result['l5'] if a.get('en') and a.get('cn'))
                write_log(f'  有效L5(en+cn非空): {valid_count}')
            analysis = _analyze_l5_issues(result['l5'])
            format_issues = _check_format_issues(result['l4'], result['l5'])
            write_log(f'  分析结果: mcn={len(analysis["mcn"])}, men={len(analysis["men"])}, changes={len(analysis["changes"])}, format={len(format_issues)}')
            result['mcn'] = analysis['mcn']
            result['men'] = analysis['men']
            result['changes'] = analysis['changes']
            result['format_issues'] = format_issues
            # 保存历史（含完整结果）
            try:
                conn = get_db()
                conn.execute("INSERT INTO asset_history(filename,l4_count,l5_count,issue_count,change_count,result_json) VALUES(?,?,?,?,?,?)",
                             (fname_clean, len(result['l4']), len(result['l5']),
                              len(analysis['mcn'])+len(analysis['men'])+len(format_issues), len(analysis['changes']),
                              json.dumps(result, ensure_ascii=False)))
                conn.commit()
                conn.close()
            except: pass
            self._log_op('解析文档', f'{fname_clean} → L4:{len(result["l4"])}个, L5:{len(result["l5"])}条, 命名问题:{len(analysis["mcn"])+len(analysis["men"])}个, 格式问题:{len(format_issues)}个')
            self._send_json(200, result)
        except Exception as e:
            self._send_json(500, {'error': str(e)})
        finally:
            os.unlink(tmp.name)

    # ========== 数据资产整改：分析规范问题 ==========
    def _handle_asset_analyze(self, data):
        l5 = data.get('l5', [])
        result = _analyze_l5_issues(l5)
        # 更新最近一条历史记录
        try:
            conn = get_db()
            conn.execute("UPDATE asset_history SET issue_count=?, change_count=? WHERE id=(SELECT MAX(id) FROM asset_history)",
                         (len(result['mcn']) + len(result['men']), len(result['changes'])))
            conn.commit()
            conn.close()
        except: pass
        self._log_op('分析L5规范问题', f'一词多义:{len(result["mcn"])}个, 一义多词:{len(result["men"])}个, 修改:{len(result["changes"])}处')
        self._send_json(200, result)

    # ========== 数据资产整改：导出 Excel ==========
    def _handle_asset_export(self, data):
        try:
            import xlsxwriter
        except ImportError:
            self._send_json(500, {'error': '缺少 xlsxwriter，请运行: pip install xlsxwriter'})
            return

        excel_bytes = _export_asset_excel(data)
        self.send_response(200)
        self._cors()
        self.send_header('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        self.send_header('Content-Disposition', 'attachment; filename="asset_report.xlsx"')
        self.send_header('Content-Length', str(len(excel_bytes)))
        self.end_headers()
        self.wfile.write(excel_bytes)
        self._log_op('导出整改版Excel', f'{len(excel_bytes)} bytes')


# ========== 资产整改核心函数（类外部） ==========
import re
import io
from collections import defaultdict

def _parse_docx_file(filepath):
    """快速解析 docx（用 zipfile+xml，比 python-docx 快 5-10 倍）"""
    import zipfile
    import xml.etree.ElementTree as ET

    WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    z = zipfile.ZipFile(filepath)

    # 读取 document.xml
    doc_xml = z.read('word/document.xml')
    root = ET.fromstring(doc_xml)
    body = root.find(f'{{{WNS}}}body')

    def get_para_text(p_el):
        return ''.join(p_el.itertext()).strip()

    def get_table_data(tbl_el):
        rows = []
        for tr in tbl_el.findall(f'{{{WNS}}}tr'):
            cells = []
            for tc in tr.findall(f'{{{WNS}}}tc'):
                cells.append(''.join(tc.itertext()).strip())
            rows.append(cells)
        return rows

    # 按顺序遍历 body 子元素
    elements = []
    for child in body:
        tag = child.tag.split('}')[1] if '}' in child.tag else child.tag
        if tag == 'p':
            text = get_para_text(child)
            if text:
                elements.append(('P', text))
        elif tag == 'tbl':
            rows = get_table_data(child)
            elements.append(('T', rows))

    z.close()

    # 解析逻辑（与之前相同）
    l4_all, l5_all = [], []
    cur_module, state, cur_list_tbls, cur_def_idx = '未知模块', None, [], 0

    for i, el in enumerate(elements):
        if el[0] == 'T' and state:
            rows = el[1]
            if len(rows) < 2: continue
            hdr = rows[0]
            if state == 'list' and any('英文' in h for h in hdr):
                for row in rows[1:]:
                    if len(row) >= 3 and row[1]:
                        l4_all.append({'mod': cur_module, 'en': row[1], 'cn': row[2],
                                       'pk': row[4] if len(row) > 4 else '',
                                       'fk': row[5] if len(row) > 5 else ''})
                        cur_list_tbls.append((row[1], row[2]))
                state = None
            elif state == 'def' and any('字段' in h or '英文' in h for h in hdr):
                te = cur_list_tbls[cur_def_idx][0] if cur_def_idx < len(cur_list_tbls) else f'unknown_{cur_def_idx}'
                tc = cur_list_tbls[cur_def_idx][1] if cur_def_idx < len(cur_list_tbls) else ''
                cur_def_idx += 1
                for row in rows[1:]:
                    if len(row) >= 3 and row[1]:
                        l5_all.append({'mod': cur_module, 'tbl_en': te, 'tbl_cn': tc,
                                       'en': row[1], 'cn': row[2],
                                       'pkfk': row[3] if len(row) > 3 else '',
                                       'tp': row[5] if len(row) > 5 else '',
                                       'len': row[6] if len(row) > 6 else '',
                                       'null': row[7] if len(row) > 7 else ''})
            continue

        if el[0] != 'P': continue
        text = el[1]
        if re.match(r'^1\.\d+', text) and any(c.isdigit() for c in text[-3:]):
            continue
        if text == 'DB清单':
            state = 'list'
        elif text == 'DB定义':
            state = 'def'; cur_def_idx = 0
        elif text in ('ER图', '数据字典'):
            state = None
        elif text == '数据库设计':
            state = None
        elif state == 'def':
            is_next = any(elements[j][0] == 'P' and elements[j][1] in ('ER图', 'DB清单')
                          for j in range(i + 1, min(i + 3, len(elements))))
            if is_next:
                cur_module = text.split('（')[0].split('(')[0].strip()
                cur_list_tbls = []; state = None
        elif not any(k in text for k in ['DB清单', 'DB定义', 'ER图', '数据字典', '目录', '修订',
                                          '编制', '审核', '批准', '烟台', '江苏', '需求', '第一卷',
                                          '数据库设计', '项 目']):
            if len(text) > 2 and not text[0].isdigit():
                cur_module = text.split('（')[0].split('(')[0].strip()
                cur_list_tbls = []; state = None

    return {'l4': l4_all, 'l5': l5_all, 'modules': list(set(a['mod'] for a in l5_all))}


def _parse_xlsx_file(filepath):
    """解析 xlsx 文件（如华新细部设计-品质.xlsx 的 02 L1-L5数据资产 sheet）"""
    import zipfile
    import xml.etree.ElementTree as ET

    z = zipfile.ZipFile(filepath)
    ns_s = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    ns_r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

    # 读取 shared strings
    shared = []
    try:
        ss_raw = z.read('xl/sharedStrings.xml')
        ss_root = ET.fromstring(ss_raw)
        for si in ss_root.findall(f'.//{{{ns_s}}}si'):
            shared.append(''.join(t.text or '' for t in si.findall(f'.//{{{ns_s}}}t')))
    except: pass

    # 找到包含 L5 数据的 sheet（优先找 "L1-L5" 或 "数据资产"）
    wb_raw = z.read('xl/workbook.xml')
    wb_root = ET.fromstring(wb_raw)
    rels_raw = z.read('xl/_rels/workbook.xml.rels')
    rels_root = ET.fromstring(rels_raw)

    target_rid = None
    for s in wb_root.findall(f'.//{{{ns_s}}}sheet'):
        name = s.attrib.get('name', '')
        if 'L5' in name or 'L1-L5' in name or '数据资产' in name:
            target_rid = s.attrib.get(f'{{{ns_r}}}id')
            break
    if not target_rid:
        # 用第一个非目录 sheet
        sheets = wb_root.findall(f'.//{{{ns_s}}}sheet')
        if len(sheets) > 1:
            target_rid = sheets[1].attrib.get(f'{{{ns_r}}}id')
        elif sheets:
            target_rid = sheets[0].attrib.get(f'{{{ns_r}}}id')

    sheet_file = None
    if target_rid:
        for rel in rels_root:
            if rel.attrib.get('Id') == target_rid:
                sheet_file = 'xl/' + rel.attrib['Target']
                break

    if not sheet_file:
        z.close()
        return {'l4': [], 'l5': [], 'modules': []}

    sheet_raw = z.read(sheet_file)
    sheet_root = ET.fromstring(sheet_raw)
    z.close()

    # 解析单元格
    def col_idx(col_str):
        r = 0
        for c in col_str:
            r = r * 26 + (ord(c) - ord('A') + 1)
        return r - 1

    rows_data = {}
    for row in sheet_root.findall(f'.//{{{ns_s}}}sheetData/{{{ns_s}}}row'):
        rn = int(row.attrib['r'])
        for cell in row.findall(f'{{{ns_s}}}c'):
            ref = cell.attrib.get('r', '')
            m = re.match(r'([A-Z]+)(\d+)', ref)
            if not m: continue
            ci = col_idx(m.group(1))
            v_el = cell.find(f'{{{ns_s}}}v')
            val = v_el.text if v_el is not None else ''
            if cell.attrib.get('t') == 's' and val:
                idx = int(val)
                val = shared[idx] if idx < len(shared) else val
            if rn not in rows_data:
                rows_data[rn] = {}
            rows_data[rn][ci] = (val or '').strip()

    # 提取 L4 和 L5
    # 尝试自动检测列：找表头行中包含 "属性-英文" 或 "逻辑实体" 的列
    l4_all, l5_all = [], []
    l4_en_col, l4_cn_col, l5_en_col, l5_cn_col = None, None, None, None
    mod_col, type_col, len_col = None, None, None

    if 1 in rows_data:
        hdr = rows_data[1]
        for ci, val in hdr.items():
            vl = val.lower().replace(' ', '')
            if '逻辑实体' in val and '英文' in val: l4_en_col = ci
            elif '逻辑实体' in val and '中文' in val: l4_cn_col = ci
            elif '属性' in val and '英文' in val and 'L5' in val.upper(): l5_en_col = ci
            elif '属性' in val and '中文' in val and 'L5' in val.upper(): l5_cn_col = ci
            elif val in ('L1', 'L2') or '模块' in val: mod_col = ci
            elif '数据类型' in val or val == '类型': type_col = ci
            elif '字段长度' in val or val == '长度': len_col = ci

    # 如果没找到，用默认列号（华新格式：4=L4英文, 5=L4中文, 9=L5英文, 10=L5中文）
    if l4_en_col is None: l4_en_col = 4
    if l4_cn_col is None: l4_cn_col = 5
    if l5_en_col is None: l5_en_col = 9
    if l5_cn_col is None: l5_cn_col = 10

    seen_l4 = set()
    for rn in sorted(rows_data.keys()):
        if rn <= 1: continue
        row = rows_data[rn]
        l4e = row.get(l4_en_col, '')
        l4c = row.get(l4_cn_col, '')
        l5e = row.get(l5_en_col, '')
        l5c = row.get(l5_cn_col, '')
        mod = row.get(mod_col, '') if mod_col is not None else row.get(1, '')

        if l4e and l4e not in seen_l4:
            seen_l4.add(l4e)
            l4_all.append({'mod': mod, 'en': l4e, 'cn': l4c, 'pk': '', 'fk': ''})

        if l5e or l5c:
            l5_all.append({
                'mod': mod, 'tbl_en': l4e, 'tbl_cn': l4c,
                'en': l5e, 'cn': l5c, 'pkfk': '',
                'tp': row.get(type_col, '') if type_col is not None else row.get(13, ''),
                'len': row.get(len_col, '') if len_col is not None else row.get(15, ''),
                'null': ''
            })

    return {'l4': l4_all, 'l5': l5_all, 'modules': list(set(a['mod'] for a in l5_all))}


def _analyze_l5_issues(l5):
    """分析一词多义 / 一义多词"""
    en2cn = defaultdict(lambda: defaultdict(int))
    cn2en = defaultdict(lambda: defaultdict(int))
    for a in l5:
        if a.get('en') and a.get('cn'):
            en2cn[a['en']][a['cn']] += 1
            cn2en[a['cn']][a['en']] += 1
    mcn = {e: dict(c) for e, c in en2cn.items() if len(c) > 1}
    men = {c: dict(e) for c, e in cn2en.items() if len(e) > 1}
    cf = {}
    for en, cns in mcn.items():
        best = max(cns, key=cns.get)
        for cn in cns:
            if cn != best: cf[(en, cn)] = best
    ef = {}
    for cn, ens in men.items():
        best = max(ens, key=ens.get)
        for en in ens:
            if en != best: ef[(en, cn)] = best
    changes = []
    for i, a in enumerate(l5):
        oe, oc = a.get('en', ''), a.get('cn', '')
        nc = cf.get((oe, oc), oc)
        ne = ef.get((oe, nc), ef.get((oe, oc), oe))
        if ne != oe or nc != oc:
            changes.append({'i': i, 'oe': oe, 'ne': ne, 'oc': oc, 'nc': nc,
                            'mod': a.get('mod', ''), 'tbl': a.get('tbl_cn', '')})
    return {'mcn': mcn, 'men': men, 'changes': changes}


def _check_format_issues(l4, l5):
    """校验数据库设计文档的格式规范性"""
    issues = []

    # 1. 表英文名重复
    tbl_en_seen = {}
    for e in l4:
        en = e.get('en', '').strip()
        if not en: continue
        en_lower = en.lower()
        if en_lower in tbl_en_seen:
            issues.append({'type': '表名重复', 'severity': 'error',
                           'msg': f'表英文名 [{en}] 重复出现，模块: {e.get("mod","")} 和 {tbl_en_seen[en_lower]}'})
        else:
            tbl_en_seen[en_lower] = e.get('mod', '')

    # 2. 表中文名重复
    tbl_cn_seen = {}
    for e in l4:
        cn = e.get('cn', '').strip()
        if not cn: continue
        if cn in tbl_cn_seen:
            issues.append({'type': '表名重复', 'severity': 'error',
                           'msg': f'表中文名 [{cn}] 重复出现，模块: {e.get("mod","")} 和 {tbl_cn_seen[cn]}'})
        else:
            tbl_cn_seen[cn] = e.get('mod', '')

    # 3. 空表名
    for e in l4:
        if not e.get('en', '').strip():
            issues.append({'type': '空表名', 'severity': 'error',
                           'msg': f'表英文名为空，中文名: {e.get("cn","")}, 模块: {e.get("mod","")}'})
        if not e.get('cn', '').strip():
            issues.append({'type': '空表名', 'severity': 'error',
                           'msg': f'表中文名为空，英文名: {e.get("en","")}, 模块: {e.get("mod","")}'})

    # 4. 表名命名规范（英文名应为小写+下划线）
    import re as _re
    for e in l4:
        en = e.get('en', '').strip()
        if en and not _re.match(r'^[a-zA-Z][a-zA-Z0-9_]*$', en):
            issues.append({'type': '表名不规范', 'severity': 'warning',
                           'msg': f'表英文名 [{en}] 含非法字符（应为字母+数字+下划线）'})

    # 5. 同表内字段英文名重复
    tbl_fields = defaultdict(list)
    for a in l5:
        tbl = a.get('tbl_en', '')
        en = a.get('en', '').strip()
        if tbl and en:
            tbl_fields[(tbl, en.lower())].append(a)
    for (tbl, en_lower), items in tbl_fields.items():
        if len(items) > 1:
            issues.append({'type': '字段名重复', 'severity': 'error',
                           'msg': f'表 [{tbl}] 中字段 [{items[0]["en"]}] 重复 {len(items)} 次'})

    # 6. 空字段名
    for a in l5:
        if not a.get('en', '').strip():
            issues.append({'type': '空字段名', 'severity': 'error',
                           'msg': f'字段英文名为空，中文名: {a.get("cn","")}, 表: {a.get("tbl_cn","")}'})
        if not a.get('cn', '').strip():
            en = a.get('en', '')
            if en.lower() not in ('id', 'uuid'):  # id 类字段允许无中文名
                issues.append({'type': '空字段名', 'severity': 'warning',
                               'msg': f'字段中文名为空，英文名: {en}, 表: {a.get("tbl_cn","")}'})

    # 7. VARCHAR/CHAR 类型缺少长度
    for a in l5:
        tp = (a.get('tp', '') or '').upper().strip()
        length = (a.get('len', '') or '').strip()
        if tp in ('VARCHAR', 'CHAR', 'NVARCHAR', 'NCHAR'):
            if not length or length == '0':
                issues.append({'type': '类型缺长度', 'severity': 'error',
                               'msg': f'字段 [{a.get("en","")}]({a.get("cn","")}) 类型为 {tp} 但未指定长度，表: {a.get("tbl_cn","")}'})

    # 8. 字段名命名规范
    for a in l5:
        en = a.get('en', '').strip()
        if en and not _re.match(r'^[a-zA-Z][a-zA-Z0-9_]*$', en):
            issues.append({'type': '字段名不规范', 'severity': 'warning',
                           'msg': f'字段 [{en}] 含非法字符，表: {a.get("tbl_cn","")}'})

    # 9. 数据类型不规范（常见拼写错误或非标准类型）
    valid_types = {'VARCHAR', 'CHAR', 'NVARCHAR', 'NCHAR', 'INT', 'INTEGER', 'BIGINT', 'SMALLINT', 'TINYINT',
                   'DECIMAL', 'NUMERIC', 'FLOAT', 'DOUBLE', 'REAL', 'DATE', 'DATETIME', 'TIMESTAMP', 'TIME',
                   'TEXT', 'LONGTEXT', 'MEDIUMTEXT', 'BLOB', 'LONGBLOB', 'BOOLEAN', 'BIT', 'CLOB', 'NUMBER'}
    for a in l5:
        tp = (a.get('tp', '') or '').upper().strip()
        if tp and tp not in valid_types:
            # 去掉括号内容再检查
            base_tp = _re.sub(r'\(.*\)', '', tp).strip()
            if base_tp and base_tp not in valid_types:
                issues.append({'type': '类型不规范', 'severity': 'warning',
                               'msg': f'字段 [{a.get("en","")}] 类型 [{tp}] 不在标准类型列表中，表: {a.get("tbl_cn","")}'})

    # 10. 主键缺失检查（每张表至少应有一个主键字段）
    tbl_has_pk = set()
    for a in l5:
        pkfk = (a.get('pkfk', '') or '').upper().strip()
        if 'PK' in pkfk or '主' in pkfk or 'Y' == pkfk:
            tbl_has_pk.add(a.get('tbl_en', ''))
    for e in l4:
        en = e.get('en', '').strip()
        if en and en not in tbl_has_pk:
            # 也检查 l4 里的 pk 字段
            if not e.get('pk', '').strip():
                issues.append({'type': '主键缺失', 'severity': 'warning',
                               'msg': f'表 [{en}]({e.get("cn","")}) 未定义主键字段'})

    return issues


def _export_asset_excel(data):
    """生成整改版 Excel 字节流"""
    import xlsxwriter
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf)
    H = wb.add_format({'bold': 1, 'bg_color': '#4472C4', 'font_color': '#FFF', 'border': 1, 'text_wrap': 1, 'valign': 'vcenter', 'align': 'center', 'font_size': 11})
    C = wb.add_format({'border': 1, 'text_wrap': 1, 'valign': 'vcenter', 'font_size': 10})
    Y = wb.add_format({'border': 1, 'text_wrap': 1, 'valign': 'vcenter', 'font_size': 10, 'bg_color': '#FFFF00', 'font_color': '#CC0000', 'bold': 1})
    N = wb.add_format({'border': 1, 'text_wrap': 1, 'valign': 'vcenter', 'font_size': 9, 'bg_color': '#FFF2CC', 'font_color': '#806000'})

    l4, l5 = data.get('l4', []), data.get('l5', [])
    mcn, men, changes = data.get('mcn', {}), data.get('men', {}), data.get('changes', [])
    chg_map = {ch['i']: ch for ch in changes}
    for i, a in enumerate(l5):
        ch = chg_map.get(i)
        a['en2'] = ch['ne'] if ch else a.get('en', '')
        a['cn2'] = ch['nc'] if ch else a.get('cn', '')

    # S1: L4
    s1 = wb.add_worksheet('L4 逻辑实体清单')
    for c, h in enumerate(['序号', 'L1', 'L2(模块)', '逻辑实体-英文(L4)', '逻辑实体-中文(L4)', '主键', '关联表-外键']):
        s1.write(0, c, h, H)
    for i, e in enumerate(l4):
        s1.write(i+1, 0, i+1, C); s1.write(i+1, 1, '品质', C); s1.write(i+1, 2, e.get('mod', ''), C)
        s1.write(i+1, 3, e.get('en', ''), C); s1.write(i+1, 4, e.get('cn', ''), C)
        s1.write(i+1, 5, e.get('pk', ''), C); s1.write(i+1, 6, e.get('fk', ''), C)
    for c, w in enumerate([6, 6, 18, 38, 24, 15, 30]): s1.set_column(c, c, w)
    s1.autofilter(0, 0, len(l4), 6); s1.freeze_panes(1, 0)

    # S2: L5 整改版
    s2 = wb.add_worksheet('L5 逻辑属性(整改版)')
    for c, h in enumerate(['序号', 'L1', 'L2(模块)', '逻辑实体-英文(L4)', '逻辑实体-中文(L4)',
                           '属性-英文(L5)', '属性-中文(L5)', '源字段(英文)', '源字段(中文)',
                           '数据类型', '字段长度', '空否', '是否修改', '修改说明']):
        s2.write(0, c, h, H)
    for i, a in enumerate(l5):
        r = i + 1; ch = chg_map.get(i)
        ec = ch and ch['oe'] != ch['ne']; cc = ch and ch['oc'] != ch['nc']
        s2.write(r, 0, i+1, C); s2.write(r, 1, '品质', C); s2.write(r, 2, a.get('mod', ''), C)
        s2.write(r, 3, a.get('tbl_en', ''), C); s2.write(r, 4, a.get('tbl_cn', ''), C)
        s2.write(r, 5, a['en2'], Y if ec else C); s2.write(r, 6, a['cn2'], Y if cc else C)
        s2.write(r, 7, a.get('en', ''), C); s2.write(r, 8, a.get('cn', ''), C)
        s2.write(r, 9, a.get('tp', ''), C); s2.write(r, 10, a.get('len', ''), C); s2.write(r, 11, a.get('null', ''), C)
        if ch:
            p, d = [], []
            if ec: p.append('英文'); d.append(f"{ch['oe']}→{ch['ne']}")
            if cc: p.append('中文'); d.append(f"{ch['oc']}→{ch['nc']}")
            s2.write(r, 12, '是(' + '+'.join(p) + ')', Y); s2.write(r, 13, '; '.join(d), N)
        else:
            s2.write(r, 12, '', C); s2.write(r, 13, '', C)
    for c, w in enumerate([6, 6, 16, 35, 20, 24, 18, 24, 18, 10, 8, 6, 12, 40]): s2.set_column(c, c, w)
    s2.autofilter(0, 0, len(l5), 13); s2.freeze_panes(1, 0)

    # S3: 修改明细
    s3 = wb.add_worksheet('修改明细清单')
    for c, h in enumerate(['序号', '模块', '表名', '问题类型', '原英文', '→新英文', '原中文', '→新中文']):
        s3.write(0, c, h, H)
    for i, ch in enumerate(changes):
        p = []
        if ch['oc'] != ch['nc']: p.append('一词多义')
        if ch['oe'] != ch['ne']: p.append('一义多词')
        s3.write(i+1, 0, i+1, C); s3.write(i+1, 1, ch.get('mod', ''), C); s3.write(i+1, 2, ch.get('tbl', ''), C)
        s3.write(i+1, 3, '+'.join(p), N); s3.write(i+1, 4, ch['oe'], C)
        s3.write(i+1, 5, ch['ne'], Y if ch['oe'] != ch['ne'] else C)
        s3.write(i+1, 6, ch['oc'], C); s3.write(i+1, 7, ch['nc'], Y if ch['oc'] != ch['nc'] else C)
    for c, w in enumerate([6, 16, 22, 12, 28, 28, 20, 20]): s3.set_column(c, c, w)
    s3.autofilter(0, 0, len(changes), 7); s3.freeze_panes(1, 0)

    # S4: 整改规则
    s4 = wb.add_worksheet('整改规则汇总')
    for c, h in enumerate(['序号', '问题类型', '属性名', '值列表', '统一为', '影响数']):
        s4.write(0, c, h, H)
    ri = 0
    for en, cns in sorted(mcn.items()):
        best = max(cns, key=cns.get); ri += 1
        s4.write(ri, 0, ri, C); s4.write(ri, 1, '一词多义', N); s4.write(ri, 2, en, C)
        s4.write(ri, 3, ' / '.join(cns.keys()), C); s4.write(ri, 4, best, Y)
        s4.write(ri, 5, sum(v for k, v in cns.items() if k != best), C)
    for cn, ens in sorted(men.items()):
        best = max(ens, key=ens.get); ri += 1
        s4.write(ri, 0, ri, C); s4.write(ri, 1, '一义多词', N); s4.write(ri, 2, cn, C)
        s4.write(ri, 3, ' / '.join(ens.keys()), C); s4.write(ri, 4, best, Y)
        s4.write(ri, 5, sum(v for k, v in ens.items() if k != best), C)
    for c, w in enumerate([6, 10, 28, 42, 28, 10]): s4.set_column(c, c, w)
    s4.autofilter(0, 0, ri, 5); s4.freeze_panes(1, 0)

    wb.close()
    return buf.getvalue()


if __name__ == '__main__':
    print('=' * 40)
    print('  词库词根管理系统 - 后端服务 v3')
    print('  含数据资产整改功能')
    print('=' * 40)
    print(f'脚本目录: {SCRIPT_DIR}')
    if USE_PG:
        print(f'数据库: PostgreSQL (Railway)')
    else:
        print(f'数据库: SQLite ({DB_FILE})')
    print(f'日志文件: {LOG_FILE}')
    get_db()
    print(f'数据库已就绪')
    write_log('服务启动')
    server = http.server.HTTPServer(('0.0.0.0', PORT), Handler)
    print(f'服务已启动: http://localhost:{PORT}')
    print('按 Ctrl+C 停止服务')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\n服务已停止')
        server.server_close()
