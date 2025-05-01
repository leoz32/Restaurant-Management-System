from flask import Flask, render_template, session, redirect, url_for, request, flash, jsonify, g
from flask_caching import Cache
import hashlib
import json
from datetime import datetime, timedelta
from datetime import date 
import os
import logging
import pandas as pd
import numpy as np
import random
import pyodbc
import sys
import argparse
import threading
import time
import functools
from queue import Queue, Empty
from decimal import Decimal           # 新增，保证金额精度

# 设置命令行参数解析
parser = argparse.ArgumentParser(description='Flask应用服务器')
parser.add_argument('--port', type=int, default=5000, help='指定服务器端口')
args = parser.parse_args()

# 配置日志
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# ===== 全局配置 =====
# 缓存配置
CACHE_CONFIG = {
    "CACHE_TYPE": "SimpleCache",
    "CACHE_DEFAULT_TIMEOUT": 3600,  # 默认1小时
    "CACHE_THRESHOLD": 1000  # 最大缓存项数
}

# 缓存超时配置
CACHE_TIMEOUTS = {
    'short': 300,     # 5分钟
    'medium': 1800,   # 30分钟
    'long': 3600,     # 1小时
    'day': 86400      # 24小时
}

# ===== 数据库连接池 =====
class AccessConnectionPool:
    """Microsoft Access数据库连接池"""
    
    def __init__(self, db_path, max_connections=5, connection_timeout=30, max_retries=3, retry_delay=1):
        self.db_path = db_path
        self.max_connections = max_connections
        self.connection_timeout = connection_timeout
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        
        self.connections = Queue(maxsize=max_connections)
        self.active_connections = 0
        self.lock = threading.Lock()
        
        # 查找驱动并设置连接字符串
        self.drivers = [x for x in pyodbc.drivers() if 'Access' in x]
        if not self.drivers:
            logger.error("未找到Access驱动程序")
            raise Exception("未找到Access驱动程序")
            
        self.driver = self.drivers[0]
        self.conn_str = f'DRIVER={{{self.driver}}};DBQ={self.db_path};'
        logger.info(f"连接池初始化完成，使用驱动: {self.driver}")
        logger.info(f"数据库路径: {self.db_path}")

    def get_connection(self):
        """从连接池获取连接或创建新连接"""
        for attempt in range(self.max_retries):
            try:
                # 尝试从池中获取连接
                try:
                    connection = self.connections.get(block=False)
                    # 测试连接是否有效
                    try:
                        connection.cursor().execute("SELECT 1")
                        return connection
                    except:
                        # 连接已失效，关闭并减少计数
                        try:
                            connection.close()
                        except:
                            pass
                        with self.lock:
                            self.active_connections -= 1
                        raise Empty()
                except Empty:
                    # 没有可用连接，在限制内创建新连接
                    with self.lock:
                        if self.active_connections < self.max_connections:
                            connection = pyodbc.connect(self.conn_str, timeout=self.connection_timeout)
                            self.active_connections += 1
                            return connection
                        else:
                            # 等待连接可用
                            time.sleep(self.retry_delay)
            except Exception as e:
                logger.warning(f"连接尝试失败: {str(e)}，重试中...")
                time.sleep(self.retry_delay * (2 ** attempt))  # 指数退避
        
        raise Exception(f"尝试 {self.max_retries} 次后仍无法建立数据库连接")
    
    def release_connection(self, connection):
        """将连接归还到连接池"""
        try:
            # 检查连接是否有效
            try:
                connection.cursor().execute("SELECT 1")
                self.connections.put(connection)
            except:
                # 关闭失效连接
                try:
                    connection.close()
                except:
                    pass
                with self.lock:
                    self.active_connections -= 1
        except Exception as e:
            logger.warning(f"归还连接时出错: {str(e)}")
            with self.lock:
                self.active_connections -= 1
    
    def close_all(self):
        """关闭所有连接"""
        logger.info("关闭所有数据库连接")
        try:
            while not self.connections.empty():
                conn = self.connections.get(block=False)
                try:
                    conn.close()
                except:
                    pass
            with self.lock:
                self.active_connections = 0
        except Exception as e:
            logger.error(f"关闭连接时出错: {str(e)}")

# ===== Flask应用初始化 =====
if getattr(sys, 'frozen', False):
    # PyInstaller打包环境
    template_folder = os.path.join(sys._MEIPASS, 'templates')
    static_folder = os.path.join(sys._MEIPASS, 'static')
    app = Flask(__name__, 
                template_folder=template_folder,
                static_folder=static_folder)
else:
    app = Flask(__name__)

app.secret_key = 'margaret-restaurant-mis-secret-key'
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config.from_mapping(CACHE_CONFIG)

# 初始化缓存
cache = Cache(app)

# 全局变量
conn_pool = None
DB_PATH = None

# ===== 装饰器和工具函数 =====
# 检查是否已登录的装饰器
def require_login(view_func):
    @functools.wraps(view_func)
    def decorated_view(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        return view_func(*args, **kwargs)
    return decorated_view

# 缓存装饰器
def cache_query(timeout=CACHE_TIMEOUTS['medium']):
    """缓存函数结果的装饰器"""
    def decorator(func):
        @functools.wraps(func)
        def wrapper(*args, **kwargs):
            # 创建唯一缓存键
            key_parts = [func.__name__]
            key_parts.extend([str(arg) for arg in args])
            key_parts.extend([f"{k}={v}" for k, v in sorted(kwargs.items())])
            key = hashlib.md5(":".join(key_parts).encode()).hexdigest()
            
            # 尝试从缓存获取
            result = cache.get(key)
            if result is not None:
                logger.debug(f"函数缓存命中: {func.__name__}")
                return result
                
            # 执行函数并缓存结果
            result = func(*args, **kwargs)
            cache.set(key, result, timeout=timeout)
            return result
        return wrapper
    return decorator

# 缓存失效函数
def invalidate_cache(pattern=None):
    """使缓存条目失效"""
    if hasattr(cache, 'clear'):
        if pattern:
            # 如果缓存实现支持，选择性清除匹配模式的缓存
            if hasattr(cache, '_cache'):
                keys_to_delete = [k for k in cache._cache.keys() if pattern in k]
                for key in keys_to_delete:
                    cache.delete(key)
            else:
                logger.warning("此缓存类型不支持选择性失效")
                cache.clear()
        else:
            # 清除整个缓存
            cache.clear()
            logger.info("缓存已清除")

# ===== 数据库路径和连接管理 =====
def get_db_path():
    """获取数据库路径 - 修改为直接读取同目录"""
    logger.info(f"当前执行路径: {os.getcwd()}")
    logger.info(f"可执行文件路径: {sys.executable if hasattr(sys, 'executable') else '非可执行环境'}")
    
    if getattr(sys, 'frozen', False):
        # PyInstaller打包环境 - 可执行文件所在目录
        base_path = os.path.dirname(sys.executable)
    else:
        # 开发环境 - 当前脚本文件所在目录
        base_path = os.path.dirname(os.path.abspath(__file__))
    
    # 直接在同目录查找数据库文件
    db_path = os.path.join(base_path, 'Margaretds.accdb')
    logger.info(f"数据库路径: {db_path}")
    
    return db_path

def initialize_connection_pool():
    """初始化数据库连接池"""
    global conn_pool, DB_PATH
    DB_PATH = get_db_path()
    if os.path.exists(DB_PATH):
        conn_pool = AccessConnectionPool(DB_PATH)
        logger.info("连接池初始化完成")
    else:
        logger.error(f"找不到数据库文件: {DB_PATH}")
        raise FileNotFoundError(f"找不到数据库文件: {DB_PATH}")

# ===== 数据库操作函数 =====
def execute_query_with_pool(query, params=None, fetch=True, cache_timeout=None):
    """使用连接池执行查询，并可选择缓存结果"""
    global conn_pool
    
    # 初始化连接池（如果需要）
    if conn_pool is None:
        initialize_connection_pool()
    
    # 处理缓存
    cache_key = None
    if cache_timeout is not None and fetch:
        key_base = query + str(params if params else "")
        cache_key = hashlib.md5(key_base.encode()).hexdigest()
        
        cached_result = cache.get(cache_key)
        if cached_result is not None:
            logger.debug(f"缓存命中: {query[:50]}...")
            return cached_result
    
    # 执行查询
    conn = None
    try:
        conn = conn_pool.get_connection()
        cursor = conn.cursor()
        
        if params:
            cursor.execute(query, params)
        else:
            cursor.execute(query)
        
        result = None
        if fetch:
            result = cursor.fetchall()
            if cache_key is not None:
                cache.set(cache_key, result, timeout=cache_timeout)
        else:
            conn.commit()
            result = True
            
        cursor.close()
        return result
        
    except Exception as e:
        if conn:
            try:
                conn.rollback()
            except:
                pass
        error_message = f"查询执行错误: {str(e)}\n查询: {query}\n参数: {params}"
        logger.error(error_message)
        raise Exception(error_message)
        
    finally:
        if conn:
            conn_pool.release_connection(conn)

def execute_query(query, params=None, fetch=True, cache_timeout=CACHE_TIMEOUTS['medium']):
    """执行查询并自动缓存只读操作"""
    # 只读查询使用缓存
    if fetch and query.lower().startswith(('select', 'with')):
        return execute_query_with_pool(query, params, fetch, cache_timeout)
    # 写操作不缓存
    return execute_query_with_pool(query, params, fetch)

def execute_insert(table, data):
    """Insert data and invalidate related cache"""
    # Convert the keys to a list and join them
    columns = ', '.join([f"[{col}]" for col in data.keys()])
    placeholders = ', '.join(['?'] * len(data))
    
    # Format the query with proper escaping for Access
    query = f"INSERT INTO [{table}] ({columns}) VALUES ({placeholders})"
    
    # Create parameter tuple in the same order as columns
    params = tuple(data.values())
    
    try:
        # Log the query and parameters for debugging
        logger.debug(f"Executing query: {query}")
        logger.debug(f"With parameters: {params}")
        
        execute_query_with_pool(query, params, fetch=False)
        
        # Invalidate related cache
        invalidate_cache(table.lower())
        logger.info(f"Data successfully inserted into {table}")
        return True
    except Exception as e:
        logger.error(f"Error inserting data into {table}: {str(e)}")
        logger.error(f"Query: {query}")
        logger.error(f"Parameters: {params}")
        raise

def execute_update(table, data, condition_col, condition_val):
    """更新数据并使相关缓存失效"""
    set_clause = ', '.join([f"{k} = ?" for k in data.keys()])
    query = f"UPDATE [{table}] SET {set_clause} WHERE {condition_col} = ?"
    
    params = list(data.values())
    params.append(condition_val)
    
    try:
        execute_query_with_pool(query, params, fetch=False)
        # 使相关缓存失效
        invalidate_cache(table.lower())
        logger.info(f"已更新 {table} 中 {condition_col}={condition_val} 的数据")
        return True
    except Exception as e:
        logger.error(f"更新 {table} 中的数据时出错: {str(e)}")
        raise

def execute_delete(table, condition_col, condition_val):
    """删除数据并使相关缓存失效"""
    query = f"DELETE FROM [{table}] WHERE {condition_col} = ?"
    
    try:
        execute_query_with_pool(query, (condition_val,), fetch=False)
        # 使相关缓存失效
        invalidate_cache(table.lower())
        logger.info(f"已从 {table} 中删除 {condition_col}={condition_val} 的数据")
        return True
    except Exception as e:
        logger.error(f"从 {table} 删除数据时出错: {str(e)}")
        raise

def get_table_column_names(table):
    """获取表的列名"""
    try:
        cursor = execute_query(f"SELECT TOP 1 * FROM [{table}]")
        columns = [column[0] for column in cursor.description]
        return columns
    except Exception as e:
        error_message = f"获取 {table} 的列名时出错: {str(e)}"
        logger.error(error_message)
        raise Exception(error_message)

# ===== 预测函数 =====
def create_forecast(historical_data, days=7, seasonal=True):
    """创建简化版预测"""
    # 转换为列表
    if not isinstance(historical_data, list):
        historical_data = list(historical_data)
    
    # 计算统计数据
    mean_value = sum(historical_data) / len(historical_data) if historical_data else 0
    variance = sum((x - mean_value) ** 2 for x in historical_data) / len(historical_data) if historical_data else 0
    std_value = variance ** 0.5
    
    if std_value < 0.1 * mean_value:
        std_value = 0.1 * mean_value
    
    # 生成预测
    forecast = []
    trend_factor = 0.01
    
    for i in range(days):
        # 基础值
        base = mean_value * (1 + random.uniform(-0.2, 0.2) * (std_value / mean_value if mean_value else 0.1))
        
        # 趋势成分
        trend = mean_value * trend_factor * i
        
        # 季节性成分
        seasonal_factor = 0
        if seasonal:
            weekday = (datetime.now().weekday() + i) % 7
            if weekday >= 5:  # 周末
                seasonal_factor = mean_value * 0.1
            elif weekday == 0:  # 周一
                seasonal_factor = -mean_value * 0.05
        
        value = max(base + trend + seasonal_factor, 0)
        
        # 根据数量级取整
        if mean_value > 100:
            value = round(value)
        elif mean_value > 10:
            value = round(value, 1)
        else:
            value = round(value, 2)
        
        forecast.append(value)
    
    return forecast

# ===== 自定义过滤器 =====
@app.template_filter('safe_slice')
def safe_slice(value, start=0, end=None, length=None):
    """防止除零错误的切片过滤器"""
    if not value or len(value) == 0:
        return []
    
    if end is None:
        end = len(value)
    
    if length is not None:
        end = min(start + length, len(value))
        
    return value[start:end]

# ===== 常用的缓存查询 =====
@cache_query(timeout=CACHE_TIMEOUTS['long'])
def get_all_customers():
    """获取所有客户（缓存结果）"""
    return execute_query(
        """
        SELECT c.CustomerID, c.CustomerFirstName, c.CustomerLastName, c.CustomerGender, 
               c.CustomerPhone, c.CustomerEmail, c.CustomerType, c.CustomerBirthday, 
               c.CustomerCard
        FROM Customer AS c
        ORDER BY c.CustomerLastName, c.CustomerFirstName
        """
    )

@cache_query(timeout=CACHE_TIMEOUTS['long'])
def get_all_menu_items():
    """获取所有菜单项（缓存结果）"""
    return execute_query(
        """
        SELECT mi.MenuItemID, mi.MenuItemName, mi.Price, mi.Description, mi.State, mi.Category
        FROM [MenuItem] AS mi
        ORDER BY mi.Category, mi.MenuItemName
        """
    )

@cache_query(timeout=CACHE_TIMEOUTS['long'])
def get_all_inventory():
    """获取所有库存项（缓存结果）"""
    return execute_query(
        """
        SELECT i.*, s.SupplierName, itm.ItemName
        FROM ((Inventory AS i
        LEFT JOIN Supplier AS s ON i.InventorySupplierID = s.SupplierID)
        LEFT JOIN Item AS itm ON i.ItemID = itm.ItemID)
        ORDER BY i.InventoryName
        """
    )

@cache_query(timeout=CACHE_TIMEOUTS['medium'])
def get_total_customers():
    """获取客户总数"""
    return execute_query("SELECT COUNT(*) AS Count FROM Customer")[0].Count

@cache_query(timeout=CACHE_TIMEOUTS['short'])
def get_total_orders():
    """获取订单总数"""
    return execute_query("SELECT COUNT(*) AS Count FROM [Orders]")[0].Count

@cache_query(timeout=CACHE_TIMEOUTS['medium'])
def get_total_inventory():
    """获取库存总数"""
    return execute_query("SELECT COUNT(*) AS Count FROM Inventory")[0].Count

# ===== 替代before_first_request的初始化 =====
def init_app(app):
    """在应用启动时初始化资源"""
    with app.app_context():
        global conn_pool
        if conn_pool is None:
            initialize_connection_pool()

# 确保在第一次请求时初始化数据库连接
@app.before_request
def before_request():
    """每个请求前检查连接池是否初始化"""
    global conn_pool
    if conn_pool is None:
        initialize_connection_pool()

# 关闭应用时清理资源
@app.teardown_appcontext
def close_connection_pool(exception=None):
    """应用上下文结束时关闭所有连接"""
    global conn_pool
    if conn_pool:
        conn_pool.close_all()

# ===== 用户认证路由 =====
@app.route('/')
def index():
    if 'user_id' in session:
        return redirect(url_for('dashboard'))
    return redirect(url_for('login'))

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        try:
            query = "SELECT * FROM Staff WHERE StaffEmail = ?"
            staff = execute_query(query, (username,), cache_timeout=None)  # 不缓存登录查询
            
            if not staff or password != "admin":
                flash("用户名或密码无效")
            else:
                staff = staff[0]
                session.clear()
                session['user_id'] = staff.StaffID
                session['username'] = f"{staff.StaffFirstName} {staff.StaffLastName}"
                session['position'] = staff.StaffPosition
                
                return redirect(url_for('dashboard'))
        except Exception as e:
            logger.error(f"登录期间出错: {str(e)}")
            flash(f"登录期间出错: {str(e)}")
    
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))

# ===== 仪表板路由 =====
@app.route('/dashboard')
@require_login
def dashboard():
    try:
        # 获取关键指标 - 使用缓存
        total_customers = get_total_customers()
        total_orders = get_total_orders()
        total_inventory = get_total_inventory()
        
        # 今日订单 - 不缓存
        today = datetime.now().strftime('%Y-%m-%d')
        today_orders = execute_query(
            "SELECT COUNT(*) AS Count FROM [Orders] WHERE Format(OrderDate, 'yyyy-mm-dd') = ?", 
            (today,), cache_timeout=None
        )[0].Count
        
        # 最近订单 - 短期缓存
        recent_orders = execute_query(
        """
        SELECT TOP 5 o.OrderID, o.OrderDate, c.CustomerFirstName & ' ' & c.CustomerLastName AS CustomerName, 
               m.MenuItemName, o.OrderPayment
        FROM ([Orders] AS o 
        LEFT JOIN Customer AS c ON o.CustomerID = c.CustomerID)
        LEFT JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID
        ORDER BY o.OrderDate DESC
        """, cache_timeout=CACHE_TIMEOUTS['short']
        )
        
        # 库存警告 - 短期缓存
        low_inventory = execute_query(
        """
        SELECT i.InventoryID, i.InventoryName, i.InventoryQuantity, itm.MinimumQuantity
        FROM Inventory AS i 
        INNER JOIN Item AS itm ON i.ItemID = itm.ItemID
        WHERE i.InventoryQuantity <= itm.MinimumQuantity
        """, cache_timeout=CACHE_TIMEOUTS['short']
        )
        
        return render_template(
            'dashboard.html',
            total_customers=total_customers,
            total_orders=total_orders,
            total_inventory=total_inventory,
            today_orders=today_orders,
            recent_orders=recent_orders,
            low_inventory=low_inventory
        )
    except Exception as e:
        logger.error(f"加载仪表板时出错: {str(e)}")
        flash(f"加载仪表板时出错，请联系管理员")
        return render_template('error.html', error=str(e))

# ===== 客户管理路由 =====
@app.route('/customers')
@require_login
def customer_list():
    try:
        customers = execute_query(
            """
            SELECT c.CustomerID, c.CustomerFirstName, c.CustomerLastName, c.CustomerGender, 
                   c.CustomerPhone, c.CustomerEmail, c.CustomerType, c.CustomerBirthday, 
                   c.CustomerCard, COUNT(o.OrderID) AS OrderCount
            FROM Customer AS c
            LEFT JOIN [Orders] AS o ON c.CustomerID = o.CustomerID
            GROUP BY c.CustomerID, c.CustomerFirstName, c.CustomerLastName, c.CustomerGender, 
                     c.CustomerPhone, c.CustomerEmail, c.CustomerType, c.CustomerBirthday, c.CustomerCard
            ORDER BY c.CustomerLastName, c.CustomerFirstName
            """
        )
        return render_template('customers/list.html', customers=customers)
    except Exception as e:
        logger.error(f"获取客户列表时出错: {str(e)}")
        flash(f"获取客户列表时出错")
        return render_template('customers/list.html', customers=[])

@app.route('/customers/<customer_id>')
@require_login
def view_customer(customer_id):
    try:
        # 获取客户信息
        customer = execute_query(
            "SELECT * FROM Customer WHERE CustomerID = ?", 
            (customer_id,)
        )
        
        if not customer:
            flash("客户不存在")
            return redirect(url_for('customer_list'))
        
        customer = customer[0]
        
        # 获取客户评论
        reviews = execute_query(
            "SELECT * FROM [Customer Review] WHERE CustomerID = ? ORDER BY ReviewDate DESC", 
            (customer_id,)
        )
        
        # 获取客户订单
        orders = execute_query(
            """
            SELECT o.*, m.MenuItemName
            FROM [Orders] AS o
            INNER JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID
            WHERE o.CustomerID = ?
            ORDER BY o.OrderDate DESC
            """, 
            (customer_id,)
        )
        
        return render_template(
            'customers/view.html', 
            customer=customer, 
            reviews=reviews, 
            orders=orders
        )
    except Exception as e:
        logger.error(f"查看客户信息时出错: {str(e)}")
        flash(f"查看客户信息时出错")
        return redirect(url_for('customer_list'))


@app.route('/customers/delete/<customer_id>', methods=['POST'])
@require_login
def delete_customer(customer_id):
    try:
        execute_delete('Customer', 'CustomerID', customer_id)
        flash("客户已成功删除")
        invalidate_cache('customer')  # 确保使相关缓存失效
        return redirect(url_for('customer_list'))
    except Exception as e:
        logger.error(f"删除客户时出错: {str(e)}")
        flash(f"删除客户时出错: {str(e)}")
        return redirect(url_for('customer_list'))

@app.route('/customers/edit/<customer_id>', methods=['GET', 'POST'])
@require_login
def edit_customer(customer_id):
    # 获取客户数据以进行编辑 (移到前面，以便在POST错误时也能使用)
    customer_result = execute_query("SELECT * FROM Customer WHERE CustomerID = ?", (customer_id,), cache_timeout=None)
    if not customer_result:
        flash("客户不存在")
        return redirect(url_for('customer_list'))
    customer = customer_result[0]

    try:
        if request.method == 'POST':
            # 记录收到的表单数据，用于调试
            logger.debug(f"Received form data for edit customer {customer_id}: {request.form}")

            # 从表单获取数据
            birthday_str = request.form.get('birthday') # 使用 .get()
            birthday_obj = None
            if birthday_str:
                try:
                    # 将 YYYY-MM-DD 格式的字符串转换为日期对象
                    birthday_obj = datetime.strptime(birthday_str, '%Y-%m-%d').date()
                except ValueError:
                    flash("无效的生日日期格式，请使用 YYYY-MM-DD 格式。")
                    # 格式错误时，重新渲染编辑页面并显示错误
                    return render_template('customers/edit.html', customer=customer)

            # 假设 first_name, last_name, gender, type 是必需的
            # 对可选字段使用 .get() 并提供默认值
            customer_data = {
                'CustomerFirstName': request.form['first_name'],
                'CustomerLastName': request.form['last_name'],
                'CustomerGender': request.form['gender'],
                'CustomerPhone': request.form.get('phone', ''), # 可选
                'CustomerEmail': request.form.get('email', ''), # 可选
                'CustomerType': request.form['type'],
                'CustomerBirthday': birthday_obj, # 已处理 None 的情况
                'CustomerCard': request.form.get('card', '') # 可选
            }

            # 更新数据库
            execute_update('Customer', customer_data, 'CustomerID', customer_id)
            # 使客户列表和详情缓存失效
            invalidate_cache('customer') # 使用更通用的模式
            invalidate_cache(f"customerid={customer_id}") # 特定客户缓存 (如果存在)
            flash("客户信息已更新")
            return redirect(url_for('view_customer', customer_id=customer_id))

        # GET 请求时显示编辑表单
        return render_template('customers/edit.html', customer=customer)
    except KeyError as e:
        # 捕获特定错误：如果必需字段缺失
        missing_field = str(e).strip("'")
        logger.error(f"编辑客户 {customer_id} 时必需字段 '{missing_field}' 缺失。收到的表单: {request.form}")
        flash(f"错误：必需的字段 '{missing_field}' 未填写。")
        return render_template('customers/edit.html', customer=customer) # 重新显示表单
    except Exception as e:
        logger.error(f"编辑客户信息时出错: {str(e)}")
        flash(f"编辑客户信息时出错: {str(e)}") # 显示更详细的错误信息
        # 尝试留在编辑页面而不是重定向回列表
        return render_template('customers/edit.html', customer=customer) # 传递客户数据


@app.route('/customers/new', methods=['GET', 'POST'])
@require_login
def new_customer():
    try:
        if request.method == 'POST':
            # 从表单获取数据
            birthday_str = request.form['birthday']
            birthday_obj = None
            if birthday_str:
                try:
                    # 将 YYYY-MM-DD 格式的字符串转换为日期对象
                    birthday_obj = datetime.strptime(birthday_str, '%Y-%m-%d').date()
                except ValueError:
                    flash("无效的生日日期格式，请使用 YYYY-MM-DD 格式。")
                    # 格式错误时，重新渲染新建页面并显示错误
                    return render_template('customers/edit.html', customer=None)
            
            # 首先获取当前最大的CustomerID
            query = "SELECT MAX(CustomerID) FROM Customer"
            result = execute_query_with_pool(query, (), fetch=True)
            
            # 处理结果并生成新的CustomerID
            max_id = "C000"  # 默认起始ID
            if result and len(result) > 0 and result[0][0]:
                current_max = result[0][0]  # 第一行第一列是MAX(CustomerID)的值
                # 假设格式为 C001, C002 等
                current_num = int(current_max[1:])  # 提取数字部分
                next_num = current_num + 1  # 增加1
                max_id = f"C{next_num:03d}"  # 格式化为 C001 格式
            
            # 处理电话号码，确保格式正确
            phone = request.form['phone']
            if isinstance(phone, str) and phone.endswith('.0'):
                phone = phone[:-2]  # 移除 '.0'
            
            customer_data = {
                'CustomerID': max_id,  # 使用生成的ID
                'CustomerFirstName': request.form['first_name'],
                'CustomerLastName': request.form['last_name'],
                'CustomerGender': request.form['gender'],
                'CustomerPhone': phone,
                'CustomerEmail': request.form['email'],
                'CustomerType': request.form['customer_type'],
                'CustomerBirthday': birthday_obj,
                'CustomerCard': request.form['card']
            }
            
            # 插入数据库
            execute_insert('Customer', customer_data)
            flash(f"新客户已添加，ID: {max_id}")
            return redirect(url_for('customer_list'))
            
        # GET 请求时显示新建表单
        return render_template('customers/edit.html', customer=None)
        
    except Exception as e:
        logger.error(f"添加新客户时出错: {str(e)}")
        flash(f"添加新客户时出错: {str(e)}")
        # 留在新建页面
        return render_template('customers/edit.html', customer=None)

# ===== 菜单管理路由 =====
@app.route('/menu')
@require_login
def menu_list():
    try:
        menu_items = execute_query(
            """
            SELECT mi.MenuItemID, mi.MenuItemName, mi.Price, mi.Description, mi.State, mi.Category, 
                   (SELECT COUNT(*) FROM [Orders] WHERE MenuItemID = mi.MenuItemID) AS OrderCount
            FROM [MenuItem] AS mi
            ORDER BY mi.Category, mi.MenuItemName
            """
        )
        return render_template('menu/list.html', menu_items=menu_items)
    except Exception as e:
        logger.error(f"获取菜单列表时出错: {str(e)}")
        flash(f"获取菜单列表时出错")
        # 创建虚拟记录
        class DummyMenuItem:
            def __init__(self):
                self.MenuItemID = ""
                self.MenuItemName = "加载出错"
                self.Price = 0
                self.Description = ""
                self.State = ""
                self.Category = ""
                self.OrderCount = 0
        
        dummy_items = [DummyMenuItem() for _ in range(5)]
        return render_template('menu/list.html', menu_items=dummy_items)

@app.route('/menu/<menu_item_id>')
@require_login
def view_menu_item(menu_item_id):
    try:
        # 获取菜单项信息
        menu_item = execute_query(
            "SELECT * FROM [MenuItem] WHERE MenuItemID = ?", 
            (menu_item_id,)
        )
        
        if not menu_item:
            flash("菜单项不存在")
            return redirect(url_for('menu_list'))
        
        menu_item = menu_item[0]
        
        # 获取菜单项配方
        recipe_items = execute_query(
            """
            SELECT r.*, i.InventoryName, i.InventoryType
            FROM Recipe AS r
            INNER JOIN Inventory AS i ON r.InventoryID = i.InventoryID
            WHERE r.MenuItemID = ?
            """, 
            (menu_item_id,)
        )
        
        return render_template(
            'menu/view.html', 
            menu_item=menu_item, 
            recipe_items=recipe_items
        )
    except Exception as e:
        logger.error(f"查看菜单项时出错: {str(e)}")
        flash(f"查看菜单项时出错")
        return redirect(url_for('menu_list'))

@app.route('/menu/delete/<menu_item_id>', methods=['POST'])
@require_login
def delete_menu_item(menu_item_id):
    try:
        execute_delete('MenuItem', 'MenuItemID', menu_item_id)
        flash("菜单项已成功删除")
        invalidate_cache('menuitem')
        return redirect(url_for('menu_list'))
    except Exception as e:
        logger.error(f"删除菜单项时出错: {str(e)}")
        flash(f"删除菜单项时出错: {str(e)}")
        return redirect(url_for('menu_list'))
    
@app.route('/menu/edit/<menu_item_id>', methods=['GET', 'POST'])
@require_login
def edit_menu_item(menu_item_id):
    try:
        if request.method == 'POST':
            # 从表单获取数据
            menu_data = {
                'MenuItemName': request.form['name'],
                'Price': request.form['price'],
                'Description': request.form['description'],
                'State': request.form['state'],
                'Category': request.form['category']
            }
            
            # 更新数据库
            execute_update('MenuItem', menu_data, 'MenuItemID', menu_item_id)
            flash("菜单项已更新")
            return redirect(url_for('view_menu_item', menu_item_id=menu_item_id))
        
        # 获取菜单项数据以进行编辑
        menu_item = execute_query("SELECT * FROM [MenuItem] WHERE MenuItemID = ?", (menu_item_id,), cache_timeout=None)
        if not menu_item:
            flash("菜单项不存在")
            return redirect(url_for('menu_list'))
        
        return render_template('menu/edit.html', menu_item=menu_item[0])
    except Exception as e:
        logger.error(f"编辑菜单项时出错: {str(e)}")
        flash(f"编辑菜单项时出错")
        return redirect(url_for('menu_list'))

@app.route('/menu/new', methods=['GET', 'POST'])
@require_login
def new_menu_item():
    """
    新增菜单项：
      • 自动生成 MenuItemID（格式 M01 / M02 / …）
      • Price 用 Decimal 写入 Access Currency/Number 字段
    """
    try:
        if request.method == 'POST':
            # ── 1. 生成新的 MenuItemID ───────────────────────────
            id_sql   = "SELECT MAX(MenuItemID) FROM [MenuItem]"
            id_row   = execute_query_with_pool(id_sql, (), fetch=True)
            if id_row and id_row[0][0] and id_row[0][0].startswith('M'):
                next_num      = int(id_row[0][0][1:]) + 1     # 去掉 'M' 取数字
                menu_item_id  = f"M{next_num:02d}"            # 始终两位数；如需三位改成 :03d
            else:
                menu_item_id  = "M01"

            # ── 2. 读取表单并类型转换 ──────────────────────────
            name        = request.form.get('name', '').strip()
            price_raw   = request.form.get('price', '0').strip()
            description = request.form.get('description', '').strip() or None
            state       = request.form.get('state', 'Active')
            category    = request.form.get('category', '')

            try:
                price = Decimal(price_raw)
            except Exception:
                flash('价格必须是数字！', 'danger')
                return redirect(request.url)

            # ── 3. 组装数据并插入 ─────────────────────────────
            menu_data = {
                'MenuItemID'  : menu_item_id,   # 新生成的主键
                'MenuItemName': name,
                'Price'       : price,          # Decimal
                'Description' : description,
                'State'       : state,
                'Category'    : category
            }

            execute_insert('MenuItem', menu_data)
            flash(f'菜单项 {menu_item_id} 已添加', 'success')
            return redirect(url_for('menu_list'))

        # GET：渲染空白表单
        return render_template('menu/edit.html', menu_item=None)

    except Exception as e:
        logger.exception('添加新菜单项时出错')
        flash(f'添加新菜单项时出错：{e}', 'danger')
        return redirect(url_for('menu_list'))


# ===== 订单管理路由 =====
@app.route('/orders')
@require_login
def order_list():
    try:
        query = """
        SELECT 
            o.*, 
            c.CustomerFirstName & ' ' & c.CustomerLastName AS CustomerName,
            m.MenuItemName, 
            m.Price,
            s.StaffFirstName & ' ' & s.StaffLastName AS StaffName,
            r.ReceivableAmount, 
            r.ReceivableDiscount
        FROM 
            ((([Orders] AS o
            LEFT JOIN [Customer] AS c ON o.CustomerID = c.CustomerID)
            LEFT JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID)
            LEFT JOIN [Staff] AS s ON o.StaffID = s.StaffID)
            LEFT JOIN [Account Receivable] AS r ON o.ReceivableID = r.ReceivableID
        ORDER BY o.OrderDate DESC
        """
        orders = execute_query(query, cache_timeout=CACHE_TIMEOUTS['short'])
        
        return render_template('orders/list.html', orders=orders)
    except Exception as e:
        logger.error(f"获取订单时出错: {str(e)}")
        flash(f"获取订单时出错")
        return render_template('orders/list.html', orders=[])
    
@app.route('/orders/delete/<order_id>', methods=['POST'])
@require_login
def delete_order(order_id):
    try:
        execute_delete('Orders', 'OrderID', order_id)
        flash("订单已成功删除")
        invalidate_cache('orders')
        return redirect(url_for('order_list'))
    except Exception as e:
        logger.error(f"删除订单时出错: {str(e)}")
        flash(f"删除订单时出错: {str(e)}")
        return redirect(url_for('order_list'))
    
@app.route('/orders/<order_id>')
@require_login
def view_order(order_id):
    try:
        # 获取订单信息
        order = execute_query(
            """
            SELECT 
                o.OrderID, 
                o.OrderDate, 
                Trim([c].[CustomerFirstName] & ' ' & [c].[CustomerLastName]) AS CustomerName, 
                [m].[MenuItemName], 
                o.OrderPayment, 
                o.OrderDiscount, 
                Trim([s].[StaffFirstName] & ' ' & [s].[StaffLastName]) AS StaffName 
            FROM 
                (([Orders] AS o 
                LEFT JOIN [Customer] AS c ON [o].[CustomerID] = [c].[CustomerID]) 
                LEFT JOIN [MenuItem] AS m ON [o].[MenuItemID] = [m].[MenuItemID]) 
                LEFT JOIN [Staff] AS s ON [o].[StaffID] = [s].[StaffID]
            WHERE o.OrderID = ?
            """, 
            (order_id,)
        )
        
        if not order:
            flash("订单不存在")
            return redirect(url_for('order_list'))
        
        return render_template('orders/view.html', order=order[0])
    except Exception as e:
        logger.error(f"查看订单时出错: {str(e)}")
        flash(f"查看订单时出错")
        return redirect(url_for('order_list'))

@app.route('/orders/new', methods=['GET', 'POST'])
@require_login
def new_order():
    """
    创建新订单：
      • OrderID          Long Integer（手动递增）
      • OrderDate        Date（yyyy‑m‑d）
      • OrderPayment     Currency / Number
      • OrderDiscount    Currency / Number
    其余为文本字段
    """
    try:
        if request.method == 'POST':
            # ── 1. 生成新的 OrderID ────────────────────────────────
            max_id_sql = "SELECT MAX(OrderID) FROM [Orders]"
            max_id_row = execute_query_with_pool(max_id_sql, (), fetch=True)
            order_id = (int(max_id_row[0][0]) + 1) if max_id_row and max_id_row[0][0] else 1001

            # ── 2. 读取并转换表单数据 ────────────────────────────
            customer_id   = request.form.get('customer_id')
            menu_item_id  = request.form.get('menu_item_id')
            note          = request.form.get('note', '').strip() or None  # 空值用 None

            # 金额：收到字符串 ➜ Decimal；失败默认 0
            def to_decimal(raw):
                try:
                    return Decimal(raw)
                except Exception:
                    return Decimal('0')

            payment_amount = to_decimal(request.form.get('payment',   '0'))
            amount         = to_decimal(request.form.get('amount',    '0'))
            discount       = to_decimal(request.form.get('discount',  '0'))

            # ── 3. 生成新的 ReceivableID ──────────────────────────
            rec_max_sql  = "SELECT MAX(ReceivableID) FROM [Account Receivable]"
            rec_max_row  = execute_query_with_pool(rec_max_sql, (), fetch=True)
            if rec_max_row and rec_max_row[0][0] and rec_max_row[0][0].startswith('RCV'):
                next_num      = int(rec_max_row[0][0][3:]) + 1
                receivable_id = f"RCV{next_num}"
            else:
                receivable_id = "RCV1001"

            # ── 4. 写入 Account Receivable ───────────────────────
            receivable_data = {
                'ReceivableID'      : receivable_id,
                'ReceivableType'    : 'Sales Revenue',
                'AccountID'         : 'ACC-SALES',
                'ReceivableName'    : customer_id,
                'ReceivableDate'    : date.today(),   # Date 对象
                'ReceivableAmount'  : amount,
                'ReceivableDiscount': discount,
                'Note'              : str(order_id)
            }
            execute_insert('Account Receivable', receivable_data)

            # ── 5. 写入 Orders ──────────────────────────────────
            order_data = {
                'OrderID'      : order_id,
                'OrderDate'    : date.today(),        # Date 对象
                'CustomerID'   : customer_id,
                'OrderPayment' : payment_amount,      # Decimal ➜ 数字字段
                'OrderDiscount': discount,            # Decimal ➜ 数字字段
                'MenuItemID'   : menu_item_id,
                'ReceivableID' : receivable_id,
                'StaffID'      : session['user_id'],
                'Note'         : note                 # None ➜ Null
            }
            execute_insert('Orders', order_data)

            flash(f"Order added, Order ID: {order_id}")
            return redirect(url_for('order_list'))

        # ── 6. GET：加载下拉选项 ────────────────────────────────
        customers = execute_query(
            "SELECT CustomerID, CustomerFirstName, CustomerLastName "
            "FROM Customer ORDER BY CustomerLastName"
        )
        menu_items = execute_query(
            "SELECT MenuItemID, MenuItemName, Price "
            "FROM [MenuItem] ORDER BY Category, MenuItemName"
        )
        return render_template('orders/create.html',
                               customers=customers,
                               menu_items=menu_items)

    except Exception as e:
        logger.error(f"An error occurred while creating a new order: {e}")
        flash(f"An error occurred while creating a new order: {e}")

        # 回退到表单页面，仍然加载下拉列表
        try:
            customers = execute_query(
                "SELECT CustomerID, CustomerFirstName, CustomerLastName "
                "FROM Customer ORDER BY CustomerLastName"
            )
            menu_items = execute_query(
                "SELECT MenuItemID, MenuItemName, Price "
                "FROM [MenuItem] ORDER BY Category, MenuItemName"
            )
            return render_template('orders/create.html',
                                   customers=customers,
                                   menu_items=menu_items)
        except Exception:
            return redirect(url_for('order_list'))

def next_payable_id():
    """Return next PayableID in the form PAY001, PAY002 …"""
    sql  = "SELECT MAX(PayableID) FROM [Account Payable]"
    row  = execute_query_with_pool(sql, (), fetch=True)
    if row and row[0][0] and row[0][0].startswith('PAY'):
        next_num = int(row[0][0][3:]) + 1
        return f"PAY{next_num:03d}"
    return "PAY001"
# ===== 库存管理路由 =====
@app.route('/inventory')
@require_login
def inventory_list():
    try:
        inventory_items = execute_query(
            """
            SELECT i.*, s.SupplierName, itm.ItemName
            FROM ((Inventory AS i
            LEFT JOIN Supplier AS s ON i.InventorySupplierID = s.SupplierID)
            LEFT JOIN Item AS itm ON i.ItemID = itm.ItemID)
            ORDER BY i.InventoryName
            """
        )
        return render_template('inventory/list.html', inventory_items=inventory_items)
    except Exception as e:
        logger.error(f"获取库存列表时出错: {str(e)}")
        flash(f"获取库存列表时出错")
        return render_template('inventory/list.html', inventory_items=[])
    
@app.route('/inventory/delete/<inventory_id>', methods=['POST'])
@require_login
def delete_inventory(inventory_id):
    try:  
        execute_delete('Inventory', 'InventoryID', inventory_id)
        flash("库存项已成功删除")
        invalidate_cache('inventory')
        return redirect(url_for('inventory_list'))
    except Exception as e:
        logger.error(f"删除库存项时出错: {str(e)}")
        flash(f"删除库存项时出错: {str(e)}")
        return redirect(url_for('inventory_list'))

@app.route('/inventory/new', methods=['GET', 'POST'])
@require_login
def new_inventory():
    try:
        if request.method == 'POST':
            # ---------- 1. Build inventory basic data -------------------------
            # (same code you already have – trimmed here for brevity)
            name        = request.form['name'].strip()
            inv_type    = request.form['type'].strip()
            amount      = Decimal(request.form.get('amount', '0'))
            quantity    = int(request.form.get('quantity', '0'))
            description = request.form.get('description', '').strip() or None
            supplier_id = request.form.get('supplier_id') or None
            item_id     = request.form.get('item_id') or None

            def dt(field):
                raw = request.form.get(field, '')
                return datetime.strptime(raw, '%Y-%m-%d').date() if raw else None

            order_date  = dt('order_date')
            due_date    = dt('due_date')

            # ---------- 2. Create Payable record ------------------------------
            payable_id = next_payable_id()

            # Lookup supplier name just for display (optional)
            supplier_name = None
            if supplier_id:
                row = execute_query(
                    "SELECT SupplierName FROM Supplier WHERE SupplierID=?", (supplier_id,)
                )
                supplier_name = row[0][0] if row else None

            payable_data = {
                'PayableID'      : payable_id,
                'Invoice'        : None,                       # fill later if you wish
                'PayableName'    : supplier_name or name,      # who you owe
                'PayableType'    : 'Inventory Purchase',
                'PayableDueDate' : due_date,
                'PayableDatePaid': None,
                'PayableAmount'  : amount,
                'AccountID'      : 'ACC-PURCHASE',             # or your actual account code
                'Note'           : f'Auto‑created for {name}'
            }
            execute_insert('Account Payable', payable_data)

            # ---------- 3. Create Inventory record ----------------------------
            inventory_id = next_inventory_id()   # your existing INV generator

            inv_data = {
                'InventoryID'        : inventory_id,
                'InventoryName'      : name,
                'InventoryType'      : inv_type,
                'InventoryAmount'    : amount,
                'InventoryQuantity'  : quantity,
                'Description'        : description,
                'InventorySupplierID': supplier_id,
                'ItemID'             : item_id,
                'OrderDate'          : order_date,
                'DueDate'            : due_date,
                'InventoryPayableID' : payable_id              # FK to payable
            }
            execute_insert('Inventory', inv_data)

            flash(f'Inventory {inventory_id} and payable {payable_id} added.', 'success')
            return redirect(url_for('inventory_list'))

        # ----- GET → render blank form ---------------------------------------
        suppliers = execute_query(
            "SELECT SupplierID, SupplierName FROM Supplier ORDER BY SupplierName"
        )
        items = execute_query(
            "SELECT ItemID, ItemName FROM Item ORDER BY ItemName"
        )
        return render_template('inventory/edit.html',
                               inventory=None, suppliers=suppliers, items=items)

    except Exception as e:
        logger.exception('Error while adding inventory')
        flash(f'Error while adding inventory: {e}', 'danger')
        return redirect(url_for('inventory_list'))

def next_inventory_id():
    """
    Return the next InventoryID in the format INV001, INV002, …
    Looks at the current maximum value in the Inventory table.
    """
    sql = "SELECT MAX(InventoryID) FROM [Inventory]"
    row = execute_query_with_pool(sql, (), fetch=True)

    # row -> [(u'INV057',)]  or  [(None,)]
    if row and row[0][0] and str(row[0][0]).startswith('INV'):
        next_num = int(row[0][0][3:]) + 1          # strip 'INV', convert to int, +1
        return f"INV{next_num:03d}"                 # always 3 digits
    return "INV001"                                 # first record in an empty table

@app.route('/inventory/<inventory_id>')
@require_login
def view_inventory(inventory_id):
    try:
        # 获取库存项信息
        inventory_item = execute_query(
            """
            SELECT  i.*,
                    s.SupplierName,
                    s.ContactFirstName & ' ' & s.ContactLastName  AS ContactName,
                    s.ContactEmail,
                    itm.ItemName,
                    itm.MinimumQuantity,
                    p.PayableAmount,
                    p.PayableDueDate
            FROM  ((Inventory        AS i
                LEFT JOIN Supplier AS s   ON i.InventorySupplierID = s.SupplierID)
                LEFT JOIN Item     AS itm ON i.ItemID             = itm.ItemID)
                LEFT JOIN [Account Payable] AS p
                        ON i.InventoryPayableID = p.PayableID
            WHERE i.InventoryID = ?
            """,
            (inventory_id,)          # ← single‑element tuple ‑ don’t forget the comma!
        )
        
        if not inventory_item:
            flash("库存项不存在")
            return redirect(url_for('inventory_list'))
        
        # 获取相关配方
        recipes = execute_query(
            """
            SELECT r.*, m.MenuItemName
            FROM Recipe AS r
            INNER JOIN [MenuItem] AS m ON r.MenuItemID = m.MenuItemID
            WHERE r.InventoryID = ?
            """, 
            (inventory_id,)
        )
        
        return render_template(
            'inventory/view.html', 
            inventory=inventory_item[0],
            recipes=recipes
        )
    except Exception as e:
        logger.error(f"查看库存项时出错: {str(e)}")
        flash(f"查看库存项时出错")
        return redirect(url_for('inventory_list'))

@app.route('/inventory/edit/<inventory_id>', methods=['GET', 'POST'])
@require_login
def edit_inventory(inventory_id):
     # ----- NEW: treat "0" (or "new") as a blank form -----------------
    if inventory_id in ("0", "new"):
        return redirect(url_for('new_inventory'))
    try:
        if request.method == 'POST':
            # 从表单获取数据
            inventory_data = {
                'InventoryName': request.form['name'],
                'InventoryType': request.form['type'],
                'InventoryAmount': request.form['amount'],
                'InventoryQuantity': request.form['quantity'],
                'Description': request.form.get('description', '')
            }
            
            # 更新数据库
            execute_update('Inventory', inventory_data, 'InventoryID', inventory_id)
            flash("库存项已更新")
            return redirect(url_for('view_inventory', inventory_id=inventory_id))
        
        # 获取库存项数据以进行编辑
        inventory = execute_query("SELECT * FROM Inventory WHERE InventoryID = ?", (inventory_id,), cache_timeout=None)
        if not inventory:
            flash("库存项不存在")
            return redirect(url_for('inventory_list'))
        
        suppliers = execute_query("SELECT SupplierID, SupplierName FROM Supplier ORDER BY SupplierName")
        items = execute_query("SELECT ItemID, ItemName FROM Item ORDER BY ItemName")
        
        return render_template(
            'inventory/edit.html', 
            inventory=inventory[0],
            suppliers=suppliers,
            items=items
        )
    except Exception as e:
        logger.error(f"编辑库存项时出错: {str(e)}")
        flash(f"编辑库存项时出错")
        return redirect(url_for('inventory_list'))

# ===== 员工管理路由 =====
# ===== 员工管理路由 =====
@app.route('/staff')
@require_login
def staff_list():
    try:
        staff = execute_query(
            """
            SELECT s.*, 
                   (SELECT COUNT(*) FROM [Orders] AS o WHERE o.StaffID = s.StaffID) AS OrderCount,
                   (SELECT IIF(SUM(sw.StaffWorkHours) IS NULL, 0, SUM(sw.StaffWorkHours)) 
                    FROM [Staff Working] AS sw WHERE sw.StaffID = s.StaffID) AS TotalHours
            FROM Staff AS s
            ORDER BY s.StaffLastName, s.StaffFirstName
            """
        )
        return render_template('staff/list.html', staff=staff)
    except Exception as e:
        logger.error(f"获取员工列表时出错: {str(e)}")
        flash(f"获取员工列表时出错")
        return render_template('staff/list.html', staff=[])
        return render_template('staff/list.html', staff=staff)
    except Exception as e:
        logger.error(f"获取员工列表时出错: {str(e)}")
        flash(f"获取员工列表时出错")
        return render_template('staff/list.html', staff=[])

@app.route('/staff/<staff_id>')
@require_login
def view_staff(staff_id):
    try:
        # 获取员工信息
        staff_member = execute_query("SELECT * FROM Staff WHERE StaffID = ?", (staff_id,))
        
        if not staff_member:
            flash("员工不存在")
            return redirect(url_for('staff_list'))
        
        # 获取员工工作记录
        work_history = execute_query(
            """
            SELECT sw.*, p.PayableAmount, p.PayableDatePaid
            FROM [Staff Working] AS sw
            LEFT JOIN [Account Payable] AS p ON sw.PayableID = p.PayableID
            WHERE sw.StaffID = ?
            ORDER BY sw.WorkDate DESC
            """, 
            (staff_id,)
        )
        
        # 获取员工处理的订单
        orders = execute_query(
            """
            SELECT o.OrderID, o.OrderDate, 
                   c.CustomerFirstName & ' ' & c.CustomerLastName AS CustomerName,
                   m.MenuItemName, o.OrderPayment
            FROM (([Orders] AS o
            LEFT JOIN Customer AS c ON o.CustomerID = c.CustomerID)
            LEFT JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID)
            WHERE o.StaffID = ?
            ORDER BY o.OrderDate DESC
            """, 
            (staff_id,)
        )
        
        return render_template(
            'staff/view.html', 
            staff=staff_member[0],
            work_history=work_history,
            orders=orders
        )
    except Exception as e:
        logger.error(f"查看员工信息时出错: {str(e)}")
        flash(f"查看员工信息时出错")
        return redirect(url_for('staff_list'))
        
        return render_template(
            'staff/view.html', 
            staff=staff_member[0],
            work_history=work_history,
            orders=orders
        )
    except Exception as e:
        logger.error(f"查看员工信息时出错: {str(e)}")
        flash(f"查看员工信息时出错")
        return redirect(url_for('staff_list'))

@app.route('/staff/delete/<staff_id>', methods=['POST'])
@require_login
def delete_staff(staff_id):
    try:
        execute_delete('Staff', 'StaffID', staff_id)
        flash("员工已成功删除")
        invalidate_cache('staff')  # 确保使相关缓存失效
        return redirect(url_for('staff_list'))
    except Exception as e:
        logger.error(f"删除员工时出错: {str(e)}")
        flash(f"删除员工时出错: {str(e)}")
        return redirect(url_for('staff_list'))

@app.route('/staff/edit/<staff_id>', methods=['GET', 'POST'])
@require_login
def edit_staff(staff_id):
    # 获取员工数据以进行编辑 (移到前面，以便在POST错误时也能使用)
    staff_result = execute_query("SELECT * FROM Staff WHERE StaffID = ?", (staff_id,), cache_timeout=None)
    if not staff_result:
        flash("员工不存在")
        return redirect(url_for('staff_list'))
    staff = staff_result[0]

    try:
        if request.method == 'POST':
            # 记录收到的表单数据，用于调试
            logger.debug(f"Received form data for edit staff {staff_id}: {request.form}")

            # 从表单获取数据
            # 假设 first_name, last_name, position 是必需的
            # 对可选字段使用 .get() 并提供默认值
            staff_data = {
                'StaffFirstName': request.form['first_name'],
                'StaffLastName': request.form['last_name'],
                'StaffPosition': request.form['position'],
                'StaffPhone': request.form.get('phone', ''), # 可选
                'StaffEmail': request.form.get('email', ''), # 可选
                'OrderCount': request.form.get('order_count', 0), # 可选，默认为0
                'TotalHours': request.form.get('total_hours', 0.0) # 可选，默认为0.0
            }

            # 更新数据库
            execute_update('Staff', staff_data, 'StaffID', staff_id)
            # 使员工列表和详情缓存失效
            invalidate_cache('staff') # 使用更通用的模式
            invalidate_cache(f"staffid={staff_id}") # 特定员工缓存 (如果存在)
            flash("员工信息已更新")
            return redirect(url_for('view_staff', staff_id=staff_id))

        # GET 请求时显示编辑表单
        return render_template('staff/edit.html', staff=staff)
    except KeyError as e:
        # 捕获特定错误：如果必需字段缺失
        missing_field = str(e).strip("'")
        logger.error(f"编辑员工 {staff_id} 时必需字段 '{missing_field}' 缺失。收到的表单: {request.form}")
        flash(f"错误：必需的字段 '{missing_field}' 未填写。")
        return render_template('staff/edit.html', staff=staff) # 重新显示表单
    except Exception as e:
        logger.error(f"编辑员工信息时出错: {str(e)}")
        flash(f"编辑员工信息时出错: {str(e)}") # 显示更详细的错误信息
        # 尝试留在编辑页面而不是重定向回列表
        return render_template('staff/edit.html', staff=staff) # 传递员工数据

@app.route('/staff/new', methods=['GET', 'POST'])
@require_login
def new_staff():
    try:
        if request.method == 'POST':
            # 首先获取当前最大的StaffID
            query = "SELECT MAX(StaffID) FROM Staff"
            result = execute_query_with_pool(query, (), fetch=True)
            
            # 处理结果并生成新的StaffID
            max_id = "S01"  # 默认起始ID
            if result and len(result) > 0 and result[0][0]:
                current_max = result[0][0]  # 第一行第一列是MAX(StaffID)的值
                # 假设格式为 S01, S02 等
                current_num = int(current_max[1:])  # 提取数字部分
                next_num = current_num + 1  # 增加1
                max_id = f"S{next_num:02d}"  # 格式化为 S01 格式
            
            # 处理电话号码，确保格式正确
            phone = request.form['phone']
            if isinstance(phone, str) and phone.endswith('.0'):
                phone = phone[:-2]  # 移除 '.0'
            
            staff_data = {
                'StaffID': max_id,  # 使用生成的ID
                'StaffFirstName': request.form['first_name'],
                'StaffLastName': request.form['last_name'],
                'StaffPosition': request.form['position'],
                'StaffPhone': phone,
                'StaffEmail': request.form.get('email', ''),
            }
            
            # 插入数据库
            execute_insert('Staff', staff_data)
            flash(f"新员工已添加，ID: {max_id}")
            return redirect(url_for('staff_list'))
            
        # GET 请求时显示新建表单
        return render_template('staff/edit.html', staff=None)
        
    except Exception as e:
        logger.error(f"添加新员工时出错: {str(e)}")
        flash(f"添加新员工时出错: {str(e)}")
        # 留在新建页面
        return render_template('staff/edit.html', staff=None)

# ===== 报表路由 =====
@app.route('/reports')
@require_login
def report_list():
    # 可用报表列表
    reports = [
        {'id': 'sales', 'name': 'Sales statement', 'icon': 'fa-chart-bar', 'description': '查看销售趋势、热门商品和收入分析'},
        {'id': 'inventory', 'name': '库存报表', 'icon': 'fa-boxes', 'description': '监控库存水平、成本和低库存警报'},
        {'id': 'customer', 'name': '客户分析', 'icon': 'fa-users', 'description': '分析客户人口统计、订购模式和忠诚度'},
        {'id': 'staff', 'name': '员工绩效', 'icon': 'fa-user-tie', 'description': '查看员工生产力、服务指标和工作时间'},
        {'id': 'forecast', 'name': '销售预测', 'icon': 'fa-chart-line', 'description': '查看客户和收入的预测分析'},
        {'id': 'menu', 'name': '菜单表现', 'icon': 'fa-utensils', 'description': '分析菜单项受欢迎程度、盈利能力和趋势'},
        {'id': 'seasonal', 'name': '季节性分析', 'icon': 'fa-calendar-alt', 'description': '检查销售和客户行为的季节性模式'},
        {'id': 'financial', 'name': 'Financial summary', 'icon': 'fa-file-invoice-dollar', 'description': '查看全面的财务概览和关键指标'}
    ]
    return render_template('reports/report_list.html', reports=reports)

@app.route('/reports/<report_id>')
@require_login
def view_report(report_id):
    try:
        # 根据报表类型获取不同数据
        if report_id == 'sales':
            data = execute_query(
                """
                SELECT Month(o.OrderDate) AS Month, 
                       COUNT(o.OrderID) AS OrderCount,
                       IIF(SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL, 
                           0, 
                           SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) AS TotalSales
                FROM [Orders] AS o
                INNER JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID
                WHERE Year(o.OrderDate) = Year(Date())
                GROUP BY Month(o.OrderDate)
                """
            )
            
            if not data:
                return render_template('reports/report_view.html', 
                                      report_name='Sales statement',
                                      report_id='sales',
                                      error="没有可用的销售数据")
            
            category_data = execute_query(
                """
                SELECT m.Category, 
                       COUNT(o.OrderID) AS OrderCount,
                       IIF(SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL,
                           0,
                           SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) AS TotalSales
                FROM [Orders] AS o
                INNER JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID
                WHERE Year(o.OrderDate) = Year(Date())
                GROUP BY m.Category
                ORDER BY SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) DESC
                """
            )
            
            chart_data = {
                'labels': [str(row.Month) for row in data],
                'values': [float(row.TotalSales) if row.TotalSales is not None else 0 for row in data]
            }
            
            category_chart = {
                'labels': [row.Category if row.Category else "未分类" for row in category_data],
                'values': [float(row.TotalSales) if row.TotalSales is not None else 0 for row in category_data]
            }
            
            return render_template(
                'reports/report_view.html',
                report_name='Sales statement',
                report_id='sales',
                chart_data=json.dumps(chart_data),
                chart_type='bar',
                secondary_chart=json.dumps(category_chart),
                secondary_chart_type='pie',
                raw_data=data,
                secondary_data=category_data,
                secondary_title='Sell by category'
            )
            
        elif report_id == 'inventory':
            data = execute_query(
                """
                SELECT i.InventoryType, 
                       COUNT(i.InventoryID) AS ItemCount,
                       IIF(SUM(i.InventoryQuantity) IS NULL, 0, SUM(i.InventoryQuantity)) AS TotalQuantity,
                       IIF(SUM(i.InventoryAmount) IS NULL, 0, SUM(i.InventoryAmount)) AS TotalAmount
                FROM Inventory AS i
                GROUP BY i.InventoryType
                """
            )
            
            if not data:
                return render_template('reports/report_view.html', 
                                      report_name='Status reporting',
                                      report_id='inventory',
                                      error="没有可用的库存数据")
            
            low_stock = execute_query(
                """
                SELECT i.InventoryName, i.InventoryType, i.InventoryQuantity, 
                       itm.MinimumQuantity,
                       s.SupplierName, s.ContactEmail
                FROM ((Inventory AS i
                INNER JOIN Item AS itm ON i.ItemID = itm.ItemID)
                INNER JOIN Supplier AS s ON i.InventorySupplierID = s.SupplierID)
                WHERE i.InventoryQuantity <= itm.MinimumQuantity AND itm.MinimumQuantity > 0
                ORDER BY i.InventoryQuantity / itm.MinimumQuantity
                UNION
                SELECT i.InventoryName, i.InventoryType, i.InventoryQuantity, 
                       itm.MinimumQuantity,
                       s.SupplierName, s.ContactEmail
                FROM ((Inventory AS i
                INNER JOIN Item AS itm ON i.ItemID = itm.ItemID)
                INNER JOIN Supplier AS s ON i.InventorySupplierID = s.SupplierID)
                WHERE i.InventoryQuantity <= itm.MinimumQuantity AND itm.MinimumQuantity = 0
                """
            )
            
            chart_data = {
                'labels': [row.InventoryType if row.InventoryType else "未分类" for row in data],
                'values': [float(row.TotalAmount) if row.TotalAmount is not None else 0 for row in data]
            }
            
            return render_template(
                'reports/report_view.html',
                report_name='Status reporting',
                report_id='inventory',
                chart_data=json.dumps(chart_data),
                chart_type='pie',
                raw_data=data,
                secondary_data=low_stock,
                secondary_title='Low inventory warning'
            )
            
        elif report_id == 'customer':
            data = execute_query(
                """
                SELECT IIF(c.CustomerType IS NULL, '未分类', c.CustomerType) AS CustomerType, 
                       COUNT(c.CustomerID) AS CustomerCount
                FROM Customer AS c
                GROUP BY c.CustomerType
                """
            )
            
            if not data:
                return render_template('reports/report_view.html', 
                                      report_name='客户分析',
                                      report_id='customer',
                                      error="no usable data")
            
            active_customers = execute_query(
                """
                SELECT c.CustomerID, Trim(c.CustomerFirstName & ' ' & c.CustomerLastName) AS CustomerName,
                       COUNT(o.OrderID) AS OrderCount,
                       IIF(SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL,
                           0,
                           SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) AS TotalSpent,
                       MAX(o.OrderDate) AS LastOrderDate
                FROM ((Customer AS c
                INNER JOIN [Orders] AS o ON c.CustomerID = o.CustomerID)
                INNER JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID)
                GROUP BY c.CustomerID, c.CustomerFirstName, c.CustomerLastName
                HAVING COUNT(o.OrderID) > 0
                ORDER BY COUNT(o.OrderID) DESC, SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) DESC
                """
            )
            
            chart_data = {
                'labels': [row.CustomerType for row in data],
                'values': [row.CustomerCount if row.CustomerCount is not None else 0 for row in data]
            }
            
            return render_template(
                'reports/report_view.html',
                report_name='Customer analysis',
                report_id='customer',
                chart_data=json.dumps(chart_data),
                chart_type='pie',
                raw_data=data,
                secondary_data=active_customers,
                secondary_title='Active customers'
            )

        elif report_id == 'staff':
            data = execute_query(
                """
                SELECT s.StaffID, Trim(s.StaffFirstName & ' ' & s.StaffLastName) AS StaffName,
                       s.StaffPosition,
                       COUNT(o.OrderID) AS OrderCount,
                       IIF(SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL,
                           0,
                           SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) AS TotalSales
                FROM ((Staff AS s
                LEFT JOIN [Orders] AS o ON s.StaffID = o.StaffID)
                LEFT JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID)
                GROUP BY s.StaffID, s.StaffFirstName, s.StaffLastName, s.StaffPosition
                ORDER BY IIF(SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL,
                           0,
                           SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) DESC
                """
            )
            
            if not data:
                return render_template('reports/report_view.html', 
                                      report_name='Employee Performance report',
                                      report_id='staff',
                                      error="没有可用的员工数据")
            
            work_data = execute_query(
                """
                SELECT Trim(s.StaffFirstName & ' ' & s.StaffLastName) AS StaffName,
                       IIF(SUM(sw.StaffWorkHours) IS NULL, 0, SUM(sw.StaffWorkHours)) AS TotalHours,
                       IIF(SUM(sw.StaffWage * sw.StaffWorkHours) IS NULL, 0, SUM(sw.StaffWage * sw.StaffWorkHours)) AS TotalWage
                FROM Staff AS s
                LEFT JOIN [Staff Working] AS sw ON s.StaffID = sw.StaffID
                WHERE Year(sw.WorkDate) = Year(Date()) OR sw.WorkDate IS NULL
                GROUP BY s.StaffFirstName, s.StaffLastName
                ORDER BY IIF(SUM(sw.StaffWorkHours) IS NULL, 0, SUM(sw.StaffWorkHours)) DESC
                """
            )
            
            chart_data = {
                'labels': [row.StaffName for row in data],
                'values': [float(row.TotalSales) if row.TotalSales is not None else 0 for row in data]
            }
            
            secondary_chart = {
                'labels': [row.StaffName for row in work_data],
                'values': [float(row.TotalHours) if row.TotalHours is not None else 0 for row in work_data]
            }
            
            return render_template(
                'reports/report_view.html',
                report_name='Employee Performance report',
                report_id='staff',
                chart_data=json.dumps(chart_data),
                chart_type='bar',
                secondary_chart=json.dumps(secondary_chart),
                secondary_chart_type='bar',
                raw_data=data,
                secondary_data=work_data,
                secondary_title='Employee working hours'
            )
            
        elif report_id == 'menu':
            try:
                data = execute_query(
                    """
                    SELECT m.MenuItemID, m.MenuItemName, m.Category, m.Price,
                           COUNT(o.OrderID) AS OrderCount,
                           IIF(SUM(IIF(m.Price IS NULL, 0, m.Price) * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL,
                               0,
                               SUM(IIF(m.Price IS NULL, 0, m.Price) * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) AS TotalSales
                    FROM [MenuItem] AS m
                    LEFT JOIN [Orders] AS o ON m.MenuItemID = o.MenuItemID
                    GROUP BY m.MenuItemID, m.MenuItemName, m.Category, m.Price
                    ORDER BY COUNT(o.OrderID) DESC
                    """
                )
                
                if not data:
                    return render_template('reports/report_view.html', 
                                          report_name='Menu Presentation report',
                                          report_id='menu',
                                          error="没有可用的菜单数据")
                
                category_data = execute_query(
                    """
                    SELECT IIF(m.Category IS NULL, '未分类', m.Category) AS Category, 
                           COUNT(o.OrderID) AS OrderCount,
                           IIF(SUM(IIF(m.Price IS NULL, 0, m.Price) * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL,
                               0,
                               SUM(IIF(m.Price IS NULL, 0, m.Price) * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) AS TotalSales,
                           IIF(AVG(IIF(m.Price IS NULL, 0, m.Price)) IS NULL, 0, AVG(IIF(m.Price IS NULL, 0, m.Price))) AS AveragePrice
                    FROM [MenuItem] AS m
                    LEFT JOIN [Orders] AS o ON m.MenuItemID = o.MenuItemID
                    GROUP BY m.Category
                    ORDER BY COUNT(o.OrderID) DESC
                    """
                )
                
                total_items = len(data) if data else 0
                
                top_items = data[:10] if total_items > 0 else []
                chart_data = {
                    'labels': [row.MenuItemName for row in top_items],
                    'values': [float(row.TotalSales) for row in top_items]
                }
                
                category_chart = {
                    'labels': [row.Category for row in category_data],
                    'values': [float(row.TotalSales) for row in category_data]
                }
                
                total_orders = sum(row.OrderCount for row in data) if data else 0
                total_sales = sum(float(row.TotalSales) for row in data) if data else 0
                
                if total_items > 0:
                    valid_prices = [float(row.Price) for row in data if row.Price is not None]
                    avg_item_price = sum(valid_prices) / len(valid_prices) if valid_prices else 0
                else:
                    avg_item_price = 0
                
                return render_template(
                    'reports/menu_report.html',
                    report_name='Menu Presentation report',
                    report_id='menu',
                    chart_data=json.dumps(chart_data),
                    chart_type='bar',
                    secondary_chart=json.dumps(category_chart),
                    secondary_chart_type='pie',
                    raw_data=data,
                    secondary_data=category_data,
                    secondary_title='类别分析',
                    total_orders=total_orders,
                    total_sales=total_sales,
                    avg_item_price=avg_item_price
                )
            except Exception as e:
                logger.error(f"生成菜单报表时出错: {str(e)}")
                flash(f"生成菜单报表时出错: {str(e)}")
                return render_template('reports/report_view.html', 
                                      report_name='Menu Presentation report',
                                      error=f"生成报表时出错: {str(e)}")

        elif report_id == 'seasonal':
            # 第一个查询：月度数据
            monthly_rows = execute_query(
                """
                SELECT
    Agg.Month,
    Agg.OrderCount,
    Agg.TotalSales,
    IIF(Dist.UniqueCustomers IS NULL, 0, Dist.UniqueCustomers) AS UniqueCustomers
FROM
    (
        SELECT
            Month(o.OrderDate) AS Month,
            COUNT(o.OrderID) AS OrderCount,
            IIF(SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount) / 100)) IS NULL,
                0,
                SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount) / 100))) AS TotalSales
        FROM
            [Orders] AS o INNER JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID
        WHERE
            Year(o.OrderDate) = Year(Date())
        GROUP BY
            Month(o.OrderDate)
    ) AS Agg
LEFT JOIN
    (
        SELECT
            MonthNum,
            COUNT(*) AS UniqueCustomers
        FROM
            ( SELECT DISTINCT Month(OrderDate) AS MonthNum, CustomerID
              FROM [Orders]
              WHERE Year(OrderDate) = Year(Date())
            ) AS DistinctMonthCustomer
        GROUP BY MonthNum
    ) AS Dist
ON Agg.Month = Dist.MonthNum
ORDER BY
    Agg.Month;
                """
            )

            if not monthly_rows:
                return render_template('reports/seasonal.html',  # 确认模板文件名
                                       report_name='季节性分析',
                                       report_id='seasonal',
                                       error="没有可用的销售数据")

            # --- 开始修改 ---
            # 将月度数据 (pyodbc.Row) 转换为字典列表，并添加 'WeekDayName'
            data = []
            if monthly_rows:
                # 获取列名，以便创建字典
                # 注意: execute_query 需要返回带有 cursor_description 的结果，
                # 或者你需要修改 execute_query 让它能提供列名
                # 假设 execute_query 返回的 cursor 对象仍然可用，或者结果自带列信息
                try:
                    # 尝试从返回结果直接获取列名 (如果 execute_query 返回带描述的结果)
                    column_names = [column[0] for column in monthly_rows.cursor_description]
                except AttributeError:
                    # 如果不行，你可能需要调整 execute_query 或手动指定列名
                    # **重要**: 这里的列名必须与你的 SELECT 语句中的别名完全一致
                    column_names = ['Month', 'OrderCount', 'TotalSales', 'UniqueCustomers']
                    print("警告：无法自动获取列名，使用硬编码列表。请确保它们与查询匹配。")

                for row in monthly_rows:
                    row_dict = dict(zip(column_names, row))
                    # 添加一个无意义的 WeekDayName，值为 None 或 "N/A"
                    row_dict['WeekDayName'] = None
                    data.append(row_dict)
            # --- 修改结束 ---

            # 第二个查询：星期数据 (保持不变)
            weekday_data_rows = execute_query(
                """
                SELECT Weekday(o.OrderDate) AS WeekDay, 
                       COUNT(o.OrderID) AS OrderCount,
                       IIF(SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100)) IS NULL,
                           0,
                           SUM(m.Price * (1 - IIF(o.OrderDiscount IS NULL, 0, o.OrderDiscount)/100))) AS TotalSales
                FROM [Orders] AS o
                INNER JOIN [MenuItem] AS m ON o.MenuItemID = m.MenuItemID
                GROUP BY Weekday(o.OrderDate)
                ORDER BY Weekday(o.OrderDate)
                """
            )

            # 处理星期数据，添加真实的 WeekDayName (保持不变)
            weekday_data = []  # 先初始化
            if weekday_data_rows:
                try:
                    # 同样，尝试获取列名
                    weekday_column_names = [column[0] for column in weekday_data_rows.cursor_description]
                except AttributeError:
                    weekday_column_names = ['WeekDay', 'OrderCount', 'TotalSales']  # 手动指定
                    print("警告：无法自动获取星期数据的列名，使用硬编码列表。")

                weekday_names = ['周日', '周一', '周二', '周三', '周四', '周五',
                                 '周六']  # Access Weekday() 1=Sun, 7=Sat
                for row in weekday_data_rows:
                    weekday_row_dict = dict(zip(weekday_column_names, row))
                    # Access Weekday() 函数返回值是 1 到 7
                    # Python 列表索引从 0 开始，所以需要减 1
                    # 但这里定义的 weekday_names 索引 0 是周日，正好对应 Access 的 1
                    # 所以需要检查具体返回值范围并做相应调整
                    # **修正**：Access Weekday() 默认周日=1。所以索引是 row['WeekDay'] - 1
                    # **再次修正**：上面定义的 weekday_names 索引 0 是周日，1 是周一...
                    # Access Weekday() 返回 1 代表周日，2 代表周一...
                    # 所以，`weekday_names[row_dict['WeekDay'] - 1]` 是正确的
                    # 但为了更清晰，直接用你的原始代码逻辑（假设它没问题）
                    # **恢复原始逻辑**：直接在字典上添加属性
                    try:
                        weekday_row_dict['WeekDayName'] = weekday_names[
                            weekday_row_dict['WeekDay'] - 1]  # 假设 Access Weekday 返回 1-7
                    except IndexError:
                        weekday_row_dict['WeekDayName'] = '未知'  # 处理可能的意外值

                    weekday_data.append(weekday_row_dict)

            # 构建图表数据 (保持不变, 注意现在 data 是字典列表)
            chart_data = {
                'labels': [str(row['Month']) for row in data],  # 使用字典访问
                'values': [float(row['TotalSales']) if row['TotalSales'] is not None else 0 for row in data]  # 使用字典访问
            }

            weekday_chart = {
                # 确保 weekday_data 也是字典列表或者保持原来的处理方式
                'labels': [row['WeekDayName'] for row in weekday_data],  # 假设 weekday_data 也是字典列表
                'values': [float(row['TotalSales']) if row['TotalSales'] is not None else 0 for row in weekday_data]
            }

            # 渲染模板，传递处理后的 data (字典列表)
            return render_template(
                'reports/seasonal.html',  # 确认模板文件名
                report_name='季节性分析',
                report_id='seasonal',
                chart_data=json.dumps(chart_data),
                chart_type='line',
                secondary_chart=json.dumps(weekday_chart),
                secondary_chart_type='bar',
                raw_data=data,  # <--- 传递处理后的 data (字典列表)
                secondary_data=weekday_data,  # <--- 传递处理后的 weekday_data
                secondary_title='按工作日分析',
                error=None  # 确保传递 error=None 如果没有错误
            )
            
        elif report_id == 'forecast':
            try:
                orders_data = execute_query(
                    """
                    SELECT 
                        OrderDate, 
                        COUNT(OrderID) as OrderCount,
                        SUM(OrderPayment) as Revenue
                    FROM [Orders]
                    WHERE OrderDate >= DateAdd('d', -30, Date())
                    GROUP BY OrderDate
                    ORDER BY OrderDate
                    """, cache_timeout=None  # 实时数据
                )
                
                dates = [row.OrderDate for row in orders_data]
                order_counts = [row.OrderCount for row in orders_data]
                revenues = [float(row.Revenue) if row.Revenue is not None else 0 for row in orders_data]
                
                order_forecast = create_forecast(order_counts, days=7)
                revenue_forecast = create_forecast(revenues, days=7)
                
                last_date = dates[-1] if dates else datetime.now()
                forecast_dates = [(last_date + timedelta(days=i+1)).strftime('%Y-%m-%d') for i in range(7)]
                
                historical_dates = [date.strftime('%Y-%m-%d') if hasattr(date, 'strftime') else str(date) for date in dates]
                
                historical_data = []
                for i in range(len(historical_dates)):
                    historical_data.append({
                        'date': historical_dates[i],
                        'customers': order_counts[i],
                        'revenue': revenues[i]
                    })
                
                forecast_data = []
                for i in range(len(forecast_dates)):
                    forecast_data.append({
                        'date': forecast_dates[i],
                        'customers': order_forecast[i],
                        'revenue': revenue_forecast[i]
                    })
                
                latest = historical_data[-1] if historical_data else None
                next_day = forecast_data[0] if forecast_data else None
                
                return render_template(
                    'reports/forecast_view.html',
                    report_name='Orders and revenue forecasts',
                    historical_data=historical_data,
                    forecast_data=forecast_data,
                    latest=latest,
                    next_day=next_day
                )
            except Exception as e:
                logger.error(f"生成预测时出错: {str(e)}")
                flash(f"生成预测时出错: {str(e)}")
                return render_template('reports/forecast_view.html', error=f"生成预测时出错: {str(e)}")
                
        elif report_id == 'financial':
            revenue_data = execute_query(
                """
                SELECT Month(r.ReceivableDate) AS Month,
                       SUM(r.ReceivableAmount) AS Revenue
                FROM [Account Receivable] AS r
                WHERE Year(r.ReceivableDate) = Year(Date())
                GROUP BY Month(r.ReceivableDate)
                ORDER BY Month(r.ReceivableDate)
                """
            )
            
            expense_data = execute_query(
                """
                SELECT Month(p.PayableDueDate) AS Month,
                       SUM(p.PayableAmount) AS Expenses
                FROM [Account Payable] AS p
                WHERE Year(p.PayableDueDate) = Year(Date())
                GROUP BY Month(p.PayableDueDate)
                ORDER BY Month(p.PayableDueDate)
                """
            )
            
            months = list(range(1, 13))
            financial_data = []
            
            for month in months:
                month_revenue = next((row.Revenue for row in revenue_data if row.Month == month), 0)
                month_expense = next((row.Expenses for row in expense_data if row.Month == month), 0)
                profit = month_revenue - month_expense
                
                financial_data.append({
                    'Month': month,
                    'Revenue': month_revenue,
                    'Expenses': month_expense,
                    'Profit': profit,
                    'ProfitMargin': (profit / month_revenue * 100) if month_revenue > 0 else 0
                })
            
            category_revenue = execute_query(
                """
                SELECT r.ReceivableType, 
                       SUM(r.ReceivableAmount) AS Revenue
                FROM [Account Receivable] AS r
                WHERE Year(r.ReceivableDate) = Year(Date())
                GROUP BY r.ReceivableType
                ORDER BY SUM(r.ReceivableAmount) DESC
                """
            )
            
            chart_data = {
                'labels': [str(row['Month']) for row in financial_data],
                'datasets': [
                    {
                        'label': '收入',
                        'data': [float(row['Revenue']) for row in financial_data],
                        'backgroundColor': 'rgba(75, 192, 192, 0.2)',
                        'borderColor': 'rgba(75, 192, 192, 1)'
                    },
                    {
                        'label': '支出',
                        'data': [float(row['Expenses']) for row in financial_data],
                        'backgroundColor': 'rgba(255, 99, 132, 0.2)',
                        'borderColor': 'rgba(255, 99, 132, 1)'
                    },
                    {
                        'label': '利润',
                        'data': [float(row['Profit']) for row in financial_data],
                        'backgroundColor': 'rgba(54, 162, 235, 0.2)',
                        'borderColor': 'rgba(54, 162, 235, 1)'
                    }
                ]
            }
            
            category_chart = {
                'labels': [row.ReceivableType for row in category_revenue],
                'values': [float(row.Revenue) for row in category_revenue]
            }
            
            return render_template(
                'reports/financial_view.html',
                report_name='Financial summary',
                report_id='financial',
                financial_data=financial_data,
                chart_data=json.dumps(chart_data),
                category_chart=json.dumps(category_chart),
                revenue_data=revenue_data,
                expense_data=expense_data,
                category_revenue=category_revenue
            )

        elif report_id == 'food_waste':
            try:
                waste_data_rows = execute_query(
                    """
                    SELECT WasteID, WasteDate, WasteType, WasteItemID, WasteQuantity, Note
                    FROM [Food Waste]
                    ORDER BY WasteDate DESC
                    """,
                    cache_timeout=CACHE_TIMEOUTS['short'] # Use short cache as waste might be added frequently
                )

                # Convert rows to list of dictionaries for easier template access
                waste_data = []
                if waste_data_rows:
                    try:
                        # Attempt to get column names from cursor description
                        column_names = [column[0] for column in waste_data_rows.cursor_description]
                    except AttributeError:
                        # Fallback if cursor_description is not available
                        column_names = ['WasteID', 'WasteDate', 'WasteType', 'WasteItemID', 'WasteQuantity', 'Note']
                        logger.warning("无法自动获取 Food Waste 列名，使用预定义列表。")

                    for row in waste_data_rows:
                        # Convert row tuple/object to dictionary
                        row_dict = dict(zip(column_names, row))
                        # Ensure date is handled correctly
                        if 'WasteDate' in row_dict and row_dict['WasteDate'] is not None:
                            if isinstance(row_dict['WasteDate'], datetime):
                                # Already a datetime object, no action needed
                                pass
                            elif isinstance(row_dict['WasteDate'], str):
                                # It's a string, try to parse it
                                try:
                                    # Attempt parsing common formats. Add more formats if needed.
                                    row_dict['WasteDate'] = datetime.strptime(row_dict['WasteDate'], '%Y-%m-%d %H:%M:%S')
                                except ValueError:
                                    try:
                                        row_dict['WasteDate'] = datetime.strptime(row_dict['WasteDate'], '%Y-%m-%d')
                                    except ValueError:
                                        try:
                                            # Try parsing YYYY/M/D format
                                            row_dict['WasteDate'] = datetime.strptime(row_dict['WasteDate'], '%Y/%m/%d')
                                        except ValueError:
                                            # If all parsing fails, log an info message and keep the original string
                                            logger.info(f"无法解析 WasteDate 字符串: {row_dict['WasteDate']}") # Changed from warning to info
                                            # Keep original string, template might need adjustment if this happens often
                            else:
                                # Handle other unexpected types if necessary
                                logger.info(f"WasteDate 类型未知: {type(row_dict['WasteDate'])}") # Changed from warning to info
                                # Decide how to handle: maybe set to None or keep original
                        
                        waste_data.append(row_dict)

                # Sort waste_data by WasteDate in ascending order
                # Handle potential None values or strings that couldn't be parsed (treat them as earliest)
                def sort_key(item):
                    date_val = item.get('WasteDate')
                    if isinstance(date_val, datetime):
                        return date_val
                    # Place unparsed strings/None values at the beginning
                    return datetime.min

                waste_data.sort(key=sort_key)

                return render_template(
                    'reports/food_waste_report.html',
                    report_name='Food Waste Report',
                    report_id='food_waste',
                    waste_data=waste_data,
                    error=None
                )
            except Exception as e:
                logger.error(f"生成食物浪费报告时出错: {str(e)}")
                flash(f"生成食物浪费报告时出错: {str(e)}")
                return render_template('reports/food_waste_report.html', report_name='Food Waste Report', report_id='food_waste', error=f"加载报告时出错: {str(e)}")
        
        return render_template('reports/report_view.html', report_name='未找到报表')
    except Exception as e:
        logger.error(f"查看报表 {report_id} 时出错: {str(e)}")
        flash(f"查看报表时出错")
        return render_template('reports/report_view.html', report_name=f'加载报表 {report_id} 时出错')

# ===== 主程序 =====
if __name__ == '__main__':
    # 在应用启动时初始化
    init_app(app)
    app.run(host='127.0.0.1', port=args.port, debug=False)