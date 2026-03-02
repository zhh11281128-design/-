import datetime
import os
import sys
import sqlite3
import logging
import re
from openpyxl import Workbook, load_workbook
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.spinner import Spinner
from kivy.uix.textinput import TextInput
from kivy.uix.popup import Popup
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.core.text import LabelBase, DEFAULT_FONT
from kivy.resources import resource_add_path, resource_find
from kivy.metrics import dp
from plyer import filechooser

# ----- 日志配置 -----
logging.basicConfig(
    filename='app.log',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

# ----- 颜色常量（便于主题调整）-----
COLOR_GREEN = (0.2, 0.7, 0.2, 1)
COLOR_BLUE = (0.3, 0.6, 0.9, 1)
COLOR_RED = (0.9, 0.2, 0.2, 1)
COLOR_GRAY = (0.7, 0.7, 0.7, 1)
COLOR_DARK_GRAY = (0.5, 0.5, 0.5, 1)

# ----- 解决中文显示问题（增强字体后备）-----
def setup_chinese_font():
    """自动检测并注册中文字体，支持多个备选字体"""
    # 常见中文字体文件名列表
    font_candidates = [
        'simhei.ttf',          # 黑体
        'msyh.ttf',            # 微软雅黑
        'simsun.ttc',          # 宋体
        'NotoSansCJK-Regular.ttc'  # 思源黑体
    ]
    # 可能的搜索路径
    search_paths = [
        os.path.dirname(__file__),
        getattr(sys, '_MEIPASS', ''),
    ]
    font_path = None
    for path in search_paths:
        for font in font_candidates:
            full_path = os.path.join(path, font)
            if os.path.exists(full_path):
                font_path = full_path
                break
        if font_path:
            break
    # 如果未找到，尝试 resource_find
    if not font_path:
        for font in font_candidates:
            res_path = resource_find(font)
            if res_path:
                font_path = res_path
                break
    if font_path:
        LabelBase.register(name='SimHei', fn_regular=font_path)
        LabelBase.register(name=DEFAULT_FONT, fn_regular=font_path)
    else:
        # 无中文字体时使用默认字体（可能显示为方块，但程序继续运行）
        print("警告：未找到中文字体，中文可能显示异常")

setup_chinese_font()

# 全局定义年份范围（2023-2050）
START_YEAR = 2023
END_YEAR = 2050
YEAR_LIST = [str(y) for y in range(START_YEAR, END_YEAR + 1)]

# 数据库表名和列定义
TABLE_NAME = 'records'
COLUMNS = ['date', 'type', 'amount', 'event', 'party', 'platform', 'other']
CREATE_TABLE_SQL = f'''
    CREATE TABLE IF NOT EXISTS {TABLE_NAME} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT NOT NULL,
        type TEXT NOT NULL,
        amount REAL NOT NULL,
        event TEXT,
        party TEXT,
        platform TEXT,
        other TEXT
    )
'''
CREATE_INDEX_SQL = f'CREATE INDEX IF NOT EXISTS idx_date ON {TABLE_NAME}(date);'

# ---------- 辅助函数 ----------
def sanitize_filename(name):
    """替换文件名中的非法字符为下划线"""
    return re.sub(r'[\\/*?:"<>|]', '_', name)

def parse_date_cell(date_cell):
    """
    将 Excel 中的日期单元格转换为标准 'YYYY-MM-DD' 字符串。
    支持 datetime 对象、字符串和数字。
    返回 None 表示无效日期。
    """
    if date_cell is None:
        return None
    if isinstance(date_cell, datetime.datetime):
        return date_cell.strftime('%Y-%m-%d')
    if isinstance(date_cell, str):
        # 取前10个字符，并尝试解析
        s = date_cell.strip()[:10]
        try:
            datetime.datetime.strptime(s, '%Y-%m-%d')
            return s
        except ValueError:
            # 尝试其他常见格式
            for fmt in ('%Y/%m/%d', '%d-%m-%Y', '%d/%m/%Y'):
                try:
                    dt = datetime.datetime.strptime(s, fmt)
                    return dt.strftime('%Y-%m-%d')
                except ValueError:
                    continue
        return None
    if isinstance(date_cell, (int, float)):
        # 处理 Excel 序列日期（简单假设为 1900 系统）
        try:
            # 仅当数字较大时认为是序列日期
            if date_cell > 10000:
                # Excel 的 1900 日期系统，注意 1900-02-29 的 bug 此处忽略
                base_date = datetime.datetime(1899, 12, 30)
                delta = datetime.timedelta(days=date_cell)
                dt = base_date + delta
                return dt.strftime('%Y-%m-%d')
        except:
            pass
    return None

def init_database(db_path):
    """如果数据库不存在，则创建表并建立索引"""
    conn = sqlite3.connect(db_path)
    c = conn.cursor()
    c.execute(CREATE_TABLE_SQL)
    c.execute(CREATE_INDEX_SQL)
    conn.commit()
    conn.close()

# ---------- 自定义日期选择控件（简化版）----------
class DateSelector(BoxLayout):
    """
    根据查询类型动态生成日期选择器。
    通过 get_date_range() 返回 (start_date, end_date) 元组，
    对于天和月、年，end_date 可能为 None 或计算得出。
    使用方法：绑定 type 属性变化或直接调用 set_type(type_str)
    """
    def __init__(self, query_type='天', **kwargs):
        super().__init__(orientation='vertical', size_hint_y=None, spacing=dp(2), **kwargs)
        self.query_type = query_type
        self.spinners = {}
        self.build()

    def build(self):
        self.clear_widgets()
        self.spinners.clear()

        now = datetime.datetime.now()
        current_year = now.year if START_YEAR <= now.year <= END_YEAR else START_YEAR
        current_month = now.month
        current_day = now.day

        spinner_style = {
            'size_hint_y': None,
            'height': dp(30),
            'font_size': dp(12)
        }

        if self.query_type == '天':
            h_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(30), spacing=dp(3))
            year_sp = Spinner(text=str(current_year), values=YEAR_LIST, size_hint_x=0.33, **spinner_style)
            h_layout.add_widget(year_sp)
            self.spinners['year'] = year_sp

            months = [f"{m:02d}" for m in range(1, 13)]
            month_sp = Spinner(text=f"{current_month:02d}", values=months, size_hint_x=0.33, **spinner_style)
            h_layout.add_widget(month_sp)
            self.spinners['month'] = month_sp

            days = [f"{d:02d}" for d in range(1, 32)]
            day_sp = Spinner(text=f"{current_day:02d}", values=days, size_hint_x=0.33, **spinner_style)
            h_layout.add_widget(day_sp)
            self.spinners['day'] = day_sp
            self.add_widget(h_layout)

        elif self.query_type == '月':
            h_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(30), spacing=dp(3))
            year_sp = Spinner(text=str(current_year), values=YEAR_LIST, size_hint_x=0.5, **spinner_style)
            h_layout.add_widget(year_sp)
            self.spinners['year'] = year_sp

            months = [f"{m:02d}" for m in range(1, 13)]
            month_sp = Spinner(text=f"{current_month:02d}", values=months, size_hint_x=0.5, **spinner_style)
            h_layout.add_widget(month_sp)
            self.spinners['month'] = month_sp
            self.add_widget(h_layout)

        elif self.query_type == '年':
            h_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(30))
            year_sp = Spinner(text=str(current_year), values=YEAR_LIST, size_hint_x=1, **spinner_style)
            h_layout.add_widget(year_sp)
            self.spinners['year'] = year_sp
            self.add_widget(h_layout)

        elif self.query_type == '时间段':
            # 起始日期
            start_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(30), spacing=dp(2))
            start_layout.add_widget(Label(text='从', size_hint_x=0.08, font_size=dp(12), halign='center'))
            start_year = Spinner(text=str(current_year), values=YEAR_LIST, size_hint_x=0.28, **spinner_style)
            start_layout.add_widget(start_year)
            self.spinners['start_year'] = start_year

            months = [f"{m:02d}" for m in range(1, 13)]
            start_month = Spinner(text=f"{current_month:02d}", values=months, size_hint_x=0.28, **spinner_style)
            start_layout.add_widget(start_month)
            self.spinners['start_month'] = start_month

            days = [f"{d:02d}" for d in range(1, 32)]
            start_day = Spinner(text=f"{current_day:02d}", values=days, size_hint_x=0.28, **spinner_style)
            start_layout.add_widget(start_day)
            self.spinners['start_day'] = start_day
            self.add_widget(start_layout)

            # 结束日期
            end_layout = BoxLayout(orientation='horizontal', size_hint_y=None, height=dp(30), spacing=dp(2))
            end_layout.add_widget(Label(text='到', size_hint_x=0.08, font_size=dp(12), halign='center'))
            end_year = Spinner(text=str(current_year), values=YEAR_LIST, size_hint_x=0.28, **spinner_style)
            end_layout.add_widget(end_year)
            self.spinners['end_year'] = end_year

            end_month = Spinner(text=f"{current_month:02d}", values=months, size_hint_x=0.28, **spinner_style)
            end_layout.add_widget(end_month)
            self.spinners['end_month'] = end_month

            end_day = Spinner(text=f"{current_day:02d}", values=days, size_hint_x=0.28, **spinner_style)
            end_layout.add_widget(end_day)
            self.spinners['end_day'] = end_day
            self.add_widget(end_layout)

    def set_type(self, query_type):
        if query_type != self.query_type:
            self.query_type = query_type
            self.build()

    def get_date_range(self):
        """
        返回 (start_date, end_date) 字符串。
        对于 '天'，end_date = start_date。
        对于 '月' 和 '年'，end_date 为下一天的开始（用于 date < end_date 查询）。
        """
        try:
            if self.query_type == '天':
                year = self.spinners['year'].text
                month = self.spinners['month'].text
                day = self.spinners['day'].text
                start = f"{year}-{month}-{day}"
                return start, start
            elif self.query_type == '月':
                year = self.spinners['year'].text
                month = self.spinners['month'].text
                start = f"{year}-{month}-01"
                if month == '12':
                    end = f"{int(year)+1}-01-01"
                else:
                    end = f"{year}-{int(month)+1:02d}-01"
                return start, end
            elif self.query_type == '年':
                year = self.spinners['year'].text
                start = f"{year}-01-01"
                end = f"{int(year)+1}-01-01"
                return start, end
            elif self.query_type == '时间段':
                sy = self.spinners['start_year'].text
                sm = self.spinners['start_month'].text
                sd = self.spinners['start_day'].text
                ey = self.spinners['end_year'].text
                em = self.spinners['end_month'].text
                ed = self.spinners['end_day'].text
                start = f"{sy}-{sm}-{sd}"
                end = f"{ey}-{em}-{ed}"
                return start, end
        except KeyError as e:
            raise ValueError(f"日期控件未完整初始化：{e}")
        return None, None

# ---------- 主界面 ----------
class MainScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        layout = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(10))

        self.title_label = Label(
            text='我的记账本',
            size_hint_y=0.12,
            font_size=24,
            halign='center'
        )
        layout.add_widget(self.title_label)

        # 账单选择下拉列表
        bill_select_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        bill_select_layout.add_widget(Label(text='选择账单：', size_hint_x=0.25, font_size=dp(14)))
        self.bill_spinner = Spinner(
            text='',
            values=[],
            size_hint_x=0.75,
            font_size=dp(12)
        )
        self.bill_spinner.bind(text=self.on_bill_change)
        bill_select_layout.add_widget(self.bill_spinner)
        layout.add_widget(bill_select_layout)

        # 三个操作按钮
        btn_layout = BoxLayout(size_hint_y=0.1, spacing=dp(10))
        new_btn = Button(
            text='新建',
            background_color=COLOR_GREEN,
            font_size=dp(14)
        )
        new_btn.bind(on_press=self.new_bill_popup)
        btn_layout.add_widget(new_btn)

        rename_btn = Button(
            text='重命名',
            background_color=COLOR_BLUE,
            font_size=dp(14)
        )
        rename_btn.bind(on_press=self.rename_bill_popup)
        btn_layout.add_widget(rename_btn)

        delete_btn = Button(
            text='删除',
            background_color=COLOR_RED,
            font_size=dp(14)
        )
        delete_btn.bind(on_press=self.delete_bill_popup)
        btn_layout.add_widget(delete_btn)
        layout.add_widget(btn_layout)

        # 功能按钮
        view_btn = Button(
            text='查看事件',
            background_color=COLOR_BLUE,
            size_hint_y=0.12
        )
        view_btn.bind(on_press=self.go_to_view_event)
        layout.add_widget(view_btn)

        record_btn = Button(
            text='记收支',
            background_color=(0.2, 0.8, 0.5, 1),
            size_hint_y=0.12
        )
        record_btn.bind(on_press=self.go_to_record)
        layout.add_widget(record_btn)

        import_btn = Button(
            text='导入 Excel',
            background_color=COLOR_DARK_GRAY,
            size_hint_y=0.12
        )
        import_btn.bind(on_press=self.go_to_import)
        layout.add_widget(import_btn)

        export_btn = Button(
            text='导出 Excel',
            background_color=COLOR_DARK_GRAY,
            size_hint_y=0.12
        )
        export_btn.bind(on_press=self.go_to_export)
        layout.add_widget(export_btn)

        exit_btn = Button(
            text='退出该应用',
            background_color=COLOR_RED,
            size_hint_y=0.12,
            font_size=dp(16)
        )
        exit_btn.bind(on_press=self.exit_application)
        layout.add_widget(exit_btn)

        self.add_widget(layout)

    def on_enter(self):
        app = App.get_running_app()
        app.refresh_bill_list()
        display_names = [os.path.splitext(name)[0] for name in app.bill_files]
        self.bill_spinner.values = display_names
        if app.current_bill:
            current_display = os.path.splitext(app.current_bill)[0]
            self.bill_spinner.text = current_display
            self.update_title(current_display)

    def update_title(self, display_name):
        self.title_label.text = f'我的记账本 - {display_name}'

    def on_bill_change(self, spinner, text):
        if not text:
            return
        app = App.get_running_app()
        for filename in app.bill_files:
            if os.path.splitext(filename)[0] == text:
                if app.current_bill != filename:
                    app.current_bill = filename
                    self.update_title(text)
                break

    def new_bill_popup(self, instance):
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(Label(text='请输入新账单名称：'))
        name_input = TextInput(hint_text='例如：2025年家庭账本', multiline=False)
        content.add_widget(name_input)

        btn_layout = BoxLayout(size_hint_y=0.3, spacing=dp(10))
        cancel_btn = Button(text='取消')
        ok_btn = Button(text='确定', background_color=COLOR_GREEN)
        btn_layout.add_widget(cancel_btn)
        btn_layout.add_widget(ok_btn)
        content.add_widget(btn_layout)

        popup = Popup(title='新建账单', content=content, size_hint=(0.8, 0.4))

        def on_ok(btn):
            raw_name = name_input.text.strip()
            app = App.get_running_app()
            success, msg = app.create_new_bill(raw_name)
            if success:
                display_names = [os.path.splitext(name)[0] for name in app.bill_files]
                self.bill_spinner.values = display_names
                current_display = os.path.splitext(app.current_bill)[0]
                self.bill_spinner.text = current_display
                self.update_title(current_display)
                popup.dismiss()
            else:
                name_input.text = ''
                name_input.hint_text = msg
        ok_btn.bind(on_press=on_ok)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def rename_bill_popup(self, instance):
        app = App.get_running_app()
        if not app.current_bill:
            return

        current_display = os.path.splitext(app.current_bill)[0]

        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(Label(text='请输入新账单名称：'))
        name_input = TextInput(text=current_display, multiline=False, hint_text='新名称')
        content.add_widget(name_input)

        btn_layout = BoxLayout(size_hint_y=0.3, spacing=dp(10))
        cancel_btn = Button(text='取消')
        ok_btn = Button(text='确定', background_color=COLOR_BLUE)
        btn_layout.add_widget(cancel_btn)
        btn_layout.add_widget(ok_btn)
        content.add_widget(btn_layout)

        popup = Popup(title='重命名账单', content=content, size_hint=(0.8, 0.4))

        def on_ok(btn):
            new_name = name_input.text.strip()
            if not new_name:
                name_input.hint_text = '名称不能为空'
                return
            app = App.get_running_app()
            success, msg = app.rename_current_bill(new_name)
            if success:
                display_names = [os.path.splitext(name)[0] for name in app.bill_files]
                self.bill_spinner.values = display_names
                current_display = os.path.splitext(app.current_bill)[0]
                self.bill_spinner.text = current_display
                self.update_title(current_display)
                popup.dismiss()
            else:
                name_input.text = ''
                name_input.hint_text = msg
        ok_btn.bind(on_press=on_ok)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def delete_bill_popup(self, instance):
        app = App.get_running_app()
        if not app.current_bill:
            return

        current_display = os.path.splitext(app.current_bill)[0]

        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(Label(
            text=f'确定要删除账单“{current_display}”吗？\n删除后无法恢复。\n请在下方输入“删除”以确认：',
            halign='center'
        ))
        confirm_input = TextInput(multiline=False, hint_text='输入“删除”')
        content.add_widget(confirm_input)

        btn_layout = BoxLayout(size_hint_y=0.3, spacing=dp(10))
        cancel_btn = Button(text='取消')
        ok_btn = Button(text='确认删除', background_color=COLOR_RED)
        btn_layout.add_widget(cancel_btn)
        btn_layout.add_widget(ok_btn)
        content.add_widget(btn_layout)

        popup = Popup(title='删除账单', content=content, size_hint=(0.8, 0.5))

        def on_ok(btn):
            if confirm_input.text.strip() == '删除':
                app = App.get_running_app()
                success, msg = app.delete_current_bill()
                if success:
                    display_names = [os.path.splitext(name)[0] for name in app.bill_files]
                    self.bill_spinner.values = display_names
                    if app.current_bill:
                        current_display = os.path.splitext(app.current_bill)[0]
                        self.bill_spinner.text = current_display
                        self.update_title(current_display)
                    else:
                        self.bill_spinner.text = ''
                        self.title_label.text = '我的记账本'
                    popup.dismiss()
                else:
                    confirm_input.text = ''
                    confirm_input.hint_text = msg
            else:
                confirm_input.text = ''
                confirm_input.hint_text = '请输入“删除”'
        ok_btn.bind(on_press=on_ok)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def go_to_view_event(self, instance):
        self.manager.current = 'view_event'

    def go_to_record(self, instance):
        self.manager.current = 'record'

    def go_to_import(self, instance):
        self.manager.current = 'import'

    def go_to_export(self, instance):
        self.manager.current = 'export'

    def exit_application(self, instance):
        App.get_running_app().stop()

# ---------- 记收支页面 ----------
class RecordScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        main_layout = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(8))

        title = Label(
            text='记录收支',
            size_hint_y=0.08,
            font_size=22,
            halign='center'
        )
        main_layout.add_widget(title)

        # 日期选择（使用简化的自定义控件）
        date_layout = BoxLayout(size_hint_y=0.1, spacing=dp(5))
        date_layout.add_widget(Label(text='交易日期：', size_hint_x=0.2))
        self.date_selector = DateSelector(query_type='天')
        date_layout.add_widget(self.date_selector)
        main_layout.add_widget(date_layout)

        # 收支类型
        type_layout = BoxLayout(size_hint_y=0.09, spacing=dp(5))
        type_layout.add_widget(Label(text='收支类型：', size_hint_x=0.2))
        self.type_spinner = Spinner(
            text='收入',
            values=('收入', '支出'),
            size_hint_x=0.8,
            font_size=dp(12)
        )
        type_layout.add_widget(self.type_spinner)
        main_layout.add_widget(type_layout)

        # 金额
        amount_layout = BoxLayout(size_hint_y=0.09, spacing=dp(5))
        amount_layout.add_widget(Label(text='金额（元）：', size_hint_x=0.2))
        self.amount_input = TextInput(
            hint_text='请输入数字，如 100.50',
            input_filter='float',
            size_hint_x=0.8,
            font_size=dp(12)
        )
        amount_layout.add_widget(self.amount_input)
        main_layout.add_widget(amount_layout)

        # 事件
        event_layout = BoxLayout(size_hint_y=0.09, spacing=dp(5))
        event_layout.add_widget(Label(text='因由：', size_hint_x=0.2))
        self.event_input = TextInput(
            hint_text='如 工资发放、餐饮消费',
            size_hint_x=0.8,
            font_size=dp(12)
        )
        event_layout.add_widget(self.event_input)
        main_layout.add_widget(event_layout)

        # 相关方
        party_layout = BoxLayout(size_hint_y=0.09, spacing=dp(5))
        party_layout.add_widget(Label(text='相关方：', size_hint_x=0.2))
        self.party_input = TextInput(
            hint_text='如 公司A、超市B',
            size_hint_x=0.8,
            font_size=dp(12)
        )
        party_layout.add_widget(self.party_input)
        main_layout.add_widget(party_layout)

        # 交易平台
        platform_layout = BoxLayout(size_hint_y=0.09, spacing=dp(5))
        platform_layout.add_widget(Label(text='交易平台：', size_hint_x=0.2))
        self.platform_input = TextInput(
            hint_text='如 微信支付、支付宝、招商银行',
            size_hint_x=0.8,
            font_size=dp(12)
        )
        platform_layout.add_widget(self.platform_input)
        main_layout.add_widget(platform_layout)

        # 其他
        other_layout = BoxLayout(size_hint_y=0.09, spacing=dp(5))
        other_layout.add_widget(Label(text='其他：', size_hint_x=0.2))
        self.other_input = TextInput(
            hint_text='补充说明信息',
            size_hint_x=0.8,
            font_size=dp(12)
        )
        other_layout.add_widget(self.other_input)
        main_layout.add_widget(other_layout)

        # 占位
        main_layout.add_widget(BoxLayout(size_hint_y=0.1))

        # 按钮行
        btn_row = BoxLayout(size_hint_y=0.09, spacing=dp(5))
        save_only_btn = Button(
            text='保存',
            background_color=COLOR_GREEN,
            font_size=dp(14)
        )
        save_only_btn.bind(on_press=self.save_only)
        btn_row.add_widget(save_only_btn)

        save_and_exit_btn = Button(
            text='保存并退出',
            background_color=COLOR_GREEN,
            font_size=dp(14)
        )
        save_and_exit_btn.bind(on_press=self.save_and_exit)
        btn_row.add_widget(save_and_exit_btn)
        main_layout.add_widget(btn_row)

        back_btn = Button(
            text='返回主界面',
            background_color=COLOR_GRAY,
            size_hint_y=0.09,
            font_size=dp(14)
        )
        back_btn.bind(on_press=self.go_back)
        main_layout.add_widget(back_btn)

        self.tip_label = Label(
            text='',
            size_hint_y=0.05,
            halign='center',
            color=COLOR_RED,
            font_size=dp(12)
        )
        main_layout.add_widget(self.tip_label)

        self.add_widget(main_layout)

    def _save_record(self):
        """保存记录到数据库，返回 (成功标志, 错误信息或None)"""
        amount_text = self.amount_input.text.strip()
        if not amount_text:
            return False, '请输入金额！'
        try:
            amount = float(amount_text)
            if amount <= 0:
                return False, '金额必须大于0！'
        except ValueError:
            return False, '金额格式错误！'

        try:
            start_date, _ = self.date_selector.get_date_range()  # 仅需 start_date
            date_str = start_date
        except Exception as e:
            return False, f'无效日期：{e}'

        trade_type = self.type_spinner.text
        event = self.event_input.text.strip() or ''
        party = self.party_input.text.strip() or ''
        platform = self.platform_input.text.strip() or ''
        other = self.other_input.text.strip() or ''

        app = App.get_running_app()
        db_path = app.get_current_db_path()
        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        try:
            c.execute('''
                INSERT INTO records (date, type, amount, event, party, platform, other)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (date_str, trade_type, amount, event, party, platform, other))
            conn.commit()
            self.amount_input.text = ''
            self.event_input.text = ''
            self.party_input.text = ''
            self.platform_input.text = ''
            self.other_input.text = ''
            return True, None
        except Exception as e:
            logging.exception("记录保存失败")
            return False, f'保存失败：{str(e)}'
        finally:
            conn.close()

    def save_only(self, instance):
        success, msg = self._save_record()
        if success:
            self.tip_label.text = '保存成功'
            self.tip_label.color = (0, 0.7, 0, 1)
        else:
            self.tip_label.text = msg
            self.tip_label.color = COLOR_RED

    def save_and_exit(self, instance):
        success, msg = self._save_record()
        if success:
            self.tip_label.text = ''
            self.manager.current = 'main'
        else:
            self.tip_label.text = msg
            self.tip_label.color = COLOR_RED

    def go_back(self, instance):
        self.tip_label.text = ''
        self.manager.current = 'main'

# ---------- 查看事件界面 ----------
class ViewEventScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        main_layout = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(5))

        back_btn = Button(
            text='返回主界面',
            size_hint_y=None,
            height=dp(35),
            background_color=COLOR_GRAY
        )
        back_btn.bind(on_press=self.go_back)
        main_layout.add_widget(back_btn)

        self.type_spinner = Spinner(
            text='天',
            values=('天', '月', '年', '时间段'),
            size_hint_y=None,
            height=dp(35)
        )
        self.type_spinner.bind(text=self.on_type_change)
        main_layout.add_widget(self.type_spinner)

        self.date_selector = DateSelector(query_type='天')
        main_layout.add_widget(self.date_selector)

        self.income_type_spinner = Spinner(
            text='全部',
            values=('全部', '收入', '支出'),
            size_hint_y=None,
            height=dp(35)
        )
        main_layout.add_widget(self.income_type_spinner)

        query_btn = Button(
            text='查询',
            size_hint_y=None,
            height=dp(35),
            background_color=COLOR_BLUE
        )
        query_btn.bind(on_press=self.query_events)
        main_layout.add_widget(query_btn)

        from kivy.uix.scrollview import ScrollView
        scroll_view = ScrollView(
            size_hint=(1, 1),
            do_scroll_x=False,
            do_scroll_y=True
        )
        self.result_label = Label(
            text='请选择查询条件并点击查询',
            size_hint_y=None,
            size_hint_x=1,
            halign='left',
            valign='top',
            padding=dp(5)
        )
        self.result_label.bind(
            texture_size=lambda instance, value: setattr(instance, 'height', value[1])
        )
        self.result_label.bind(
            width=lambda instance, value: setattr(instance, 'text_size', (value - dp(10), None))
        )
        scroll_view.add_widget(self.result_label)
        main_layout.add_widget(scroll_view)

        self.add_widget(main_layout)

    def on_type_change(self, instance, text):
        self.date_selector.set_type(text)

    def go_back(self, instance):
        self.manager.current = 'main'

    def query_events(self, instance):
        query_type = self.type_spinner.text
        income_filter = self.income_type_spinner.text

        try:
            start_date, end_date = self.date_selector.get_date_range()
        except Exception as e:
            self.result_label.text = f"日期错误：{e}"
            return

        app = App.get_running_app()
        db_path = app.get_current_db_path()
        if not os.path.exists(db_path):
            self.result_label.text = "当前账单数据库不存在，请先记录一笔收支。"
            return

        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        try:
            sql = "SELECT date, type, amount, event, party, platform, other FROM records"
            conditions = []
            params = []

            if query_type in ('天', '时间段'):
                conditions.append("date BETWEEN ? AND ?")
                params.extend([start_date, end_date])
            elif query_type in ('月', '年'):
                conditions.append("date >= ? AND date < ?")
                params.extend([start_date, end_date])

            if income_filter == '收入':
                conditions.append("type = '收入'")
            elif income_filter == '支出':
                conditions.append("type = '支出'")

            if conditions:
                sql += " WHERE " + " AND ".join(conditions)

            sql += " ORDER BY date"

            c.execute(sql, params)
            rows = c.fetchall()

            if not rows:
                self.result_label.text = "没有找到符合条件的记录。"
                return

            lines = []
            total_income = 0.0
            total_expense = 0.0

            for row in rows:
                date_val, typ, amount, event, party, platform, other = row
                date_val = date_val or ''
                typ = typ or ''
                amount = amount or 0.0
                event = event or ''
                party = party or ''
                platform = platform or ''
                other = other or ''

                if typ == '收入':
                    total_income += amount
                elif typ == '支出':
                    total_expense += amount

                line = f'您于"{date_val}"，"{typ}"{amount:.2f}元，因由："{event}"，相关方："{party}"，交易平台："{platform}"，其他："{other}"'
                lines.append(line)

            lines.append("")
            lines.append(f"总收入：{total_income:.2f} 元")
            lines.append(f"总支出：{total_expense:.2f} 元")
            lines.append(f"结余：{total_income - total_expense:.2f} 元")

            self.result_label.text = "\n".join(lines)

        except Exception as e:
            logging.exception("查询失败")
            self.result_label.text = f"查询出错：{str(e)}"
        finally:
            conn.close()

# ---------- 导入页面 ----------
class ImportScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.selected_file = None

        main_layout = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(10))

        title = Label(text='导入 Excel', size_hint_y=0.1, font_size=22, halign='center')
        main_layout.add_widget(title)

        # 导入模式
        mode_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        mode_layout.add_widget(Label(text='导入模式：', size_hint_x=0.25, font_size=dp(14)))
        self.mode_spinner = Spinner(
            text='导入成为新账单',
            values=('导入成为新账单', '与指定账单合并为新账单', '并入指定账单'),
            size_hint_x=0.75,
            font_size=dp(12)
        )
        self.mode_spinner.bind(text=self.on_mode_change)
        mode_layout.add_widget(self.mode_spinner)
        main_layout.add_widget(mode_layout)

        # 选择外部文件
        file_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        file_layout.add_widget(Label(text='外部文件：', size_hint_x=0.25, font_size=dp(14)))
        self.file_path_label = Label(text='未选择', size_hint_x=0.5, font_size=dp(12), color=COLOR_DARK_GRAY)
        file_layout.add_widget(self.file_path_label)
        choose_file_btn = Button(text='浏览', size_hint_x=0.25, font_size=dp(12))
        choose_file_btn.bind(on_press=self.choose_file)
        file_layout.add_widget(choose_file_btn)
        main_layout.add_widget(file_layout)

        # 目标账单选择
        self.target_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        self.target_layout.add_widget(Label(text='目标账单：', size_hint_x=0.25, font_size=dp(14)))
        self.target_spinner = Spinner(
            text='',
            values=[],
            size_hint_x=0.75,
            font_size=dp(12)
        )
        self.target_layout.add_widget(self.target_spinner)
        main_layout.add_widget(self.target_layout)

        # 新账单名输入
        self.new_name_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        self.new_name_layout.add_widget(Label(text='新账单名：', size_hint_x=0.25, font_size=dp(14)))
        self.new_name_input = TextInput(
            hint_text='例如：合并账本',
            size_hint_x=0.75,
            font_size=dp(12),
            multiline=False
        )
        self.new_name_layout.add_widget(self.new_name_input)
        main_layout.add_widget(self.new_name_layout)

        main_layout.add_widget(BoxLayout(size_hint_y=0.1))

        # 执行导入按钮
        import_btn = Button(
            text='执行导入',
            size_hint_y=0.08,
            background_color=COLOR_GREEN,
            font_size=dp(14)
        )
        import_btn.bind(on_press=self.do_import)
        main_layout.add_widget(import_btn)

        back_btn = Button(
            text='返回主界面',
            size_hint_y=0.08,
            background_color=COLOR_GRAY,
            font_size=dp(14)
        )
        back_btn.bind(on_press=self.go_back)
        main_layout.add_widget(back_btn)

        self.tip_label = Label(
            text='',
            size_hint_y=0.05,
            halign='center',
            color=COLOR_RED,
            font_size=dp(12)
        )
        main_layout.add_widget(self.tip_label)

        self.add_widget(main_layout)
        self.on_mode_change(None, self.mode_spinner.text)

    def on_enter(self):
        app = App.get_running_app()
        display_names = [os.path.splitext(name)[0] for name in app.bill_files]
        self.target_spinner.values = display_names
        if app.current_bill:
            self.target_spinner.text = os.path.splitext(app.current_bill)[0]

    def on_mode_change(self, instance, mode):
        # 通过 opacity 和 disabled 增强反馈
        if mode == '导入成为新账单':
            self.target_layout.disabled = True
            self.target_layout.opacity = 0.3
            self.new_name_layout.disabled = False
            self.new_name_layout.opacity = 1
        elif mode == '与指定账单合并为新账单':
            self.target_layout.disabled = False
            self.target_layout.opacity = 1
            self.new_name_layout.disabled = False
            self.new_name_layout.opacity = 1
        elif mode == '并入指定账单':
            self.target_layout.disabled = False
            self.target_layout.opacity = 1
            self.new_name_layout.disabled = True
            self.new_name_layout.opacity = 0.3

    def choose_file(self, instance):
        try:
            filechooser.open_file(on_selection=self.on_file_selected, filters=["*.xlsx"])
        except Exception as e:
            # 如果 filechooser 不可用，提示手动输入路径
            self.show_manual_path_popup()

    def show_manual_path_popup(self):
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(Label(text='文件选择器不可用，请手动输入Excel文件路径：'))
        path_input = TextInput(hint_text='例如：C:\\myfile.xlsx', multiline=False)
        content.add_widget(path_input)
        btn_layout = BoxLayout(size_hint_y=0.3, spacing=dp(10))
        cancel_btn = Button(text='取消')
        ok_btn = Button(text='确定', background_color=COLOR_GREEN)
        btn_layout.add_widget(cancel_btn)
        btn_layout.add_widget(ok_btn)
        content.add_widget(btn_layout)
        popup = Popup(title='手动输入路径', content=content, size_hint=(0.9, 0.4))
        def on_ok(btn):
            path = path_input.text.strip()
            if path and os.path.exists(path):
                self.on_file_selected([path])
                popup.dismiss()
            else:
                path_input.text = ''
                path_input.hint_text = '文件不存在，请重新输入'
        ok_btn.bind(on_press=on_ok)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def on_file_selected(self, selection):
        if selection:
            self.selected_file = selection[0]
            self.file_path_label.text = os.path.basename(self.selected_file)
            self.file_path_label.color = (0,0,0,1)
        else:
            self.selected_file = None
            self.file_path_label.text = '未选择'
            self.file_path_label.color = COLOR_DARK_GRAY

    def _read_excel_to_records(self, filepath):
        """读取Excel文件，返回记录列表，并验证必要列，日期规范化"""
        wb = load_workbook(filepath, data_only=True)
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        headers = rows[0]
        # 期望的列名映射
        expected = ['时间', '收支', '金额', '事件', '相关方', '交易平台', '其他']
        col_indices = {}
        for i, h in enumerate(headers):
            if h:
                h_norm = h.strip()
                for exp in expected:
                    if h_norm == exp:
                        col_indices[exp] = i
                        break
        required = ['时间', '收支', '金额']
        for r in required:
            if r not in col_indices:
                raise ValueError(f"Excel中缺少必要列：{r}")

        records = []
        for row in rows[1:]:
            if not any(row):
                continue
            # 日期解析
            date_cell = row[col_indices['时间']] if col_indices['时间'] < len(row) else None
            date_str = parse_date_cell(date_cell)
            if not date_str:
                continue  # 无效日期则跳过该行

            type_cell = row[col_indices['收支']] if col_indices['收支'] < len(row) else ''
            type_str = str(type_cell).strip() if type_cell else ''
            if type_str not in ('收入', '支出'):
                continue

            amount_cell = row[col_indices['金额']] if col_indices['金额'] < len(row) else 0
            try:
                amount = float(amount_cell) if amount_cell not in (None, '') else 0.0
                if amount <= 0:
                    continue  # 金额必须为正
            except:
                continue

            def get_val(idx):
                if idx < len(row):
                    val = row[idx]
                    return str(val).strip() if val else ''
                return ''
            event = get_val(col_indices.get('事件', -1))
            party = get_val(col_indices.get('相关方', -1))
            platform = get_val(col_indices.get('交易平台', -1))
            other = get_val(col_indices.get('其他', -1))

            records.append({
                'date': date_str,
                'type': type_str,
                'amount': amount,
                'event': event,
                'party': party,
                'platform': platform,
                'other': other
            })
        return records

    def _insert_records_with_transaction(self, conn, records):
        """在事务中插入多条记录"""
        c = conn.cursor()
        c.execute("BEGIN")
        try:
            for rec in records:
                c.execute('''
                    INSERT INTO records (date, type, amount, event, party, platform, other)
                    VALUES (?, ?, ?, ?, ?, ?, ?)
                ''', (rec['date'], rec['type'], rec['amount'], rec['event'],
                      rec['party'], rec['platform'], rec['other']))
            conn.commit()
        except Exception as e:
            conn.rollback()
            raise e

    def do_import(self, instance):
        mode = self.mode_spinner.text
        app = App.get_running_app()

        if not self.selected_file:
            self.tip_label.text = '请先选择外部 Excel 文件'
            return

        try:
            records = self._read_excel_to_records(self.selected_file)
        except Exception as e:
            logging.exception("读取Excel失败")
            self.tip_label.text = f'读取外部文件失败：{str(e)}'
            return

        if not records:
            self.tip_label.text = 'Excel中没有有效数据'
            return

        try:
            if mode == '导入成为新账单':
                raw_name = self.new_name_input.text.strip()
                if not raw_name:
                    self.tip_label.text = '请输入新账单名称'
                    return
                safe_name = sanitize_filename(raw_name)
                filename = safe_name + '.db'
                filepath = os.path.join(app.bills_dir, filename)
                if os.path.exists(filepath):
                    self.tip_label.text = '该账单名称已存在，请换一个'
                    return
                init_database(filepath)
                conn = sqlite3.connect(filepath)
                self._insert_records_with_transaction(conn, records)
                conn.close()
                app.refresh_bill_list()
                app.current_bill = filename
                self.tip_label.text = f'导入成功，已切换到新账单“{safe_name}”'

            elif mode == '与指定账单合并为新账单':
                target_display = self.target_spinner.text
                if not target_display:
                    self.tip_label.text = '请选择目标账单'
                    return
                raw_new = self.new_name_input.text.strip()
                if not raw_new:
                    self.tip_label.text = '请输入新账单名称'
                    return
                safe_new = sanitize_filename(raw_new)
                target_file = None
                for f in app.bill_files:
                    if os.path.splitext(f)[0] == target_display:
                        target_file = f
                        break
                if not target_file:
                    self.tip_label.text = '目标账单不存在'
                    return

                # 读取目标数据库所有记录
                target_path = os.path.join(app.bills_dir, target_file)
                conn_src = sqlite3.connect(target_path)
                c_src = conn_src.cursor()
                c_src.execute("SELECT date, type, amount, event, party, platform, other FROM records")
                src_rows = c_src.fetchall()
                conn_src.close()

                # 创建新数据库
                new_filename = safe_new + '.db'
                new_path = os.path.join(app.bills_dir, new_filename)
                if os.path.exists(new_path):
                    self.tip_label.text = '该账单名称已存在，请换一个'
                    return
                init_database(new_path)
                conn_new = sqlite3.connect(new_path)
                # 开启事务
                conn_new.execute("BEGIN")
                try:
                    c_new = conn_new.cursor()
                    for row in src_rows:
                        c_new.execute('''
                            INSERT INTO records (date, type, amount, event, party, platform, other)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        ''', row)
                    for rec in records:
                        c_new.execute('''
                            INSERT INTO records (date, type, amount, event, party, platform, other)
                            VALUES (?, ?, ?, ?, ?, ?, ?)
                        ''', (rec['date'], rec['type'], rec['amount'], rec['event'],
                              rec['party'], rec['platform'], rec['other']))
                    conn_new.commit()
                except Exception as e:
                    conn_new.rollback()
                    raise e
                finally:
                    conn_new.close()
                app.refresh_bill_list()
                app.current_bill = new_filename
                self.tip_label.text = f'合并成功，已切换到新账单“{safe_new}”'

            elif mode == '并入指定账单':
                target_display = self.target_spinner.text
                if not target_display:
                    self.tip_label.text = '请选择目标账单'
                    return
                target_file = None
                for f in app.bill_files:
                    if os.path.splitext(f)[0] == target_display:
                        target_file = f
                        break
                if not target_file:
                    self.tip_label.text = '目标账单不存在'
                    return

                target_path = os.path.join(app.bills_dir, target_file)
                conn = sqlite3.connect(target_path)
                self._insert_records_with_transaction(conn, records)
                conn.close()
                self.tip_label.text = f'并入成功，账单“{target_display}”已更新'

            # 清空选择
            self.selected_file = None
            self.file_path_label.text = '未选择'
            self.file_path_label.color = COLOR_DARK_GRAY
            self.tip_label.color = (0, 0.7, 0, 1)

        except Exception as e:
            logging.exception("导入失败")
            self.tip_label.text = f'导入失败：{str(e)}'
            self.tip_label.color = COLOR_RED

    def go_back(self, instance):
        self.manager.current = 'main'

# ---------- 导出页面 ----------
class ExportScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        main_layout = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(10))

        title = Label(text='导出账单副本', size_hint_y=0.1, font_size=22, halign='center')
        main_layout.add_widget(title)

        # 查询类型选择
        type_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        type_layout.add_widget(Label(text='查询类型：', size_hint_x=0.25, font_size=dp(14)))
        self.type_spinner = Spinner(
            text='天',
            values=('天', '月', '年', '时间段'),
            size_hint_x=0.75,
            font_size=dp(12)
        )
        self.type_spinner.bind(text=self.on_type_change)
        type_layout.add_widget(self.type_spinner)
        main_layout.add_widget(type_layout)

        self.date_selector = DateSelector(query_type='天')
        main_layout.add_widget(self.date_selector)

        # 收支类型选择
        income_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        income_layout.add_widget(Label(text='收支类型：', size_hint_x=0.25, font_size=dp(14)))
        self.income_spinner = Spinner(
            text='全部',
            values=('全部', '收入', '支出'),
            size_hint_x=0.75,
            font_size=dp(12)
        )
        income_layout.add_widget(self.income_spinner)
        main_layout.add_widget(income_layout)

        # 导出格式选择（新增）
        format_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        format_layout.add_widget(Label(text='导出格式：', size_hint_x=0.25, font_size=dp(14)))
        self.format_spinner = Spinner(
            text='Excel (.xlsx)',
            values=('Excel (.xlsx)', 'CSV (.csv)'),
            size_hint_x=0.75,
            font_size=dp(12)
        )
        format_layout.add_widget(self.format_spinner)
        main_layout.add_widget(format_layout)

        # 目标文件夹选择
        folder_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        folder_layout.add_widget(Label(text='目标文件夹：', size_hint_x=0.25, font_size=dp(14)))
        self.folder_label = Label(
            text='未选择',
            size_hint_x=0.5,
            font_size=dp(12),
            color=COLOR_DARK_GRAY,
            shorten=True,
            shorten_from='right'
        )
        folder_layout.add_widget(self.folder_label)
        browse_btn = Button(text='浏览', size_hint_x=0.25, font_size=dp(12))
        browse_btn.bind(on_press=self.choose_folder)
        folder_layout.add_widget(browse_btn)
        main_layout.add_widget(folder_layout)

        # 文件名输入
        name_layout = BoxLayout(size_hint_y=0.08, spacing=dp(5))
        name_layout.add_widget(Label(text='文件名：', size_hint_x=0.25, font_size=dp(14)))
        self.filename_input = TextInput(
            text='',
            size_hint_x=0.75,
            font_size=dp(12),
            multiline=False,
            hint_text='例如：我的账单_导出'
        )
        name_layout.add_widget(self.filename_input)
        main_layout.add_widget(name_layout)

        main_layout.add_widget(BoxLayout(size_hint_y=0.05))

        # 导出按钮
        export_btn = Button(
            text='执行导出',
            size_hint_y=0.1,
            background_color=COLOR_GREEN,
            font_size=dp(16)
        )
        export_btn.bind(on_press=self.do_export)
        main_layout.add_widget(export_btn)

        back_btn = Button(
            text='返回主界面',
            size_hint_y=0.08,
            background_color=COLOR_GRAY,
            font_size=dp(14)
        )
        back_btn.bind(on_press=self.go_back)
        main_layout.add_widget(back_btn)

        self.tip_label = Label(
            text='',
            size_hint_y=0.05,
            halign='center',
            color=COLOR_RED,
            font_size=dp(12)
        )
        main_layout.add_widget(self.tip_label)

        self.add_widget(main_layout)

        self.selected_folder = None

    def on_enter(self):
        app = App.get_running_app()
        if app.current_bill:
            default_name = os.path.splitext(app.current_bill)[0]
            self.filename_input.text = default_name
        else:
            self.filename_input.text = ''

    def on_type_change(self, instance, text):
        self.date_selector.set_type(text)

    def choose_folder(self, instance):
        try:
            filechooser.choose_dir(on_selection=self.on_folder_selected)
        except Exception as e:
            self.show_manual_folder_popup()

    def show_manual_folder_popup(self):
        content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
        content.add_widget(Label(text='文件夹选择器不可用，请手动输入目标文件夹路径：'))
        path_input = TextInput(hint_text='例如：C:\\myfolder', multiline=False)
        content.add_widget(path_input)
        btn_layout = BoxLayout(size_hint_y=0.3, spacing=dp(10))
        cancel_btn = Button(text='取消')
        ok_btn = Button(text='确定', background_color=COLOR_GREEN)
        btn_layout.add_widget(cancel_btn)
        btn_layout.add_widget(ok_btn)
        content.add_widget(btn_layout)
        popup = Popup(title='手动输入路径', content=content, size_hint=(0.9, 0.4))
        def on_ok(btn):
            path = path_input.text.strip()
            if path and os.path.isdir(path):
                self.on_folder_selected([path])
                popup.dismiss()
            else:
                path_input.text = ''
                path_input.hint_text = '文件夹不存在，请重新输入'
        ok_btn.bind(on_press=on_ok)
        cancel_btn.bind(on_press=popup.dismiss)
        popup.open()

    def on_folder_selected(self, selection):
        if selection:
            self.selected_folder = selection[0]
            self.folder_label.text = self.selected_folder
            self.folder_label.color = (0, 0, 0, 1)
        else:
            self.selected_folder = None
            self.folder_label.text = '未选择'
            self.folder_label.color = COLOR_DARK_GRAY

    def do_export(self, instance):
        if not self.selected_folder:
            self.tip_label.text = '请先选择目标文件夹'
            self.tip_label.color = COLOR_RED
            return

        raw_filename = self.filename_input.text.strip()
        if not raw_filename:
            self.tip_label.text = '请输入文件名'
            self.tip_label.color = COLOR_RED
            return

        safe_filename = sanitize_filename(raw_filename)
        export_format = self.format_spinner.text
        if export_format == 'Excel (.xlsx)':
            if not safe_filename.lower().endswith('.xlsx'):
                safe_filename += '.xlsx'
        else:  # CSV
            if not safe_filename.lower().endswith('.csv'):
                safe_filename += '.csv'

        full_path = os.path.join(self.selected_folder, safe_filename)

        # 检查文件是否已存在
        if os.path.exists(full_path):
            # 弹出确认覆盖对话框
            content = BoxLayout(orientation='vertical', spacing=dp(10), padding=dp(10))
            content.add_widget(Label(text=f'文件“{safe_filename}”已存在，是否覆盖？'))
            btn_layout = BoxLayout(size_hint_y=0.3, spacing=dp(10))
            cancel_btn = Button(text='取消')
            ok_btn = Button(text='覆盖', background_color=COLOR_RED)
            btn_layout.add_widget(cancel_btn)
            btn_layout.add_widget(ok_btn)
            content.add_widget(btn_layout)
            popup = Popup(title='确认覆盖', content=content, size_hint=(0.8, 0.3))

            def on_ok(btn):
                popup.dismiss()
                self._perform_export(full_path, export_format)
            ok_btn.bind(on_press=on_ok)
            cancel_btn.bind(on_press=popup.dismiss)
            popup.open()
        else:
            self._perform_export(full_path, export_format)

    def _perform_export(self, full_path, export_format):
        """实际执行导出（文件已确保可写入）"""
        query_type = self.type_spinner.text
        income_filter = self.income_spinner.text

        try:
            start_date, end_date = self.date_selector.get_date_range()
        except Exception as e:
            self.tip_label.text = f"日期错误：{e}"
            return

        app = App.get_running_app()
        db_path = app.get_current_db_path()
        if not os.path.exists(db_path):
            self.tip_label.text = "当前账单数据库不存在，无数据可导出"
            return

        conn = sqlite3.connect(db_path)
        c = conn.cursor()
        try:
            sql = "SELECT date, type, amount, event, party, platform, other FROM records"
            conditions = []
            params = []

            if query_type in ('天', '时间段'):
                conditions.append("date BETWEEN ? AND ?")
                params.extend([start_date, end_date])
            elif query_type in ('月', '年'):
                conditions.append("date >= ? AND date < ?")
                params.extend([start_date, end_date])

            if income_filter == '收入':
                conditions.append("type = '收入'")
            elif income_filter == '支出':
                conditions.append("type = '支出'")

            if conditions:
                sql += " WHERE " + " AND ".join(conditions)

            sql += " ORDER BY date"

            c.execute(sql, params)
            rows = c.fetchall()

            if not rows:
                self.tip_label.text = "没有符合条件的记录，无法导出"
                return

            if export_format == 'Excel (.xlsx)':
                wb = Workbook()
                ws = wb.active
                ws.title = '导出结果'
                headers = ['时间', '收支', '金额', '事件', '相关方', '交易平台', '其他']
                ws.append(headers)
                for row in rows:
                    ws.append(row)
                wb.save(full_path)
            else:  # CSV
                import csv
                with open(full_path, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.writer(f)
                    writer.writerow(['时间', '收支', '金额', '事件', '相关方', '交易平台', '其他'])
                    writer.writerows(rows)

            self.tip_label.text = f"导出成功！\n文件保存至：{full_path}"
            self.tip_label.color = (0, 0.7, 0, 1)

        except Exception as e:
            logging.exception("导出失败")
            self.tip_label.text = f"导出失败：{str(e)}"
            self.tip_label.color = COLOR_RED
        finally:
            conn.close()

    def go_back(self, instance):
        self.manager.current = 'main'

# ---------- 应用入口 ----------
class MyApp(App):
    def build(self):
        self.bills_dir = os.path.join(os.path.dirname(__file__), 'bills')
        os.makedirs(self.bills_dir, exist_ok=True)

        self.bill_files = []          # 存储 .db 文件名列表
        self.current_bill = None
        self.refresh_bill_list()

        sm = ScreenManager()
        sm.add_widget(MainScreen(name='main'))
        sm.add_widget(ViewEventScreen(name='view_event'))
        sm.add_widget(RecordScreen(name='record'))
        sm.add_widget(ImportScreen(name='import'))
        sm.add_widget(ExportScreen(name='export'))
        return sm

    def refresh_bill_list(self):
        self.bill_files = [f for f in os.listdir(self.bills_dir) if f.endswith('.db')]
        if not self.bill_files:
            default_name = '我的账单.db'
            default_path = os.path.join(self.bills_dir, default_name)
            init_database(default_path)
            self.bill_files = [default_name]
        if self.current_bill is None or self.current_bill not in self.bill_files:
            self.current_bill = self.bill_files[0]

    def create_new_bill(self, raw_name):
        if not raw_name.strip():
            return False, "文件名不能为空"
        safe_name = sanitize_filename(raw_name.strip())
        filename = safe_name + '.db'
        filepath = os.path.join(self.bills_dir, filename)
        if os.path.exists(filepath):
            return False, "文件已存在"
        init_database(filepath)
        self.refresh_bill_list()
        self.current_bill = filename
        return True, "创建成功"

    def rename_current_bill(self, raw_name):
        if not self.current_bill:
            return False, "没有选中任何账单"
        new_name = sanitize_filename(raw_name.strip())
        if not new_name:
            return False, "名称不能为空"
        if new_name.lower().endswith('.db'):
            new_name = new_name[:-3]
        new_filename = new_name + '.db'
        if new_filename == self.current_bill:
            return False, "新名称与当前名称相同"
        new_path = os.path.join(self.bills_dir, new_filename)
        if os.path.exists(new_path):
            return False, "该名称已存在，请换一个"
        old_path = os.path.join(self.bills_dir, self.current_bill)
        try:
            os.rename(old_path, new_path)
        except Exception as e:
            logging.exception("重命名失败")
            return False, f"重命名失败：{str(e)}"
        self.current_bill = new_filename
        self.refresh_bill_list()
        return True, "重命名成功"

    def delete_current_bill(self):
        if not self.current_bill:
            return False, "没有选中任何账单"
        if len(self.bill_files) == 1:
            return False, "至少需要保留一个账单，无法删除最后一个"
        filepath = os.path.join(self.bills_dir, self.current_bill)
        try:
            os.remove(filepath)
        except Exception as e:
            logging.exception("删除失败")
            return False, f"删除失败：{str(e)}"
        self.refresh_bill_list()
        self.current_bill = self.bill_files[0]
        return True, "删除成功"

    def get_current_db_path(self):
        return os.path.join(self.bills_dir, self.current_bill)

if __name__ == '__main__':
    MyApp().run()
