"""
员工档案管理系统 - 主程序
基于 Python + PySimpleGUI + SQLite
"""

import PySimpleGUI as sg
from database import EmployeeDatabase
import csv
from datetime import datetime

# 设置主题
sg.theme('LightBlue2')

class EmployeeManagementApp:
    def __init__(self):
        self.db = EmployeeDatabase()
        self.employees = []
        self.search_results = []
        
    def get_main_window(self):
        """创建主窗口"""
        
        # 定义表格列
        headings = ['工号', '姓名', '性别', '出生日期', '部门', 
                   '职位', '入职日期', '电话', '邮箱', '学历']
        
        # 定义布局
        layout = [
            [sg.Text('员工档案管理系统', font=('Arial', 18, 'bold'))],
            [sg.HorizontalSeparator()],
            
            # 搜索栏
            [
                sg.Text('搜索：', size=(8, 1)),
                sg.InputText(size=(30, 1), key='-SEARCH-'),
                sg.Button('搜索', size=(10, 1)),
                sg.Button('重置', size=(10, 1)),
                sg.Button('刷新', size=(10, 1)),
            ],
            
            # 表格
            [
                sg.Table(
                    values=[],
                    headings=headings,
                    max_col_widths=(12, 12, 8, 15, 12, 12, 15, 12, 20, 10),
                    auto_size_columns=False,
                    display_row_numbers=False,
                    num_rows=15,
                    key='-TABLE-',
                    enable_events=True,
                    vertical_scroll_only=False,
                )
            ],
            
            # 按钮栏
            [
                sg.Button('添加员工', size=(12, 1)),
                sg.Button('编辑选中', size=(12, 1)),
                sg.Button('删除选中', size=(12, 1)),
                sg.Button('导出 CSV', size=(12, 1)),
                sg.Button('统计报告', size=(12, 1)),
                sg.Button('退出', size=(12, 1)),
            ],
            
            # 状态栏
            [sg.Text('', key='-STATUS-', size=(80, 1))]
        ]
        
        return sg.Window('员工档案管理系统', layout, finalize=True)
    
    def get_employee_window(self, employee=None):
        """创建员工编辑窗口"""
        
        if employee:
            # 编辑模式
            title = '编辑员工'
            default_values = employee
        else:
            # 添加模式
            title = '添加新员工'
            default_values = ['', '', '男', '', '', '技术部', '', '', '', '', '']
        
        layout = [
            [sg.Text(title, font=('Arial', 14, 'bold'))],
            [sg.HorizontalSeparator()],
            
            [sg.Text('工号：', size=(12, 1)), 
             sg.InputText(default_values[0], size=(30, 1), key='-ID-', 
                         disabled=bool(employee))],
            
            [sg.Text('姓名：', size=(12, 1)), 
             sg.InputText(default_values[1], size=(30, 1), key='-NAME-')],
            
            [sg.Text('性别：', size=(12, 1)), 
             sg.Combo(['男', '女'], default_value=default_values[2], 
                     size=(28, 1), key='-GENDER-')],
            
            [sg.Text('出生日期：', size=(12, 1)), 
             sg.InputText(default_values[3], size=(30, 1), key='-BIRTH-',
                         tooltip='格式: YYYY-MM-DD')],
            
            [sg.Text('身份证号：', size=(12, 1)), 
             sg.InputText(default_values[4], size=(30, 1), key='-ID-CARD-')],
            
            [sg.Text('部门：', size=(12, 1)), 
             sg.Combo(['技术部', '销售部', '人力资源部', '财务部', '市场部'], 
                     default_value=default_values[5], size=(28, 1), key='-DEPT-')],
            
            [sg.Text('职位：', size=(12, 1)), 
             sg.InputText(default_values[6], size=(30, 1), key='-POSITION-')],
            
            [sg.Text('入职日期：', size=(12, 1)), 
             sg.InputText(default_values[7], size=(30, 1), key='-JOIN-DATE-',
                         tooltip='格式: YYYY-MM-DD')],
            
            [sg.Text('电话：', size=(12, 1)), 
             sg.InputText(default_values[8], size=(30, 1), key='-PHONE-')],
            
            [sg.Text('邮箱：', size=(12, 1)), 
             sg.InputText(default_values[9], size=(30, 1), key='-EMAIL-')],
            
            [sg.Text('学历：', size=(12, 1)), 
             sg.Combo(['高中', '大专', '本科', '研究生'], 
                     default_value=default_values[10], size=(28, 1), key='-EDUCATION-')],
            
            [sg.HorizontalSeparator()],
            
            [
                sg.Button('保存', size=(12, 1)),
                sg.Button('取消', size=(12, 1)),
            ]
        ]
        
        return sg.Window(title, layout, finalize=True)
    
    def refresh_table(self, employees=None):
        """刷新表格数据"""
        if employees is None:
            employees = self.db.get_all_employees()
        
        table_data = [list(emp) for emp in employees]
        return table_data
    
    def run(self):
        """运行应用"""
        window = self.get_main_window()
        
        # 初始加载数据
        self.employees = self.db.get_all_employees()
        table_data = self.refresh_table(self.employees)
        window['-TABLE-'].update(values=table_data)
        window['-STATUS-'].update(f'已加载 {len(self.employees)} 条记录')
        
        while True:
            event, values = window.read()
            
            if event == sg.WINDOW_CLOSED or event == '退出':
                break
            
            # 搜索功能
            elif event == '搜索':
                search_text = values['-SEARCH-'].strip()
                if not search_text:
                    sg.popup_error('请输入搜索内容！')
                    continue
                
                self.search_results = self.db.search_employees(search_text)
                if self.search_results:
                    table_data = self.refresh_table(self.search_results)
                    window['-TABLE-'].update(values=table_data)
                    window['-STATUS-'].update(f'搜索结果：找到 {len(self.search_results)} 条记录')
                else:
                    sg.popup_info('未找到匹配的员工记录')
                    window['-STATUS-'].update('搜索无结果')
            
            # 重置搜索
            elif event == '重置':
                window['-SEARCH-'].update('')
                self.employees = self.db.get_all_employees()
                table_data = self.refresh_table(self.employees)
                window['-TABLE-'].update(values=table_data)
                window['-STATUS-'].update(f'已加载 {len(self.employees)} 条记录')
            
            # 刷新数据
            elif event == '刷新':
                self.employees = self.db.get_all_employees()
                table_data = self.refresh_table(self.employees)
                window['-TABLE-'].update(values=table_data)
                window['-STATUS-'].update(f'已刷新，共 {len(self.employees)} 条记录')
            
            # 添加员工
            elif event == '添加员工':
                sub_window = self.get_employee_window()
                while True:
                    sub_event, sub_values = sub_window.read()
                    
                    if sub_event == sg.WINDOW_CLOSED or sub_event == '取消':
                        sub_window.close()
                        break
                    
                    if sub_event == '保存':
                        # 验证输入
                        if not sub_values['-ID-'] or not sub_values['-NAME-']:
                            sg.popup_error('工号和姓名不能为空！')
                            continue
                        
                        employee = (
                            sub_values['-ID-'],
                            sub_values['-NAME-'],
                            sub_values['-GENDER-'],
                            sub_values['-BIRTH-'],
                            sub_values['-ID-CARD-'],
                            sub_values['-DEPT-'],
                            sub_values['-POSITION-'],
                            sub_values['-JOIN-DATE-'],
                            sub_values['-PHONE-'],
                            sub_values['-EMAIL-'],
                            sub_values['-EDUCATION-'],
                            ''
                        )
                        
                        if self.db.add_employee(employee):
                            sg.popup_ok('员工添加成功！')
                            sub_window.close()
                            
                            # 刷新主窗口
                            self.employees = self.db.get_all_employees()
                            table_data = self.refresh_table(self.employees)
                            window['-TABLE-'].update(values=table_data)
                            window['-STATUS-'].update(f'已加载 {len(self.employees)} 条记录')
                            break
                        else:
                            sg.popup_error('添加失败！该工号已存在。')
            
            # 编辑员工
            elif event == '编辑选中':
                selected = values['-TABLE-']
                if not selected:
                    sg.popup_error('请先选择要编辑的员工！')
                    continue
                
                row_idx = selected[0]
                employee = self.employees[row_idx] if not values['-SEARCH-'] else self.search_results[row_idx]
                
                sub_window = self.get_employee_window(employee)
                while True:
                    sub_event, sub_values = sub_window.read()
                    
                    if sub_event == sg.WINDOW_CLOSED or sub_event == '取消':
                        sub_window.close()
                        break
                    
                    if sub_event == '保存':
                        updated_employee = (
                            employee[0],  # 工号不变
                            sub_values['-NAME-'],
                            sub_values['-GENDER-'],
                            sub_values['-BIRTH-'],
                            sub_values['-ID-CARD-'],
                            sub_values['-DEPT-'],
                            sub_values['-POSITION-'],
                            sub_values['-JOIN-DATE-'],
                            sub_values['-PHONE-'],
                            sub_values['-EMAIL-'],
                            sub_values['-EDUCATION-'],
                            ''
                        )
                        
                        if self.db.update_employee(employee[0], updated_employee):
                            sg.popup_ok('员工信息更新成功！')
                            sub_window.close()
                            
                            # 刷新主窗口
                            self.employees = self.db.get_all_employees()
                            table_data = self.refresh_table(self.employees)
                            window['-TABLE-'].update(values=table_data)
                            break
                        else:
                            sg.popup_error('更新失败！')
            
            # 删除员工
            elif event == '删除选中':
                selected = values['-TABLE-']
                if not selected:
                    sg.popup_error('请先选择要删除的员工！')
                    continue
                
                row_idx = selected[0]
                employee = self.employees[row_idx] if not values['-SEARCH-'] else self.search_results[row_idx]
                
                if sg.popup_yes_no(f'确定删除员工 {employee[1]} 吗？') == 'Yes':
                    if self.db.delete_employee(employee[0]):
                        sg.popup_ok('员工删除成功！')
                        
                        # 刷新主窗口
                        self.employees = self.db.get_all_employees()
                        table_data = self.refresh_table(self.employees)
                        window['-TABLE-'].update(values=table_data)
                        window['-STATUS-'].update(f'已加载 {len(self.employees)} 条记录')
                    else:
                        sg.popup_error('删除失败！')
            
            # 导出 CSV
            elif event == '导出 CSV':
                filename = sg.popup_get_file('保存为', save_as=True, 
                                            file_types=(('CSV 文件', '*.csv'),))
                if filename:
                    try:
                        with open(filename, 'w', newline='', encoding='utf-8') as f:
                            writer = csv.writer(f)
                            writer.writerow(['工号', '姓名', '性别', '出生日期', '身份证号', 
                                           '部门', '职位', '入职日期', '电话', '邮箱', '学历', '备注'])
                            for emp in self.employees:
                                writer.writerow(emp)
                        
                        sg.popup_ok(f'数据已导出到 {filename}')
                        window['-STATUS-'].update(f'已导出 {len(self.employees)} 条记录')
                    except Exception as e:
                        sg.popup_error(f'导出失败: {str(e)}')
            
            # 统计报告
            elif event == '统计报告':
                stats = self.db.get_statistics()
                
                report = '员工统计报告\n' + '='*40 + '\n\n'
                report += f'总员工数: {stats["total"]}\n\n'
                report += '按部门统计:\n'
                report += '-'*40 + '\n'
                
                for dept, count in stats['by_department'].items():
                    report += f'{dept}: {count} 人\n'
                
                sg.popup_scrolled(report, title='统计报告', size=(50, 20))
        
        window.close()

if __name__ == '__main__':
    app = EmployeeManagementApp()
    app.run()
