#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
表單驗證模組
提供各種表單欄位的驗證功能
"""

import re
import datetime
from typing import Dict, List, Tuple, Any, Optional


class FormValidator:
    """表單驗證器類別"""
    
    def __init__(self):
        # 預定義的驗證規則
        self.validation_rules = {
            'required': self.validate_required,
            'email': self.validate_email,
            'phone': self.validate_phone,
            'id_number': self.validate_id_number,
            'date': self.validate_date,
            'time': self.validate_time,
            'number': self.validate_number,
            'text_length': self.validate_text_length,
            'choice': self.validate_choice,
            'employee_id': self.validate_employee_id
        }
        
        # 錯誤訊息模板
        self.error_messages = {
            'required': "此欄位為必填",
            'email': "請輸入有效的電子郵件格式",
            'phone': "請輸入有效的電話號碼格式",
            'id_number': "請輸入有效的身分證字號格式",
            'date': "請輸入有效的日期格式 (YYYY-MM-DD)",
            'time': "請輸入有效的時間格式 (HH:MM)",
            'number': "請輸入有效的數字",
            'text_length': "文字長度超出限制",
            'choice': "請選擇有效的選項",
            'employee_id': "請輸入有效的員工編號格式"
        }
    
    def validate_field(self, field_name: str, value: Any, rules: List[Dict]) -> Tuple[bool, List[str]]:
        """
        驗證單一欄位
        
        Args:
            field_name: 欄位名稱
            value: 欄位值
            rules: 驗證規則列表
        
        Returns:
            Tuple[bool, List[str]]: (是否通過驗證, 錯誤訊息列表)
        """
        errors = []
        
        # 轉換為字串並去除空白
        str_value = str(value).strip() if value is not None else ""
        
        for rule in rules:
            rule_type = rule.get('type')
            rule_params = rule.get('params', {})
            custom_message = rule.get('message')
            
            if rule_type in self.validation_rules:
                is_valid, error_msg = self.validation_rules[rule_type](
                    str_value, rule_params
                )
                
                if not is_valid:
                    error_message = custom_message or self.error_messages.get(rule_type, "驗證失敗")
                    errors.append(f"{field_name}: {error_message}")
        
        return len(errors) == 0, errors
    
    def validate_form(self, form_data: Dict[str, Any], validation_schema: Dict[str, List[Dict]]) -> Tuple[bool, Dict[str, List[str]]]:
        """
        驗證整個表單
        
        Args:
            form_data: 表單資料
            validation_schema: 驗證規則架構
        
        Returns:
            Tuple[bool, Dict[str, List[str]]]: (是否通過驗證, 錯誤訊息字典)
        """
        all_errors = {}
        is_form_valid = True
        
        for field_name, rules in validation_schema.items():
            field_value = form_data.get(field_name, "")
            is_field_valid, field_errors = self.validate_field(field_name, field_value, rules)
            
            if not is_field_valid:
                all_errors[field_name] = field_errors
                is_form_valid = False
        
        return is_form_valid, all_errors
    
    def validate_required(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證必填欄位"""
        if not value or value.strip() == "":
            return False, "此欄位為必填"
        return True, ""
    
    def validate_email(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證電子郵件格式"""
        if not value:
            return True, ""  # 空值由required規則處理
        
        email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}
        if not re.match(email_pattern, value):
            return False, "請輸入有效的電子郵件格式"
        return True, ""
    
    def validate_phone(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證電話號碼格式"""
        if not value:
            return True, ""
        
        # 移除所有非數字字符
        phone_digits = re.sub(r'\D', '', value)
        
        # 台灣手機號碼格式驗證
        mobile_pattern = r'^09\d{8}
        # 台灣市話格式驗證
        landline_pattern = r'^0[2-8]\d{7,8}
        
        if re.match(mobile_pattern, phone_digits) or re.match(landline_pattern, phone_digits):
            return True, ""
        
        return False, "請輸入有效的電話號碼格式 (如: 0912345678 或 02-12345678)"
    
    def validate_id_number(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證台灣身分證字號格式"""
        if not value:
            return True, ""
        
        # 台灣身分證字號格式: 1個英文字母 + 9個數字
        if len(value) != 10:
            return False, "身分證字號必須為10碼"
        
        if not re.match(r'^[A-Z][12]\d{8}, value.upper()):
            return False, "請輸入正確的身分證字號格式"
        
        # 驗證檢查碼
        if not self._validate_taiwan_id_checksum(value.upper()):
            return False, "身分證字號檢查碼錯誤"
        
        return True, ""
    
    def _validate_taiwan_id_checksum(self, id_number: str) -> bool:
        """驗證台灣身分證字號檢查碼"""
        # 英文字母對應數字表
        letter_map = {
            'A': 10, 'B': 11, 'C': 12, 'D': 13, 'E': 14, 'F': 15, 'G': 16,
            'H': 17, 'I': 34, 'J': 18, 'K': 19, 'L': 20, 'M': 21, 'N': 22,
            'O': 35, 'P': 23, 'Q': 24, 'R': 25, 'S': 26, 'T': 27, 'U': 28,
            'V': 29, 'W': 32, 'X': 30, 'Y': 31, 'Z': 33
        }
        
        first_letter = id_number[0]
        if first_letter not in letter_map:
            return False
        
        # 計算檢查碼
        letter_value = letter_map[first_letter]
        total = (letter_value // 10) + (letter_value % 10) * 9
        
        for i in range(1, 9):
            total += int(id_number[i]) * (9 - i)
        
        checksum = (10 - (total % 10)) % 10
        return checksum == int(id_number[9])
    
    def validate_date(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證日期格式"""
        if not value:
            return True, ""
        
        # 支援的日期格式
        date_formats = [
            '%Y-%m-%d',      # 2024-01-01
            '%Y/%m/%d',      # 2024/01/01
            '%Y.%m.%d',      # 2024.01.01
            '%d/%m/%Y',      # 01/01/2024
            '%d-%m-%Y',      # 01-01-2024
        ]
        
        for date_format in date_formats:
            try:
                parsed_date = datetime.datetime.strptime(value, date_format)
                
                # 檢查日期範圍
                min_year = params.get('min_year', 1900)
                max_year = params.get('max_year', 2100)
                
                if not (min_year <= parsed_date.year <= max_year):
                    return False, f"日期年份必須在 {min_year} 到 {max_year} 之間"
                
                return True, ""
            except ValueError:
                continue
        
        return False, "請輸入有效的日期格式 (如: 2024-01-01)"
    
    def validate_time(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證時間格式"""
        if not value:
            return True, ""
        
        # 支援的時間格式
        time_formats = [
            '%H:%M',         # 14:30
            '%H:%M:%S',      # 14:30:00
            '%I:%M %p',      # 02:30 PM
        ]
        
        for time_format in time_formats:
            try:
                datetime.datetime.strptime(value, time_format)
                return True, ""
            except ValueError:
                continue
        
        return False, "請輸入有效的時間格式 (如: 14:30)"
    
    def validate_number(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證數字格式"""
        if not value:
            return True, ""
        
        try:
            num_value = float(value)
            
            # 檢查最小值
            if 'min_value' in params and num_value < params['min_value']:
                return False, f"數值不能小於 {params['min_value']}"
            
            # 檢查最大值
            if 'max_value' in params and num_value > params['max_value']:
                return False, f"數值不能大於 {params['max_value']}"
            
            # 檢查是否為整數
            if params.get('integer_only', False) and not num_value.is_integer():
                return False, "請輸入整數"
            
            return True, ""
        except ValueError:
            return False, "請輸入有效的數字"
    
    def validate_text_length(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證文字長度"""
        if not value:
            return True, ""
        
        min_length = params.get('min_length', 0)
        max_length = params.get('max_length', float('inf'))
        
        if len(value) < min_length:
            return False, f"文字長度至少需要 {min_length} 個字元"
        
        if len(value) > max_length:
            return False, f"文字長度不能超過 {max_length} 個字元"
        
        return True, ""
    
    def validate_choice(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證選擇項目"""
        if not value:
            return True, ""
        
        valid_choices = params.get('choices', [])
        if value not in valid_choices:
            return False, f"請選擇以下其中一項: {', '.join(valid_choices)}"
        
        return True, ""
    
    def validate_employee_id(self, value: str, params: Dict) -> Tuple[bool, str]:
        """驗證員工編號格式"""
        if not value:
            return True, ""
        
        # 預設員工編號格式: EMP + 3位數字 (如: EMP001)
        pattern = params.get('pattern', r'^EMP\d{3})
        
        if not re.match(pattern, value):
            return False, "請輸入正確的員工編號格式 (如: EMP001)"
        
        return True, ""


class EmployeeFormValidator(FormValidator):
    """員工表單專用驗證器"""
    
    def __init__(self):
        super().__init__()
        
        # 員工表單驗證架構
        self.basic_info_schema = {
            'employee_id': [
                {'type': 'required'},
                {'type': 'employee_id'}
            ],
            'name': [
                {'type': 'required'},
                {'type': 'text_length', 'params': {'min_length': 2, 'max_length': 50}}
            ],
            'id_number': [
                {'type': 'required'},
                {'type': 'id_number'}
            ],
            'gender': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['男', '女']}}
            ],
            'birth_date': [
                {'type': 'required'},
                {'type': 'date', 'params': {'min_year': 1930, 'max_year': 2010}}
            ],
            'phone': [
                {'type': 'required'},
                {'type': 'phone'}
            ],
            'email': [
                {'type': 'email'}
            ],
            'department': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['人事部', '財務部', '業務部', '技術部', '行政部']}}
            ],
            'position': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['經理', '副理', '主任', '專員', '助理']}}
            ],
            'hire_date': [
                {'type': 'required'},
                {'type': 'date', 'params': {'min_year': 2000}}
            ],
            'work_location': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['台北', '台中', '高雄', '新竹']}}
            ],
            'employment_type': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['正職', '約聘', '派遣', '兼職']}}
            ]
        }
        
        self.performance_schema = {
            'year': [
                {'type': 'required'},
                {'type': 'number', 'params': {'min_value': 2020, 'max_value': 2030, 'integer_only': True}}
            ],
            'annual_rating': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['優', '良', '可', '差']}}
            ],
            'first_half': [
                {'type': 'choice', 'params': {'choices': ['優', '良', '可', '差']}}
            ],
            'second_half': [
                {'type': 'choice', 'params': {'choices': ['優', '良', '可', '差']}}
            ]
        }
        
        self.attendance_schema = {
            'date': [
                {'type': 'required'},
                {'type': 'date'}
            ],
            'start_time': [
                {'type': 'required'},
                {'type': 'time'}
            ],
            'end_time': [
                {'type': 'required'},
                {'type': 'time'}
            ],
            'status': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['正常', '遲到', '早退', '曠職', '請假']}}
            ]
        }
        
        self.leave_request_schema = {
            'leave_type': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['年假', '病假', '事假', '婚假', '喪假', '產假', '陪產假']}}
            ],
            'start_date': [
                {'type': 'required'},
                {'type': 'date'}
            ],
            'end_date': [
                {'type': 'required'},
                {'type': 'date'}
            ],
            'days': [
                {'type': 'required'},
                {'type': 'number', 'params': {'min_value': 0.5, 'max_value': 365}}
            ],
            'apply_date': [
                {'type': 'required'},
                {'type': 'date'}
            ],
            'status': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['待審核', '已核准', '已拒絕']}}
            ],
            'reason': [
                {'type': 'required'},
                {'type': 'text_length', 'params': {'min_length': 5, 'max_length': 500}}
            ]
        }
        
        self.overtime_request_schema = {
            'overtime_date': [
                {'type': 'required'},
                {'type': 'date'}
            ],
            'start_time': [
                {'type': 'required'},
                {'type': 'time'}
            ],
            'end_time': [
                {'type': 'required'},
                {'type': 'time'}
            ],
            'hours': [
                {'type': 'required'},
                {'type': 'number', 'params': {'min_value': 0.5, 'max_value': 12}}
            ],
            'overtime_type': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['平日加班', '假日加班', '國定假日加班']}}
            ],
            'apply_date': [
                {'type': 'required'},
                {'type': 'date'}
            ],
            'status': [
                {'type': 'required'},
                {'type': 'choice', 'params': {'choices': ['待審核', '已核准', '已拒絕']}}
            ],
            'reason': [
                {'type': 'required'},
                {'type': 'text_length', 'params': {'min_length': 5, 'max_length': 500}}
            ]
        }
    
    def validate_basic_info(self, form_data: Dict[str, Any]) -> Tuple[bool, Dict[str, List[str]]]:
        """驗證基本資料"""
        return self.validate_form(form_data, self.basic_info_schema)
    
    def validate_performance(self, form_data: Dict[str, Any]) -> Tuple[bool, Dict[str, List[str]]]:
        """驗證考績資料"""
        return self.validate_form(form_data, self.performance_schema)
    
    def validate_attendance(self, form_data: Dict[str, Any]) -> Tuple[bool, Dict[str, List[str]]]:
        """驗證出勤資料"""
        is_valid, errors = self.validate_form(form_data, self.attendance_schema)
        
        # 額外驗證：結束時間必須晚於開始時間
        if 'start_time' in form_data and 'end_time' in form_data:
            start_time = form_data['start_time']
            end_time = form_data['end_time']
            
            if start_time and end_time:
                try:
                    start_dt = datetime.datetime.strptime(start_time, '%H:%M')
                    end_dt = datetime.datetime.strptime(end_time, '%H:%M')
                    
                    if end_dt <= start_dt:
                        if 'end_time' not in errors:
                            errors['end_time'] = []
                        errors['end_time'].append("結束時間必須晚於開始時間")
                        is_valid = False
                except ValueError:
                    pass  # 時間格式錯誤已由其他驗證處理
        
        return is_valid, errors
    
    def validate_leave_request(self, form_data: Dict[str, Any]) -> Tuple[bool, Dict[str, List[str]]]:
        """驗證請假申請"""
        is_valid, errors = self.validate_form(form_data, self.leave_request_schema)
        
        # 額外驗證：結束日期必須不早於開始日期
        if 'start_date' in form_data and 'end_date' in form_data:
            start_date = form_data['start_date']
            end_date = form_data['end_date']
            
            if start_date and end_date:
                try:
                    start_dt = datetime.datetime.strptime(start_date, '%Y-%m-%d')
                    end_dt = datetime.datetime.strptime(end_date, '%Y-%m-%d')
                    
                    if end_dt < start_dt:
                        if 'end_date' not in errors:
                            errors['end_date'] = []
                        errors['end_date'].append("結束日期不能早於開始日期")
                        is_valid = False
                except ValueError:
                    pass
        
        return is_valid, errors
    
    def validate_overtime_request(self, form_data: Dict[str, Any]) -> Tuple[bool, Dict[str, List[str]]]:
        """驗證加班申請"""
        is_valid, errors = self.validate_form(form_data, self.overtime_request_schema)
        
        # 額外驗證：結束時間必須晚於開始時間
        if 'start_time' in form_data and 'end_time' in form_data:
            start_time = form_data['start_time']
            end_time = form_data['end_time']
            
            if start_time and end_time:
                try:
                    start_dt = datetime.datetime.strptime(start_time, '%H:%M')
                    end_dt = datetime.datetime.strptime(end_time, '%H:%M')
                    
                    if end_dt <= start_dt:
                        if 'end_time' not in errors:
                            errors['end_time'] = []
                        errors['end_time'].append("結束時間必須晚於開始時間")
                        is_valid = False
                except ValueError:
                    pass
        
        return is_valid, errors


# 工具函數
def validate_employee_form(form_type: str, form_data: Dict[str, Any]) -> Tuple[bool, Dict[str, List[str]]]:
    """
    快速驗證員工表單
    
    Args:
        form_type: 表單類型 ('basic_info', 'performance', 'attendance', 'leave', 'overtime')
        form_data: 表單資料
    
    Returns:
        Tuple[bool, Dict[str, List[str]]]: (是否通過驗證, 錯誤訊息字典)
    """
    validator = EmployeeFormValidator()
    
    if form_type == 'basic_info':
        return validator.validate_basic_info(form_data)
    elif form_type == 'performance':
        return validator.validate_performance(form_data)
    elif form_type == 'attendance':
        return validator.validate_attendance(form_data)
    elif form_type == 'leave':
        return validator.validate_leave_request(form_data)
    elif form_type == 'overtime':
        return validator.validate_overtime_request(form_data)
    else:
        return False, {'form_type': ['不支援的表單類型']}


def format_validation_errors(errors: Dict[str, List[str]]) -> str:
    """
    格式化驗證錯誤訊息
    
    Args:
        errors: 錯誤訊息字典
    
    Returns:
        str: 格式化的錯誤訊息
    """
    if not errors:
        return ""
    
    error_lines = []
    for field, field_errors in errors.items():
        for error in field_errors:
            error_lines.append(f"• {error}")
    
    return "\n".join(error_lines)


# 測試用主程式
if __name__ == "__main__":
    print("表單驗證模組測試")
    
    # 測試基本資料驗證
    test_basic_data = {
        'employee_id': 'EMP001',
        'name': '張三',
        'id_number': 'A123456789',
        'gender': '男',
        'birth_date': '1990-01-01',
        'phone': '0912345678',
        'email': 'zhang@example.com',
        'department': '技術部',
        'position': '工程師',
        'hire_date': '2020-01-15',
        'work_location': '台北',
        'employment_type': '正職'
    }
    
    is_valid, errors = validate_employee_form('basic_info', test_basic_data)
    print(f"基本資料驗證結果: {'通過' if is_valid else '失敗'}")
    if not is_valid:
        print("錯誤訊息:")
        print(format_validation_errors(errors))
    
    # 測試錯誤資料
    test_invalid_data = {
        'employee_id': '',  # 必填但空白
        'name': '張',       # 長度不足
        'id_number': 'INVALID',  # 格式錯誤
        'phone': '123',     # 格式錯誤
        'email': 'invalid-email'  # 格式錯誤
    }
    
    is_valid, errors = validate_employee_form('basic_info', test_invalid_data)
    print(f"\n錯誤資料驗證結果: {'通過' if is_valid else '失敗'}")
    if not is_valid:
        print("錯誤訊息:")
        print(format_validation_errors(errors))