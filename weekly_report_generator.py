#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Hanwha BNCP Weekly Security Report Generator v3.0
- Claude AI API integration for Korean → English translation
- Auto-defaults (PSD H01/H02, IQD currency, personnel categorization)
- Auto-formatting to match report template conventions
"""

import os
import re
import json
import secrets
import zipfile
import tempfile
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

# ============================================================
# CONFIGURATION
# ============================================================
TEMPLATE_PATH = r"C:\Users\user\OneDrive - Harlow Group\한화 - 00.2026년\6.참고자료\자동화 작업\AI학습문서\업무보고서\Hanwha BNCP Weekly Report from 11 Mar 2025 to 17 Mar 2026(양식).docx"
OUTPUT_DIR = r"C:\Users\user\OneDrive - Harlow Group\00.2025년\6.참고자료\5.기타\AI 작업\AI Output\Weekly Report"
CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "weekly_report_config.json")

WEEKDAYS = ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]
MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

TRAINING_CATEGORIES = ["Quarterly", "Monthly &amp; Weekly", "Weapon Training", "PSD Training", "Other"]
TRAINING_LABELS = ["Quarterly", "Monthly & Weekly", "Weapon Training", "PSD Training", "Other (TBT)"]

VEHICLES = [
    {"plate": "189", "assigned": "Hanwha PMT"},
    {"plate": "19558", "assigned": "Hanwha PMT"},
]

# ============================================================
# KOREAN → ENGLISH TRANSLATION DICTIONARY
# ============================================================
KO_EN_DICT = {
    # Personnel
    "결근": "absent", "병가": "sick leave", "연차": "annual leave", "휴가": "leave",
    "무급휴가": "unpaid leave", "유급휴가": "paid leave", "진료": "medical treatment",
    "병원": "hospital", "출장": "business trip", "복귀": "return", "예정": "scheduled",
    "재택근무": "remote work", "교대": "shift change", "보충": "replacement",
    "대체": "replacement", "인원": "personnel", "가족": "family", "장례": "funeral",
    "경조사": "family event", "개인사유": "personal reason",
    # Vehicle
    "주유": "Refuel", "급유": "Refuel", "정비": "maintenance", "차량": "vehicle",
    "견인": "towing", "수리": "repair", "고장": "breakdown", "운행중지": "taken out of service",
    "엔진오일": "engine oil", "타이어": "tire",
    # Security
    "훈련": "training", "교육": "training", "무기": "weapon", "탄약": "ammunition",
    "경비": "guard", "가드": "guard", "경호": "escort", "순찰": "patrol",
    "출입": "access", "통제": "control", "점검": "inspection", "확인": "confirmed",
    # Operations
    "철수": "evacuation", "대피": "evacuation", "항공편": "flight",
    "민간항공편": "commercial flight", "공항": "airport", "입국": "entry",
    "출국": "departure", "이동": "movement", "제한": "restriction",
    # General
    "시행": "implemented", "완료": "completed", "진행": "in progress",
    "대기": "standby", "모니터링": "monitoring", "보고": "reported",
    "제출": "submitted", "승인": "approved", "요청": "requested",
    "실시": "conducted", "방문": "visited", "매카닉": "mechanic",
    "본사": "headquarters", "사무실": "office",
    # Time
    "오전": "AM", "오후": "PM", "일간": "day(s)", "주간": "week(s)",
    "월": "month", "주": "week", "일": "day",
    # Finance
    "지출": "expenditure", "잔액": "balance", "수령": "received",
    "플로트": "Float", "현금": "cash",
}

# Common full phrase translations
KO_EN_PHRASES = {
    "한국으로 출국": "departed to Korea",
    "한국에서 복귀": "returned from Korea",
    "복귀 예정": "scheduled to return",
    "병원 진료": "medical treatment at hospital",
    "가족 장례": "family funeral ceremony",
    "개인 사유": "personal reason",
    "본사로 견인": "to be towed to headquarters",
    "운행 중지": "taken out of service",
    "차량 상태": "vehicle condition",
    "상태가 좋지 않": "in poor condition",
    "아무런 답변이 없": "no response has been received",
    "사태 추이": "situation developments",
    "철수 계획": "evacuation plan",
    "민간항공편이 재개": "commercial flights resume",
    "즉시 철수": "immediate evacuation",
    "대사관": "Embassy",
    "한국 대사관": "Korean Embassy",
    "실장님": "Director",
    "매니저": "Manager",
    "드라이빙 어세스먼트": "Driving Assessment",
}


# Global API client (initialized when API key is set)
_api_client = None
_api_key = ""

def init_api_client(api_key):
    """Initialize the Anthropic API client."""
    global _api_client, _api_key
    if not HAS_ANTHROPIC or not api_key:
        _api_client = None
        return False
    try:
        _api_client = anthropic.Anthropic(api_key=api_key)
        _api_key = api_key
        return True
    except Exception:
        _api_client = None
        return False

def translate_ko_to_en(text):
    """Translate Korean text to English using Claude API."""
    if not text:
        return text
    # No Korean characters → return as-is
    if not any('\uac00' <= c <= '\ud7a3' for c in text):
        return text
    # If API not available, return original text with a warning marker
    if not _api_client:
        return text

    try:
        response = _api_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": text}],
            system="""You are a translator for a security company's weekly report at Hanwha BNCP project site in Iraq.
Translate the Korean input to professional English suitable for a formal security weekly report.

Rules:
- Keep person names, location names, and ID numbers (like L1005) exactly as-is
- Use security/military terminology (e.g., PSD, SSG, NSTR, AMS HQ)
- PSD teams are always H01 and H02 unless stated otherwise
- Currency is always IQD (Iraqi Dinar)
- Keep dates in "DD Mon" format (e.g., "11 Mar")
- Be concise and professional
- If the input has "- " bullet points, keep them as "- " bullets
- Output ONLY the translated English text, no explanations
- Do NOT add any extra information not in the original text"""
        )
        return response.content[0].text.strip()
    except Exception as e:
        print(f"API translation error: {e}")
        return text  # Return original on error

def translate_all_fields(data_dict):
    """Batch translate all text fields in a data dictionary to reduce API calls."""
    if not _api_client:
        return data_dict

    # Collect all Korean texts that need translation
    texts_to_translate = []
    text_paths = []  # Track where each text came from

    def collect_texts(obj, path=""):
        if isinstance(obj, str) and any('\uac00' <= c <= '\ud7a3' for c in obj):
            texts_to_translate.append(obj)
            text_paths.append(path)
        elif isinstance(obj, dict):
            for k, v in obj.items():
                collect_texts(v, f"{path}.{k}")
        elif isinstance(obj, list):
            for i, v in enumerate(obj):
                collect_texts(v, f"{path}[{i}]")

    collect_texts(data_dict)

    if not texts_to_translate:
        return data_dict

    # Batch translate in a single API call
    numbered_texts = "\n".join(f"[{i+1}] {t}" for i, t in enumerate(texts_to_translate))
    try:
        response = _api_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            messages=[{"role": "user", "content": numbered_texts}],
            system="""You are a translator for a security company's weekly report at Hanwha BNCP project site in Iraq.
Translate each numbered Korean text to professional English for a formal security weekly report.

Rules:
- Keep person names, location names, ID numbers (L1005, RV920 etc.) exactly as-is
- Use security/military terminology (PSD, SSG, NSTR, AMS HQ)
- PSD teams are always H01 and H02 unless stated otherwise
- Currency is always IQD (Iraqi Dinar)
- Dates in "DD Mon" format (e.g., "11 Mar")
- Be concise and professional
- Keep "- " bullet formatting if present
- Output format: [1] translated text\n[2] translated text\n...
- Output ONLY translations with their numbers, no extra text
- Do NOT add information not in the original"""
        )
        result_text = response.content[0].text.strip()

        # Parse results
        translations = {}
        for line in result_text.split("\n"):
            m = re.match(r'\[(\d+)\]\s*(.*)', line)
            if m:
                idx = int(m.group(1)) - 1
                translations[idx] = m.group(2).strip()

        # Apply translations back to data
        def apply_translations(obj, path=""):
            if isinstance(obj, str) and path in text_paths:
                idx = text_paths.index(path)
                return translations.get(idx, obj)
            elif isinstance(obj, dict):
                return {k: apply_translations(v, f"{path}.{k}") for k, v in obj.items()}
            elif isinstance(obj, list):
                return [apply_translations(v, f"{path}[{i}]") for i, v in enumerate(obj)]
            return obj

        return apply_translations(data_dict)

    except Exception as e:
        print(f"Batch translation error: {e}")
        return data_dict


def auto_format_finance(amount_str):
    """Auto-add IQD prefix and format amount."""
    amount_str = amount_str.strip()
    if not amount_str:
        return ""
    # Remove existing IQD if present
    amount_str = amount_str.replace("IQD", "").strip()
    # Try to format as number with commas
    try:
        num = int(amount_str.replace(",", "").replace(" ", ""))
        return f"IQD {num:,}"
    except ValueError:
        return f"IQD {amount_str}"


def xml_escape(text):
    """Escape special characters for XML."""
    text = text.replace("&", "&amp;")
    text = text.replace("<", "&lt;")
    text = text.replace(">", "&gt;")
    text = text.replace('"', "&quot;")
    text = text.replace("'", "&apos;")
    return text


def gen_id():
    return secrets.token_hex(4).upper()


def get_report_period(reference_date=None):
    if reference_date is None:
        reference_date = datetime.now().date()
    weekday = reference_date.weekday()
    days_since_wed = (weekday - 2) % 7
    start = reference_date - timedelta(days=days_since_wed)
    end = start + timedelta(days=6)
    return start, end


# ============================================================
# CONFIG MANAGER
# ============================================================
class ConfigManager:
    DEFAULTS = {
        "api_key": "",
        "vehicle_189_mileage": "", "vehicle_189_next_service": "", "vehicle_189_comments": "Serviced on 27 December 2023",
        "vehicle_19558_mileage": "", "vehicle_19558_next_service": "", "vehicle_19558_comments": "Serviced on 21 March 2024",
    }

    def __init__(self):
        self.data = dict(self.DEFAULTS)
        self.load()

    def load(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                self.data.update(json.load(f))
        except (FileNotFoundError, json.JSONDecodeError):
            pass

    def save(self):
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.data, f, ensure_ascii=False, indent=2)

    def get(self, key):
        return self.data.get(key, "")

    def set(self, key, value):
        self.data[key] = value


# ============================================================
# XML TEMPLATES
# ============================================================
class XmlTemplates:
    @staticmethod
    def day_header(day_num, month_abbr, weekday):
        pid, tid = gen_id(), gen_id()
        return f'''<w:p w14:paraId="{pid}" w14:textId="{tid}" w:rsidR="00486C4C" w:rsidRDefault="00486C4C" w:rsidP="00486C4C">
            <w:pPr><w:widowControl w:val="0"/><w:tabs><w:tab w:val="left" w:pos="270"/></w:tabs>
              <w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/>
              <w:spacing w:line="254" w:lineRule="atLeast"/><w:ind w:left="360"/><w:jc w:val="both"/>
              <w:rPr><w:rFonts w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
            </w:pPr>
            <w:r><w:rPr><w:rFonts w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
              <w:t>{day_num} {month_abbr}, {weekday}</w:t></w:r></w:p>'''

    @staticmethod
    def bullet_item(text, num_id="8"):
        pid, tid = gen_id(), gen_id()
        escaped = xml_escape(text)
        return f'''<w:p w14:paraId="{pid}" w14:textId="{tid}" w:rsidR="00486C4C" w:rsidRDefault="00486C4C" w:rsidP="00486C4C">
            <w:pPr><w:pStyle w:val="a4"/><w:widowControl w:val="0"/>
              <w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>
              <w:tabs><w:tab w:val="left" w:pos="270"/></w:tabs>
              <w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/>
              <w:spacing w:line="254" w:lineRule="atLeast"/><w:jc w:val="both"/>
              <w:rPr><w:rFonts w:cs="Calibri"/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
            </w:pPr>
            <w:r><w:rPr><w:rFonts w:cs="Calibri"/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
              <w:t>{escaped}</w:t></w:r></w:p>'''

    @staticmethod
    def blank_separator():
        pid, tid = gen_id(), gen_id()
        return f'''<w:p w14:paraId="{pid}" w14:textId="{tid}" w:rsidR="00486C4C" w:rsidRDefault="00486C4C" w:rsidP="00486C4C">
            <w:pPr><w:widowControl w:val="0"/><w:tabs><w:tab w:val="left" w:pos="270"/></w:tabs>
              <w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/>
              <w:spacing w:line="254" w:lineRule="atLeast"/><w:ind w:left="360"/><w:jc w:val="both"/>
              <w:rPr><w:rFonts w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
            </w:pPr></w:p>'''

    @staticmethod
    def training_cell_para(text):
        pid, tid = gen_id(), gen_id()
        escaped = xml_escape(text)
        return f'''<w:p w14:paraId="{pid}" w14:textId="{tid}" w:rsidR="006322AC" w:rsidRDefault="006322AC" w:rsidP="006322AC">
            <w:pPr><w:jc w:val="both"/><w:rPr><w:rFonts w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr>
            <w:r><w:rPr><w:rFonts w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
              <w:t>{escaped}</w:t></w:r></w:p>'''

    @staticmethod
    def issues_row(issue, summary, actions, col_widths=(2382, 3714, 3728)):
        rpid = gen_id()
        cells = []
        for val, w in zip([issue, summary, actions], col_widths):
            pid, tid = gen_id(), gen_id()
            cells.append(f'''<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="{pid}" w14:textId="{tid}" w:rsidR="004E3951" w:rsidRDefault="004E3951">
            <w:pPr><w:widowControl w:val="0"/><w:autoSpaceDE w:val="0"/><w:autoSpaceDN w:val="0"/><w:adjustRightInd w:val="0"/><w:spacing w:line="254" w:lineRule="atLeast"/><w:jc w:val="both"/></w:pPr>
            <w:r><w:t>{xml_escape(val)}</w:t></w:r></w:p></w:tc>''')
        return f'''<w:tr w:rsidR="004E3951" w14:paraId="{rpid}" w14:textId="{gen_id()}" w:rsidTr="003503B3">
        <w:trPr><w:trHeight w:val="350"/></w:trPr>{"".join(cells)}</w:tr>'''

    @staticmethod
    def issues_empty_row(col_widths=(2382, 3714, 3728)):
        rpid = gen_id()
        cells = []
        for w in col_widths:
            pid = gen_id()
            cells.append(f'''<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="{pid}" w14:textId="{gen_id()}" w:rsidR="004E3951" w:rsidRDefault="004E3951"/></w:tc>''')
        return f'''<w:tr w:rsidR="004E3951" w14:paraId="{rpid}" w14:textId="{gen_id()}" w:rsidTr="003503B3">
        <w:trPr><w:trHeight w:val="350"/></w:trPr>{"".join(cells)}</w:tr>'''

    @staticmethod
    def finance_row(date_str, pr_num, description, amount, balance):
        rpid = gen_id()
        cells = []
        widths = [1389, 1418, 3969, 1417, 1447]
        for val, w in zip([date_str, pr_num, description, amount, balance], widths):
            pid = gen_id()
            cells.append(f'''<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>
            <w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders></w:tcPr>
          <w:p w14:paraId="{pid}" w14:textId="{gen_id()}" w:rsidR="00C96240" w:rsidRDefault="00C96240">
            <w:pPr><w:jc w:val="center"/><w:rPr><w:rFonts w:cs="Calibri"/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr></w:pPr>
            <w:r><w:rPr><w:rFonts w:cs="Calibri"/><w:bCs/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
              <w:t>{xml_escape(val)}</w:t></w:r></w:p></w:tc>''')
        return f'''<w:tr w:rsidR="00C96240" w14:paraId="{rpid}" w14:textId="{gen_id()}" w:rsidTr="00917E72">
        <w:trPr><w:trHeight w:val="227"/></w:trPr>{"".join(cells)}</w:tr>'''

    @staticmethod
    def finance_empty_row():
        rpid = gen_id()
        cells = []
        for w in [1389, 1418, 3969, 1417, 1447]:
            pid = gen_id()
            cells.append(f'''<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/>
            <w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders></w:tcPr>
          <w:p w14:paraId="{pid}" w14:textId="{gen_id()}" w:rsidR="00C96240" w:rsidRDefault="00C96240"><w:pPr><w:jc w:val="center"/></w:pPr></w:p></w:tc>''')
        return f'''<w:tr w:rsidR="00C96240" w14:paraId="{rpid}" w14:textId="{gen_id()}" w:rsidTr="00917E72">
        <w:trPr><w:trHeight w:val="227"/></w:trPr>{"".join(cells)}</w:tr>'''

    @staticmethod
    def client_feedback_row(issue, summary_lines, actions_lines):
        rpid = gen_id()
        pid1 = gen_id()
        issue_cell = f'''<w:tc><w:tcPr><w:tcW w:w="1843" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="{pid1}" w14:textId="{gen_id()}" w:rsidR="0011734F" w:rsidRDefault="0011734F">
            <w:pPr><w:jc w:val="center"/></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/></w:rPr>
              <w:t>{xml_escape(issue)}</w:t></w:r></w:p></w:tc>'''
        summary_paras = ""
        for line in summary_lines:
            pid = gen_id()
            summary_paras += f'''<w:p w14:paraId="{pid}" w14:textId="{gen_id()}" w:rsidR="0011734F" w:rsidRDefault="0011734F">
            <w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/></w:rPr></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/></w:rPr>
              <w:t>{xml_escape(line)}</w:t></w:r></w:p>'''
        summary_cell = f'''<w:tc><w:tcPr><w:tcW w:w="5245" w:type="dxa"/></w:tcPr>{summary_paras}</w:tc>'''
        actions_paras = ""
        for line in actions_lines:
            pid = gen_id()
            actions_paras += f'''<w:p w14:paraId="{pid}" w14:textId="{gen_id()}" w:rsidR="0011734F" w:rsidRDefault="0011734F">
            <w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/></w:rPr></w:pPr>
            <w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/><w:color w:val="000000"/></w:rPr>
              <w:t>{xml_escape(line)}</w:t></w:r></w:p>'''
        actions_cell = f'''<w:tc><w:tcPr><w:tcW w:w="2594" w:type="dxa"/></w:tcPr>{actions_paras}</w:tc>'''
        return f'''<w:tr w:rsidR="0011734F" w14:paraId="{rpid}" w14:textId="{gen_id()}" w:rsidTr="00724B80">
        <w:trPr><w:trHeight w:val="325"/></w:trPr>{issue_cell}{summary_cell}{actions_cell}</w:tr>'''

    @staticmethod
    def client_feedback_empty_row():
        rpid = gen_id()
        cells = ""
        for w in [1843, 5245, 2594]:
            pid = gen_id()
            cells += f'''<w:tc><w:tcPr><w:tcW w:w="{w}" w:type="dxa"/></w:tcPr>
          <w:p w14:paraId="{pid}" w14:textId="{gen_id()}" w:rsidR="0011734F" w:rsidRDefault="0011734F"/></w:tc>'''
        return f'''<w:tr w:rsidR="0011734F" w14:paraId="{rpid}" w14:textId="{gen_id()}" w:rsidTr="00724B80">
        <w:trPr><w:trHeight w:val="325"/></w:trPr>{cells}</w:tr>'''


# ============================================================
# DOCX GENERATOR ENGINE
# ============================================================
class DocxGenerator:
    def __init__(self, template_path, output_dir):
        self.template_path = template_path
        self.output_dir = output_dir

    def generate(self, data):
        start = data["period_start"]
        end = data["period_end"]
        fname = f"Hanwha BNCP Weekly Report from {start.day} {MONTH_ABBR[start.month-1]} to {end.day} {MONTH_ABBR[end.month-1]} {end.year}.docx"
        output_path = os.path.join(self.output_dir, fname)

        with tempfile.TemporaryDirectory() as tmpdir:
            work_dir = os.path.join(tmpdir, "work")
            with zipfile.ZipFile(self.template_path, "r") as z:
                z.extractall(work_dir)

            xml_path = os.path.join(work_dir, "word", "document.xml")
            with open(xml_path, "r", encoding="utf-8") as f:
                content = f.read()

            content = self._modify_header(content, data)
            content = self._modify_weekly_summary(content, data)
            content = self._modify_training(content, data)
            content = self._modify_issues(content, data)
            content = self._modify_mileage(content, data)
            content = self._modify_finance(content, data)
            content = self._modify_client_feedback(content, data)

            with open(xml_path, "w", encoding="utf-8") as f:
                f.write(content)

            os.makedirs(self.output_dir, exist_ok=True)
            self._repack(work_dir, output_path)
        return output_path

    def _modify_header(self, content, data):
        """Update MONTH and PERIOD using paragraph-level text replacement."""
        start = data["period_start"]
        end = data["period_end"]
        month_name = MONTHS[start.month - 1]
        month_abbr = MONTH_ABBR[start.month - 1]

        # MONTH: Replace "January 2026" (or any month+year) with new value
        # Find the exact <w:t> containing the month name
        for m in MONTHS:
            old_month_text = f"<w:t>{m} {start.year - 1}</w:t>"
            new_month_text = f"<w:t>{month_name} {start.year}</w:t>"
            if old_month_text in content:
                content = content.replace(old_month_text, new_month_text, 1)
                break
            old_month_text2 = f"<w:t>{m} {start.year}</w:t>"
            if old_month_text2 in content:
                content = content.replace(old_month_text2, new_month_text, 1)
                break

        # PERIOD: Find the paragraph containing "PERIOD:" and replace date values
        period_idx = content.index("PERIOD:")
        period_p_start = content.rfind("<w:p ", 0, period_idx)
        period_p_end = content.index("</w:p>", period_idx) + 6
        period_para = content[period_p_start:period_p_end]

        # Extract all <w:t> values from this paragraph
        t_elements = list(re.finditer(r'(<w:t[^>]*>)([^<]*)(</w:t>)', period_para))

        # The text sequence is: "PERIOD:", " ", "4", " ", "Mar", " ", "to ", "10", " ", "Mar", " 202", "6"
        # We need to replace the date components
        new_period_texts = {
            # Find the first number (start day) - it's after "PERIOD:" and spaces
            # Find "to " - marks the boundary between start and end dates
        }

        # Rebuild paragraph with new dates
        found_period = False
        date_parts_start = []  # indices of t_elements for start date
        date_parts_end = []    # indices of t_elements for end date
        found_to = False
        found_year = False

        for i, m in enumerate(t_elements):
            txt = m.group(2)
            if "PERIOD:" in txt:
                found_period = True
                continue
            if not found_period:
                continue
            if "to" in txt:
                found_to = True
                continue
            if not found_to:
                date_parts_start.append(i)
            else:
                date_parts_end.append(i)

        # Now replace: rebuild the entire paragraph with correct dates
        # Simpler approach: replace specific text patterns within the paragraph
        new_para = period_para

        # Replace year (might be split as " 202" + "6")
        new_para = re.sub(r'(<w:t[^>]*>)\s*\d{3,4}(</w:t>)', f'\\1 {start.year}\\2', new_para, count=0)

        # Remove stray single-digit year fragments
        # Better approach: rebuild the date text completely
        # Find all <w:t> elements and their positions, then replace the date parts

        # Actually, let's use a cleaner approach:
        # 1. Find the PERIOD paragraph
        # 2. Remove all runs after "PERIOD:" except the first space
        # 3. Insert a single run with the complete period text

        # Find the run containing "PERIOD:"
        period_run_end = period_para.index("PERIOD:")
        period_run_end = period_para.index("</w:r>", period_run_end) + 6

        # Get the formatting from the next run for styling
        next_run = re.search(r'<w:r[^>]*>(.*?)</w:r>', period_para[period_run_end:], re.DOTALL)
        if next_run:
            rpr_match = re.search(r'<w:rPr>.*?</w:rPr>', next_run.group(1), re.DOTALL)
            rpr = rpr_match.group() if rpr_match else ""
        else:
            rpr = '<w:rPr><w:rFonts w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>'

        # Get the paragraph closing
        ppr_end = period_para.index("</w:pPr>") + 8 if "</w:pPr>" in period_para else 0

        # Build new paragraph: keep pPr + PERIOD run + new date run
        p_open_end = period_para.index(">", period_p_start - period_p_start) + 1  # This won't work, let me simplify

        # Simplest: just keep everything up to and including the PERIOD: run, then add a new run with the full date
        before_period_runs = period_para[:period_run_end]
        date_text = f" {start.day} {month_abbr} to {end.day} {MONTH_ABBR[end.month-1]} {end.year}"
        new_run = f'<w:r><w:rPr><w:rFonts w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>{date_text}</w:t></w:r>'
        new_para = before_period_runs + new_run + "</w:p>"

        content = content[:period_p_start] + new_para + content[period_p_end:]
        return content

    def _modify_weekly_summary(self, content, data):
        start = data["period_start"]
        shifts = data.get("shift_changes", [])
        daily_extras = data.get("daily_extras", {})

        marker = "WEEKLY SUMMARY"
        marker_idx = content.index(marker)
        tbl_start = content.index("<w:tbl>", marker_idx)
        tbl_end = content.index("</w:tbl>", tbl_start) + 8

        table_xml = content[tbl_start:tbl_end]
        first_tr_end = table_xml.index("</w:tr>") + 7
        second_tr_start = table_xml.index("<w:tr", first_tr_end)
        second_tr_end = table_xml.index("</w:tr>", second_tr_start) + 7

        row_xml = table_xml[second_tr_start:second_tr_end]
        tc_start = row_xml.index("<w:tc>")
        tc_end = row_xml.rindex("</w:tc>") + 7
        tc_xml = row_xml[tc_start:tc_end]
        tcpr_end = tc_xml.index("</w:tcPr>") + 9
        tc_header = tc_xml[:tcpr_end]

        paragraphs = []
        num_ids = ["8", "9", "10", "11", "12", "13", "14"]

        for i in range(7):
            day_date = start + timedelta(days=i)
            day_num = day_date.day
            day_abbr = MONTH_ABBR[day_date.month - 1]
            weekday = WEEKDAYS[day_date.weekday()]
            num_id = num_ids[i] if i < len(num_ids) else "8"

            if i > 0:
                paragraphs.append(XmlTemplates.blank_separator())

            paragraphs.append(XmlTemplates.day_header(day_num, day_abbr, weekday))
            paragraphs.append(XmlTemplates.bullet_item("Daily check - PSD & Static Guard: NSTR", num_id))

            # Shift changes
            for sc in shifts:
                if sc["date"] == day_date.day:
                    paragraphs.append(XmlTemplates.bullet_item(
                        f"SSG shift change completed (Shift #{sc['shift']})", num_id))

            # Routine items
            if day_date.weekday() == 4:  # Friday
                paragraphs.append(XmlTemplates.bullet_item("Submitted Meal Request to Hanwha", num_id))
                paragraphs.append(XmlTemplates.bullet_item("Submitted Weekly Weapon & Ammunition Status to AMS HQ", num_id))
            elif day_date.weekday() == 1:  # Tuesday
                paragraphs.append(XmlTemplates.bullet_item("Submitted Weekly Mileage & Weekly Security Report to AMS HQ", num_id))

            # Extra items (translated to English)
            day_key = str(day_date.day)
            if day_key in daily_extras and daily_extras[day_key].strip():
                for line in daily_extras[day_key].strip().split("\n"):
                    if line.strip():
                        translated = translate_ko_to_en(line.strip())
                        paragraphs.append(XmlTemplates.bullet_item(translated, num_id))

        new_tc = tc_header + "\n" + "\n".join(paragraphs) + "\n</w:tc>"
        new_row = row_xml[:tc_start] + new_tc + row_xml[tc_end:]
        new_table = table_xml[:second_tr_start] + new_row + table_xml[second_tr_end:]
        content = content[:tbl_start] + new_table + content[tbl_end:]
        return content

    def _modify_training(self, content, data):
        training = data.get("training", {})
        marker_idx = content.index("TRAINING")
        tbl_start = content.index("<w:tbl>", marker_idx)
        tbl_end = content.index("</w:tbl>", tbl_start) + 8
        table_xml = content[tbl_start:tbl_end]

        rows = list(re.finditer(r'<w:tr\b[^>]*>.*?</w:tr>', table_xml, re.DOTALL))
        for i in range(len(rows) - 1, 0, -1):
            row_match = rows[i]
            row_xml = row_match.group()

            cat_match = None
            for cat in TRAINING_CATEGORIES:
                if cat in row_xml:
                    cat_match = cat
                    break
            if cat_match is None:
                continue

            cat_idx = TRAINING_CATEGORIES.index(cat_match)
            cat_label = TRAINING_LABELS[cat_idx]
            entries = training.get(cat_label, "").strip()

            first_tc_end = row_xml.index("</w:tc>") + 7
            second_tc_start = row_xml.index("<w:tc>", first_tc_end)
            second_tc_end = row_xml.rindex("</w:tc>") + 7
            old_tc = row_xml[second_tc_start:second_tc_end]

            tcpr_match = re.search(r'<w:tcPr>.*?</w:tcPr>', old_tc, re.DOTALL)
            tcpr = tcpr_match.group() if tcpr_match else '<w:tcPr><w:tcW w:w="7949" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>'

            if entries:
                paras = []
                for line in entries.split("\n"):
                    if line.strip():
                        paras.append(XmlTemplates.training_cell_para(translate_ko_to_en(line.strip())))
                new_tc = f"<w:tc>{tcpr}{''.join(paras)}</w:tc>"
            else:
                new_tc = f"<w:tc>{tcpr}{XmlTemplates.training_cell_para('N/A')}</w:tc>"

            new_row = row_xml[:second_tc_start] + new_tc + row_xml[second_tc_end:]
            table_xml = table_xml[:row_match.start()] + new_row + table_xml[row_match.end():]

        content = content[:tbl_start] + table_xml + content[tbl_end:]
        return content

    def _modify_issues(self, content, data):
        issues = data.get("issues", [])
        marker_idx = content.index("Issues")
        check_area = content[max(0, marker_idx-200):marker_idx]
        if "5.1" not in check_area:
            marker_idx = content.index("Issues", marker_idx + 1)
        tbl_start = content.index("<w:tbl>", marker_idx)
        tbl_end = content.index("</w:tbl>", tbl_start) + 8
        table_xml = content[tbl_start:tbl_end]

        rows = list(re.finditer(r'<w:tr\b[^>]*>.*?</w:tr>', table_xml, re.DOTALL))
        if not rows:
            return content
        header_row = rows[0].group()

        new_rows = [header_row]
        if issues:
            for iss in issues:
                new_rows.append(XmlTemplates.issues_row(
                    translate_ko_to_en(iss.get("issue", "")),
                    translate_ko_to_en(iss.get("summary", "")),
                    translate_ko_to_en(iss.get("actions", ""))
                ))
        new_rows.append(XmlTemplates.issues_empty_row())

        tbl_grid_end = table_xml.index("</w:tblGrid>") + 12
        new_table = table_xml[:tbl_grid_end] + "\n" + "\n".join(new_rows) + "\n</w:tbl>"
        content = content[:tbl_start] + new_table + content[tbl_end:]
        return content

    def _modify_mileage(self, content, data):
        mileage = data.get("mileage", {})
        for plate, values in mileage.items():
            if not values.get("current"):
                continue
            plate_marker = f">{plate}<"
            idx = content.find(plate_marker)
            if idx == -1:
                continue
            tr_start = content.rfind("<w:tr", 0, idx)
            tr_end = content.index("</w:tr>", idx) + 7
            row_xml = content[tr_start:tr_end]

            tcs = list(re.finditer(r'<w:tc>.*?</w:tc>', row_xml, re.DOTALL))
            if len(tcs) < 4:
                continue

            # Mileage (col 2)
            if values.get("current"):
                old_tc = tcs[1].group()
                new_tc = re.sub(r'<w:t[^>]*>[^<]*</w:t>', f'<w:t>{values["current"]}</w:t>', old_tc, count=1)
                row_xml = row_xml[:tcs[1].start()] + new_tc + row_xml[tcs[1].end():]
                tcs = list(re.finditer(r'<w:tc>.*?</w:tc>', row_xml, re.DOTALL))

            # Next Service (col 3)
            if values.get("next_service") and len(tcs) >= 3:
                old_tc = tcs[2].group()
                new_tc = re.sub(r'<w:t[^>]*>[^<]*</w:t>', f'<w:t>{values["next_service"]}</w:t>', old_tc)
                row_xml = row_xml[:tcs[2].start()] + new_tc + row_xml[tcs[2].end():]
                tcs = list(re.finditer(r'<w:tc>.*?</w:tc>', row_xml, re.DOTALL))

            # Comments (col 4)
            if values.get("comments") and len(tcs) >= 4:
                old_tc = tcs[3].group()
                new_tc = re.sub(r'<w:t[^>]*>[^<]*</w:t>', f'<w:t>{xml_escape(values["comments"])}</w:t>', old_tc, count=1)
                row_xml = row_xml[:tcs[3].start()] + new_tc + row_xml[tcs[3].end():]

            content = content[:tr_start] + row_xml + content[tr_end:]
        return content

    def _modify_finance(self, content, data):
        finance = data.get("finance", [])
        marker_idx = content.index("5.8 Finance")
        tbl_start = content.index("<w:tbl>", marker_idx)
        tbl_end = content.index("</w:tbl>", tbl_start) + 8
        table_xml = content[tbl_start:tbl_end]

        rows = list(re.finditer(r'<w:tr\b[^>]*>.*?</w:tr>', table_xml, re.DOTALL))
        if not rows:
            return content
        header_row = rows[0].group()

        new_rows = [header_row]
        if finance:
            for fin in finance:
                new_rows.append(XmlTemplates.finance_row(
                    fin.get("date", ""), fin.get("pr_number", ""),
                    translate_ko_to_en(fin.get("description", "")),
                    auto_format_finance(fin.get("amount", "")),
                    auto_format_finance(fin.get("balance", ""))
                ))
        new_rows.append(XmlTemplates.finance_empty_row())

        tbl_grid_end = table_xml.index("</w:tblGrid>") + 12
        new_table = table_xml[:tbl_grid_end] + "\n" + "\n".join(new_rows) + "\n</w:tbl>"
        content = content[:tbl_start] + new_table + content[tbl_end:]
        return content

    def _modify_client_feedback(self, content, data):
        feedback = data.get("client_feedback", [])
        marker_idx = content.index("6. Client Feedback")
        tbl_start = content.index("<w:tbl>", marker_idx)
        tbl_end = content.index("</w:tbl>", tbl_start) + 8
        table_xml = content[tbl_start:tbl_end]

        rows = list(re.finditer(r'<w:tr\b[^>]*>.*?</w:tr>', table_xml, re.DOTALL))
        if not rows:
            return content
        header_row = rows[0].group()

        new_rows = [header_row]
        if feedback:
            for fb in feedback:
                summary_lines = [translate_ko_to_en(l.strip()) for l in fb.get("summary", "").split("\n") if l.strip()]
                actions_lines = [translate_ko_to_en(l.strip()) for l in fb.get("actions", "").split("\n") if l.strip()]
                new_rows.append(XmlTemplates.client_feedback_row(
                    translate_ko_to_en(fb.get("issue", "")),
                    summary_lines if summary_lines else [""],
                    actions_lines if actions_lines else [""]
                ))
        new_rows.append(XmlTemplates.client_feedback_empty_row())

        tbl_grid_end = table_xml.index("</w:tblGrid>") + 12
        new_table = table_xml[:tbl_grid_end] + "\n" + "\n".join(new_rows) + "\n</w:tbl>"
        content = content[:tbl_start] + new_table + content[tbl_end:]
        return content

    def _repack(self, work_dir, output_path):
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for root, dirs, files in os.walk(work_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, work_dir)
                    zf.write(file_path, arcname)


# ============================================================
# GUI APPLICATION
# ============================================================
class WeeklyReportApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Hanwha BNCP Weekly Report Generator v3.0")
        self.root.geometry("950x800")
        self.config = ConfigManager()
        self.generator = DocxGenerator(TEMPLATE_PATH, OUTPUT_DIR)
        self._init_api()
        self._build_ui()
        self._set_current_week()

    def _init_api(self):
        """Initialize API from saved key."""
        saved_key = self.config.get("api_key")
        if saved_key and HAS_ANTHROPIC:
            init_api_client(saved_key)

    def _build_ui(self):
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # API Key bar at top
        api_frame = ttk.LabelFrame(main, text="Claude API (Korean → English auto-translation)", padding=5)
        api_frame.pack(fill=tk.X, pady=(0, 5))

        api_row = ttk.Frame(api_frame)
        api_row.pack(fill=tk.X)
        ttk.Label(api_row, text="API Key:").pack(side=tk.LEFT)
        self.api_key_var = tk.StringVar(value=self.config.get("api_key"))
        self.api_key_entry = ttk.Entry(api_row, textvariable=self.api_key_var, width=50, show="*")
        self.api_key_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Button(api_row, text="Save & Test", command=self._save_api_key).pack(side=tk.LEFT, padx=2)
        self.api_status = ttk.Label(api_row, text="● Connected" if _api_client else "○ Not connected",
                                     foreground="green" if _api_client else "gray")
        self.api_status.pack(side=tk.LEFT, padx=10)

        self.notebook = ttk.Notebook(main)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self._build_tab_basic()
        self._build_tab_training()
        self._build_tab_issues_feedback()
        self._build_tab_vehicle_finance()

        bottom = ttk.Frame(main)
        bottom.pack(fill=tk.X)
        self.status_var = tk.StringVar(value="Ready" + (" (AI Translation ON)" if _api_client else " (AI Translation OFF)"))
        ttk.Label(bottom, textvariable=self.status_var).pack(side=tk.LEFT)
        ttk.Button(bottom, text="  Generate Report  ", command=self._generate_report).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom, text="This Week", command=self._set_current_week).pack(side=tk.RIGHT, padx=5)

    def _save_api_key(self):
        """Save and test the API key."""
        key = self.api_key_var.get().strip()
        if not key:
            messagebox.showwarning("Warning", "Please enter an API key.")
            return
        if not HAS_ANTHROPIC:
            messagebox.showerror("Error", "anthropic package not installed.\nRun: pip install anthropic")
            return
        success = init_api_client(key)
        if success:
            # Quick test
            try:
                test = translate_ko_to_en("테스트")
                self.config.set("api_key", key)
                self.config.save()
                self.api_status.config(text="● Connected", foreground="green")
                self.status_var.set("Ready (AI Translation ON)")
                messagebox.showinfo("Success", f"API connected! Test: '테스트' → '{test}'")
            except Exception as e:
                self.api_status.config(text="● Error", foreground="red")
                messagebox.showerror("Error", f"API test failed: {e}")
        else:
            self.api_status.config(text="○ Failed", foreground="red")
            messagebox.showerror("Error", "Failed to initialize API client. Check your key.")

    # ---- Tab 1: Basic Info + Daily Summary ----
    def _build_tab_basic(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Basic / Daily Summary")

        # Period
        pf = ttk.LabelFrame(tab, text="Report Period (Wed~Tue)", padding=10)
        pf.pack(fill=tk.X, pady=3)
        row = ttk.Frame(pf)
        row.pack(fill=tk.X)
        ttk.Label(row, text="Start (WED):").pack(side=tk.LEFT)
        self.start_year = ttk.Combobox(row, width=6, values=list(range(2025, 2030)))
        self.start_year.pack(side=tk.LEFT, padx=2)
        self.start_month = ttk.Combobox(row, width=4, values=list(range(1, 13)))
        self.start_month.pack(side=tk.LEFT, padx=2)
        self.start_day = ttk.Combobox(row, width=4, values=list(range(1, 32)))
        self.start_day.pack(side=tk.LEFT, padx=2)
        ttk.Label(row, text="  ~  End (TUE):").pack(side=tk.LEFT, padx=(20, 0))
        self.end_label = ttk.Label(row, text="", font=("Calibri", 10, "bold"))
        self.end_label.pack(side=tk.LEFT, padx=5)
        for w in [self.start_year, self.start_month, self.start_day]:
            w.bind("<<ComboboxSelected>>", self._update_end_date)

        # Shift changes
        sf = ttk.LabelFrame(tab, text="SSG Shift Changes (4-day rotation)", padding=10)
        sf.pack(fill=tk.X, pady=3)
        self.shift_entries = []
        for i in range(3):
            row = ttk.Frame(sf)
            row.pack(fill=tk.X, pady=1)
            ttk.Label(row, text=f"#{i+1} Day:").pack(side=tk.LEFT)
            dv = tk.StringVar()
            ttk.Entry(row, textvariable=dv, width=5).pack(side=tk.LEFT, padx=2)
            ttk.Label(row, text="Shift#:").pack(side=tk.LEFT, padx=(10, 0))
            sv = tk.StringVar()
            ttk.Combobox(row, textvariable=sv, width=4, values=["1", "2"]).pack(side=tk.LEFT, padx=2)
            self.shift_entries.append((dv, sv))

        # Daily extras
        df = ttk.LabelFrame(tab, text="Daily Events (routine auto-filled, add extras only - Korean OK, auto-translated)", padding=10)
        df.pack(fill=tk.BOTH, expand=True, pady=3)

        canvas = tk.Canvas(df)
        sb = ttk.Scrollbar(df, orient="vertical", command=canvas.yview)
        self.daily_frame = ttk.Frame(canvas)
        self.daily_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self.daily_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        self.daily_texts = {}
        self.daily_labels = {}
        for i in range(7):
            lbl = ttk.Label(self.daily_frame, text=f"Day {i+1}", font=("Calibri", 9, "bold"))
            lbl.pack(anchor=tk.W, pady=(4, 0))
            self.daily_labels[i] = lbl
            txt = tk.Text(self.daily_frame, height=2, width=90, font=("Calibri", 9))
            txt.pack(fill=tk.X, pady=(0, 2))
            self.daily_texts[i] = txt

    def _update_end_date(self, event=None):
        try:
            y, m, d = int(self.start_year.get()), int(self.start_month.get()), int(self.start_day.get())
            start = datetime(y, m, d).date()
            end = start + timedelta(days=6)
            self.end_label.config(text=f"{end.day} {MONTH_ABBR[end.month-1]} {end.year}")
            for i in range(7):
                dd = start + timedelta(days=i)
                wd = WEEKDAYS[dd.weekday()]
                routine = "Daily check"
                if dd.weekday() == 4:
                    routine += " + Meal + Weapon"
                elif dd.weekday() == 1:
                    routine += " + Mileage + Security Report"
                self.daily_labels[i].config(text=f"{dd.day} {MONTH_ABBR[dd.month-1]}, {wd}  [Auto: {routine}]")
        except (ValueError, TypeError):
            pass

    def _set_current_week(self):
        today = datetime.now().date()
        start, end = get_report_period(today)
        self.start_year.set(str(start.year))
        self.start_month.set(str(start.month))
        self.start_day.set(str(start.day))
        self._update_end_date()

    # ---- Tab 2: Training ----
    def _build_tab_training(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Training")

        ttk.Label(tab, text="Enter training per category. Default PSD: H01, H02.\nFormat: - [Date] : [Group] ([Topic])",
                  font=("Calibri", 9)).pack(anchor=tk.W, pady=5)

        self.training_texts = {}
        for label in TRAINING_LABELS:
            lf = ttk.LabelFrame(tab, text=label, padding=5)
            lf.pack(fill=tk.X, pady=2)
            txt = tk.Text(lf, height=2, width=90, font=("Calibri", 9))
            txt.pack(fill=tk.X)
            txt.insert("1.0", "N/A")
            self.training_texts[label] = txt

    # ---- Tab 3: Issues + Client Feedback (combined) ----
    def _build_tab_issues_feedback(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Issues / Client Feedback")

        # 5.1 Issues
        lf1 = ttk.LabelFrame(tab, text="5.1 Issues (Korean OK - auto-translated to English)", padding=5)
        lf1.pack(fill=tk.BOTH, expand=True, pady=3)

        self.issues_tree = ttk.Treeview(lf1, columns=("issue", "summary", "actions"), show="headings", height=3)
        self.issues_tree.heading("issue", text="Issue")
        self.issues_tree.heading("summary", text="Summary")
        self.issues_tree.heading("actions", text="Actions")
        self.issues_tree.column("issue", width=150)
        self.issues_tree.column("summary", width=400)
        self.issues_tree.column("actions", width=250)
        self.issues_tree.pack(fill=tk.BOTH, expand=True)
        bf1 = ttk.Frame(lf1)
        bf1.pack(fill=tk.X, pady=3)
        ttk.Button(bf1, text="+ Add", command=self._add_issue).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf1, text="- Remove", command=lambda: self._del_item(self.issues_tree)).pack(side=tk.LEFT, padx=2)

        # 6. Client Feedback - FREE TEXT input
        lf2 = ttk.LabelFrame(tab, text="6. Client Feedback (Free text - auto-structured into Issue/Summary/Actions)", padding=5)
        lf2.pack(fill=tk.BOTH, expand=True, pady=3)

        ttk.Label(lf2, text="Write freely in Korean or English. The app will auto-generate title, summary bullets, and action items.",
                  font=("Calibri", 8, "italic")).pack(anchor=tk.W)
        self.feedback_text = tk.Text(lf2, height=8, width=90, font=("Calibri", 9))
        self.feedback_text.pack(fill=tk.BOTH, expand=True)

    def _add_issue(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Issue")
        dlg.geometry("600x250")
        dlg.transient(self.root)
        dlg.grab_set()
        entries = []
        for label in ["Issue (title)", "Summary", "Actions"]:
            ttk.Label(dlg, text=label + ":").pack(anchor=tk.W, padx=10, pady=(5, 0))
            e = ttk.Entry(dlg, width=70)
            e.pack(padx=10, fill=tk.X)
            entries.append(e)

        def save():
            self.issues_tree.insert("", tk.END, values=tuple(e.get() for e in entries))
            dlg.destroy()
        ttk.Button(dlg, text="Add", command=save).pack(pady=10)

    def _del_item(self, tree):
        for item in tree.selection():
            tree.delete(item)

    # ---- Tab 4: Vehicle + Finance ----
    def _build_tab_vehicle_finance(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Mileage / Finance")

        # Mileage
        lf1 = ttk.LabelFrame(tab, text="5.4 Vehicle Mileage (values saved between sessions)", padding=10)
        lf1.pack(fill=tk.X, pady=3)

        self.mileage_vars = {}
        for v in VEHICLES:
            plate = v["plate"]
            row = ttk.Frame(lf1)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=f"VPN {plate}:", width=12, font=("Calibri", 10, "bold")).pack(side=tk.LEFT)
            ttk.Label(row, text="Mileage:").pack(side=tk.LEFT, padx=(10, 0))
            mv = tk.StringVar(value=self.config.get(f"vehicle_{plate}_mileage"))
            ttk.Entry(row, textvariable=mv, width=10).pack(side=tk.LEFT, padx=2)
            ttk.Label(row, text="Next Svc:").pack(side=tk.LEFT, padx=(10, 0))
            nv = tk.StringVar(value=self.config.get(f"vehicle_{plate}_next_service"))
            ttk.Entry(row, textvariable=nv, width=10).pack(side=tk.LEFT, padx=2)
            ttk.Label(row, text="Comments:").pack(side=tk.LEFT, padx=(10, 0))
            cv = tk.StringVar(value=self.config.get(f"vehicle_{plate}_comments"))
            ttk.Entry(row, textvariable=cv, width=30).pack(side=tk.LEFT, padx=2)
            self.mileage_vars[plate] = {"current": mv, "next_service": nv, "comments": cv}

        # Finance
        lf2 = ttk.LabelFrame(tab, text="5.8 Finance (IQD auto-added, just enter number)", padding=10)
        lf2.pack(fill=tk.BOTH, expand=True, pady=3)

        self.finance_tree = ttk.Treeview(lf2, columns=("date", "pr", "desc", "amount", "balance"), show="headings", height=5)
        self.finance_tree.heading("date", text="Date")
        self.finance_tree.heading("pr", text="PR/RV#")
        self.finance_tree.heading("desc", text="Description")
        self.finance_tree.heading("amount", text="Amount")
        self.finance_tree.heading("balance", text="Balance")
        self.finance_tree.column("date", width=100)
        self.finance_tree.column("pr", width=80)
        self.finance_tree.column("desc", width=250)
        self.finance_tree.column("amount", width=120)
        self.finance_tree.column("balance", width=120)
        self.finance_tree.pack(fill=tk.BOTH, expand=True)
        bf = ttk.Frame(lf2)
        bf.pack(fill=tk.X, pady=3)
        ttk.Button(bf, text="+ Add", command=self._add_finance).pack(side=tk.LEFT, padx=2)
        ttk.Button(bf, text="- Remove", command=lambda: self._del_item(self.finance_tree)).pack(side=tk.LEFT, padx=2)

    def _add_finance(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Add Finance Entry")
        dlg.geometry("500x300")
        dlg.transient(self.root)
        dlg.grab_set()
        fields = ["Date (e.g. 01 Mar)", "PR/RV Number", "Description (Korean OK)", "Amount (number only, IQD auto)", "Balance (number only, IQD auto)"]
        entries = []
        for label in fields:
            ttk.Label(dlg, text=label + ":").pack(anchor=tk.W, padx=10, pady=(3, 0))
            e = ttk.Entry(dlg, width=50)
            e.pack(padx=10, fill=tk.X)
            entries.append(e)

        def save():
            vals = [e.get() for e in entries]
            self.finance_tree.insert("", tk.END, values=tuple(vals))
            dlg.destroy()
        ttk.Button(dlg, text="Add", command=save).pack(pady=10)

    # ---- Data Collection ----
    def _collect_data(self):
        data = {}
        try:
            y, m, d = int(self.start_year.get()), int(self.start_month.get()), int(self.start_day.get())
            data["period_start"] = datetime(y, m, d).date()
            data["period_end"] = data["period_start"] + timedelta(days=6)
        except (ValueError, TypeError):
            raise ValueError("Please enter a valid date.")

        # Shifts
        shifts = []
        for dv, sv in self.shift_entries:
            day, shift = dv.get().strip(), sv.get().strip()
            if day and shift:
                shifts.append({"date": int(day), "shift": int(shift)})
        data["shift_changes"] = shifts

        # Daily extras
        start = data["period_start"]
        daily_extras = {}
        for i in range(7):
            dd = start + timedelta(days=i)
            text = self.daily_texts[i].get("1.0", tk.END).strip()
            if text:
                daily_extras[str(dd.day)] = text
        data["daily_extras"] = daily_extras

        # Training
        training = {}
        for label in TRAINING_LABELS:
            text = self.training_texts[label].get("1.0", tk.END).strip()
            if text and text != "N/A":
                training[label] = text
        data["training"] = training

        # Issues
        issues = []
        for item in self.issues_tree.get_children():
            v = self.issues_tree.item(item, "values")
            issues.append({"issue": v[0], "summary": v[1], "actions": v[2]})
        data["issues"] = issues

        # Mileage
        mileage = {}
        for plate in [v["plate"] for v in VEHICLES]:
            mileage[plate] = {
                "current": self.mileage_vars[plate]["current"].get().strip(),
                "next_service": self.mileage_vars[plate]["next_service"].get().strip(),
                "comments": self.mileage_vars[plate]["comments"].get().strip(),
            }
            for k, v in mileage[plate].items():
                self.config.set(f"vehicle_{plate}_{k}", v)
        data["mileage"] = mileage

        # Finance
        finance = []
        for item in self.finance_tree.get_children():
            v = self.finance_tree.item(item, "values")
            finance.append({"date": v[0], "pr_number": v[1], "description": v[2], "amount": v[3], "balance": v[4]})
        data["finance"] = finance

        # Client Feedback - parse free text into structured format
        fb_text = self.feedback_text.get("1.0", tk.END).strip()
        feedback = []
        if fb_text:
            feedback = self._parse_client_feedback(fb_text)
        data["client_feedback"] = feedback

        return data

    def _parse_client_feedback(self, text):
        """Parse free-form text into structured Client Feedback entries."""
        text = translate_ko_to_en(text)
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if not lines:
            return []

        # Generate a title from the first line or first sentence
        first_line = lines[0]
        # Try to extract a short title (first phrase or up to first comma/period)
        title_match = re.match(r'^(.{10,50}?)[,.\-]', first_line)
        if title_match:
            title = title_match.group(1).strip()
        else:
            title = first_line[:50].strip()
            if len(first_line) > 50:
                title += "..."

        # Format as summary bullets
        summary_lines = []
        actions_lines = []
        for line in lines:
            if not line.startswith("- "):
                line = "- " + line
            # Detect action-oriented lines
            action_keywords = ["monitor", "prepar", "standby", "ready", "await", "confirm", "ensur", "coordinate", "ready", "alert"]
            is_action = any(kw in line.lower() for kw in action_keywords)
            if is_action:
                actions_lines.append(line)
            else:
                summary_lines.append(line)

        if not actions_lines:
            actions_lines = ["- Monitoring situation continuously."]

        return [{"issue": title, "summary": "\n".join(summary_lines), "actions": "\n".join(actions_lines)}]

    # ---- Generate ----
    def _generate_report(self):
        try:
            self.status_var.set("Collecting data...")
            self.root.update()
            data = self._collect_data()

            # AI Translation step
            if _api_client:
                self.status_var.set("Translating Korean → English (AI)...")
                self.root.update()
                data = translate_all_fields(data)

            self.status_var.set("Generating report...")
            self.root.update()
            output_path = self.generator.generate(data)
            self.config.save()

            self.status_var.set("Report generated successfully!")
            messagebox.showinfo("Success", f"Report generated:\n\n{output_path}")
            os.startfile(output_path)
        except Exception as e:
            self.status_var.set("Error occurred")
            messagebox.showerror("Error", f"Error:\n\n{str(e)}")
            import traceback
            traceback.print_exc()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    app = WeeklyReportApp()
    app.run()
