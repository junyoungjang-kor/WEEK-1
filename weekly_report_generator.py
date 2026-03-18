#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Hanwha BNCP Weekly Security Report Generator v4.0
- Claude AI API integration for Korean → English translation
- Bullet-point (개조식) professional security report style
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

# Translation system prompt for security report style
TRANSLATION_SYSTEM_PROMPT = """You are a translator for a security company's weekly report at Hanwha BNCP project site in Iraq.
Translate Korean input to professional English suitable for a formal security weekly report.

CRITICAL RULES:
- Use BULLET-POINT style (개조식): short, concise phrases. NOT full sentences.
  Example: "L1005 Ihsan Ali - absent for 3 days due to medical treatment" (NOT "L1005 Ihsan Ali was absent from work for a period of three days because he needed to receive medical treatment at a hospital.")
- Keep person names, location names, ID numbers (L1005, RV920) EXACTLY as-is
- Use standard security/military terminology: PSD, SSG, NSTR, AMS HQ, PMT
- PSD teams are H01 and H02 unless stated otherwise
- Currency is always IQD (Iraqi Dinar)
- Dates in "DD Mon" format (e.g., "11 Mar")
- Be concise: remove unnecessary words, keep only essential information
- If input has bullet markers (- or ·), keep them as "- " bullets
- Output ONLY the translated English text, no explanations or extra information
- Do NOT add any information not present in the original text
- For personnel status: use format "ID Name - status description"
- For training: use format "DD Mon - Team/Group conducted [activity]"
"""

BATCH_TRANSLATION_PROMPT = """You are a translator for a security company's weekly report at Hanwha BNCP project site in Iraq.
Translate each numbered Korean text to professional English for a formal security weekly report.

CRITICAL RULES:
- Use BULLET-POINT style (개조식): short, concise phrases. NOT full sentences.
  Example: "L1005 Ihsan Ali - 3 days absent, medical treatment" (NOT "L1005 Ihsan Ali has been absent from his duties for three days to receive medical treatment.")
- Keep person names, location names, ID numbers (L1005, RV920) EXACTLY as-is
- Security terminology: PSD, SSG, NSTR, AMS HQ, PMT
- PSD teams: H01, H02
- Currency: IQD
- Dates: "DD Mon" format
- Be maximally concise
- Keep "- " bullet formatting if present
- Output format: [1] translated text
[2] translated text
...
- Output ONLY translations with their numbers, no extra text
- Do NOT add information not in the original
- If a numbered item has multiple lines, translate each line and keep them under the same number separated by newlines"""


# Global API client
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


def has_korean(text):
    """Check if text contains Korean characters."""
    if not text or not isinstance(text, str):
        return False
    return any('\uac00' <= c <= '\ud7a3' for c in text)


def translate_ko_to_en(text):
    """Translate Korean text to English using Claude API (single item)."""
    if not text or not has_korean(text):
        return text
    if not _api_client:
        return text
    try:
        response = _api_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1024,
            messages=[{"role": "user", "content": text}],
            system=TRANSLATION_SYSTEM_PROMPT
        )
        return response.content[0].text.strip()
    except Exception as e:
        print(f"API translation error: {e}")
        return text


def translate_all_fields(data_dict):
    """Batch translate all Korean text fields in a data dictionary.
    Falls back to individual translation if batch fails."""
    if not _api_client:
        return data_dict

    # Collect all Korean texts that need translation
    texts_to_translate = []
    text_paths = []

    def collect_texts(obj, path=""):
        if isinstance(obj, str) and has_korean(obj):
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

    # Try batch translation first
    translations = {}
    numbered_texts = "\n".join(f"[{i+1}] {t}" for i, t in enumerate(texts_to_translate))
    try:
        response = _api_client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=4096,
            messages=[{"role": "user", "content": numbered_texts}],
            system=BATCH_TRANSLATION_PROMPT
        )
        result_text = response.content[0].text.strip()

        # Parse results - handle multi-line translations per number
        current_idx = None
        current_lines = []
        for line in result_text.split("\n"):
            m = re.match(r'\[(\d+)\]\s*(.*)', line)
            if m:
                # Save previous item
                if current_idx is not None:
                    translations[current_idx] = "\n".join(current_lines).strip()
                current_idx = int(m.group(1)) - 1
                current_lines = [m.group(2).strip()]
            elif current_idx is not None:
                current_lines.append(line)
        # Save last item
        if current_idx is not None:
            translations[current_idx] = "\n".join(current_lines).strip()

    except Exception as e:
        print(f"Batch translation error: {e}")

    # Fallback: individually translate any items that weren't in batch result
    for i, text in enumerate(texts_to_translate):
        if i not in translations or not translations[i]:
            translations[i] = translate_ko_to_en(text)

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


def auto_format_finance(amount_str):
    """Auto-add IQD prefix and format amount."""
    amount_str = amount_str.strip()
    if not amount_str:
        return ""
    amount_str = amount_str.replace("IQD", "").strip()
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
    """Generates .docx by modifying template XML.
    NOTE: All translation must be done BEFORE calling generate().
    This class does NOT translate - it only writes pre-translated data."""

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
        """Update MONTH and PERIOD in header."""
        start = data["period_start"]
        end = data["period_end"]
        month_name = MONTHS[start.month - 1]
        month_abbr = MONTH_ABBR[start.month - 1]

        # MONTH: Find any "MonthName YYYY" pattern in <w:t> and replace
        for m in MONTHS:
            for year in range(2024, 2031):
                old_text = f"<w:t>{m} {year}</w:t>"
                if old_text in content:
                    new_text = f"<w:t>{month_name} {start.year}</w:t>"
                    content = content.replace(old_text, new_text, 1)
                    break
                # Also check with xml:space preserve
                old_text2 = f'<w:t xml:space="preserve">{m} {year}</w:t>'
                if old_text2 in content:
                    new_text2 = f'<w:t xml:space="preserve">{month_name} {start.year}</w:t>'
                    content = content.replace(old_text2, new_text2, 1)
                    break
            else:
                continue
            break

        # PERIOD: Find paragraph containing "PERIOD:" and rebuild date portion
        period_idx = content.find("PERIOD:")
        if period_idx == -1:
            return content

        period_p_start = content.rfind("<w:p ", 0, period_idx)
        period_p_end = content.index("</w:p>", period_idx) + 6
        period_para = content[period_p_start:period_p_end]

        # Find the run containing "PERIOD:" and keep everything up to its closing </w:r>
        period_run_end = period_para.index("PERIOD:")
        period_run_end = period_para.index("</w:r>", period_run_end) + 6

        # Keep paragraph up to end of PERIOD: run, then add single new run with date text
        before_period_runs = period_para[:period_run_end]
        date_text = f" {start.day} {month_abbr} to {end.day} {MONTH_ABBR[end.month-1]} {end.year}"
        new_run = f'<w:r><w:rPr><w:rFonts w:cs="Calibri"/><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr><w:t>{date_text}</w:t></w:r>'
        new_para = before_period_runs + new_run + "</w:p>"

        content = content[:period_p_start] + new_para + content[period_p_end:]
        return content

    def _modify_weekly_summary(self, content, data):
        """Build weekly summary section - NO translation here, data is pre-translated."""
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

            # Shift changes - compare full date (year, month, day), not just day number
            for sc in shifts:
                sc_date = sc.get("full_date")
                if sc_date and sc_date == day_date:
                    paragraphs.append(XmlTemplates.bullet_item(
                        f"SSG shift change completed (Shift #{sc['shift']})", num_id))
                elif not sc_date and sc.get("date") == day_date.day:
                    # Fallback: compare day only (backward compatibility)
                    paragraphs.append(XmlTemplates.bullet_item(
                        f"SSG shift change completed (Shift #{sc['shift']})", num_id))

            # Routine items
            if day_date.weekday() == 4:  # Friday
                paragraphs.append(XmlTemplates.bullet_item("Submitted Meal Request to Hanwha", num_id))
                paragraphs.append(XmlTemplates.bullet_item("Submitted Weekly Weapon & Ammunition Status to AMS HQ", num_id))
            elif day_date.weekday() == 1:  # Tuesday
                paragraphs.append(XmlTemplates.bullet_item("Submitted Weekly Mileage & Weekly Security Report to AMS HQ", num_id))

            # Extra items - already translated, just write directly
            day_key = str(day_date.day)
            if day_key in daily_extras and daily_extras[day_key].strip():
                for line in daily_extras[day_key].strip().split("\n"):
                    if line.strip():
                        paragraphs.append(XmlTemplates.bullet_item(line.strip(), num_id))

        new_tc = tc_header + "\n" + "\n".join(paragraphs) + "\n</w:tc>"
        new_row = row_xml[:tc_start] + new_tc + row_xml[tc_end:]
        new_table = table_xml[:second_tr_start] + new_row + table_xml[second_tr_end:]
        content = content[:tbl_start] + new_table + content[tbl_end:]
        return content

    def _modify_training(self, content, data):
        """Modify training table - data is pre-translated."""
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
                        paras.append(XmlTemplates.training_cell_para(line.strip()))
                new_tc = f"<w:tc>{tcpr}{''.join(paras)}</w:tc>"
            else:
                new_tc = f"<w:tc>{tcpr}{XmlTemplates.training_cell_para('N/A')}</w:tc>"

            new_row = row_xml[:second_tc_start] + new_tc + row_xml[second_tc_end:]
            table_xml = table_xml[:row_match.start()] + new_row + table_xml[row_match.end():]

        content = content[:tbl_start] + table_xml + content[tbl_end:]
        return content

    def _modify_issues(self, content, data):
        """Modify issues table - data is pre-translated."""
        issues = data.get("issues", [])

        # Safe search for "Issues" marker near "5.1"
        marker_idx = content.find("Issues")
        if marker_idx == -1:
            return content

        check_area = content[max(0, marker_idx - 200):marker_idx]
        if "5.1" not in check_area:
            next_idx = content.find("Issues", marker_idx + 1)
            if next_idx == -1:
                return content  # Cannot find proper Issues section, skip safely
            marker_idx = next_idx

        tbl_start = content.find("<w:tbl>", marker_idx)
        if tbl_start == -1:
            return content

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
                    iss.get("issue", ""),
                    iss.get("summary", ""),
                    iss.get("actions", "")
                ))
        new_rows.append(XmlTemplates.issues_empty_row())

        tbl_grid_end = table_xml.index("</w:tblGrid>") + 12
        new_table = table_xml[:tbl_grid_end] + "\n" + "\n".join(new_rows) + "\n</w:tbl>"
        content = content[:tbl_start] + new_table + content[tbl_end:]
        return content

    def _modify_mileage(self, content, data):
        """Modify vehicle mileage table."""
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

            # Use more robust tc matching (handles both <w:tc> and <w:tc ...>)
            tcs = list(re.finditer(r'<w:tc[ >].*?</w:tc>', row_xml, re.DOTALL))
            if len(tcs) < 4:
                continue

            # Mileage (col 2)
            if values.get("current"):
                old_tc = tcs[1].group()
                new_tc = re.sub(r'<w:t[^>]*>[^<]*</w:t>', f'<w:t>{values["current"]}</w:t>', old_tc, count=1)
                row_xml = row_xml[:tcs[1].start()] + new_tc + row_xml[tcs[1].end():]
                tcs = list(re.finditer(r'<w:tc[ >].*?</w:tc>', row_xml, re.DOTALL))

            # Next Service (col 3)
            if values.get("next_service") and len(tcs) >= 3:
                old_tc = tcs[2].group()
                new_tc = re.sub(r'<w:t[^>]*>[^<]*</w:t>', f'<w:t>{values["next_service"]}</w:t>', old_tc)
                row_xml = row_xml[:tcs[2].start()] + new_tc + row_xml[tcs[2].end():]
                tcs = list(re.finditer(r'<w:tc[ >].*?</w:tc>', row_xml, re.DOTALL))

            # Comments (col 4)
            if values.get("comments") and len(tcs) >= 4:
                old_tc = tcs[3].group()
                new_tc = re.sub(r'<w:t[^>]*>[^<]*</w:t>', f'<w:t>{xml_escape(values["comments"])}</w:t>', old_tc, count=1)
                row_xml = row_xml[:tcs[3].start()] + new_tc + row_xml[tcs[3].end():]

            content = content[:tr_start] + row_xml + content[tr_end:]
        return content

    def _modify_finance(self, content, data):
        """Modify finance table - data is pre-translated."""
        finance = data.get("finance", [])

        # Safe search for finance marker
        marker_idx = content.find("5.8 Finance")
        if marker_idx == -1:
            # Try split version: "5.8" near "Finance"
            marker_idx = content.find("Finance")
            if marker_idx == -1:
                return content
            check_area = content[max(0, marker_idx - 100):marker_idx]
            if "5.8" not in check_area:
                return content

        tbl_start = content.find("<w:tbl>", marker_idx)
        if tbl_start == -1:
            return content

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
                    fin.get("description", ""),
                    auto_format_finance(fin.get("amount", "")),
                    auto_format_finance(fin.get("balance", ""))
                ))
        new_rows.append(XmlTemplates.finance_empty_row())

        tbl_grid_end = table_xml.index("</w:tblGrid>") + 12
        new_table = table_xml[:tbl_grid_end] + "\n" + "\n".join(new_rows) + "\n</w:tbl>"
        content = content[:tbl_start] + new_table + content[tbl_end:]
        return content

    def _modify_client_feedback(self, content, data):
        """Modify client feedback table - data is pre-translated."""
        feedback = data.get("client_feedback", [])

        marker_idx = content.find("6. Client Feedback")
        if marker_idx == -1:
            marker_idx = content.find("Client Feedback")
            if marker_idx == -1:
                return content

        tbl_start = content.find("<w:tbl>", marker_idx)
        if tbl_start == -1:
            return content

        tbl_end = content.index("</w:tbl>", tbl_start) + 8
        table_xml = content[tbl_start:tbl_end]

        rows = list(re.finditer(r'<w:tr\b[^>]*>.*?</w:tr>', table_xml, re.DOTALL))
        if not rows:
            return content
        header_row = rows[0].group()

        new_rows = [header_row]
        if feedback:
            for fb in feedback:
                summary_lines = [l.strip() for l in fb.get("summary", "").split("\n") if l.strip()]
                actions_lines = [l.strip() for l in fb.get("actions", "").split("\n") if l.strip()]
                new_rows.append(XmlTemplates.client_feedback_row(
                    fb.get("issue", ""),
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
        self.root.title("Hanwha BNCP Weekly Report Generator v4.0")
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
        api_frame = ttk.LabelFrame(main, text="Claude API (한국어 → 영어 자동 번역)", padding=5)
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
            try:
                test = translate_ko_to_en("보안 점검 완료")
                self.config.set("api_key", key)
                self.config.save()
                self.api_status.config(text="● Connected", foreground="green")
                self.status_var.set("Ready (AI Translation ON)")
                messagebox.showinfo("Success", f"API connected!\n\nTest: '보안 점검 완료' → '{test}'")
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
        df = ttk.LabelFrame(tab, text="Daily Events (한국어 입력 → 영어 자동 번역, 개조식 변환)", padding=10)
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

        ttk.Label(tab, text="한국어 입력 가능 → 자동 영어 번역\n형식 예: 3월 21일 psd 2개팀 드라이빙 어세스먼트 실시",
                  font=("Calibri", 9)).pack(anchor=tk.W, pady=5)

        self.training_texts = {}
        for label in TRAINING_LABELS:
            lf = ttk.LabelFrame(tab, text=label, padding=5)
            lf.pack(fill=tk.X, pady=2)
            txt = tk.Text(lf, height=2, width=90, font=("Calibri", 9))
            txt.pack(fill=tk.X)
            txt.insert("1.0", "N/A")
            self.training_texts[label] = txt

    # ---- Tab 3: Issues + Client Feedback ----
    def _build_tab_issues_feedback(self):
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Issues / Client Feedback")

        # 5.1 Issues
        lf1 = ttk.LabelFrame(tab, text="5.1 Issues (한국어 입력 → 영어 자동 번역)", padding=5)
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

        # 6. Client Feedback
        lf2 = ttk.LabelFrame(tab, text="6. Client Feedback (자유 텍스트 → Issue/Summary/Actions 자동 구조화)", padding=5)
        lf2.pack(fill=tk.BOTH, expand=True, pady=3)

        ttk.Label(lf2, text="한국어로 자유롭게 작성하세요. 자동으로 제목/요약/조치사항으로 분류됩니다.",
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
        for label in ["Issue (제목)", "Summary (요약)", "Actions (조치사항)"]:
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
        lf1 = ttk.LabelFrame(tab, text="5.4 Vehicle Mileage (값 자동 저장됨)", padding=10)
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
        lf2 = ttk.LabelFrame(tab, text="5.8 Finance (숫자만 입력, IQD 자동 추가)", padding=10)
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
        fields = ["Date (예: 01 Mar)", "PR/RV Number", "Description (한국어 가능)", "Amount (숫자만, IQD 자동)", "Balance (숫자만, IQD 자동)"]
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

        # Shifts - store full_date for accurate comparison
        shifts = []
        start = data["period_start"]
        for dv, sv in self.shift_entries:
            day_str, shift = dv.get().strip(), sv.get().strip()
            if day_str and shift:
                day_num = int(day_str)
                # Determine full date: find which day in the week matches
                full_date = None
                for i in range(7):
                    dd = start + timedelta(days=i)
                    if dd.day == day_num:
                        full_date = dd
                        break
                shifts.append({"date": day_num, "shift": int(shift), "full_date": full_date})
        data["shift_changes"] = shifts

        # Daily extras
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

        # Client Feedback - parse free text
        fb_text = self.feedback_text.get("1.0", tk.END).strip()
        feedback = []
        if fb_text:
            feedback = self._parse_client_feedback(fb_text)
        data["client_feedback"] = feedback

        return data

    def _parse_client_feedback(self, text):
        """Parse free-form text into structured Client Feedback entries.
        Translation happens later in translate_all_fields()."""
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        if not lines:
            return []

        # Generate title from first line
        first_line = lines[0]
        title_match = re.match(r'^(.{10,50}?)[,.\-]', first_line)
        if title_match:
            title = title_match.group(1).strip()
        else:
            title = first_line[:50].strip()
            if len(first_line) > 50:
                title += "..."

        # All lines as summary bullets
        summary_lines = []
        for line in lines:
            if not line.startswith("- "):
                line = "- " + line
            summary_lines.append(line)

        # Default action
        actions_lines = ["- Monitoring situation continuously."]

        return [{"issue": title, "summary": "\n".join(summary_lines), "actions": "\n".join(actions_lines)}]

    # ---- Generate ----
    def _generate_report(self):
        try:
            self.status_var.set("Collecting data...")
            self.root.update()
            data = self._collect_data()

            # AI Translation step - translate ALL Korean text at once
            if _api_client:
                self.status_var.set("Translating Korean → English (AI 개조식 번역 중)...")
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
