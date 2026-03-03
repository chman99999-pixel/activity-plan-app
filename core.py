"""
core.py — 월별 활동계획서 자동 작성 비즈니스 로직

달력형 계획서(.xls)를 파싱하여 이용자별 엑셀 활동계획서(.xlsx)에
활동 내용을 자동으로 입력한다.
"""
import io
import re
from datetime import date

import xlrd
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Alignment


# ─── 1. 달력 파싱 ───

def parse_calendar(xls_bytes: bytes):
    """달력 .xls에서 날짜별 활동 내용을 추출한다.

    Returns:
        activities: {날짜(int): [시간대별 문자열, ...]}
        holidays: set of 날짜(int) — 대체공휴일 등
        month: int — 해당 월
        year: int — 해당 연도
    """
    wb = xlrd.open_workbook(file_contents=xls_bytes)
    sheet = wb.sheet_by_index(0)
    sheet_name = sheet.name

    # 월 추출
    month = None
    for ch in sheet_name:
        if ch.isdigit():
            if month is None:
                month = int(ch)
            else:
                month = month * 10 + int(ch)
    if month and month > 12:
        month = month % 100

    # 연도 추출 시도
    year = None
    for row in range(min(3, sheet.nrows)):
        for col in range(sheet.ncols):
            val = str(sheet.cell(row, col).value)
            if '20' in val:
                for word in val.split():
                    cleaned = ''.join(c for c in word if c.isdigit())
                    if len(cleaned) == 4 and cleaned.startswith('20'):
                        year = int(cleaned)
                        break
            if year:
                break
        if year:
            break
    if not year:
        year = date.today().year

    activities = {}
    holidays = set()
    time_slots = [
        '09:00~10:00', '10:00~11:00', '11:00~12:00',
        '12:00~13:00', '13:00~14:00', '14:00~15:00', '15:00~16:00'
    ]

    row = 0
    while row < sheet.nrows:
        row_vals = [str(sheet.cell(row, c).value).strip() for c in range(sheet.ncols)]
        is_week_header = any('주차' in v for v in row_vals)

        if is_week_header:
            date_row = None
            time_start_row = None
            for dr in range(row + 1, min(row + 4, sheet.nrows)):
                vals = [str(sheet.cell(dr, c).value).strip() for c in range(sheet.ncols)]
                if any('시간' in v for v in vals):
                    continue
                has_date = False
                for c in range(2, 7):
                    v = str(sheet.cell(dr, c).value).strip()
                    nums = ''.join(ch for ch in v.split('일')[0].split('(')[0] if ch.isdigit())
                    if nums and 1 <= int(nums) <= 31:
                        has_date = True
                        break
                if has_date:
                    date_row = dr
                    time_start_row = dr + 1
                    break

            if date_row is None:
                row += 1
                continue

            dates_in_week = {}
            for c in range(2, 7):
                v = str(sheet.cell(date_row, c).value).strip()
                if not v:
                    continue
                nums = ''.join(ch for ch in v.split('일')[0].split('(')[0] if ch.isdigit())
                if nums and 1 <= int(nums) <= 31:
                    d = int(nums)
                    dates_in_week[c] = d
                    if '대체공휴일' in v or '공휴일' in v:
                        holidays.add(d)

            for c, d in dates_in_week.items():
                if d in holidays:
                    continue
                first_val = str(sheet.cell(time_start_row, c).value).strip() if time_start_row < sheet.nrows else ''
                if '대체공휴일' in first_val or '공휴일' in first_val:
                    holidays.add(d)
                    continue

                day_activities = []
                for slot_idx, slot_time in enumerate(time_slots):
                    r = time_start_row + slot_idx
                    if r >= sheet.nrows:
                        break
                    val = str(sheet.cell(r, c).value).strip()
                    if val and val != 'None':
                        day_activities.append(f"{slot_time} {val}")
                    else:
                        if slot_time == '12:00~13:00':
                            day_activities.append(f"{slot_time} 점심식사 및 위생지원")
                        else:
                            day_activities.append(f"{slot_time} ")

                if day_activities:
                    activities[d] = day_activities

            row = time_start_row + len(time_slots)
        else:
            row += 1

    return activities, holidays, month, year


# ─── 2. 이용자 감지 ───

def detect_users(xlsx_bytes: bytes) -> list[str]:
    """템플릿 xlsx의 시트명에서 이용자 이름을 추출한다."""
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    users = []
    for sn in wb.sheetnames:
        # 시트명에서 한글 이름 추출 (예: "26.03월 계획서-유정빈" → "유정빈")
        m = re.search(r'-([가-힣]{2,4})$', sn.strip())
        if m:
            users.append(m.group(1))
    wb.close()
    return users


# ─── 3. 행 수 검증 ───

def count_available_rows(xlsx_bytes: bytes) -> int:
    """첫 번째 이용자 시트에서 데이터 입력 가능한 행 수를 확인한다.
    행 9부터 시작, 행 29까지 (30행은 수식)."""
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    if not wb.sheetnames:
        wb.close()
        return 0
    ws = wb[wb.sheetnames[0]]
    # 기본적으로 9~29행 = 21행이 입력 가능
    # 실제 행 수를 세서 반환
    count = 0
    for row in range(9, 30):  # 행 9~29
        try:
            cell = ws.cell(row=row, column=1)
            count += 1
        except Exception:
            break
    wb.close()
    return count


# ─── 4. 원본 폰트 정보 확인 ───

def get_font_info(xlsx_bytes: bytes):
    """원본 xlsx에서 D열 폰트 정보를 확인한다."""
    wb = load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb[wb.sheetnames[0]]
    cell = ws.cell(row=9, column=4)
    font_name = cell.font.name or '맑은 고딕'
    font_size = cell.font.size or 14
    wb.close()
    return font_name, font_size


# ─── 5. 오후 송영 시간 계산 ───

def _get_last_end_hour(day_activities: list[str]) -> int:
    """하루 활동 목록에서 마지막 활동의 종료 시간(시)을 반환한다.
    예: '15:00~16:00 활동' → 16, '16:00~17:00 활동' → 17
    기본값: 16
    """
    last_hour = 16
    for act in day_activities:
        m = re.match(r'\d{2}:\d{2}~(\d{2}):\d{2}', act)
        if m:
            last_hour = int(m.group(1))
    return last_hour


# ─── 6. 활동계획 입력 ───

def fill_sheets(template_bytes: bytes, activities: dict, holidays: set,
                user_config: dict, provider: str, month: int, year: int):
    """이용자별 시트에 활동계획을 입력한다.

    Returns:
        output_bytes: 완성된 xlsx 바이트
        results: 이용자별 결과 목록
        working_days: 활동일 목록
        formulas_ok: 수식 보존 여부
    """
    wb = load_workbook(io.BytesIO(template_bytes))

    font_name, font_size = get_font_info(template_bytes)
    normal_font = InlineFont(rFont=font_name, sz=font_size)
    red_font = InlineFont(rFont=font_name, sz=font_size, color='FFFF0000')

    def make_cell_value(text):
        if '(협)' not in text:
            return text
        parts = []
        segments = text.split('(협)')
        for i, seg in enumerate(segments):
            if seg:
                parts.append(TextBlock(normal_font, seg))
            if i < len(segments) - 1:
                parts.append(TextBlock(red_font, '(협)'))
        return CellRichText(*parts)

    # 활동일 목록 (공휴일 제외, 정렬)
    working_days = sorted([d for d in activities.keys() if d not in holidays])

    # 요일 매핑
    dow_names_kr = ['월', '화', '수', '목', '금', '토', '일']
    dow_map = {}
    for d in working_days:
        try:
            dt = date(year, month, d)
            dow_map[d] = dow_names_kr[dt.weekday()]
        except ValueError:
            dow_map[d] = ''

    results = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]

        # 이용자 이름 매칭
        user_name = None
        for name in user_config:
            if name in sheet_name:
                user_name = name
                break
        if not user_name:
            continue

        config = user_config[user_name]
        has_오전송영 = config.get('오전송영', False)
        오전송영시간 = config.get('오전송영시간', '08:30~09:00 송영')
        has_오후송영 = config.get('오후송영', False)
        오후송영시간 = config.get('오후송영시간', '16:00~16:30 송영')
        shuttle_count = int(has_오전송영) + int(has_오후송영)
        row_height = 167 + 21 * shuttle_count

        for i, d in enumerate(working_days):
            row = 9 + i

            # A: 날짜, B: 요일, C: 제공자
            for col, val in [(1, d), (2, dow_map.get(d, '')), (3, provider)]:
                cell = ws.cell(row=row, column=col)
                if not isinstance(cell, MergedCell):
                    cell.value = val

            # D: 활동계획
            day_acts = activities.get(d, [])
            lines = []
            if has_오전송영:
                lines.append(오전송영시간)
            lines.extend(day_acts)
            if has_오후송영:
                lines.append(오후송영시간)
            full_text = '\n'.join(lines)

            cell_d = ws.cell(row=row, column=4)
            if not isinstance(cell_d, MergedCell):
                cell_d.value = make_cell_value(full_text)
                cell_d.alignment = Alignment(wrap_text=True, vertical='center')

            ws.row_dimensions[row].height = row_height

        # 남은 행 비우기
        for row in range(9 + len(working_days), 30):
            for col in range(1, 16):
                cell = ws.cell(row=row, column=col)
                if isinstance(cell, MergedCell):
                    continue
                val = cell.value
                if val and isinstance(val, str) and val.startswith('='):
                    continue  # 수식 보존
                cell.value = None

        ws.column_dimensions['D'].width = 30

        results.append({
            'name': user_name,
            'sheet': sheet_name,
            '오전송영': has_오전송영,
            '오후송영': has_오후송영,
            'days': len(working_days)
        })

    # 저장
    buf = io.BytesIO()
    wb.save(buf)
    output_bytes = buf.getvalue()

    # 수식 보존 검증
    wb_check = load_workbook(io.BytesIO(output_bytes))
    formulas_ok = True
    for sn in wb_check.sheetnames:
        ws = wb_check[sn]
        for col in [12, 13, 14, 15]:
            val = ws.cell(row=30, column=col).value
            if val and isinstance(val, str) and val.startswith('='):
                continue
            else:
                formulas_ok = False
    wb_check.close()

    return output_bytes, results, working_days, formulas_ok
