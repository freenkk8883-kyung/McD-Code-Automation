import streamlit as st
import pandas as pd
import re
import io
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment, Font, Protection
from openpyxl.formatting.rule import FormulaRule
from openpyxl.worksheet.datavalidation import DataValidation
from collections import Counter

st.set_page_config(page_title="McD 코드 자동화", page_icon="🍟", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }
.hero { background: linear-gradient(135deg, #DA291C 0%, #C0241A 100%); color: white; padding: 2rem 2.5rem; border-radius: 16px; margin-bottom: 2rem; }
.hero h1 { margin: 0; font-size: 1.7rem; font-weight: 700; }
.hero p  { margin: 0.3rem 0 0; font-size: 0.9rem; opacity: 0.85; }
.card { background: white; border-radius: 12px; padding: 1.5rem; margin-bottom: 1rem; border: 1px solid #EBEBEB; }
.card h3 { margin: 0 0 1rem; font-size: 0.95rem; font-weight: 600; }
.badge-ok      { background:#E8F5E9; color:#2E7D32; padding:3px 10px; border-radius:20px; font-size:0.78rem; font-weight:600; }
.badge-neutral { background:#F3F4F6; color:#374151; padding:3px 10px; border-radius:20px; font-size:0.78rem; font-weight:600; }
.stat-box { background:#FAFAFA; border:1px solid #EBEBEB; border-radius:10px; padding:1rem; text-align:center; }
.stat-num { font-size:1.8rem; font-weight:700; color:#DA291C; }
.stat-label { font-size:0.78rem; color:#6B7280; margin-top:2px; }
.stButton > button { background:#DA291C !important; color:white !important; border:none !important; border-radius:8px !important; font-weight:600 !important; width:100% !important; }
.stButton > button:hover { background:#B71C1C !important; }
hr { border:none; border-top:1px solid #F0F0F0; margin:1.5rem 0; }
</style>
""", unsafe_allow_html=True)


def safe(val):
    s = str(val).strip() if pd.notna(val) else ''
    return '' if s in ('nan', 'NaN') else s

def get_device_code(d):
    d = d.upper().replace(' ', '')
    if not d or d in ('', '-'): return 'A'
    if 'PC' in d and 'MOBILE' in d: return 'A'
    if 'PC' in d: return 'P'
    if 'MOBILE' in d: return 'M'
    if 'TV' in d or 'CTV' in d: return 'C'
    return 'A'

def parse_targeting_lines(raw, target_map):
    if not raw or raw in ('-', 'NonTargeting', 'Non-Targeting', 'nan'):
        return [{'gender': 'P', 'age': '1865', 'targeting': 'non', 'note': 'A'}]
    t = raw.replace('\n', ' ').strip()
    t_masked = re.sub(r'\([^)]*\)', lambda m: '(' + 'X' * len(m.group()[1:-1]) + ')', t)
    split_indices = [0]
    for m in re.finditer(r'(?<!\d)\s*(?=\d{1,2}[\.\)](?!\d))', t_masked):
        if m.start() > 0: split_indices.append(m.start())
    split_indices.append(len(t))
    parts = []
    for i in range(len(split_indices) - 1):
        chunk = t[split_indices[i]:split_indices[i+1]].strip()
        chunk = re.sub(r'^\d{1,2}[\.\)]\s*', '', chunk).strip()
        if chunk: parts.append(chunk)
    if not parts: parts = [t]
    results = []
    for item in parts:
        gm = re.search(r'\b([PMF])(\d{4})\b', item)
        if gm: gender, age = gm.group(1), gm.group(2)
        else:
            gm2 = re.search(r'\b([PMF])\b', item)
            gender = gm2.group(1) if gm2 else 'P'
            am2 = re.search(r'\b(\d{4})\b', item)
            age = am2.group(1) if am2 else '1865'
        cleaned = re.sub(r'[PMF]\d{4}\+?', '', item).strip().lstrip('+').strip()
        cleaned = re.sub(r'\([^)]*\)', '', cleaned).strip()
        cleaned = re.sub(r'^\[타겟팅\]\s*', '', cleaned).strip()
        first_kw = cleaned.split(',')[0].split('_')[0].strip()
        code = target_map.get(first_kw.upper()) or target_map.get(cleaned.upper()) or (first_kw if first_kw else cleaned)
        results.append({'gender': gender, 'age': age, 'targeting': code, 'note': 'A'})
    return results

def parse_creative_combined(raw):
    """
    G열 Creative에서 소재명 + 가로세로 + 초수 모두 추출
    소재명 컬럼(H열)이 없을 때 사용
    """
    if not raw or raw in ('-', 'nan', ''):
        return [{'name': '', 'orientation': 'H', 'seconds': ''}]
    lines = [l.strip() for l in raw.replace('\\n', '\n').split('\n')
             if l.strip() and l.strip() != '-']
    expanded = []
    for line in lines:
        if '가로/세로' in line or '세로/가로' in line:
            expanded.append(line.replace('가로/세로', '가로').replace('세로/가로', '가로'))
            expanded.append(line.replace('가로/세로', '세로').replace('세로/가로', '세로'))
        else:
            expanded.append(line)
    results = []
    for line in expanded:
        sec_m = re.search(r'(\d+)[\'\'""""]', line)
        seconds = sec_m.group(1) if sec_m else ''
        orientation = 'V' if ('세로' in line or 'Vertical' in line or 'vertical' in line) else 'H'
        name = line
        name = re.sub(r'\d+[\'\'""""]', '', name)
        name = re.sub(r'\([^)]*\)', '', name)
        name = re.sub(r'가로/?세로|세로/?가로|세로|가로', '', name)
        name = re.sub(r'\s+', ' ', name).strip().strip('&').strip('-').strip()
        results.append({'name': name, 'orientation': orientation, 'seconds': seconds})
    return results or [{'name': '', 'orientation': 'H', 'seconds': ''}]

def parse_creative_format_only(raw):
    """G열 Creative에서 가로세로 + 초수만 추출 (소재명은 H열에서 따로 가져올 때)"""
    if not raw or raw in ('-', 'nan', ''):
        return [{'orientation': 'H', 'seconds': ''}]
    lines = [l.strip() for l in raw.replace('\\n', '\n').split('\n')
             if l.strip() and l.strip() != '-']
    expanded = []
    for line in lines:
        if '가로/세로' in line or '세로/가로' in line:
            expanded.append(line.replace('가로/세로', '가로').replace('세로/가로', '가로'))
            expanded.append(line.replace('가로/세로', '세로').replace('세로/가로', '세로'))
        else:
            expanded.append(line)
    results = []
    for line in expanded:
        sec_m = re.search(r'(\d+)[\'\'""""]', line)
        seconds = sec_m.group(1) if sec_m else ''
        orientation = 'V' if ('세로' in line or 'Vertical' in line or 'vertical' in line) else 'H'
        results.append({'orientation': orientation, 'seconds': seconds})
    return results or [{'orientation': 'H', 'seconds': ''}]

def parse_creative_names(raw):
    """H열 소재명 컬럼 파싱 (날짜 패턴이면 빈칸)"""
    if not raw or raw in ('-', 'nan', ''):
        return ['']
    date_pattern = re.compile(r'^\d{1,4}[-/]\d{1,2}|^\d{4}-\d{2}-\d{2}|W\d+|^\d+/\d+\s*[~(]')
    if date_pattern.search(raw.strip()):
        return ['']
    items = re.split(r'[\n,]', raw.replace('\\n', '\n'))
    results = []
    for item in items:
        item = re.sub(r'\([^)]*\)', '', item)
        item = re.sub(r'\d+%?', '', item).strip()
        if item: results.append(item)
    return results if results else ['']

def load_data_raw(tool_bytes):
    wb = load_workbook(io.BytesIO(tool_bytes), data_only=True)
    sheet_names = wb.sheetnames
    media_map, product_map, target_map = {}, {}, {}

    # ── Media 시트: C열=매체명, G열=최종CODE ──────────────────────
    if 'Media' in sheet_names:
        ws_m = wb['Media']
        for r in range(2, ws_m.max_row + 1):
            name = ws_m.cell(row=r, column=3).value  # C열
            code = ws_m.cell(row=r, column=7).value  # G열
            if name and code and str(code).strip() not in ('최종 CODE', ''):
                media_map[str(name).strip().upper()] = str(code).strip()
    # ── 구버전 호환: DATA RAW 단일 시트 (기존 포맷) ─────────────
    elif 'DATA RAW' in sheet_names:
        df = pd.read_excel(io.BytesIO(tool_bytes), sheet_name='DATA RAW',
                           header=None, engine='openpyxl')
        for _, r in df.iloc[1:].iterrows():
            if safe(r[2]) and safe(r[6]):
                media_map[safe(r[2]).upper()] = safe(r[6])

    # ── Product 시트: C열=상품명, G열=최종CODE ────────────────────
    if 'Product' in sheet_names:
        ws_p = wb['Product']
        for r in range(2, ws_p.max_row + 1):
            name = ws_p.cell(row=r, column=3).value  # C열
            code = ws_p.cell(row=r, column=7).value  # G열
            if name and code and str(code).strip() not in ('최종 CODE', ''):
                product_map[str(name).strip().upper()] = str(code).strip()
    elif 'DATA RAW' in sheet_names:
        df = pd.read_excel(io.BytesIO(tool_bytes), sheet_name='DATA RAW',
                           header=None, engine='openpyxl')
        for _, r in df.iloc[1:].iterrows():
            if safe(r[9]) and safe(r[13]):
                product_map[safe(r[9]).upper()] = safe(r[13])

    # ── DATA RAW 시트: F열=타겟팅, G열=최종CODE ──────────────────
    if 'DATA RAW' in sheet_names:
        ws_r = wb['DATA RAW']
        # v2: F열=타겟팅, G열=최종CODE
        # 구버전: col19=타겟팅, col20=최종CODE
        header_f = ws_r.cell(row=1, column=6).value
        if header_f and '타겟팅' in str(header_f):
            # v2 구조
            for r in range(2, ws_r.max_row + 1):
                tgt  = ws_r.cell(row=r, column=6).value  # F: 타겟팅
                code = ws_r.cell(row=r, column=7).value  # G: 최종CODE
                if tgt and code and str(code).strip() not in ('최종 CODE', ''):
                    target_map[str(tgt).strip().upper()] = str(code).strip()
        else:
            # 구버전 구조
            df = pd.read_excel(io.BytesIO(tool_bytes), sheet_name='DATA RAW',
                               header=None, engine='openpyxl')
            for _, r in df.iloc[1:].iterrows():
                if safe(r[19]) and safe(r[20]):
                    target_map[safe(r[19]).upper()] = safe(r[20])

    def get_media_code(media):
        return media_map.get(media.strip().upper(), '')

    def get_product_code(adtype):
        return product_map.get(adtype.strip().upper(), '')

    return get_media_code, get_product_code, target_map

def parse_media_mix(mm_bytes):
    df = pd.read_excel(io.BytesIO(mm_bytes), sheet_name='Media Mix', header=None)
    header_row = None
    for i in range(20):
        vals = [safe(v) for v in df.iloc[i]]
        if 'Media' in vals and 'Ad type' in vals:
            header_row = i
            break
    if header_row is None:
        raise ValueError("헤더행을 찾을 수 없어요. 'Media'와 'Ad type' 컬럼을 확인해주세요.")

    col_media    = next((i for i, v in enumerate(df.iloc[header_row]) if safe(v) == 'Media'), None)
    col_adtype   = next((i for i, v in enumerate(df.iloc[header_row]) if safe(v) == 'Ad type'), None)
    col_device   = next((i for i, v in enumerate(df.iloc[header_row]) if safe(v) == 'Device'), None)
    col_targeting= next((i for i, v in enumerate(df.iloc[header_row]) if safe(v) == 'Targeting'), None)
    col_creative = next((i for i, v in enumerate(df.iloc[header_row]) if safe(v) == 'Creative'), None)

    # 소재명 컬럼 감지: Creative 다음 컬럼이 Schedule/Exp 시작이 아닌 경우
    col_creative2 = None
    if col_creative is not None:
        next_col = col_creative + 1
        next_hdr = safe(df.iloc[header_row, next_col]) if next_col < len(df.columns) else ''
        if next_hdr not in ('Schedule', 'Exp. Imps', 'CTR') and not next_hdr.startswith('Exp'):
            col_creative2 = next_col

    # 소재명 컬럼 존재 여부 플래그
    has_name_col = col_creative2 is not None

    date_raw  = safe(df.iloc[3, 4])
    month     = date_raw.split('~')[0].split('/')[0].zfill(2)
    date_code = '26' + month
    title = safe(df.iloc[1, 1])
    camp  = re.sub(r'\s*Campaign\s*', '', re.sub(r'^\d{4}\s+', '', re.sub(r'_Media Mix$', '', title).strip()).strip(), flags=re.IGNORECASE).strip()
    cname = camp.replace(' ', '_')

    data = df.iloc[header_row+1:].copy()
    data[col_media] = data[col_media].ffill()
    for col in [col_device, col_targeting, col_creative]:
        if col is not None: data[col] = data[col].ffill()
    if col_creative2: data[col_creative2] = data[col_creative2].ffill()

    def is_valid(r):
        adtype = safe(r[col_adtype]); media = safe(r[col_media])
        if not adtype or not media: return False
        if 'Total' in media or 'total' in media: return False
        try:
            float(adtype.replace(',', ''))
            return False
        except: pass
        if adtype == '-': return False
        return True

    actual = data[data.apply(is_valid, axis=1)].copy()
    return actual, date_code, camp, cname, col_media, col_adtype, col_device, col_targeting, col_creative, col_creative2, has_name_col

def build_code_rows(actual, date_code, camp, cname,
                    col_media, col_adtype, col_device,
                    col_targeting, col_creative, col_creative2, has_name_col,
                    get_media_code, get_product_code, target_map):
    code_rows = []
    for _, r in actual.iterrows():
        media     = safe(r[col_media])
        device_raw= safe(r[col_device]) if col_device else ''
        adtype    = safe(r[col_adtype]).replace('\n', ' ')
        tgt_raw   = safe(r[col_targeting]).replace('\n', ' ') if col_targeting else ''
        cr_raw    = safe(r[col_creative]) if col_creative else ''
        cr_name   = safe(r[col_creative2]) if col_creative2 else ''

        m_code = get_media_code(media)
        p_code = get_product_code(adtype)
        dev    = get_device_code(device_raw)
        tgt_list = parse_targeting_lines(tgt_raw, target_map)

        if has_name_col:
            # BA Q1 방식: G열=포맷(가로세로/초수), H열=소재명
            fmt_list  = parse_creative_format_only(cr_raw)
            name_list = parse_creative_names(cr_name)
            creative_list = [
                {'name': n, 'orientation': f['orientation'], 'seconds': f['seconds']}
                for f in fmt_list for n in name_list
            ]
        else:
            # QPC/McCrispy 방식: G열에서 소재명+포맷 모두 추출
            creative_list = parse_creative_combined(cr_raw)

        for tgt in tgt_list:
            for cr in creative_list:
                j_code = f"{date_code}_{m_code}_{p_code}_{cname}" if (m_code and p_code) else ''
                o_code = f"{tgt['gender']}_{tgt['age']}_{tgt['targeting']}_{tgt['note']}"
                u_code = f"{dev}_{cr['name']}_{cr['orientation']}_{cr['seconds']}_A"
                full   = f"{j_code}{o_code}{u_code}" if j_code else ''
                code_rows.append({
                    'date': date_code, 'media': media, 'product': adtype, 'campaign': camp,
                    'd_code': date_code, 'm_code': m_code, 'p_code': p_code, 'c_code': cname,
                    'j_code': j_code, 'gender': tgt['gender'], 'age': tgt['age'],
                    'targeting': tgt['targeting'], 'note': tgt['note'], 'o_code': o_code,
                    'device': dev, 'creative': cr['name'], 'orient': cr['orientation'],
                    'seconds': cr['seconds'], 'param': 'A', 'u_code': u_code,
                    'full': full, 'missing': not (m_code and p_code),
                })
    return code_rows

def get_formula_for_row(template_formulas, new_row, template_row=10):
    """template_row 수식을 new_row용으로 행번호만 치환 (절대참조 $는 유지)"""
    result = {}
    for col_letter, formula in template_formulas.items():
        new_formula = re.sub(
            r'(?<=[A-Z])(\d+)(?!\d)',
            lambda m: str(new_row) if int(m.group()) == template_row else m.group(),
            formula
        )
        result[col_letter] = new_formula
    return result

def write_excel(tool_bytes, code_rows):
    wb = load_workbook(io.BytesIO(tool_bytes))
    ws = wb['CODE']

    # 셀 잠금 해제 → 유효성검사는 유지하되 셀 편집 가능하도록
    # (원본 파일 셀이 전부 locked=True → 출력 파일에서도 수정 불가 원인)
    ws.protection.sheet = False
    for row in ws.iter_rows():
        for cell in row:
            cell.protection = Protection(locked=False)

    # 10행 수식 템플릿 추출
    # #REF! 오류 수식은 올바른 시트 참조로 교체
    sheet_names = wb.sheetnames
    media_sheet   = 'Media'   if 'Media'   in sheet_names else 'DATA RAW'
    product_sheet = 'Product' if 'Product' in sheet_names else 'DATA RAW'

    # 시트별 XLOOKUP 수식 정의
    if 'Media' in sheet_names:
        # v2: Media/Product 시트 분리 구조
        g_formula = f"=_xlfn.XLOOKUP(C10,{media_sheet}!$C$2:$C$5000,{media_sheet}!$G$2:$G$5000,\"\")"
        h_formula = f"=_xlfn.XLOOKUP(D10,{product_sheet}!$C$2:$C$5000,{product_sheet}!$G$2:$G$5000,\"\")"
    else:
        # 구버전: DATA RAW 단일 시트
        g_formula = "=_xlfn.XLOOKUP(C10,'DATA RAW'!$C$2:$C$4000,'DATA RAW'!$G$2:$G$4000,\"\")"
        h_formula = "=_xlfn.XLOOKUP(D10,'DATA RAW'!$J$2:$J$4000,'DATA RAW'!$N$2:$N$4000,\"\")"

    template_formulas = {}
    for col in range(1, 28):
        cell = ws.cell(row=10, column=col)
        if cell.value and str(cell.value).startswith('='):
            formula = str(cell.value)
            # #REF! 오류 수식 교체
            if col == 7 and '#REF!' in formula:
                formula = g_formula
            elif col == 8 and '#REF!' in formula:
                formula = h_formula
            template_formulas[cell.column_letter] = formula

    # 기존 데이터 삭제 (10행부터)
    for r in range(ws.max_row, 9, -1): ws.delete_rows(r)

    # 8행 헤더 음영
    header_fill = PatternFill(fill_type='solid', fgColor='FFC000')
    header_font = Font(bold=True, color='FFFFFF')
    headers = {
        2:'날짜', 3:'매체', 4:'상품', 5:'캠페인',
        6:'날짜', 7:'매체', 8:'상품', 9:'캠페인', 10:'CODE',
        11:'성별', 12:'연령', 13:'타겟팅', 14:'비고', 15:'CODE',
        16:'Device', 17:'소재', 18:'가로세로', 19:'초수', 20:'매개변수',
        21:'CODE', 22:'Full Code Name',
    }
    for col, name in headers.items():
        cell = ws.cell(row=8, column=col, value=name)
        cell.fill = header_fill; cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    miss_fill = PatternFill(fill_type='solid', fgColor='FFD7D7')
    ok_fill   = PatternFill(fill_type='solid', fgColor='FFFFFF')
    no_fill   = PatternFill(fill_type=None)

    # 가운데 정렬 컬럼 (J=10, O=15, U=21, V=22 제외)
    center_cols = {2,3,4,5,6,7,8,9,11,12,13,14,16,17,18,19,20}
    # 조건부서식 담당 컬럼 (K~N, P~T)
    cf_cols = set(range(11,15)) | set(range(16,21))

    for i, d in enumerate(code_rows):
        r    = 10 + i
        fill = miss_fill if d['missing'] else ok_fill

        # 값 기입 (B, C, D, E, K, L, M, N, P, Q, R, S, T)
        value_map = {
            2: d['date'],      3: d['media'],    4: d['product'],  5: d['campaign'],
            11: d['gender'],   12: d['age'],     13: d['targeting'], 14: d['note'],
            16: d['device'],   17: d['creative'], 18: d['orient'],
            19: d['seconds'],  20: d['param'],
        }
        for col, val in value_map.items():
            cell = ws.cell(row=r, column=col, value=val)
            cell.fill = no_fill if col in cf_cols else fill
            align_h = 'center' if col in center_cols else 'left'
            cell.alignment = Alignment(horizontal=align_h, vertical='center', wrap_text=False)

        # 수식 기입 (F, G, H, I, J, O, U, V)
        row_formulas = get_formula_for_row(template_formulas, r)
        for col_letter, formula in row_formulas.items():
            col_idx = ord(col_letter) - ord('A') + 1
            cell = ws.cell(row=r, column=col_idx, value=formula)
            # G(7)=매체코드, H(8)=상품코드 missing이면 분홍
            if col_idx in (7, 8) and d['missing']:
                cell.fill = miss_fill
            elif col_idx in cf_cols:
                cell.fill = no_fill
            else:
                cell.fill = fill
            align_h = 'center' if col_idx in center_cols else 'left'
            cell.alignment = Alignment(horizontal=align_h, vertical='center', wrap_text=False)

    # 조건부서식: K~N, P~T열 빈칸이면 빨간 음영, 값 채우면 자동 해제
    red_fill   = PatternFill(fill_type='solid', fgColor='FFCCCC')
    total_rows = 10 + len(code_rows) - 1
    cf_ranges  = [
        (f'K10:N{total_rows}', 'K10'),
        (f'P10:T{total_rows}', 'P10'),
    ]
    for rng, first_cell in cf_ranges:
        rule = FormulaRule(
            formula=[f'LEN(TRIM({first_cell}))=0'],
            fill=red_fill,
            stopIfTrue=False,
        )
        ws.conditional_formatting.add(rng, rule)

    # 유효성검사 직접 추가 (openpyxl이 원본 파일의 확장 유효성검사를 읽지 못해 소실되므로 코드로 재생성)
    # Media/Product 시트 실제 데이터 마지막 행 동적 탐색
    def last_data_row(sheet, col=3):
        for r in range(sheet.max_row, 1, -1):
            if sheet.cell(row=r, column=col).value:
                return r
        return 2

    media_last   = last_data_row(wb['Media'])   if 'Media'   in wb.sheetnames else 2
    product_last = last_data_row(wb['Product']) if 'Product' in wb.sheetnames else 2

    ws.data_validations.dataValidation = []  # 기존 깨진 유효성검사 초기화

    dv_configs = [
        # (열, 참조범위, 설명)
        ('C', f"Media!$C$2:$C${media_last}",    '매체명'),
        ('D', f"Product!$C$2:$C${product_last}", '상품명'),
        ('K', '"P,M,F"',                            '성별'),
        ('P', '"A,P,M,C"',                          'Device'),
        ('R', '"H,V"',                              '가로세로'),
    ]
    total_rows = 10 + len(code_rows) - 1
    for col_letter, formula, _ in dv_configs:
        dv = DataValidation(type='list', formula1=formula, allow_blank=True, showErrorMessage=False)
        dv.sqref = f'{col_letter}10:{col_letter}{total_rows}'
        ws.add_data_validation(dv)

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out.read()


# ── UI ────────────────────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>🍟 맥도날드 코드 자동화 툴</h1>
    <p>Media Mix 파일을 업로드하면 가이드북 기준으로 CODE 시트를 자동 생성합니다</p>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="card"><h3>📊 Media Mix 파일</h3>', unsafe_allow_html=True)
    mm_file = st.file_uploader("Media Mix", type=['xlsx'], key='mm', label_visibility='collapsed')
    st.markdown(f'<span class="badge-ok">✓ 업로드 완료</span>' if mm_file else '<span class="badge-neutral">파일을 선택해주세요</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="card"><h3>⚙️ 자동화 작업용 파일 (DATA RAW)</h3>', unsafe_allow_html=True)
    tool_file = st.file_uploader("자동화 작업용", type=['xlsx'], key='tool', label_visibility='collapsed')
    st.markdown(f'<span class="badge-ok">✓ 업로드 완료</span>' if tool_file else '<span class="badge-neutral">파일을 선택해주세요</span>', unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown('<hr>', unsafe_allow_html=True)

if mm_file and tool_file:
    if st.button('▶  코드 자동 생성하기'):
        with st.spinner('분석 중...'):
            try:
                mm_bytes   = mm_file.read()
                tool_bytes = tool_file.read()

                get_media_code, get_product_code, target_map = load_data_raw(tool_bytes)
                actual, date_code, camp, cname, *cols = parse_media_mix(mm_bytes)
                has_name_col = cols[-1]  # 마지막 값이 has_name_col 플래그
                code_rows = build_code_rows(actual, date_code, camp, cname, *cols,
                                            get_media_code, get_product_code, target_map)

                total    = len(code_rows)
                ok_cnt   = sum(1 for d in code_rows if not d['missing'])
                miss_cnt = total - ok_cnt
                miss_media   = sorted(set(d['media']   for d in code_rows if not d['m_code']))
                miss_product = sorted(set(d['product'] for d in code_rows if not d['p_code']))

                c1, c2, c3 = st.columns(3)
                with c1: st.markdown(f'<div class="stat-box"><div class="stat-num">{total}</div><div class="stat-label">총 생성 행수</div></div>', unsafe_allow_html=True)
                with c2: st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#2E7D32">{ok_cnt}</div><div class="stat-label">코드 완성</div></div>', unsafe_allow_html=True)
                with c3: st.markdown(f'<div class="stat-box"><div class="stat-num" style="color:#C62828">{miss_cnt}</div><div class="stat-label">빈칸 처리</div></div>', unsafe_allow_html=True)

                cr_mode = "BA Q1 방식 (G열=포맷, H열=소재명)" if has_name_col else "QPC 방식 (G열=소재명+포맷 통합)"
                st.markdown(f'<div style="margin:1rem 0;display:flex;gap:0.5rem;flex-wrap:wrap;"><span class="badge-neutral">캠페인: {camp}</span><span class="badge-neutral">날짜코드: {date_code}</span><span class="badge-neutral">소재 파싱: {cr_mode}</span></div>', unsafe_allow_html=True)

                if miss_media or miss_product:
                    with st.expander(f"⚠️ DATA RAW에 없는 항목 — 매체 {len(miss_media)}개 / 상품 {len(miss_product)}개"):
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown("**매체 코드 없음**")
                            for m in miss_media: st.markdown(f"- {m}")
                        with c2:
                            st.markdown("**상품 코드 없음**")
                            for p in miss_product: st.markdown(f"- {p}")

                with st.expander("📋 매체별 행수 요약", expanded=True):
                    summary = []
                    for media, cnt in Counter(d['media'] for d in code_rows).items():
                        miss = sum(1 for d in code_rows if d['media'] == media and d['missing'])
                        summary.append({'매체': media, '총 행수': cnt, '완성': cnt-miss, '빈칸': miss})
                    st.dataframe(pd.DataFrame(summary), use_container_width=True, hide_index=True)

                st.markdown("**결과 미리보기**")
                preview = [{'매체': d['media'], '상품': d['product'], 'CODE(J)': d['j_code'] or '⬜ 빈칸',
                            '성별': d['gender'], '연령': d['age'], '타겟팅': d['targeting'],
                            'Device': d['device'], '소재': d['creative'], '가로세로': d['orient'], '초수': d['seconds']}
                           for d in code_rows]
                st.dataframe(pd.DataFrame(preview), use_container_width=True, height=380)

                result_bytes = write_excel(tool_bytes, code_rows)
                st.download_button(
                    label='⬇️  결과 파일 다운로드',
                    data=result_bytes,
                    file_name=f'McD_코드자동화_{camp}_{date_code}.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )

            except Exception as e:
                st.error(f"오류가 발생했어요: {str(e)}")
                st.exception(e)
else:
    st.markdown('<div style="text-align:center;padding:2.5rem;color:#9CA3AF;background:#FAFAFA;border-radius:12px;"><div style="font-size:2.5rem">⬆️</div><div style="margin-top:0.8rem;font-size:0.95rem">두 파일을 모두 업로드하면 실행 버튼이 활성화됩니다</div></div>', unsafe_allow_html=True)

st.markdown('<hr>', unsafe_allow_html=True)
with st.expander("ℹ️ 사용 가이드"):
    st.markdown("""
    **파일 준비**
    - `Media Mix 파일` — 캠페인 미디어 믹스 엑셀 파일
    - `자동화 작업용 파일` — DATA RAW가 포함된 코드 마스터 파일

    **소재 파싱 방식 (자동 감지)**
    - **BA Q1 방식**: Creative(G열)=포맷, 다음 컬럼(H열)=소재명 → 두 컬럼 분리 파싱
    - **QPC/McCrispy 방식**: Creative(G열)에 소재명+포맷 통합 → G열에서 모두 추출

    **Alias 기능**
    - DATA RAW의 Alias 컬럼에 미디어믹스 표기명을 쉼표로 구분해서 입력하면 자동 매칭

    **결과 파일**
    - 코드 완성 행: 흰 배경 / 미매핑 행: 연빨간 배경
    """)
