import io
import re
from collections import defaultdict

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SHIFTS = ['בוקר', 'צהריים', 'ערב', 'לילה']
OUTPUT_COLUMNS = ['יום','תאריך','משמרת','מנמ"ש','ווטסאפ','צנזור 1','צנזור 2','מנהל ידע',"צ'קינג 11","צ'קינג 12","צ'קינג 13","צ'קינג 14","צ'קינג 24I",'רח"ל','דסקאי 1','דסקאי 2','ערוץ 11','ערוץ 12','ערוץ 13','ערוץ 14','מדור','חופש','אחרי לילה']
MANUAL_ONLY_COLUMNS = {'דסקאי 1','דסקאי 2','מדור','חופש','אחרי לילה'}
CENSOR_ASSIGNMENT_ORDER = ['ווטסאפ','צנזור 1','צנזור 2',"צ'קינג 11","צ'קינג 12","צ'קינג 13","צ'קינג 14","צ'קינג 24I",'רח"ל']
CENSOR_OVERFLOW_COLUMN = 'צנזור 2'
CHANNEL_COLUMNS = ['ערוץ 11','ערוץ 12','ערוץ 13','ערוץ 14']
_CENSOR_COLUMNS = CENSOR_ASSIGNMENT_ORDER + CHANNEL_COLUMNS
ROLE_TO_COLUMNS = {'צנזור':_CENSOR_COLUMNS,'סדיר (קבע/אע"צ)':_CENSOR_COLUMNS,'מנהל משמרת':['מנמ"ש'],'מנהל ידע':['מנהל ידע','צנזור 1','צנזור 2'],'מילואים':_CENSOR_COLUMNS}
CENSOR_ROLES = {'צנזור','סדיר (קבע/אע"צ)','מילואים'}
SHIFT_ALIASES = {'אחרי משמרת לילה':'לילה'}
IGNORE_TOKENS = {'חופש','מרוכז','יום חופש מרוכז'}

# צבעי מרקר עדינים. כל עובד מקבל צבע קבוע לפי סדר הופעתו בקובץ.
EMPLOYEE_HIGHLIGHT_COLORS = ['FFF2CC','DDEBF7','E2F0D9','FCE4D6','E4DFEC','EADCF8','D9EAD3','F4CCCC','CFE2F3','FCE5CD','D9E1F2','EAD1DC']

def get_role(service_type, main_role):
    for value in (main_role, service_type):
        if value is not None:
            text = str(value).strip()
            if text and text.lower() != 'nan': return text
    return None

def parse_availability(cell_value):
    if cell_value is None: return set()
    text = str(cell_value).strip()
    if not text or text.lower() == 'nan': return set()
    shifts = set()
    for raw in text.split(','):
        token = raw.strip()
        if not token or token in IGNORE_TOKENS: continue
        token = SHIFT_ALIASES.get(token, token)
        if token in SHIFTS or token == 'ערוץ': shifts.add(token)
    return shifts

def extract_days_from_requests(columns):
    pattern = re.compile(r"זמינות\s*([א-ת])(?:['׳״\"])?\s*[-–—]?\s*(\d{1,2}[./-]\d{1,2}(?:[./-]\d{2,4})?)")
    days=[]
    for col in columns:
        m=pattern.search(str(col).strip())
        if m: days.append((col,m.group(1),m.group(2)))
    return days

def build_schedule_workbook(requests_bytes):
    df=pd.read_excel(io.BytesIO(requests_bytes))
    day_cols=extract_days_from_requests(df.columns)
    if not day_cols:
        raise ValueError('לא נמצאו עמודות זמינות בקובץ (בפורמט "זמינות א\' 23.08"). בדקי את כותרות הקובץ.')

    employees=[]; regular_index=defaultdict(list); censor_index=defaultdict(list); channel_index=defaultdict(list)
    for row in df.to_dict(orient='records'):
        first=str(row.get('שם פרטי','') or '').strip(); last=str(row.get('שם משפחה','') or '').strip()
        full_name=f'{first} {last}'.strip()
        if not full_name or full_name.lower()=='nan nan': continue
        role=get_role(row.get('סוג שירות'),row.get('תפקיד עיקרי במפעל'))
        availability={}; is_censor=role in CENSOR_ROLES
        for col,day_letter,date in day_cols:
            shifts=parse_availability(row.get(col)); availability[day_letter]=shifts
            if is_censor and 'ערוץ' in shifts: channel_index[day_letter].append(full_name)
            for shift in shifts:
                if shift=='ערוץ': continue
                if is_censor: censor_index[(day_letter,shift)].append(full_name)
                else:
                    for out in ROLE_TO_COLUMNS.get(role,[]):
                        if out not in CHANNEL_COLUMNS: regular_index[(out,day_letter,shift)].append(full_name)
        employees.append({'name':full_name,'role':role,'availability':availability})

    employee_colors={}
    for i,e in enumerate(employees):
        employee_colors.setdefault(e['name'],EMPLOYEE_HIGHLIGHT_COLORS[len(employee_colors)%len(EMPLOYEE_HIGHLIGHT_COLORS)])

    wb=Workbook(); ws=wb.active; ws.title='סידור עבודה'
    header_font=Font(bold=True); header_fill=PatternFill(start_color='D9E1F2',end_color='D9E1F2',fill_type='solid')
    thin=Side(style='thin',color='BFBFBF'); border=Border(left=thin,right=thin,top=thin,bottom=thin)
    header_align=Alignment(horizontal='center',vertical='center',wrap_text=True)
    cell_align=Alignment(horizontal='right',vertical='center',wrap_text=True)
    for c,title in enumerate(OUTPUT_COLUMNS,1):
        cell=ws.cell(1,c,title); cell.font=header_font; cell.fill=header_fill; cell.alignment=header_align; cell.border=border

    row_idx=2
    for col,day_letter,date in day_cols:
        first_row_of_day=row_idx
        for shift in SHIFTS:
            values=defaultdict(list)
            for out in OUTPUT_COLUMNS:
                if out in MANUAL_ONLY_COLUMNS or out in {'יום','תאריך','משמרת'} or out in CENSOR_ASSIGNMENT_ORDER or out in CHANNEL_COLUMNS: continue
                values[out]=list(regular_index.get((out,day_letter,shift),[]))

            available=list(censor_index.get((day_letter,shift),[]))
            for i,out in enumerate(CENSOR_ASSIGNMENT_ORDER):
                if i<len(available): values[out].append(available[i])
            if len(available)>len(CENSOR_ASSIGNMENT_ORDER):
                values[CENSOR_OVERFLOW_COLUMN].extend(available[len(CENSOR_ASSIGNMENT_ORDER):])

            if shift=='צהריים':
                channel=list(channel_index.get(day_letter,[]))
                for i,out in enumerate(CHANNEL_COLUMNS):
                    if i<len(channel): values[out].append(channel[i])
                if len(channel)>len(CHANNEL_COLUMNS): values[CHANNEL_COLUMNS[-1]].extend(channel[len(CHANNEL_COLUMNS):])

            # כמה שורות צריך כדי שכל עובד יקבל תא נפרד באותה עמודת תפקיד.
            height=max([len(v) for v in values.values()] or [1])
            start=row_idx; end=row_idx+height-1

            for offset in range(height):
                r=row_idx+offset
                for c,colname in enumerate(OUTPUT_COLUMNS,1):
                    value=None
                    if colname=='יום': value=day_letter if r==first_row_of_day else None
                    elif colname=='תאריך': value=date if r==first_row_of_day else None
                    elif colname=='משמרת': value=shift if offset==0 else None
                    else:
                        names=values.get(colname,[])
                        value=names[offset] if offset<len(names) else None
                    cell=ws.cell(r,c,value); cell.border=border; cell.alignment=cell_align
                    if value and colname not in {'יום','תאריך','משמרת'} and value in employee_colors:
                        color=employee_colors[value]
                        cell.fill=PatternFill(start_color=color,end_color=color,fill_type='solid')
                        cell.font=Font(color='000000')

            # מאחדים רק את יום/תאריך/משמרת כדי שהעובדים יישארו בשורות נפרדות באותה משבצת תפקיד.
            if height>1:
                for colname in ('משמרת',):
                    c=OUTPUT_COLUMNS.index(colname)+1; ws.merge_cells(start_row=start,start_column=c,end_row=end,end_column=c)
                    ws.cell(start,c).alignment=header_align
            row_idx=end+1

        # יום ותאריך מאוחדים על פני כל ארבע המשמרות של אותו יום.
        last_row_of_day=row_idx-1
        if last_row_of_day>first_row_of_day:
            for colname in ('יום','תאריך'):
                c=OUTPUT_COLUMNS.index(colname)+1
                ws.merge_cells(start_row=first_row_of_day,start_column=c,end_row=last_row_of_day,end_column=c)
                ws.cell(first_row_of_day,c).alignment=header_align

    for c,title in enumerate(OUTPUT_COLUMNS,1): ws.column_dimensions[get_column_letter(c)].width=max(10,len(title)+4)
    ws.freeze_panes='D2'; ws.sheet_view.rightToLeft=True
    return wb,employees,day_cols
