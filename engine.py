"""ECHS NMI Excel engine. Stateless; no persistence or deduplication of raw records."""
from __future__ import annotations
import copy, io, re, zipfile
from collections import Counter
from datetime import date, datetime
from pathlib import Path
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties

EXPECTED=['Unit Name','Sponsor','IP / OP','Admi No. / Invoice No.','Claim ID','Patient Name','Claim Amount','NMI Date','NMI In Detail','Action Remark','NMI Status','Other Remark','Other Remark','Days to Pending','Aging']
RAW_COLUMNS=['Claim Id','Hospital Name','Patient Name','Claim Amount','NMI Date','Latest User Remarks','I/O','Admi No. / Invoice No.','Source File']
ALLOWED_RAW={'.xls','.xlsx','.xlsb','.csv'}

def norm(v): return re.sub('[^a-z0-9]','',str(v or '').strip().lower())
def standard(v):
    k=norm(v)
    return {'claimnumber':'claimid','unitname':'hospitalname','ipop':'io','inpatientoutpatient':'io','admissionno':'admissionno','admissionnumber':'admissionno','admino':'admissionno','invoiceno':'admissionno','adminoinvoiceno':'admissionno'}.get(k,k)
def blank(v):
    if v is None:return True
    try:return bool(pd.isna(v))
    except (TypeError,ValueError):return False

def cid(v):
    if blank(v) or isinstance(v,bool):return None
    s=str(v).strip().replace(',','').replace('\u200b','').lstrip("'")
    if not re.fullmatch(r'\d+(?:\.0+)?',s):return None
    n=int(s.split('.')[0]);return n if 0<n and len(str(n))<=15 else None

def dt(v):
    if blank(v):return None
    if isinstance(v,(datetime,date,pd.Timestamp)):return datetime(v.year,v.month,v.day)
    if isinstance(v,(int,float)) and not isinstance(v,bool):
        if 20000<=float(v)<=80000:
            try:return pd.to_datetime(float(v),unit='D',origin='1899-12-30').to_pydatetime().replace(hour=0,minute=0,second=0,microsecond=0)
            except (ValueError,OverflowError):return None
        return None
    s=str(v).strip()
    for fmt in ('%d-%b-%Y','%d-%B-%Y','%d-%m-%Y','%d/%m/%Y','%Y-%m-%d','%d-%b-%y','%d/%m/%y','%Y-%m-%d %H:%M:%S'):
        try:return datetime.strptime(s,fmt)
        except ValueError:pass
    return None

def txt(v):
    if blank(v):return None
    return re.sub(r'\s+',' ',str(v)).strip() or None

def amount(v):
    if blank(v):return None
    try:return float(str(v).replace(',','').replace('₹','').strip())
    except ValueError:return None

def read_table(filename, binary):
    ext=Path(filename).suffix.lower(); b=io.BytesIO(binary)
    if ext=='.xls':return pd.read_excel(b,header=None,dtype=object,engine='xlrd')
    if ext in ('.xlsx','.xlsm'):return pd.read_excel(b,header=None,dtype=object,engine='openpyxl')
    if ext=='.xlsb':return pd.read_excel(b,header=None,dtype=object,engine='pyxlsb')
    if ext=='.csv':
        try:return pd.read_csv(b,header=None,dtype=object,encoding='utf-8-sig')
        except UnicodeDecodeError:
            b.seek(0);return pd.read_csv(b,header=None,dtype=object,encoding='cp1252')
    raise ValueError(f'{filename}: Unsupported file format')

def extract_raw(filename,binary):
    df=read_table(filename,binary)
    head=None;pos={}
    for i in range(min(35,len(df))):
        names=[standard(x) for x in df.iloc[i]]
        if 'claimid' in names and 'hospitalname' in names and ('nmidate' in names or 'patientname' in names):
            head=i
            for j,k in enumerate(names):
                if k and k not in pos:pos[k]=j
            break
    if head is None:raise ValueError(f'{filename}: Claim Id / Hospital Name header not found in first 35 rows.')
    out=[];invalid=0;invalid_dates=0
    for ridx in range(head+1,len(df)):
        row=df.iloc[ridx]
        def get(k):return row.iloc[pos[k]] if k in pos else None
        val=get('claimid')
        if blank(val) or not str(val).strip():continue
        number=cid(val)
        if number is None:invalid+=1;continue
        date_val=dt(get('nmidate'))
        if date_val is None and txt(get('nmidate')):invalid_dates+=1
        out.append({'Claim Id':number,'Hospital Name':txt(get('hospitalname')),'Patient Name':txt(get('patientname')),'Claim Amount':amount(get('claimamount')),'NMI Date':date_val,'Latest User Remarks':txt(get('latestuserremarks')),'I/O':txt(get('io')),'Admi No. / Invoice No.':txt(get('admissionno')),'_source':filename,'_source_row':ridx+1})
    return out,{'file':filename,'header_row':head+1,'rows':len(out),'invalid_ids':invalid,'invalid_dates':invalid_dates,'io_found':'io' in pos}

def raw_parse(files):
    if not files:raise ValueError('Upload at least one Raw NMI Excel file.')
    allrows=[];stats=[]
    for name,binary in files:
        if Path(name).suffix.lower() not in ALLOWED_RAW:raise ValueError(f'{name}: unsupported file extension')
        rows,stat=extract_raw(name,binary)
        allrows.extend(rows);stats.append(stat)
    if not allrows:raise ValueError('No valid Claim IDs found in Raw files.')
    unique={}
    for row in allrows:
        old=unique.get(row['Claim Id'])
        if old is None or (row['NMI Date'] or datetime.min)>=(old['NMI Date'] or datetime.min): unique[row['Claim Id']]=row
    return sorted(allrows,key=lambda r:r['Claim Id']),unique,stats

def master_load(binary,filename):
    if binary is None:
        wb=Workbook(); ws=wb.active;ws.title='NMI Master';header=1
        for index,name in enumerate(EXPECTED,1):ws.cell(1,index,name)
        style_header(ws,1)
        return wb,ws,header
    ext=Path(filename).suffix.lower()
    if ext not in ('.xlsx','.xlsm'):raise ValueError('Master must be .xlsx or .xlsm.')
    wb=load_workbook(io.BytesIO(binary),keep_vba=(ext=='.xlsm'))
    for ws in wb.worksheets:
        for hr in range(1,min(ws.max_row,30)+1):
            vals=[norm(ws.cell(hr,j).value) for j in range(1,16)]
            if vals==[norm(x) for x in EXPECTED]:return wb,ws,hr
    raise ValueError('Master A:O headers do not match the 15 required columns. Original sheet was not modified.')

def master_index(ws,hr):
    existing={};bad=[];oldrows=0
    for rr in range(hr+1,ws.max_row+1):
        v=ws.cell(rr,5).value
        if blank(v) or not str(v).strip():continue
        number=cid(v)
        if number is None:bad.append(rr);continue
        existing.setdefault(number,[]).append(rr);oldrows+=1
    if bad:raise ValueError(f'Invalid Claim ID in Master rows: {bad[:15]}'+(' ...' if len(bad)>15 else ''))
    return existing,oldrows

def preview_data(rows,unique,stats,wb,ws,hr,asof):
    existing,oldrows=master_index(ws,hr);active=set(unique)
    duplicate_master=sum(len(v)-1 for v in existing.values())
    absent=set(existing)-active
    old_pending_to_done=sum(1 for k in absent for r in existing[k] if str(ws.cell(r,11).value or '').strip().lower()=='pending')
    changed=sum(1 for k,locs in existing.items() for r in locs if str(ws.cell(r,11).value or '').strip()!=('Pending' if k in active else 'Done'))
    pending=sum(len(locs) for k,locs in existing.items() if k in active)+len(active-set(existing))
    done=sum(len(locs) for k,locs in existing.items() if k not in active)
    return {'raw_rows':len(rows),'raw_unique_ids':len(active),'raw_duplicates':len(rows)-len(active),'raw_files':len(stats),'master_old_rows':oldrows,'master_duplicate_rows':duplicate_master,'master_duplicate_ids':sum(len(v)>1 for v in existing.values()),'new_records':len(active-set(existing)),'absent_ids':len(absent),'pending_to_done':old_pending_to_done,'status_changed':changed,'final_pending':pending,'final_done':done,'final_total':oldrows+len(active-set(existing)),'report_date':asof.strftime('%d-%m-%Y'),'files':stats,'warning':('Some previously Pending claims will become Done. Confirm every region/hospital Raw report is included.' if old_pending_to_done else ''),'sample': [{'claim_id':r['Claim Id'],'hospital':r['Hospital Name'],'patient':r['Patient Name'],'date':r['NMI Date'].strftime('%d-%m-%Y') if r['NMI Date'] else '—','type':{'I':'IP','O':'OP'}.get((r['I/O'] or '').strip().upper(),(r['I/O'] or '').strip().upper() or '—')} for r in rows[:8]]}

def style_header(ws,hr=1):
    ws.sheet_view.showGridLines=False;ws.freeze_panes=f'A{hr+1}'
    ws.row_dimensions[hr].height=30
    for c in ws[hr]:
        c.fill=PatternFill('solid',fgColor='122744');c.font=Font(name='Calibri',size=11,bold=True,color='FFFFFF');c.alignment=Alignment(wrap_text=True,vertical='center')
    for i in range(1,ws.max_column+1):
        name=str(ws.cell(hr,i).value or '').lower()
        ws.column_dimensions[get_column_letter(i)].width=(40 if 'remark' in name or 'detail' in name else 30 if 'hospital' in name or 'unit' in name else 19)
    ws.auto_filter.ref=f'A{hr}:{get_column_letter(ws.max_column)}{ws.max_row}'

def formula_aging(ws,r):
    ws.cell(r,14).value=f'=IF(ISNUMBER(H{r}),Settings!$B$2-H{r},"")'
    ws.cell(r,15).value=f'=IF(N{r}="","",IF(N{r}<=15,"Below 15 Days","Above 15 Days"))'
    ws.cell(r,14).number_format='0'

def part1_book(rows,stats):
    wb=Workbook();ws=wb.active;ws.title='Combined Raw';ws.append(RAW_COLUMNS)
    for item in rows:
        ws.append([item.get(k) if k!='Source File' else item['_source'] for k in RAW_COLUMNS]);r=ws.max_row
        ws.cell(r,1).number_format='0';ws.cell(r,4).number_format='#,##0.00';ws.cell(r,5).number_format='dd-mm-yyyy'
    style_header(ws)
    au=wb.create_sheet('Import Audit');au.append(['Source File','Header Row','Data Rows','Invalid Claim IDs','Invalid NMI Dates','I/O Present'])
    for s in stats:au.append([s['file'],s['header_row'],s['rows'],s['invalid_ids'],s['invalid_dates'],'Yes' if s['io_found'] else 'No'])
    au.append(['TOTAL','',len(rows),sum(x['invalid_ids'] for x in stats),sum(x['invalid_dates'] for x in stats),''])
    au.append(['Duplicate occurrences retained','',len(rows)-len({x['Claim Id'] for x in rows})]);style_header(au)
    return workbook_bytes(wb)

def workbook_bytes(wb):
    stream=io.BytesIO();wb.save(stream);return stream.getvalue()

def create_master(wb,ws,hr,unique,stats,asof):
    existing,old=master_index(ws,hr);active=set(unique)
    for key,locs in existing.items():
        status='Pending' if key in active else 'Done'
        for r in locs:
            ws.cell(r,11).value=status
            # Existing user-managed NMI Date and details are preserved, including other remarks.
            formula_aging(ws,r)
    for key in sorted(active-set(existing)):
        item=unique[key];rr=ws.max_row+1
        if rr>hr+1:
            for c in range(1,16):
                p=ws.cell(rr-1,c);v=ws.cell(rr,c)
                if p.has_style:v._style=copy.copy(p._style)
        io_type={'I':'IP','O':'OP','IP':'IP','OP':'OP'}.get((item['I/O'] or '').strip().upper())
        mapped={1:item['Hospital Name'],2:'ECHS',3:io_type,5:key,6:item['Patient Name'],7:item['Claim Amount'],8:item['NMI Date'],9:item['Latest User Remarks'],11:'Pending'}
        # Admission / Invoice No column D intentionally stays empty per business rule.
        for col,v in mapped.items():ws.cell(rr,col).value=v
        ws.cell(rr,5).number_format='0';ws.cell(rr,7).number_format='#,##0.00';ws.cell(rr,8).number_format='dd-mm-yyyy'
        formula_aging(ws,rr)
    sw=wb['Settings'] if 'Settings' in wb else wb.create_sheet('Settings')
    settings=[('Report Controls',None),('Report / As-of Date',asof),('Raw Files Uploaded',len(stats)),('Unique Active IDs',len(active)),('Master Rows After Update',ws.max_row-hr)]
    for idx,(label,val) in enumerate(settings,1):sw.cell(idx,1).value=label;sw.cell(idx,2).value=val
    sw['B2'].number_format='dd-mm-yyyy';sw.column_dimensions['A'].width=33;sw.column_dimensions['B'].width=24
    sw['A1'].font=Font(size=15,color='122744',bold=True)
    ws.auto_filter.ref=f'A{hr}:O{ws.max_row}';ws.freeze_panes=f'A{hr+1}'
    if old==0:style_header(ws,hr)
    wb.calculation=CalcProperties(calcMode='auto',fullCalcOnLoad=True,forceFullCalc=True)
    return workbook_bytes(wb)

def make_zip(part1,master,asof,ext='xlsx'):
    stamp=datetime.now().strftime('%Y%m%d_%H%M%S')
    out=io.BytesIO()
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED) as z:
        z.writestr(f'ECHS_Part1_Combined_{stamp}.xlsx',part1)
        z.writestr(f'ECHS_NMI_Master_Updated_{stamp}.{ext}',master)
    return out.getvalue(),f'ECHS_NMI_Reports_{stamp}.zip'
