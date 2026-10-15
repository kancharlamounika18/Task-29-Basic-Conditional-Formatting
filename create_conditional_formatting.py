from pathlib import Path
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule, CellIsRule, Rule
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.utils import get_column_letter

ROOT=Path('/home/ubuntu/Basic-Conditional-Formatting')
ROOT.mkdir(parents=True,exist_ok=True)
CSV=Path('/home/ubuntu/work/superstore.csv')
OUT=ROOT/'basic_conditional_formatting.xlsx'
df=pd.read_csv(CSV)
# Use a focused retail-sales view so the conditional formatting is easy to inspect.
cols=['Row ID','Order ID','Order Date','Customer Name','Region','Category','Sub-Category','Product Name','Sales','Quantity','Discount','Profit']
data=df[cols].copy()

rules=[
 ['Sales','Top 10 values','Automatically highlight the ten highest transaction sales.','Green fill','Top10'],
 ['Profit','Bottom 10 values','Automatically highlight the ten lowest transaction profits.','Red fill','Bottom10'],
 ['Profit','Negative values','Flag loss-making transactions where Profit < 0.','Red text / light red fill','CellIs < 0'],
 ['Discount','Color scale','Show low-to-high discount levels using a 3-color scale.','Green → Yellow → Red','ColorScale'],
 ['Quantity','Data bar','Compare item quantities using proportional data bars.','Blue data bars','DataBar'],
 ['Sales','Top 10%','Highlight the highest-sales 10% of transactions.','Gold fill','Top10Percent'],
]

wb=Workbook(); ws=wb.active; ws.title='Formatted_Sales'; rs=wb.create_sheet('Rules_List'); info=wb.create_sheet('Read_Me')
navy='17365D'; blue='2F75B5'; light='D9EAF7'; pale='F4F8FB'; green='E2F0D9'; red='FCE4D6'; gold='FFF2CC'; white='FFFFFF'; gray='666666'; thin=Side(style='thin',color='D9E2F3')

def title(sheet,text,subtitle,end_col):
    sheet.merge_cells(start_row=1,start_column=1,end_row=1,end_column=end_col)
    c=sheet.cell(1,1,text); c.font=Font(size=16,bold=True,color=white); c.fill=PatternFill('solid',fgColor=navy); c.alignment=Alignment(horizontal='left'); sheet.row_dimensions[1].height=29
    sheet.merge_cells(start_row=2,start_column=1,end_row=2,end_column=end_col)
    c=sheet.cell(2,1,subtitle); c.font=Font(size=10,italic=True,color=gray); c.alignment=Alignment(wrap_text=True)
def style_header(sheet,row,n):
    for col in range(1,n+1):
        c=sheet.cell(row,col); c.font=Font(bold=True,color=white); c.fill=PatternFill('solid',fgColor=blue); c.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True); c.border=Border(bottom=thin)
    sheet.row_dimensions[row].height=25
def widths(sheet,n):
    for col in range(1,n+1):
        letter=get_column_letter(col); maxlen=max((len(str(c.value)) for c in sheet[letter] if c.value is not None),default=10)
        sheet.column_dimensions[letter].width=min(max(maxlen+2,12),36)

# Formatted data
subtitle='Sample Superstore retail sales with automatic top/bottom, loss, scale, and data-bar rules.'
title(ws,'Basic Conditional Formatting',subtitle,len(cols))
for j,c in enumerate(cols,1): ws.cell(4,j,c)
style_header(ws,4,len(cols))
for i,row in enumerate(data.itertuples(index=False,name=None),5):
    for j,v in enumerate(row,1): ws.cell(i,j,v)
last=4+len(data)
tab=Table(displayName='FormattedSalesData',ref=f'A4:{get_column_letter(len(cols))}{last}'); tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True,showColumnStripes=False); ws.add_table(tab)
ws.freeze_panes='A5'; widths(ws,len(cols))
# formats
for r in range(5,last+1):
    ws.cell(r,9).number_format='$#,##0.00'; ws.cell(r,11).number_format='0%'; ws.cell(r,12).number_format='$#,##0.00'
# Apply rules to relevant columns.
# Sales I, Quantity J, Discount K, Profit L
green_dxf=DifferentialStyle(fill=PatternFill('solid',fgColor='C6EFCE'),font=Font(color='006100'))
gold_dxf=DifferentialStyle(fill=PatternFill('solid',fgColor='FFF2CC'),font=Font(color='9C6500'))
red_dxf=DifferentialStyle(fill=PatternFill('solid',fgColor='FFC7CE'),font=Font(color='9C0006'))
ws.conditional_formatting.add(f'I5:I{last}', Rule(type='top10',rank=10,percent=False,dxf=green_dxf))
ws.conditional_formatting.add(f'L5:L{last}', Rule(type='top10',rank=10,percent=True,dxf=gold_dxf))
# Explicit bottom10 using openpyxl rule API.
ws.conditional_formatting.add(f'L5:L{last}', Rule(type='top10',rank=10,percent=False,bottom=True,dxf=red_dxf))
ws.conditional_formatting.add(f'L5:L{last}', CellIsRule(operator='lessThan',formula=['0'],fill=PatternFill('solid',fgColor='FFC7CE'),font=Font(color='9C0006')))
ws.conditional_formatting.add(f'K5:K{last}', ColorScaleRule(start_type='min',start_color='63BE7B',mid_type='percentile',mid_value=50,mid_color='FFEB84',end_type='max',end_color='F8696B'))
ws.conditional_formatting.add(f'J5:J{last}', DataBarRule(start_type='min',end_type='max',color='5B9BD5',showValue=True))
ws.conditional_formatting.add(f'I5:I{last}', Rule(type='top10',rank=10,percent=True,dxf=gold_dxf))

# Rules list sheet
headers=['Column','Rule','Purpose','Visual Result','Excel Rule Type']
title(rs,'Conditional Formatting Rules','Automatic rules applied to the Formatted_Sales sheet.',len(headers))
for j,h in enumerate(headers,1): rs.cell(4,j,h)
style_header(rs,4,len(headers))
for i,row in enumerate(rules,5):
    for j,v in enumerate(row,1): rs.cell(i,j,v)
ref=f'A4:E{4+len(rules)}'; tab=Table(displayName='ConditionalFormattingRules',ref=ref); tab.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True,showColumnStripes=False); rs.add_table(tab); rs.freeze_panes='A5'; widths(rs,5)

# Read me
info.merge_cells('A1:F1'); info['A1']='Basic Conditional Formatting'; info['A1'].font=Font(size=18,bold=True,color=white); info['A1'].fill=PatternFill('solid',fgColor=navy); info['A1'].alignment=Alignment(horizontal='left'); info.row_dimensions[1].height=32
info.merge_cells('A2:F2'); info['A2']='Retail Sales | Sample Superstore | Automatic highlighting for report readability'; info['A2'].font=Font(size=10,italic=True,color=gray)
rows=[('Dataset','Sample Superstore public retail-sales dataset'),('Records',len(data)),('Main sheet','Formatted_Sales'),('Rules sheet','Rules_List'),('Source','https://github.com/leonism/sample-superstore'),('Objective','Highlight top and bottom values automatically instead of manually coloring cells')]
for i,(k,v) in enumerate(rows,4):
    info.cell(i,1,k).font=Font(bold=True,color=navy); info.cell(i,1).fill=PatternFill('solid',fgColor=light); info.cell(i,2,v).fill=PatternFill('solid',fgColor=pale); info.cell(i,2).alignment=Alignment(wrap_text=True)
info['A12']='How to test'; info['A12'].font=Font(size=13,bold=True,color=navy)
for i,text in enumerate(['Sort or filter the Sales and Profit columns and observe the automatic highlights.','Edit a Sales, Profit, Discount, or Quantity value; conditional formatting updates automatically in Excel.','Review Rules_List for the purpose and type of each rule.'],13): info.cell(i,1,'• '+text)
info.column_dimensions['A'].width=28; info.column_dimensions['B'].width=86
for s in wb.worksheets: s.sheet_view.showGridLines=False
wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True
wb.save(OUT)
# Separate plain CSV rules list for easy submission.
pd.DataFrame(rules,columns=headers).to_csv(ROOT/'conditional_formatting_rules.csv',index=False)
print(OUT)
print('records',len(data),'rules',len(rules))
