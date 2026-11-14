"""Offline regression guards for 2026-11 snapshots and months; no Excel required."""
from decimal import Decimal as D
from wasde_automation import (
    tab_names,previous_month,compatible_sheet_names,fill_us,US_ROWS,
    validate_report_values, WORLD_REGIONS, norm
)

class Cell:
    def __init__(self,ws,row,col):self.ws,self.row,self.col=ws,row,col
    @property
    def Formula(self):return self.ws.formulas.get((self.row,self.col))
    @Formula.setter
    def Formula(self,val):self.ws.formulas[(self.row,self.col)]=val
    @property
    def Value2(self):
        f=self.Formula
        if f:
            import re
            match=re.fullmatch(r'=([A-Z])(\d+)-([A-Z])(\d+)',f)
            assert match,f
            c1,r1,c0,r0=match.groups()
            return self.ws.values[(int(r1),ord(c1)-64)]-self.ws.values[(int(r0),ord(c0)-64)]
        return self.ws.values.get((self.row,self.col))
    @Value2.setter
    def Value2(self,value):self.ws.values[(self.row,self.col)]=value
    def ClearContents(self):self.ws.formulas.pop((self.row,self.col),None);self.ws.values.pop((self.row,self.col),None)
class WS:
    def __init__(self):self.values={};self.formulas={}
    def Cells(self,row,col):return Cell(self,row,col)
    def Calculate(self):return None

def check_release_names():
    for crop in ('wheat','corn','soybean','soyoil','sugar'):
        b,w=tab_names(2026,11,crop)
        assert b=="November'26 Balance Sheet"
        assert 'November' in w and 'October' not in w
        assert 'Mexico Sugar' in w if crop=='sugar' else 'World' in w
    names=["October'26 Balance Sheet (2)","October'26 Mexico Sugar (2)",
           "November'26 Balance Sheet","November'26 Mexico Sugar", "September'26 Balance Sheet"]
    assert compatible_sheet_names(names,2026,11,'sugar','us')==["November'26 Balance Sheet"]
    assert compatible_sheet_names(names,2026,11,'sugar','world')==["November'26 Mexico Sugar"]
    assert compatible_sheet_names(names,2026,10,'sugar','world')==["October'26 Mexico Sugar (2)"]
    assert previous_month(2027,1)==(2026,12)

def check_sugar_change():
    # Taken from published Nov 10 2026 USDA report.
    data={'us_years':['2024/25','2025/26','2026/27','2026/27'], 'us':{}}
    for k in US_ROWS['sugar'].values():data['us'][k]=[D(0)]*4
    for k,prev,current in [
        ('beginningstocks',2231,2199),('production',8768,8520),
        ('deliveries',12497,12401),('food',12392,12296),
        ('other',105,105),('miscellaneous',0,0),
        ('totaluse',12522,12426),('endingstocks',1865,1723),
        ('stockstouseratio',D('14.9'),D('13.9'))]:
        data['us'][k]=[D(1),D(2),D(prev),D(current)]
    ws=WS()
    fill_us(ws,data,'sugar','Oct','Nov')
    assert ws.values[(3,5)]=='2026/27 Proj. October'
    assert ws.values[(3,6)]=='2026/27 Proj. November'
    assert ws.Cells(21,7).Value2==-96
    assert ws.Cells(22,7).Value2==-96
    assert ws.Cells(23,7).Value2==0
    assert ws.Cells(24,7).Value2==0
    assert ws.Cells(25,7).Value2==-96
    assert ws.Cells(26,7).Value2==-142
    assert round(ws.Cells(27,7).Value2,1)==-1.0
    assert ws.Cells(26,7).Formula=='=F26-E26'
    assert ws.Cells(27,7).Formula=='=F27-E27'

check_release_names()
check_sugar_change()
print('PASS: monthly tab names, November headers, all key USDA Sugar changes, year rollover.')

# Test actual new-sheet construction with a minimal Excel COM simulator. No file writes.
from wasde_automation import create_world_sheet,create_mexico_sheet,WORLD_COLUMNS

class Style:
    pass
class Borders:
    def __call__(self,n):return Style()
class MockRange:
    def __init__(self):
        self.Interior=Style();self.Font=Style();self.Borders=Borders()
        self.merged=False
    def Merge(self):self.merged=True
class MockSheet(WS):
    def __init__(self):
        super().__init__();self.Name='';self.ranges={}
    def Range(self,addr):
        if addr not in self.ranges:self.ranges[addr]=MockRange()
        return self.ranges[addr]
    def __getattr__(self,name):
        if name=='Columns':return lambda x:self.Range(x)
        raise AttributeError(name)
class MockSheets:
    def __init__(self):self.items=[]
    @property
    def Count(self):return len(self.items)
    def __iter__(self):return iter(self.items)
    def __call__(self,key):
        if isinstance(key,int):return self.items[key-1]
        return next(s for s in self.items if s.Name==key)
    def Add(self,After=None):
        s=MockSheet();self.items.append(s);return s
class MockWB:
    def __init__(self):self.Worksheets=MockSheets()

def check_world_and_mexico():
    for crop in ('wheat','corn','soybean','soyoil'):
        regions=WORLD_REGIONS[crop];n=7 if crop!='soyoil' else 6
        d={'season':'2026/27','world':{}}
        for region in regions:
            d['world'][norm(region)]={'Oct':tuple(D(10+j) for j in range(n)),
                                      'Nov':tuple(D(11+j) for j in range(n))}
        wb=MockWB();wb.Worksheets.Add().Name='Previous Sheet'
        ws=create_world_sheet(wb,"November'26 World "+crop,d,crop,'Oct','Nov')
        assert ws.Cells(3,3).Value2=='Oct' and ws.Cells(4,3).Value2=='Nov'
        assert ws.Cells(5,3).Value2=='Change'
        assert ws.Cells(3,2).Value2==regions[0]
        assert ws.Cells(5,4).Value2==1
        assert ws.Cells(5,4).Formula=='=D4-D3'
        assert ws.Cells(3+3*(len(regions)-1),2).Value2==regions[-1]
    mexico={'mexico':{
        '2025/26':{'Oct':tuple(map(D,[1123,5325,42,4306,1000,1184])),
                   'Nov':tuple(map(D,[1123,5325,41,4207,1082,1201]))},
        '2026/27':{'Oct':tuple(map(D,[1184,5377,28,4274,1193,1124])),
                   'Nov':tuple(map(D,[1201,5292,28,4274,1094,1154]))},
    }}
    wb=MockWB();wb.Worksheets.Add().Name='Previous Sheet'
    ws=create_mexico_sheet(wb,"November'26 Mexico Sugar",mexico,'Oct','Nov')
    assert ws.Cells(11,3).Value2=='Oct' and ws.Cells(12,3).Value2=='Nov'
    assert ws.Cells(13,9).Value2==17  # 1201 - 1184
    assert ws.Cells(16,5).Value2==-85 # 5292 - 5377
    assert ws.Cells(16,5).Formula=='=E15-E14'

check_world_and_mexico()
print('PASS: fresh World (4 crops) and Mexico sheets, all region rows/changes (mock).')

# Real USDA November 2026 world-wheat lines: old parser accidentally used
# Total Foreign 785.76 instead of World 827.51, a dangerous market-data bug.
import wasde_automation as module
real_excerpt='''World Wheat Supply and Use  1/
                                   2026/27 Proj.
World  3/
                      Oct  281.55  822.39  204.80  166.90  827.90  208.97  276.04
                      Nov  282.86  827.51  203.85  168.41  829.55  208.66  280.82
United States
                      Oct   25.04   41.74    3.95    2.45   30.18   20.41   20.14
                      Nov   25.04   41.74    3.95    2.45   30.10   20.41   20.22
Total Foreign
                      Oct  256.51  780.65  200.85  164.46  797.72  188.56  255.90
                      Nov  257.82  785.76  199.90  165.96  799.45  188.25  260.60
World Coarse Grain Supply and Use  1/
'''
saved_regions=module.WHEAT_REGIONS[:]
try:
    module.WHEAT_REGIONS[:]=['World']  # excerpt covers only this required region
    parsed,season=module.parse_world(real_excerpt,'wheat','Oct','Nov')
    assert parsed['world']['Nov'][1]==D('827.51')
    assert parsed['totalforeign']['Nov'][1]==D('785.76')
    assert parsed['unitedstates']['Nov'][1]==D('41.74')
finally:
    module.WHEAT_REGIONS[:]=saved_regions
print('PASS: USDA November 2026 World Wheat vs Total Foreign regression.')
