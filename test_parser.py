"""Offline unit tests for wheat, corn, soybean, soybean oil, US & Mexico sugar.
These are SHORT TEXT EXCERPTS; they test the parsing/mapping and never touch Excel.
"""
from wasde_automation import (parse_extra_us,parse_extra_world,parse_mexico,
                              US_ROWS,WORLD_REGIONS,norm)
from wasde_automation import parse_us,parse_world,parse_wasde
import test_wheat_corn as grain_cases

US_OILSEEDS = '''
U.S. Soybeans and Products Supply and Use (Domestic Measure)  1/
                                2024/25  2025/26 Est. 2026/27 Proj. 2026/27 Proj.
Item                                                           Sep           Oct
SOYBEANS
Area Planted                      87.3          81.2          86.9          86.9
Area Harvested                    86.2          80.4          85.9          85.9
Yield per Harvested Acre          50.7          53.0          52.8          53.1
Beginning Stocks                   342           325           325           315
Production                        4374          4261          4535          4562
Imports                             30            22            25            25
  Supply, Total                   4746          4608          4885          4902
Crushings                         2445          2646          2780          2780
Exports                           1892          1515          1685          1695
Seed                                70            75            73            73
Residual                            14            57            38            39
  Use, Total                      4421          4293          4575          4587
Ending Stocks                      325           315           310           315
Avg.FarmPrice ($/bu)  2/         10.00         10.50         12.00         12.00
SOYBEAN OIL
Beginning Stocks                  1551          1747          1837          1657
Production 4/                    29218         30715         32945         32945
Imports                            362           525           600           700
  Supply, Total                  31131         32987         35382         35302
Domestic Disappearance           26912         30355         33105         33105
    Biofuel 3/                    11758         14800         17800         17800
    Food, Feed & other
    Industrial                    15154         15555         15305         15305
Exports                           2472           975           400           400
    Use, Total                    29384         31330         33505         33505
Ending stocks                     1747          1657          1877          1797
Avg. Price (c/lb)  2/            47.59         64.08         70.00         70.00
SOYBEAN MEAL
Beginning Stocks                   453           398           450           450
U.S. Sugar Supply and Use  1/
'''
US_SUGAR = '''
U.S. Sugar Supply and Use  1/
                                2024/25  2025/26 Est. 2026/27 Proj. 2026/27 Proj.
Beginning Stocks                  2141          2376          2033          2231
Production 2/                     9397          9280          8839          8768
  Beet Sugar                      5370          5038          4769          4742
  Cane Sugar                      4027          4242          4071          4026
    Florida                       1932          1957          1819          1819
    Louisiana                     2095          2285          2252          2208
    Texas                            0             0             0             0
Imports                           3393          3097          3396          3388
  TRQ 3/                          1534          1299          1391          1391
  Other Program 4/                 373           399           300           300
  Non-program                     1485          1399          1705          1696
    Mexico                         504           247          1188          1188
  High-tier tariff/other           980          1152           517           508
       Total Supply               14931         14753         14268         14387
Exports                            111            25            25            25
Deliveries                       12485         12497         12546         12497
  Food                           12375         12392         12441         12392
  Other 5/                         111           105           105           105
Miscellaneous                      -41             0             0             0
       Total Use                  12555         12522         12571         12522
Ending Stocks                     2376          2231          1697          1865
Stocks to Use Ratio               18.9          17.8          13.5          14.9
Mexico Sugar Supply and Use and High Fructose Corn Syrup Consumption 1/
'''
MEXICO='''
Mexico Sugar Supply and Use and High Fructose Corn Syrup Consumption 1/
Fiscal     Beginning  Production Imports Domestic Exports Ending Stocks
Sugar
2025/26 Est.
Sep                          1123     5325       42     4278      843     1370
Oct                          1123     5325       42     4306     1000     1184
2026/27 Proj.
Sep                          1370     5377       28     4245     1387     1143
Oct                          1184     5377       28     4274     1193     1124
U.S. Cotton Supply and Use  1/
'''

def world_piece(name,region_list,field_count,stop):
    # Test actual values for important market driving totals, all countries.
    result=name+'  1/\n                              2026/27 Proj.\n'
    for region in region_list:
        a=' '.join(['2']*field_count); b=' '.join(['3']*field_count)
        if region=='World' and field_count==7:
            a='125.27 442.35 190.42 385.22 442.44 191.58 124.02'
            b='124.93 442.70 190.82 385.71 442.18 192.00 124.28'
        if region=='World' and field_count==6:
            a='6.51 75.12 13.43 73.38 14.75 6.92'
            b='6.42 75.24 13.63 73.48 14.95 6.85'
        result+=f'  {region}'+('  2/' if region=='World' else '')+'\n'
        result+=f'                      Sep   {a}\n'
        result+=f'                      Oct   {b}\n'
    return result+stop+'  1/\n'

if __name__=='__main__':
    w,wy=parse_us(grain_cases.US_WHEAT,'wheat')
    c,cy=parse_us(grain_cases.US_CORN,'corn')
    ww,ws=parse_world(grain_cases.WORLD_WHEAT,'wheat','Sep','Oct')
    wc,cs=parse_world(grain_cases.WORLD_CORN,'corn','Sep','Oct')
    assert w['exports'][-1]==750 and c['endingstocks'][-1]==1849
    assert float(ww['world']['Oct'][1])==822.39 and float(wc['world']['Oct'][-1])==280.44
    sb,sy=parse_extra_us(US_OILSEEDS,'soybean')
    so,oy=parse_extra_us(US_OILSEEDS,'soyoil')
    su,uy=parse_extra_us(US_SUGAR,'sugar')
    assert sb['production'][-2:]==(4535,4562)
    assert sb['avgfarmprice'][-1]==12
    assert so['industrial'][-2:]==(15305,15305)
    assert so['biofuel'][-1]==17800
    assert so['avgprice'][-1]==70
    assert su['endingstocks'][-1]==1865
    assert tuple(map(float,su['stockstouseratio'][-2:]))==(13.5,14.9)
    assert len(sb)>=len(US_ROWS['soybean']) and len(so)>=len(US_ROWS['soyoil'])
    assert len(su)>=len(US_ROWS['sugar'])
    mx=parse_mexico(MEXICO,'Sep','Oct')
    assert mx['2025/26']['Oct'][-1]==1184
    assert mx['2026/27']['Oct'][-1]==1124
    wsb='World Soybean Supply and Use'; wso='World Soybean Oil Supply and Use'
    soybean=world_piece(wsb,WORLD_REGIONS['soybean'],7,'World Soybean Meal Supply and Use')
    oil=world_piece(wso,WORLD_REGIONS['soyoil'],6,'U.S. Quarterly Animal Product Production')
    rows,year=parse_extra_world(soybean,'soybean','Sep','Oct')
    assert float(rows['world']['Oct'][1])==442.70 and year=='2026/27'
    oil_rows,year=parse_extra_world(oil,'soyoil','Sep','Oct')
    assert float(oil_rows['world']['Oct'][1])==75.24
    assert len(oil_rows)==12
    try:
        parse_extra_world(oil.replace('  North Africa\n','  Wrong Region\n'),'soyoil','Sep','Oct')
    except ValueError:pass
    else:raise AssertionError('Missing regions should abort parsing')
    full=('WASDE - 676                     October 2026\n'
          +grain_cases.US_WHEAT+grain_cases.US_CORN+US_OILSEEDS
          +US_SUGAR+MEXICO+grain_cases.WORLD_WHEAT
          +grain_cases.WORLD_CORN+soybean+oil+(' ' * 40001))
    # Synthetic fixtures do not contain complete World/US/Foreign reconciliations;
    # the production parser must reject that rather than publish bad workbook data.
    try:
        parse_wasde(full,2026,10)
    except ValueError as exc:
        assert 'required' in str(exc) or 'production' in str(exc)
    else:
        raise AssertionError('Incomplete World/Foreign synthetic dataset must be rejected')
    print('PASS: All 5 crop parsers, Mexico, and strict incomplete-source guards.')
