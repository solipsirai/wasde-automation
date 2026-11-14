"""Offline smoke tests for free USDA TXT parsing and 2026 October reference numbers.
These are short excerpts from the official October 2026 WASDE TXT, not a report.
"""
from wasde_automation import parse_us, parse_world, WHEAT_REGIONS, CORN_REGIONS, norm

US_WHEAT = '''
U.S. Wheat Supply and Use  1/
                             2024/25  2025/26 Est. 2026/27 Proj. 2026/27 Proj.
Item                                                           Sep           Oct
Area Planted                      46.3          45.5          42.7          42.5
Area Harvested                    38.6          37.3          32.1          31.9
Yield per Harvested Acre          51.2          53.3          47.8          48.1
Beginning Stocks                   696           855           920           920
Production                        1979          1989          1531          1534
Imports                            150           125           140           145
  Supply, Total                   2825          2969          2591          2599
Food                               969           959           960           960
Seed                                61            58            59            59
Feed and Residual                  111           124            80            90
  Domestic, Total                 1141          1141          1099          1109
Exports                            829           908           775           750
  Use, Total                      1970          2049          1874          1859
Ending Stocks                      855           920           717           740
Avg.FarmPrice ($/bu)  2/          5.52          5.06          6.40          6.30
U.S. Wheat by Class: Supply and Use
'''
US_CORN = '''
U.S. Feed Grain and Corn Supply and Use  1/
                             2024/25  2025/26 Est. 2026/27 Proj. 2026/27 Proj.
Item                                                           Sep           Oct
FEED GRAINS
Area Planted                      101.8         110.0         107.7         107.6
CORN
Area Planted                      90.9          98.7          96.8          96.8
Area Harvested                    83.0          90.9          88.5          88.5
Yield per Harvested Acre         179.3         186.5         178.5         181.2
Beginning Stocks                  1763          1551          1922          2095
Production                       14892         16964         15800         16034
Imports                             22            31            25            25
  Supply, Total                  16677         18546         17747         18154
Feed and Residual                 5437          6119          5950          6000
Food,Seed& Industrial 2/          6816          6903          6955          7005
Ethanol & by-products 3/          5437          5547          5600          5650
  Domestic, Total                12253         13022         12905         13005
Exports                           2873          3429          3275          3300
  Use, Total                     15126         16451         16180         16305
Ending Stocks                     1551          2095          1567          1849
Avg.FarmPrice ($/bu)  4/          4.24          4.16          4.80          4.70
U.S. Sorghum, Barley, and Oats Supply and Use  1/
'''

def ww(r, a, b):
    return f'    {r}\n                      Sep   {a}\n                      Oct   {b}\n'

WORLD_WHEAT = 'World Wheat Supply and Use  1/\n                               (Million Metric Tons)\n                              2026/27 Proj.\n'
world_wheat = {
    'Argentina': ('3.81 21.50 0.01 0.25 7.45 15.50 2.37', '3.81 21.50 0.01 0.25 7.45 16.00 1.87'),
    'World': ('280.60 822.43 207.48 164.73 826.75 211.77 276.29', '281.55 822.39 204.80 166.90 827.90 208.97 276.04'),
}
for region in WHEAT_REGIONS:
    a, b = world_wheat.get(region, ('1 2 3 4 5 6 7', '1 2 3 4 5 6 7'))
    WORLD_WHEAT += ww(region + ('  3/' if region=='World' else ''), a, b)
WORLD_WHEAT += '\nWorld Coarse Grain Supply and Use  1/\n'

WORLD_CORN = 'World Corn Supply and Use  1/\n                               (Million Metric Tons)\n                              2026/27 Proj.\n'
world_corn = {
    'World': ('301.38 1290.95 203.11 827.40 1320.23 210.08 272.10','307.16 1293.75 203.38 826.33 1320.47 210.72 280.44'),
    'United States': ('48.82 401.33 0.64 151.14 327.80 83.19 39.79','53.22 407.28 0.64 152.41 330.34 83.82 46.97'),
}
for region in CORN_REGIONS:
    a, b = world_corn.get(region, ('1 2 3 4 5 6 7', '1 2 3 4 5 6 7'))
    WORLD_CORN += ww(region, a, b)
WORLD_CORN += '\nWorld Rice Supply and Use  1/\n'

if __name__ == '__main__':
    wheat_us, wy = parse_us(US_WHEAT, 'wheat')
    corn_us, cy = parse_us(US_CORN, 'corn')
    wheat_world, season_w = parse_world(WORLD_WHEAT, 'wheat', 'Sep', 'Oct')
    corn_world, season_c = parse_world(WORLD_CORN, 'corn', 'Sep', 'Oct')
    assert wy == ['2024/25','2025/26','2026/27','2026/27']
    assert corn_us['endingstocks'][-2:] == (1567,1849)
    assert corn_us['ethanolbyproducts'][-2:] == (5600,5650)
    assert wheat_us['exports'][-2:] == (775,750)
    assert float(wheat_world['world']['Oct'][1]) == 822.39
    assert float(wheat_world['argentina']['Oct'][-2]) == 16
    assert float(corn_world['world']['Oct'][-1]) == 280.44
    assert float(corn_world['unitedstates']['Oct'][1]) == 407.28
    print('PASS: US wheat/corn mapping + world tables + report examples validated.')
