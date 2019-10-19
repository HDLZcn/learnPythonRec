from example.commons import Faker
from pyecharts import options as opts
from pyecharts.charts import Map,Timeline
import xlrd
wb = xlrd.open_workbook(r'D:\data\wb.xlsx')
wbs = wb.sheet_by_index(0)
range_month =   wbs.row(1)
del range_month[0]
del range_month[0]

def SAP_provience():
    SAP = []
    for n in range(5,94,3):
        SAP.append(wbs.cell(n,0).value)
    return SAP
    
def SAP_values(idx):
    idk = idx+2
    n= 0
    SAP = []
    for n in range(5,94,3):
       # print(n,idk)
        SAP.append(wbs.cell(n,idk).value*200)
    return SAP

def map_timeLine() -> Timeline:
    tl = Timeline()
    for i in range_month:
        idx = range_month.index(i)
        map0 = (
            Map()
            .add(
                "东北米长粒", [list(z) for z in zip(SAP_provience(), SAP_values(idx))], "china"
            )
            .set_global_opts(
                title_opts=opts.TitleOpts(title="Map-{}分省份长粒销量占全部销量数据".format(i)),
                visualmap_opts=opts.VisualMapOpts(max_=200),
            )
        )
        tl.add(map0, "{}月".format(i))
    return tl
map_timeLine().render(path=r"D:/data/长粒每月省占率变化.html")
