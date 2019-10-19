from pyecharts import options as opts
from pyecharts.charts import Bar, Page, Pie, Timeline
import xlrd
wb = xlrd.open_workbook(r'D:\data\wb5.xlsx')
wbs = wb.sheet_by_index(0)
range_month =   wbs.row(1)
del range_month[0]
del range_month[0]

def SAP_provience():
    SAP = []
    for n in range(2,153,5):
        SAP.append(wbs.cell(n,1).value)
    return SAP
    
def SAP_j_values(idx):
    idk = idx+2
    n= 0
    SAP = []
    for n in range(2,153,5):
       # print(n,idk)
        SAP.append(wbs.cell(n,idk).value)
    return SAP

def SAP_x_values(idx):
    idk = idx+2
    n= 0
    SAP = []
    for n in range(4,153,5):
       # print(n,idk)
        SAP.append(wbs.cell(n,idk).value)
    return SAP

def SAP_z_values(idx):
    idk = idx+2
    n= 0
    SAP = []
    for n in range(5,158,5):
       # print(n,idk)
        SAP.append(wbs.cell(n,idk).value)
    return SAP

def timeline_bar() -> Timeline:
    x = SAP_provience()
    tl = Timeline()
    for i in range_month:
        idx = range_month.index(i)
        bar = (
            Bar()
            .add_xaxis(x)
            .add_yaxis("金龙鱼", SAP_j_values(idx))
            .add_yaxis("香满园", SAP_x_values(idx))
           # .add_yaxis("总计", SAP_x_values(idx)+SAP_j_values(idx))
            
            #.add_yaxis("总计", SAP_z_values(idx))
            .set_global_opts(title_opts=opts.TitleOpts("月度分省份销量".format(i)))
        )
        tl.add(bar, "{}月".format(i))
    return tl
timeline_bar().render(path=r"D:/data/长粒每月省占率变化_xy.html")
