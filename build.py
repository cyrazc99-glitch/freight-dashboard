# Generator: reads 运价看板_v1.2.xlsx + trends.json -> index.html (self-contained, public data only)
import openpyxl, json, re, os, datetime
D=os.path.dirname(os.path.abspath(__file__))
XLSX='/workspace/运价看板_v1.2.xlsx'
wb=openpyxl.load_workbook(XLSX,data_only=True)
ws=wb['运价看板']
rows=list(ws.iter_rows(values_only=True))
hdr=[str(h) for h in rows[0]]
KEY={'大区':'region','起运地':'origin','起运类型':'originType','目的城市':'city','对应口岸/机场':'port','运输方式':'mode','货型':'cargo','计费单位':'unit',
 '参考单价':'price','币种':'ccy','价格依据':'basis','估算方法':'method','价格范围':'scope','计费重档':'tier','柜型':'equip','超尺寸/不可堆叠整机适用':'oversize',
 '主段时效(天)':'transit','时效边界':'tbound','路径':'route','中转点':'via','末段方式':'lastMode','内陆/末段费用':'inlandCost','内陆/末段费用依据':'inlandBasis',
 '内陆/末段时效':'inlandTime','全程(主段+内陆段)':'full','原末段衔接标签(v1.1)':'legacyTag','记录类型':'rtype','来源发布日期':'pub','采集日期':'collected',
 '报价有效期':'validity','有效期到期日':'validEnd','数据周期':'cycle','数据来源':'source','淡旺季备注':'season','销售使用说明':'note','v1.1原时效(天,口径未拆分)':'legacyTime'}
def col(h): return KEY.get(h)
keys=[col(h) for h in hdr]
def rng(s):
    m=re.match(r'\s*([\d,\.]+)\s*[–-]\s*([\d,\.]+)',str(s or ''))
    return [float(m.group(1).replace(',','')),float(m.group(2).replace(',',''))] if m else None
routes=[];prices={}
for r in rows[1:]:
    if not any(r): continue
    d={k:(str(v).strip() if v is not None else '') for k,v in zip(keys,r) if k}
    d['id']='-'.join([d['origin'],d['city'],d['mode'],d['cargo']])
    routes.append(d)
    prices[d['id']]={'range':rng(d['price']),'text':d['price'],'basis':d['basis']}
def kv(name): return [[str(c) if c is not None else '' for c in r] for r in wb[name].iter_rows(values_only=True) if any(c is not None for c in r)]
usage=kv('使用说明')[1:]; rules=kv('更新规则')[1:]; fields=kv('字段说明')[1:]
src=kv('数据来源明细')
srcRows=[r for r in src[2:] if r[1]]
gaps=[r[0] for r in src if re.match(r'^\d+\.\s',r[0])]
trends=json.load(open(os.path.join(D,'trends.json'),encoding='utf-8'))
now=datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8)))
data={
 'meta':{'title':'出口运价参考看板','version':'v1.2','cycle':'2026-W41','cycleRange':'约 2026-10-06～10-12（部分空运牌价至 10-21）','cadence':'每两周更新',
  'generated':now.strftime('%Y-%m-%d'),'generatedAt':now.strftime('%Y-%m-%d %H:%M')+' (UTC+8)','staleDays':28},
 'routes':routes,
 # Append a new object here each cycle (same route ids) -> sparklines extend automatically
 'snapshots':[{'cycle':'2026-W41','date':'2026-10-09','prices':prices}],
 'trends':trends,
 'about':{'usage':usage,'rules':rules,'fields':fields,'sources':srcRows,'gaps':gaps}
}
tpl=open(os.path.join(D,'template.html'),encoding='utf-8').read()
js=json.dumps(data,ensure_ascii=False,indent=1).replace('</','<\\/')
open(os.path.join(D,'index.html'),'w',encoding='utf-8').write(tpl.replace('/*__DATA__*/null',js))
print('ok',len(routes),'routes',len(trends.get('series',[])),'series')
