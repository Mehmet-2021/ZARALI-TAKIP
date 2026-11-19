"""Cok danisanli meteoroloji takip motoru. Python 3.9+; harici kutuphane gerektirmez."""
import csv, datetime as dt, json, os, re, sys, tempfile, urllib.parse, urllib.request, zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

BASE = Path(__file__).resolve().parent
SABLON = BASE / 'SABLON.xlsx'
KAYIT = BASE / 'danisanlar.csv'
CIKTI = BASE / 'Danisan_Dosyalari'
NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
RNS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
ET.register_namespace('', NS)
ET.register_namespace('r', RNS)
Q = lambda tag: '{%s}%s' % (NS, tag)

def hucre(root, adres):
    for c in root.findall('.//' + Q('c')):
        if c.get('r') == adres:
            return c
    return None

def oku(root, adres, ss):
    c = hucre(root, adres)
    if c is None: return ''
    v = c.find(Q('v'))
    if c.get('t') == 's' and v is not None: return ss[int(v.text)]
    if c.get('t') == 'inlineStr': return ''.join(x.text or '' for x in c.iter(Q('t')))
    return v.text if v is not None and v.text else ''

def yaz(root, adres, deger, numeric=False):
    c = hucre(root, adres)
    if c is None:
        num = int(re.search(r'\d+', adres).group())
        sd = root.find(Q('sheetData'))
        row = next((r for r in sd.findall(Q('row')) if int(r.get('r')) == num), None)
        if row is None:
            row = ET.Element(Q('row'), {'r':str(num)})
            for i, old in enumerate(list(sd)):
                if old.tag == Q('row') and int(old.get('r')) > num:
                    sd.insert(i, row); break
            else: sd.append(row)
        c = ET.SubElement(row, Q('c'), {'r':adres})
    for ch in list(c): c.remove(ch)
    c.attrib.pop('t', None)
    if deger is None: return
    if numeric: ET.SubElement(c, Q('v')).text = str(deger)
    else:
        c.set('t', 'inlineStr')
        ET.SubElement(ET.SubElement(c, Q('is')), Q('t')).text = str(deger)

def workbook_parts(files):
    wb = ET.fromstring(files['xl/workbook.xml'])
    rels = ET.fromstring(files['xl/_rels/workbook.xml.rels'])
    targets = {x.get('Id'): x.get('Target') for x in rels}
    out=[]
    for sh in wb.findall('.//' + Q('sheet')):
        target=targets[sh.get('{%s}id'%RNS)].lstrip('/')
        if not target.startswith('xl/'): target='xl/'+target
        out.append((sh.get('name'),target))
    return out

def strings(files):
    if 'xl/sharedStrings.xml' not in files: return []
    root=ET.fromstring(files['xl/sharedStrings.xml'])
    return [''.join(t.text or '' for t in si.iter(Q('t'))) for si in root.findall(Q('si'))]

def weather(lat,lon,start,end):
    params=urllib.parse.urlencode({'latitude':lat,'longitude':lon,'start_date':start.isoformat(),
        'end_date':end.isoformat(),'daily':'temperature_2m_max,temperature_2m_min','timezone':'Europe/Istanbul'})
    url='https://archive-api.open-meteo.com/v1/archive?'+params
    with urllib.request.urlopen(url,timeout=35) as response:
        data=json.load(response)['daily']
    return {d:(hi,lo) for d,hi,lo in zip(data['time'],data['temperature_2m_max'],data['temperature_2m_min'])}

def run_one(client,template):
    ident=client['danisan_id'].strip()
    if not re.fullmatch(r'[a-zA-Z0-9_-]{1,64}',ident): raise ValueError('Gecersiz danisan_id')
    lat=float(client['enlem'].replace(',','.')); lon=float(client['boylam'].replace(',','.'))
    if not (-90<=lat<=90 and -180<=lon<=180): raise ValueError('Koordinatlar gecersiz')
    dest=CIKTI / (ident+'.xlsx')
    if dest.exists():
        with zipfile.ZipFile(dest) as z: files={name:z.read(name) for name in z.namelist()}
    else: files=dict(template)
    parts=workbook_parts(files); ss=strings(files)
    summary=ET.fromstring(files[parts[0][1]])
    old_lat=oku(summary,'B3',ss); old_lon=oku(summary,'E3',ss)
    # Konum degismisse yeni lokasyonun tum sicaklik serisi yeniden hesaplanir.
    changed=dest.exists() and (abs(float(old_lat)-lat)>0.000001 or abs(float(old_lon)-lon)>0.000001)
    yesterday=dt.date.today()-dt.timedelta(days=1)
    roots=[]; starts=[]
    for name,path in parts[1:]:
        root=ET.fromstring(files[path]); raw=oku(root,'B5',ss)[:10]
        try: start=dt.date.fromisoformat(raw)
        except ValueError: raise ValueError('%s sayfasinda B5 baslangic tarihi gecersiz: %r'%(name,raw))
        roots.append((name,path,root,start))
        if start<=yesterday: starts.append(start)
    if not starts:
        print(ident+': Baslangic tarihi henuz gelmemis');return
    earliest=min(starts)
    # Mevcut sablon B10:D159: en fazla 150 gun. Guncel 150 gunu gostermek icin
    # eski baslangici degistirmek biyolojik toplamlarin anlamini degistireceginden REDDEDILIR.
    if (yesterday-earliest).days>=150:
        raise ValueError('150 gunluk sablon siniri asildi; tarih araligi kisaltilmadan otomatik kayit yapilmadi')
    daily=weather(lat,lon,earliest,yesterday)
    if not daily: raise ValueError('Meteoroloji servisinden veri donmedi')
    now=dt.datetime.now().strftime('%d.%m.%Y %H:%M')
    updated=0
    for name,path,root,start in roots:
        if start>yesterday: continue
        # Yalnizca dolu ve dogrulanmis gunleri degistir, diger gunlere dokunma.
        for i in range((yesterday-start).days+1):
            day=start+dt.timedelta(days=i); temps=daily.get(day.isoformat())
            if not temps or temps[0] is None or temps[1] is None: continue
            row=10+i
            yaz(root,'B'+str(row),day.strftime('%d.%m.%Y'))
            yaz(root,'C'+str(row),temps[0],True)
            yaz(root,'D'+str(row),temps[1],True)
            updated+=1
        yaz(root,'E5','Son Guncelleme: '+now)
        files[path]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    if updated==0: raise ValueError('Guncellenebilir gunluk veri bulunamadi')
    yaz(summary,'B3',lat,True);yaz(summary,'E3',lon,True)
    yaz(summary,'G3','Son Guncelleme: '+now+' | '+ident)
    files[parts[0][1]]=ET.tostring(summary,encoding='utf-8',xml_declaration=True)
    CIKTI.mkdir(exist_ok=True)
    fd,tmp=tempfile.mkstemp(suffix='.xlsx',dir=CIKTI);os.close(fd)
    try:
        with zipfile.ZipFile(tmp,'w',compression=zipfile.ZIP_DEFLATED) as z:
            for path,contents in files.items(): z.writestr(path,contents)
        with zipfile.ZipFile(tmp) as z:
            if z.testzip(): raise ValueError('Olusturulan Excel arsivi bozuk')
        os.replace(tmp,dest)
    finally:
        if os.path.exists(tmp): os.unlink(tmp)
    print('%s: %d sicaklik kaydi guncellendi%s'%(ident,updated,' (konum degisti)' if changed else ''))

def main():
    if not SABLON.exists(): raise FileNotFoundError('SABLON.xlsx bulunamadi')
    with zipfile.ZipFile(SABLON) as z: template={n:z.read(n) for n in z.namelist()}
    with KAYIT.open(encoding='utf-8-sig',newline='') as f: clients=list(csv.DictReader(f))
    ids=set();fail=0
    for c in clients:
        if c['aktif'].strip().lower() not in ('1','evet','true'): continue
        try:
            if c['danisan_id'] in ids: raise ValueError('Tekrarlanan danisan_id')
            ids.add(c['danisan_id']);run_one(c,template)
        except Exception as e:
            fail+=1;print('[HATA] %s: %s'%(c.get('danisan_id','?'),e),file=sys.stderr)
    print('Tamamlandi. Basarisiz danisan sayisi:',fail)
    return 1 if fail else 0

if __name__=='__main__': sys.exit(main())
