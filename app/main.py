import asyncio, io, os, re, smtplib, sqlite3
from contextlib import asynccontextmanager
from dataclasses import dataclass
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
import httpx
from apscheduler.schedulers.asyncio import AsyncIOScheduler
from fastapi import FastAPI, Form, Request
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from openpyxl import Workbook

DB=Path(os.getenv('DB_PATH','/tmp/cee_tenders.db')); TZ=os.getenv('DIGEST_TIMEZONE','Europe/Warsaw'); HOUR=int(os.getenv('DIGEST_HOUR','9')); scheduler=AsyncIOScheduler(timezone=TZ)
MARKETS={'UA':'Україна','PL':'Polska','RO':'România','CZ':'Česko','LT':'Lietuva','LV':'Latvija','EE':'Eesti','SK':'Slovensko','SI':'Slovenija','HR':'Hrvatska','BG':'България','HU':'Magyarország','DE':'Deutschland','FR':'France','IT':'Italia','ES':'España','EU':'EU'}
TED_CODES={'PL':'POL','RO':'ROU','CZ':'CZE','LT':'LTU','LV':'LVA','EE':'EST','SK':'SVK','SI':'SVN','HR':'HRV','BG':'BGR','HU':'HUN','DE':'DEU','FR':'FRA','IT':'ITA','ES':'ESP'}
ACTIVE={'active.tendering','active.pre-qualification','active.pre-qualification.stand-still','active.auction','active.qualification','active.awarded'}
@dataclass
class Tender: source:str; market:str; external_id:str; title:str; description:str=''; status:str='active'; amount:float=0; currency:str=''; buyer:str=''; deadline:str=''; published_at:str=''; url:str=''; cpv:str=''
def conn(): DB.parent.mkdir(parents=True,exist_ok=True);c=sqlite3.connect(DB);c.row_factory=sqlite3.Row;return c
def now():return datetime.now(timezone.utc).isoformat()
def norm(v):return re.sub(r'\s+',' ',str(v or '').casefold()).strip()
def first(v,default=''):
 if isinstance(v,list):return first(v[0],default) if v else default
 if isinstance(v,dict):
  for lang in ('eng','pol','ron','ces','lit','lav','est','slk'):
   if lang in v:return first(v[lang],default)
  return first(next(iter(v.values())),default) if v else default
 return v if v not in (None,'') else default
def init():
 with conn() as c:
  c.executescript('''CREATE TABLE IF NOT EXISTS kams(id INTEGER PRIMARY KEY,name TEXT,email TEXT,telegram_chat_id TEXT,language TEXT DEFAULT 'uk',telegram_enabled INTEGER DEFAULT 0,email_enabled INTEGER DEFAULT 0,active INTEGER DEFAULT 1);CREATE TABLE IF NOT EXISTS kam_markets(kam_id INTEGER,market TEXT,PRIMARY KEY(kam_id,market));CREATE TABLE IF NOT EXISTS filters(id INTEGER PRIMARY KEY,kam_id INTEGER,name TEXT,keywords TEXT,min_amount REAL DEFAULT 0,cpv TEXT DEFAULT '',active INTEGER DEFAULT 1);CREATE TABLE IF NOT EXISTS tenders(uid TEXT PRIMARY KEY,source TEXT,market TEXT,external_id TEXT,title TEXT,description TEXT,status TEXT,amount REAL,currency TEXT,buyer TEXT,deadline TEXT,published_at TEXT,url TEXT,cpv TEXT,first_seen TEXT,last_seen TEXT);CREATE TABLE IF NOT EXISTS matches(kam_id INTEGER,filter_id INTEGER,tender_uid TEXT,matched_keywords TEXT,created_at TEXT,PRIMARY KEY(kam_id,filter_id,tender_uid));CREATE TABLE IF NOT EXISTS scans(source TEXT PRIMARY KEY,last_run TEXT,last_status TEXT,last_count INTEGER,last_error TEXT);''')
  if not c.execute('SELECT 1 FROM kams').fetchone():
   kid=c.execute('INSERT INTO kams(name,language) VALUES(?,?)',('Vitaliy Drofyak','uk')).lastrowid
   for m in ('UA','PL','LT','LV','EE'):c.execute('INSERT INTO kam_markets VALUES(?,?)',(kid,m))
   c.execute('INSERT INTO filters(kam_id,name,keywords,cpv) VALUES(?,?,?,?)',(kid,'Memory & Storage','SSD,NVMe,M.2,DDR4,DDR5,RAM,memory,storage,flash,Goodram,Kioxia,накопичувач,пам’ять,dysk,pamięć','3023,30234'))
async def prozorro(markets):
 if 'UA' not in markets:return []
 base=os.getenv('PROZORRO_API_BASE','https://public-api.prozorro.gov.ua/api/2.5').rstrip('/')
 async with httpx.AsyncClient(follow_redirects=True,headers={'User-Agent':'CEE-Tender-Intelligence/5.0'}) as client:
  r=await client.get(f'{base}/tenders',params={'limit':100,'descending':1},timeout=40);r.raise_for_status();listing=r.json().get('data',[]);sem=asyncio.Semaphore(12)
  async def one(x):
   async with sem:
    try:q=await client.get(f'{base}/tenders/{x["id"]}',timeout=30);q.raise_for_status();return q.json().get('data',{})
    except Exception as e:print('[PROZORRO detail]',e,flush=True)
  rows=await asyncio.gather(*(one(x) for x in listing))
 out=[]
 for x in filter(None,rows):
  if x.get('status') not in ACTIVE:continue
  value=x.get('value') or {};items=x.get('items') or [];ext=x.get('tenderID') or x.get('id','');desc=' '.join(i.get('description','') for i in items);cpv=','.join(sorted({(i.get('classification') or {}).get('id','') for i in items if (i.get('classification') or {}).get('id')}))
  out.append(Tender('PROZORRO','UA',x.get('id',''),x.get('title') or desc,x.get('description') or desc,x.get('status',''),float(value.get('amount') or 0),value.get('currency','UAH'),(x.get('procuringEntity') or {}).get('name',''),(x.get('tenderPeriod') or {}).get('endDate',''),x.get('dateModified',''),f'https://prozorro.gov.ua/tender/{ext}',cpv))
 return out
async def ted(markets):
 countries=[TED_CODES[m] for m in markets if m in TED_CODES]
 if not countries and 'EU' not in markets:return []
 query='publication-date >= 20260101'+((' AND ('+' OR '.join(f'buyer-country = {x}' for x in countries)+')') if countries else '')
 fields=['notice-identifier','publication-number','notice-title','title-proc','description-proc','buyer-name','buyer-country','deadline','estimated-value-proc','estimated-value-cur-proc','classification-cpv','publication-date','notice-type','links']
 body={'query':query,'fields':fields,'page':1,'limit':100,'scope':'ACTIVE','onlyLatestVersions':True};url=os.getenv('TED_API_BASE','https://api.ted.europa.eu/v3/notices/search')
 async with httpx.AsyncClient(follow_redirects=True) as client:r=await client.post(url,json=body,timeout=60);r.raise_for_status();rows=r.json().get('notices',[])
 reverse={v:k for k,v in TED_CODES.items()};out=[]
 for x in rows:
  nid=str(first(x.get('notice-identifier')) or first(x.get('publication-number')));links=x.get('links') or {};html=(links.get('html') or links.get('htmlDirect') or {}) if isinstance(links,dict) else {};u=first(html) or (f'https://ted.europa.eu/en/notice/-/detail/{nid}' if nid else '')
  out.append(Tender('TED',reverse.get(first(x.get('buyer-country')),'EU'),nid,str(first(x.get('notice-title')) or first(x.get('title-proc')) or 'TED notice'),str(first(x.get('description-proc'))),str(first(x.get('notice-type'),'active')),float(first(x.get('estimated-value-proc'),0) or 0),str(first(x.get('estimated-value-cur-proc'),'EUR')),str(first(x.get('buyer-name'))),str(first(x.get('deadline'))),str(first(x.get('publication-date'))),u,str(first(x.get('classification-cpv')))))
 return out
async def bzp(markets):
 if 'PL' not in markets:return []
 url=os.getenv('BZP_API_BASE','https://ezamowienia.gov.pl/mo-board/api/v1/notice')
 async with httpx.AsyncClient(follow_redirects=True,headers={'Accept':'application/json'}) as client:r=await client.get(url,params={'PageSize':100,'PageNumber':1},timeout=60);r.raise_for_status();p=r.json()
 rows=p if isinstance(p,list) else p.get('items') or p.get('data') or p.get('results') or [];out=[]
 for x in rows:
  nid=str(x.get('id') or x.get('noticeId') or x.get('bzpNumber') or x.get('noticeNumber') or '');title=x.get('title') or x.get('orderName') or x.get('noticeTitle') or 'Ogłoszenie BZP';u=x.get('url') or (f'https://ezamowienia.gov.pl/mo-client-board/bzp/notice-details/{nid}' if nid else 'https://ezamowienia.gov.pl/')
  out.append(Tender('BZP','PL',nid,title,str(x.get('description') or ''),str(x.get('status') or 'active'),float(x.get('value') or x.get('amount') or 0),str(x.get('currency') or 'PLN'),x.get('organizationName') or x.get('contractingAuthorityName') or x.get('buyerName') or '',x.get('submissionDeadline') or x.get('offerDeadline') or x.get('deadline') or '',x.get('publicationDate') or x.get('publishedAt') or '',u,str(x.get('cpvCode') or '')))
 return out
def match_and_store(rows):
 new=0
 with conn() as c:
  fs=c.execute('SELECT f.*,GROUP_CONCAT(km.market) markets FROM filters f JOIN kams k ON k.id=f.kam_id LEFT JOIN kam_markets km ON km.kam_id=k.id WHERE f.active=1 AND k.active=1 GROUP BY f.id').fetchall()
  for t in rows:
   uid=f'{t.source}:{t.external_id}';c.execute('''INSERT INTO tenders VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(uid) DO UPDATE SET title=excluded.title,description=excluded.description,status=excluded.status,amount=excluded.amount,currency=excluded.currency,buyer=excluded.buyer,deadline=excluded.deadline,published_at=excluded.published_at,url=excluded.url,cpv=excluded.cpv,last_seen=excluded.last_seen''',(uid,t.source,t.market,t.external_id,t.title,t.description,t.status,t.amount,t.currency,t.buyer,t.deadline,t.published_at,t.url,t.cpv,now(),now()))
   text=norm(' '.join((t.title,t.description,t.buyer,t.cpv)))
   for f in fs:
    keys=[x.strip() for x in f['keywords'].split(',') if x.strip()];hits=[x for x in keys if norm(x) in text];cpvs=[x.strip() for x in (f['cpv'] or '').split(',') if x.strip()];cpv_ok=not cpvs or any(t.cpv.startswith(x) for x in cpvs)
    if t.market in (f['markets'] or '').split(',') and hits and cpv_ok and t.amount>=f['min_amount']:
     new+=c.execute('INSERT OR IGNORE INTO matches VALUES(?,?,?,?,?)',(f['kam_id'],f['id'],uid,','.join(hits),now())).rowcount
 return new
async def scan_all():
 with conn() as c:markets={r['market'] for r in c.execute('SELECT DISTINCT km.market FROM kam_markets km JOIN kams k ON k.id=km.kam_id WHERE k.active=1')}
 result=[]
 for code,fn in [('PROZORRO',prozorro),('TED',ted),('BZP',bzp)]:
  try:rows=await fn(markets);new=match_and_store(rows);status=f'OK: {len(rows)} read, {new} new matches';error=''
  except Exception as e:rows=[];status='ERROR';error=f'{type(e).__name__}: {e}';print(f'[{code}] {error}',flush=True)
  with conn() as c:c.execute('INSERT INTO scans VALUES(?,?,?,?,?) ON CONFLICT(source) DO UPDATE SET last_run=excluded.last_run,last_status=excluded.last_status,last_count=excluded.last_count,last_error=excluded.last_error',(code,now(),status,len(rows),error))
  result.append((code,status))
 return result
async def send_tg(chat,text):
 token=os.getenv('TELEGRAM_BOT_TOKEN','').strip()
 if not token or not chat:return
 async with httpx.AsyncClient() as client:r=await client.post(f'https://api.telegram.org/bot{token}/sendMessage',json={'chat_id':chat,'text':text[:4000],'disable_web_page_preview':True},timeout=30);r.raise_for_status()
def send_mail(to,text):
 host=os.getenv('SMTP_HOST');sender=os.getenv('SMTP_FROM') or os.getenv('SMTP_USER')
 if not host or not sender or not to:return
 m=EmailMessage();m['From']=sender;m['To']=to;m['Subject']='CEE Tender Intelligence';m.set_content(text)
 with smtplib.SMTP(host,int(os.getenv('SMTP_PORT','587'))) as s:
  s.starttls()
  if os.getenv('SMTP_USER'):s.login(os.getenv('SMTP_USER'),os.getenv('SMTP_PASSWORD',''))
  s.send_message(m)
async def digests():
 await scan_all()
 with conn() as c:kams=c.execute('SELECT * FROM kams WHERE active=1').fetchall()
 for k in kams:
  with conn() as c:rows=c.execute('SELECT DISTINCT t.* FROM tenders t JOIN matches m ON m.tender_uid=t.uid WHERE m.kam_id=? ORDER BY t.published_at DESC LIMIT 30',(k['id'],)).fetchall()
  lines=[f'CEE Tender Intelligence · {datetime.now().date()}',f'Relevant tenders: {len(rows)}']+[f'• [{r["market"]}/{r["source"]}] {r["title"][:130]}\n{r["amount"]:,.2f} {r["currency"]} · {r["deadline"][:10]}\n{r["url"]}' for r in rows[:20]];text='\n'.join(lines)
  if k['telegram_enabled'] and k['telegram_chat_id']:
   try:await send_tg(k['telegram_chat_id'],text)
   except Exception as e:print('[TELEGRAM]',e,flush=True)
  if k['email_enabled'] and k['email']:
   try:await asyncio.to_thread(send_mail,k['email'],text)
   except Exception as e:print('[EMAIL]',e,flush=True)
@asynccontextmanager
async def life(app):
 init();scheduler.add_job(digests,'cron',day_of_week='mon-fri',hour=HOUR,minute=0,id='digest',replace_existing=True);scheduler.start();yield;scheduler.shutdown(False)
app=FastAPI(title='CEE Tender Intelligence v5',lifespan=life);base=Path(__file__).parent;templates=Jinja2Templates(directory=base/'templates');app.mount('/static',StaticFiles(directory=base/'static'),name='static')
@app.get('/health')
def health():return {'status':'ok'}
@app.get('/',response_class=HTMLResponse)
def home(request:Request):
 with conn() as c:kams=c.execute('SELECT k.*,GROUP_CONCAT(km.market) markets FROM kams k LEFT JOIN kam_markets km ON km.kam_id=k.id GROUP BY k.id').fetchall();filters=c.execute('SELECT f.*,k.name kam_name FROM filters f JOIN kams k ON k.id=f.kam_id').fetchall();tenders=c.execute('SELECT * FROM tenders ORDER BY published_at DESC,first_seen DESC LIMIT 200').fetchall();scans=c.execute('SELECT * FROM scans ORDER BY source').fetchall()
 return templates.TemplateResponse('index.html',{'request':request,'kams':kams,'filters':filters,'tenders':tenders,'scans':scans,'markets':MARKETS})
@app.post('/scan')
async def scan():await scan_all();return RedirectResponse('/',303)
@app.post('/digest')
async def digest():await digests();return RedirectResponse('/',303)
@app.post('/kams')
def add_kam(name:str=Form(...),email:str=Form(''),telegram_chat_id:str=Form(''),language:str=Form('uk'),channels:list[str]=Form([]),markets:list[str]=Form([])):
 with conn() as c:
  kid=c.execute('INSERT INTO kams(name,email,telegram_chat_id,language,telegram_enabled,email_enabled) VALUES(?,?,?,?,?,?)',(name,email,telegram_chat_id,language,int('telegram' in channels),int('email' in channels))).lastrowid
  for m in markets:c.execute('INSERT OR IGNORE INTO kam_markets VALUES(?,?)',(kid,m))
 return RedirectResponse('/',303)
@app.post('/filters')
def add_filter(kam_id:int=Form(...),name:str=Form(...),keywords:str=Form(...),min_amount:float=Form(0),cpv:str=Form('')):
 with conn() as c:c.execute('INSERT INTO filters(kam_id,name,keywords,min_amount,cpv) VALUES(?,?,?,?,?)',(kam_id,name,keywords,min_amount,cpv))
 return RedirectResponse('/',303)
@app.get('/export.xlsx')
def export():
 with conn() as c:rows=c.execute('SELECT * FROM tenders ORDER BY published_at DESC').fetchall()
 wb=Workbook();ws=wb.active;ws.append(['Source','Market','ID','Title','Buyer','Amount','Currency','Deadline','Published','CPV','URL'])
 for r in rows:ws.append([r['source'],r['market'],r['external_id'],r['title'],r['buyer'],r['amount'],r['currency'],r['deadline'],r['published_at'],r['cpv'],r['url']])
 b=io.BytesIO();wb.save(b);b.seek(0);return StreamingResponse(b,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename=cee_tenders.xlsx'})
