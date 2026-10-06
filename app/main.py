import asyncio, hashlib, hmac, io, os, re, smtplib, sqlite3
from contextlib import asynccontextmanager
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
from telegram import Bot, Update
from telegram.ext import Application, CommandHandler, ContextTypes

DB=Path('/data/cee_tenders.db'); TOKEN=os.getenv('TELEGRAM_BOT_TOKEN','').strip(); API=os.getenv('PROZORRO_API_BASE','https://public-api.prozorro.gov.ua/api/2.5').rstrip('/')
TZ=os.getenv('DIGEST_TIMEZONE','Europe/Warsaw'); HOUR=int(os.getenv('DIGEST_HOUR','9')); scheduler=AsyncIOScheduler(timezone=TZ); tg=None
ACTIVE={'active.tendering','active.pre-qualification','active.pre-qualification.stand-still','active.auction','active.qualification','active.awarded'}
MARKETS=[('UA','Україна','Prozorro',1),('EU','Європейський Союз','TED',0),('PL','Польща','e-Zamówienia',0),('RO','Румунія','SEAP/SICAP',0),('CZ','Чехія','NEN',0),('LT','Литва','CVP IS',0),('LV','Латвія','EIS',0),('EE','Естонія','Riigihanked',0),('SK','Словаччина','UVO',0)]

def db(): c=sqlite3.connect(DB); c.row_factory=sqlite3.Row; return c
def now(): return datetime.now(timezone.utc).isoformat()
def init():
 DB.parent.mkdir(parents=True,exist_ok=True)
 with db() as c:
  c.executescript('''CREATE TABLE IF NOT EXISTS markets(code TEXT PRIMARY KEY,name TEXT,source TEXT,enabled INTEGER);
  CREATE TABLE IF NOT EXISTS kams(id INTEGER PRIMARY KEY,name TEXT NOT NULL,email TEXT,telegram_chat_id TEXT,language TEXT DEFAULT 'uk',telegram_enabled INTEGER DEFAULT 1,email_enabled INTEGER DEFAULT 0,active INTEGER DEFAULT 1);
  CREATE TABLE IF NOT EXISTS kam_markets(kam_id INTEGER,market_code TEXT,PRIMARY KEY(kam_id,market_code));
  CREATE TABLE IF NOT EXISTS filters(id INTEGER PRIMARY KEY,kam_id INTEGER,name TEXT,keywords TEXT,min_amount REAL DEFAULT 0,active INTEGER DEFAULT 1);
  CREATE TABLE IF NOT EXISTS tenders(id TEXT PRIMARY KEY,market_code TEXT,external_id TEXT,title TEXT,description TEXT,status TEXT,amount REAL,currency TEXT,buyer TEXT,deadline TEXT,url TEXT,first_seen TEXT);
  CREATE TABLE IF NOT EXISTS matches(kam_id INTEGER,filter_id INTEGER,tender_id TEXT,first_seen TEXT,PRIMARY KEY(kam_id,filter_id,tender_id));
  CREATE TABLE IF NOT EXISTS tracked(kam_id INTEGER,tender_id TEXT,status TEXT DEFAULT 'Review',reminder_date TEXT,note TEXT,PRIMARY KEY(kam_id,tender_id));''')
  for x in MARKETS:c.execute('INSERT OR IGNORE INTO markets VALUES(?,?,?,?)',x)
  if not c.execute('SELECT 1 FROM kams').fetchone():
   cur=c.execute('INSERT INTO kams(name,language,telegram_enabled,email_enabled) VALUES(?,?,?,?)',('Vitaliy Drofyak','uk',1,0)); kid=cur.lastrowid
   c.execute('INSERT INTO kam_markets VALUES(?,?)',(kid,'UA')); c.execute('INSERT INTO filters(kam_id,name,keywords,min_amount) VALUES(?,?,?,?)',(kid,'Пам’ять і накопичувачі','SSD,NVMe,M.2,SATA SSD,DDR4,DDR5,RAM,оперативна пам’ять,серверна пам’ять,enterprise SSD,USB Flash,microSD,Goodram,Kioxia',0))
def norm(v): return re.sub(r'\s+',' ',(v or '').lower()).strip()
def hits(text,keys): return [x.strip() for x in keys.split(',') if x.strip() and norm(x) in norm(text)]
async def scan_prozorro():
 with db() as c: fs=c.execute("SELECT f.*,k.name FROM filters f JOIN kams k ON k.id=f.kam_id JOIN kam_markets km ON km.kam_id=k.id WHERE f.active=1 AND k.active=1 AND km.market_code='UA'").fetchall()
 if not fs:return 0
 async with httpx.AsyncClient(headers={'User-Agent':'CEE-Tender-Intelligence/1.0'}) as client:
  r=await client.get(f'{API}/tenders',params={'limit':100},timeout=30); r.raise_for_status(); listing=r.json().get('data',[])
  sem=asyncio.Semaphore(10)
  async def one(x):
   async with sem:
    try:return (await client.get(f'{API}/tenders/{x["id"]}',timeout=25)).json().get('data',{})
    except:return None
  data=await asyncio.gather(*(one(x) for x in listing))
 n=0
 with db() as c:
  for raw in filter(None,data):
   if raw.get('status') not in ACTIVE:continue
   value=raw.get('value') or {}; buyer=(raw.get('procuringEntity') or {}).get('name',''); items=' '.join(i.get('description','') for i in raw.get('items',[])); text=' '.join([raw.get('title',''),raw.get('description',''),items,buyer]); tid=raw.get('id'); ext=raw.get('tenderID',tid); deadline=(raw.get('tenderPeriod') or {}).get('endDate','')
   for f in fs:
    hh=hits(text,f['keywords']); amount=value.get('amount') or 0
    if hh and amount>=f['min_amount']:
     c.execute('''INSERT INTO tenders VALUES(?,?,?,?,?,?,?,?,?,?,?,?) ON CONFLICT(id) DO UPDATE SET status=excluded.status,amount=excluded.amount,deadline=excluded.deadline''',(tid,'UA',ext,raw.get('title',''),raw.get('description',''),raw.get('status',''),amount,value.get('currency','UAH'),buyer,deadline,f'https://prozorro.gov.ua/tender/{ext}',now()))
     cur=c.execute('INSERT OR IGNORE INTO matches VALUES(?,?,?,?)',(f['kam_id'],f['id'],tid,now())); n+=cur.rowcount
 return n

def email_send(to,subject,body):
 host=os.getenv('SMTP_HOST'); sender=os.getenv('SMTP_FROM') or os.getenv('SMTP_USER')
 if not host or not to or not sender:return False
 msg=EmailMessage();msg['From']=sender;msg['To']=to;msg['Subject']=subject;msg.set_content(body)
 with smtplib.SMTP(host,int(os.getenv('SMTP_PORT','587'))) as s:
  if os.getenv('SMTP_STARTTLS','true').lower()=='true':s.starttls()
  if os.getenv('SMTP_USER'):s.login(os.getenv('SMTP_USER'),os.getenv('SMTP_PASSWORD',''))
  s.send_message(msg)
 return True

def digest_text(kam,rows,tracked):
 pl=kam['language']=='pl'; title='Poranny raport przetargowy' if pl else 'Ранковий тендерний звіт'; lines=[f'{title}: {datetime.now().date()}',('Nowe/aktywne: ' if pl else 'Нові/активні: ')+str(len(rows))]
 for r in rows[:15]:lines.append(f'• [{r["market_code"]}] {r["title"][:150]}\n  {r["amount"]:,.2f} {r["currency"]} | {r["deadline"][:10] if r["deadline"] else "-"}\n  {r["url"]}')
 if tracked:
  lines.append('\n⭐ '+('Obserwowane:' if pl else 'Відслідковувані:'))
  for r in tracked[:10]:lines.append(f'• {r["title"][:120]} | {r["status"]} | {r["reminder_date"] or r["deadline"][:10]}\n  {r["url"]}')
 return '\n'.join(lines)
async def send_digests():
 try:await scan_prozorro()
 except:pass
 with db() as c:kams=c.execute('SELECT * FROM kams WHERE active=1').fetchall()
 for k in kams:
  with db() as c:
   rows=c.execute('''SELECT DISTINCT t.* FROM tenders t JOIN matches m ON m.tender_id=t.id JOIN kam_markets km ON km.kam_id=m.kam_id AND km.market_code=t.market_code WHERE m.kam_id=? ORDER BY t.first_seen DESC LIMIT 50''',(k['id'],)).fetchall()
   tr=c.execute('SELECT t.*,x.status,x.reminder_date,x.note FROM tracked x JOIN tenders t ON t.id=x.tender_id WHERE x.kam_id=?',(k['id'],)).fetchall()
  text=digest_text(k,rows,tr)
  if k['telegram_enabled'] and TOKEN and k['telegram_chat_id']:
   bot=Bot(TOKEN)
   for i in range(0,len(text),3900):await bot.send_message(k['telegram_chat_id'],text[i:i+3900],disable_web_page_preview=True)
  if k['email_enabled'] and k['email']:
   try:await asyncio.to_thread(email_send,k['email'],'CEE Tender Intelligence',text)
   except:pass
async def tg_start(update:Update,context:ContextTypes.DEFAULT_TYPE):
 code=(context.args[0] if context.args else '').strip()
 if not code.isdigit():return await update.message.reply_text('Відкрийте профіль KAM у веб-панелі та використайте команду /start ID.')
 with db() as c:
  row=c.execute('SELECT id FROM kams WHERE id=?',(int(code),)).fetchone()
  if row:c.execute('UPDATE kams SET telegram_chat_id=? WHERE id=?',(str(update.effective_chat.id),int(code)))
 await update.message.reply_text('Telegram підключено / Telegram połączony.' if row else 'KAM не знайдений / Nie znaleziono KAM.')
async def tg_digest(update:Update,context:ContextTypes.DEFAULT_TYPE): await send_digests(); await update.message.reply_text('Готово / Gotowe')
@asynccontextmanager
async def life(app):
 global tg;init();scheduler.add_job(send_digests,'cron',day_of_week='mon-fri',hour=HOUR,minute=0,id='digest',replace_existing=True);scheduler.start()
 if TOKEN:
  tg=Application.builder().token(TOKEN).build();tg.add_handler(CommandHandler('start',tg_start));tg.add_handler(CommandHandler('digest',tg_digest));await tg.initialize();await tg.start();await tg.updater.start_polling()
 yield
 scheduler.shutdown(False)
 if tg:await tg.updater.stop();await tg.stop();await tg.shutdown()
app=FastAPI(title='CEE Tender Intelligence',lifespan=life);templates=Jinja2Templates(directory=Path(__file__).parent/'templates');app.mount('/static',StaticFiles(directory=Path(__file__).parent/'static'),name='static')
@app.get('/',response_class=HTMLResponse)
def home(request:Request):
 with db() as c:
  kams=c.execute('''SELECT k.*,GROUP_CONCAT(km.market_code) markets FROM kams k LEFT JOIN kam_markets km ON km.kam_id=k.id GROUP BY k.id ORDER BY k.name''').fetchall();markets=c.execute('SELECT * FROM markets ORDER BY code').fetchall();tenders=c.execute('SELECT * FROM tenders ORDER BY first_seen DESC LIMIT 100').fetchall();filters=c.execute('SELECT f.*,k.name kam_name FROM filters f JOIN kams k ON k.id=f.kam_id ORDER BY k.name').fetchall();tracked=c.execute('SELECT x.*,k.name kam_name,t.title FROM tracked x JOIN kams k ON k.id=x.kam_id JOIN tenders t ON t.id=x.tender_id').fetchall()
 return templates.TemplateResponse('index.html',{'request':request,'kams':kams,'markets':markets,'tenders':tenders,'filters':filters,'tracked':tracked,'token':bool(TOKEN),'smtp':bool(os.getenv('SMTP_HOST'))})
@app.post('/kams')
def add_kam(name:str=Form(...),email:str=Form(''),language:str=Form('uk'),channels:list[str]=Form([]),markets:list[str]=Form([])):
 with db() as c:
  cur=c.execute('INSERT INTO kams(name,email,language,telegram_enabled,email_enabled) VALUES(?,?,?,?,?)',(name,email,language,int('telegram' in channels),int('email' in channels)));kid=cur.lastrowid
  for m in markets:c.execute('INSERT OR IGNORE INTO kam_markets VALUES(?,?)',(kid,m))
 return RedirectResponse('/',303)
@app.post('/kams/{kid}/delete')
def del_kam(kid:int):
 with db() as c:
  for t in ['kam_markets','filters','tracked','matches']:c.execute(f'DELETE FROM {t} WHERE kam_id=?',(kid,))
  c.execute('DELETE FROM kams WHERE id=?',(kid,))
 return RedirectResponse('/',303)
@app.post('/filters')
def add_filter(kam_id:int=Form(...),name:str=Form(...),keywords:str=Form(...),min_amount:float=Form(0)):
 with db() as c:c.execute('INSERT INTO filters(kam_id,name,keywords,min_amount) VALUES(?,?,?,?)',(kam_id,name,keywords,min_amount))
 return RedirectResponse('/',303)
@app.post('/track/{tid}')
def track(tid:str,kam_id:int=Form(...),status:str=Form('Review'),reminder_date:str=Form(''),note:str=Form('')):
 with db() as c:c.execute('INSERT INTO tracked VALUES(?,?,?,?,?) ON CONFLICT(kam_id,tender_id) DO UPDATE SET status=excluded.status,reminder_date=excluded.reminder_date,note=excluded.note',(kam_id,tid,status,reminder_date,note))
 return RedirectResponse('/',303)
@app.post('/scan')
async def scan():
 try:await scan_prozorro()
 except:pass
 return RedirectResponse('/',303)
@app.post('/digest')
async def digest():await send_digests();return RedirectResponse('/',303)
@app.get('/export.xlsx')
def export():
 with db() as c:rows=c.execute('SELECT * FROM tenders ORDER BY first_seen DESC').fetchall()
 w=Workbook();s=w.active;s.append(['Market','ID','Title','Buyer','Amount','Currency','Status','Deadline','URL'])
 for r in rows:s.append([r['market_code'],r['external_id'],r['title'],r['buyer'],r['amount'],r['currency'],r['status'],r['deadline'],r['url']])
 b=io.BytesIO();w.save(b);b.seek(0);return StreamingResponse(b,media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',headers={'Content-Disposition':'attachment; filename=cee_tenders.xlsx'})
if __name__=='__main__':import uvicorn;uvicorn.run('app.main:app',host='0.0.0.0',port=8000)
