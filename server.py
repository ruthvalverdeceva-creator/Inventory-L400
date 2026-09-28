# Servidor Inventario Picking 400 - lee tu Excel y sincroniza todos los RF en tiempo real
import json,os,re,socket,threading,queue,datetime as dt
from http.server import ThreadingHTTPServer,BaseHTTPRequestHandler
import openpyxl
XLSX=os.environ.get('XLSX','CEVA_Picking_Registros.xlsx');STATE='estado.json';OUT=os.environ.get('OUT','salida.xlsx');JOURNAL='registros.jsonl';PORT=int(os.environ.get('PORT',8080))
U=lambda s:str(s if s is not None else '').strip().upper()
now=lambda:dt.datetime.now().isoformat(timespec='seconds')
lock=threading.Lock();clients=[]
def tbl(wb,n,h=0):
    r=list(wb[n].iter_rows(values_only=True));hd=r[h]
    return [{k:v for k,v in zip(hd,x) if k} for x in r[h+1:] if any(v is not None for v in x)]
def txt(v):
    return str(int(v)) if isinstance(v,float) and v==int(v) else ('' if v is None else str(v).strip())
def load_static():
    wb=openpyxl.load_workbook(XLSX,data_only=True)
    sap={U(r.get('Ubicación')):r for r in tbl(wb,'sap')};mara={txt(r.get('CÓDIGO')):r for r in tbl(wb,'MARA')}
    cant=[]
    for r in tbl(wb,'CANTIDAD'):
        ub=txt(r.get('Ubicación'));s=sap.get(U(ub),{});p=txt(s.get('Producto'));m=mara.get(p,{})
        try:pas='PASILLO %d'%int(ub[3:5])
        except:pas=txt(r.get('PASILLO'))
        cant.append({'PASILLO':pas,'Ubicación':ub,'Producto':p,'EAN':txt(m.get('EAN UMA')),'Descripción producto':txt(s.get('Descripción producto')),
                     'CAMA':txt(m.get('CONVERSIÓN CAMA / PQT')),'UMA':txt(m.get('UMA'))})
    tot={};cr=tbl(wb,'CRUCE',1)
    for r in cr:tot[U(r.get('PASILLO'))]=tot.get(U(r.get('PASILLO')),0)+1
    if not tot:
        for c in cant:tot[U(c['PASILLO'])]=tot.get(U(c['PASILLO']),0)+1
    ops=[{'NOMBRE':txt(r['NOMBRE'])} for r in tbl(wb,'OPERARIOS')]
    return wb,{'operarios':ops,'cantidad':cant},tot
wb0,STATIC,TOT=load_static()
def initial():
    t=[{'Pasillo':txt(r['Pasillo']),'Estado':txt(r.get('Estado')) or 'Pendiente','Operario':txt(r.get('Operario')),
        'Inicio':r['Inicio'].isoformat() if isinstance(r.get('Inicio'),dt.datetime) else '','Fin':r['Fin'].isoformat() if isinstance(r.get('Fin'),dt.datetime) else ''} for r in tbl(wb0,'Tiempos')]
    g=[{'Title':txt(r.get('Fecha')),**{k:r.get(k) for k in ['Operario','Pasillo','Ubicacion','Codigo','Descripcion','UMA','CamaPqt','Niveles','Saldo','Total']}} for r in tbl(wb0,'Registros')]
    return {'tiempos':t,'registros':g}
S=json.load(open(STATE,encoding='utf-8')) if os.path.exists(STATE) else initial()
ids={g.get('_id') for g in S['registros']}
if os.path.exists(JOURNAL):
    for l in open(JOURNAL,encoding='utf-8'):
        try:g=json.loads(l)
        except Exception:continue
        if g.get('_id') not in ids:S['registros'].append(g);ids.add(g.get('_id'))
def view():
    return {'tiempos':[{**t,'Total de Ubicaciones':TOT.get(U(t['Pasillo']),0)} for t in S['tiempos']],'registros':S['registros']}
def export(path):
    wb=openpyxl.load_workbook(XLSX);ws=wb['Tiempos'];h={c.value:i+1 for i,c in enumerate(ws[1])}
    for r in range(2,ws.max_row+1):
        t=next((x for x in S['tiempos'] if U(x['Pasillo'])==U(ws.cell(r,h['Pasillo']).value)),None)
        if t:
            ws.cell(r,h['Estado'],t['Estado']);ws.cell(r,h['Operario'],t['Operario'] or None)
            for k in('Inicio','Fin'):ws.cell(r,h[k],dt.datetime.fromisoformat(t[k]) if t[k] else None)
    wr=wb['Registros'];cols=['Title','Operario','Pasillo','Ubicacion','Codigo','Descripcion','UMA','CamaPqt','Niveles','Saldo','Total']
    for row in wr.iter_rows(min_row=2,max_col=11):
        for c in row:c.value=None
    for i,g in enumerate(S['registros'],2):
        for j,k in enumerate(cols,1):wr.cell(i,j,g.get(k))
    wr.tables['Registros'].ref='A1:M%d'%(max(len(S['registros']),1)+1)
    wb.calculation.fullCalcOnLoad=True;wb.save(path)
timer=[None]
def persist():
    json.dump(S,open(STATE+'.tmp','w',encoding='utf-8'),ensure_ascii=False);os.replace(STATE+'.tmp',STATE)
    if timer[0]:timer[0].cancel()
    def go():
        try:
            tmp=OUT.replace('.xlsx','.tmp.xlsx')
            with lock:export(tmp)
            os.replace(tmp,OUT)
        except Exception as e:
            print('Excel ocupado o ruta inválida (%s). Reintentando en 10s. Los registros ya están a salvo.'%e);timer[0]=threading.Timer(10,go);timer[0].start()
    timer[0]=threading.Timer(5,go);timer[0].start()
def push():
    m=json.dumps(view(),ensure_ascii=False)
    for q in clients:q.put(m)
def act(b):
    with lock:
        p=U(b.get('pasillo'));user=U(b.get('user'));t=next((x for x in S['tiempos'] if U(x['Pasillo'])==p),None)
        if not t:return {'ok':False,'error':'Pasillo no encontrado'}
        k=b['t']
        if k=='iniciar':
            if t['Estado']=='Cerrado':return {'ok':False,'error':'El %s ya está CERRADO.'%t['Pasillo']}
            if t['Estado']=='En Proceso' and U(t['Operario'])!=user:return {'ok':False,'error':'El %s está siendo trabajado por %s.'%(t['Pasillo'],t['Operario'])}
            if t['Estado']!='En Proceso':t.update(Estado='En Proceso',Operario=user,Inicio=now(),Fin='')
        elif k=='guardar':
            if t['Estado']!='En Proceso' or U(t['Operario'])!=user:return {'ok':False,'error':'No se puede guardar: el %s está %s.'%(t['Pasillo'],t['Estado'].upper())}
            ub=U(b.get('Ubicacion'))
            if not ub:return {'ok':False,'error':'Debe escanear la ubicación.'}
            if any(U(g['Pasillo'])==p and U(g['Ubicacion'])==ub for g in S['registros']):return {'ok':False,'error':'⚠️ La ubicación %s ya fue registrada.'%ub}
            rid=b.get('rid')
            if rid and any(g.get('_id')==rid for g in S['registros']):return {'ok':True,'state':view()}
            g={'_id':rid or now()+str(len(S['registros'])),'Title':dt.datetime.now().strftime('%d/%m/%Y %H:%M:%S'),'Operario':user,'Pasillo':t['Pasillo'],**{x:b.get(x) for x in ['Ubicacion','Codigo','Descripcion','UMA','CamaPqt','Niveles','Saldo','Total']}}
            with open(JOURNAL,'a',encoding='utf-8') as f:f.write(json.dumps(g,ensure_ascii=False)+'\n');f.flush();os.fsync(f.fileno())
            S['registros'].append(g)
        elif k=='finalizar':
            if U(t['Operario'])!=user:return {'ok':False,'error':'Solo %s puede finalizar este pasillo.'%t['Operario']}
            t.update(Estado='Cerrado',Fin=now())
        elif k=='admin':
            e=b['estado'];t['Estado']=e
            if e=='Pendiente':t.update(Operario='',Inicio='',Fin='')
            if e=='En Proceso':t['Fin']=''
        persist()
    push();return {'ok':True,'state':view()}
class Srv(ThreadingHTTPServer):
    request_queue_size=256;daemon_threads=True
class H(BaseHTTPRequestHandler):
    def log_message(s,*a):pass
    def send(s,code,body,ct):
        s.send_response(code);s.send_header('Content-Type',ct);s.send_header('Content-Length',str(len(body)));s.end_headers();s.wfile.write(body)
    def do_GET(s):
        if s.path=='/events':
            s.send_response(200);s.send_header('Content-Type','text/event-stream');s.send_header('Cache-Control','no-cache');s.end_headers()
            q=queue.Queue();clients.append(q);q.put(json.dumps(view(),ensure_ascii=False))
            try:
                while True:
                    try:m=q.get(timeout=15);s.wfile.write(('data: %s\n\n'%m).encode())
                    except queue.Empty:s.wfile.write(b': ping\n\n')
                    s.wfile.flush()
            except Exception:pass
            finally:clients.remove(q)
        elif s.path=='/api/init':s.send(200,json.dumps(STATIC,ensure_ascii=False).encode(),'application/json')
        elif s.path.startswith('/export'):
            with lock:export('_tmp.xlsx')
            s.send(200,open('_tmp.xlsx','rb').read(),'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        else:
            f='admin.html' if s.path.startswith('/admin') else 'index.html'
            s.send(200,open(f,'rb').read(),'text/html; charset=utf-8')
    def do_POST(s):
        b=json.loads(s.rfile.read(int(s.headers['Content-Length'])));s.send(200,json.dumps(act(b),ensure_ascii=False).encode(),'application/json')
if __name__=='__main__':
    ip=socket.gethostbyname(socket.gethostname())
    print('Guardando Excel en:',os.path.abspath(OUT));print('RF:      http://%s:%d\nSupervisor: http://%s:%d/admin\nExcel:   %s (se actualiza solo)'%(ip,PORT,ip,PORT,OUT))
    Srv(('0.0.0.0',PORT),H).serve_forever()
