# -*- coding: utf-8 -*-
"""ÇKS Askı Listesi - Windows masaüstü uygulaması."""
import os, sys, sqlite3, tempfile, threading, queue, re, unicodedata, datetime, subprocess, webbrowser
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from openpyxl import load_workbook
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

APP_TITLE = 'ÇKS Askı Listesi - Masaüstü V1.6'
COLUMNS = ['İşletme Adı','T.C. / Vergi No','İl','İlçe','Mahalle/Köy','Ada No','Parsel No',
           "ÇKS'ye Beyan Edilen Ürün","ÇKS'ye Beyan Edilen Ürün Grubu",
           "ÇKS'ye Beyan Edilen Ürün Alanı (da)", 'Tespit Edilen Ürün Grubu',
           'Tespit Edilen Ürün Alanı (da)','Beyan/Tespit Durumu','Sezon']
ALIASES = {
 'İşletme Adı':['Ad Soyad','İşletme Adı Soyadı','İşletme Adı'],
 'T.C. / Vergi No':['T.C. No','TC Vergi No','TCKN','Vergi No','T.C. / Vergi No'],
 'İl':['İl','İl Adı','İli'],
 'İlçe':['İlçe'],  # Yalnızca Excel'in İlçe sütunu; ÇKS Dilekçe İlçesi kullanılmaz.
 'Mahalle/Köy':['Mahalle/Köy','Köy/Mahalle','Köy','Köy Adı','Mahalle Köy'],
 'Ada No':['Ada','Ada No','AdaNo'], 'Parsel No':['Parsel','Parsel No','ParselNo'],
 "ÇKS'ye Beyan Edilen Ürün":['Beyan Edilen Ürün',"ÇKS'ye Beyan Edilen Ürün"],
 "ÇKS'ye Beyan Edilen Ürün Grubu":['Beyan Edilen Ürün Grubu',"ÇKS'ye Beyan Edilen Ürün Grubu"],
 "ÇKS'ye Beyan Edilen Ürün Alanı (da)":['Beyan Edilen Ürün Alanı',"ÇKS'ye Beyan Edilen Ürün Alanı (da)"],
 'Tespit Edilen Ürün Grubu':['Tespit Edilen Ürün Grubu'],
 'Tespit Edilen Ürün Alanı (da)':['Tespit Edilen Ürün Alanı (da)','Tespit Edilen Ürün Alanı'],
 'Beyan/Tespit Durumu':['Beyan/Tespit Durumu','Beyan Tespit Durumu','Durum'],
 'Sezon':['Sezon','Üretim Yılı','Yıl']
}
REQUIRED = ['İşletme Adı','İl','İlçe','Mahalle/Köy','Ada No','Parsel No','Beyan/Tespit Durumu']

def norm(s):
    s=str(s or '').replace('\u00a0',' ').strip().upper()
    s=s.replace('İ','I').replace('ı','I')
    s=unicodedata.normalize('NFKD',s)
    s=''.join(c for c in s if not unicodedata.combining(c))
    return re.sub(r'[^A-Z0-9]','',s)

LOOKUP = {norm(alias):canonical for canonical,aliases in ALIASES.items() for alias in aliases}

def find_headers(sheet):
    best = ({},0)
    for idx,row in enumerate(sheet.iter_rows(min_row=1,max_row=min(sheet.max_row or 100,100),values_only=True),1):
        mapping={}
        for j,val in enumerate(row):
            canonical=LOOKUP.get(norm(val))
            if canonical and canonical not in mapping: mapping[canonical]=j
        score=sum(x in mapping for x in REQUIRED)
        if score>best[1]: best=((idx,mapping),score)
        if score==len(REQUIRED): return idx,mapping
    return None,best[0][1] if best[0] else {}

def sval(v):
    if v is None: return ''
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip()

def mask_name(name):
    pieces=str(name).split()
    # Real names and corporate names are both masked in the public list
    return ' '.join((p[:2]+'***') if len(p)>2 else (p[:1]+'***') for p in pieces)

def mask_id(v):
    x=str(v).strip()
    if not x: return ''
    if len(x)<5: return '*'*len(x)
    return x[:2]+'*'*(len(x)-4)+x[-2:]

def area(v):
    if v in ('',None): return '-'
    try: return f'{float(str(v).replace(",",".")):.2f}'.replace('.',',')
    except (ValueError,TypeError): return str(v)

class Application(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE); self.geometry('1160x800'); self.minsize(950,650)
        self.db_path=os.path.join(tempfile.mkdtemp(prefix='cks_aski_'),'records.db')
        self.db=sqlite3.connect(self.db_path); self.setup_db()
        self.events=queue.Queue(); self.importing=False
        self.filename=tk.StringVar(value='Excel dosyası seçilmedi')
        self.year=tk.StringVar(value=str(datetime.date.today().year))
        self.province=tk.StringVar(value='Tümü'); self.district=tk.StringVar(value='Tümü'); self.village=tk.StringVar(value='Tümü')
        self.message=tk.StringVar(value='Excel yükleyerek başlayın. Başlangıçta hiçbir durum işaretli değildir.')
        self.status_vars={}; self.create_ui(); self.after(150,self.poll)
        self.protocol('WM_DELETE_WINDOW',self.on_close)

    def setup_db(self):
        cols=','.join(f'c{i} TEXT' for i in range(len(COLUMNS)))
        self.db.execute(f'CREATE TABLE records (id INTEGER PRIMARY KEY,{cols})')
        self.db.execute('CREATE INDEX idx_location ON records(c2,c3,c4)')
        self.db.execute('CREATE INDEX idx_status ON records(c12)')
        self.db.commit()

    def create_ui(self):
        theme=ttk.Style(self);theme.theme_use('clam')
        theme.configure('Title.TLabel',font=('Segoe UI',17,'bold'))
        theme.configure('Sub.TLabel',font=('Segoe UI',10),foreground='#475569')
        outer=ttk.Frame(self,padding=18);outer.pack(fill='both',expand=True)
        ttk.Label(outer,text='ÇKS BEYAN KONTROL ASKİ LİSTESİ',style='Title.TLabel').pack(anchor='w')
        ttk.Label(outer,text='Excel → il/ilçe/köy ve durum seçimi → A4 yatay PDF / yazdırma',style='Sub.TLabel').pack(anchor='w',pady=(2,14))
        head=ttk.Frame(outer);head.pack(fill='x')
        self.load_btn=ttk.Button(head,text='Excel Dosyası Yükle (.xlsx)',command=self.choose);self.load_btn.pack(side='left')
        ttk.Label(head,textvariable=self.filename).pack(side='left',padx=15)
        ttk.Label(outer,textvariable=self.message,wraplength=1080).pack(anchor='w',pady=12)
        flt=ttk.LabelFrame(outer,text='Rapor filtreleri',padding=12);flt.pack(fill='x',pady=4)
        ttk.Label(flt,text='İl').grid(row=0,column=0,sticky='w')
        self.province_cb=ttk.Combobox(flt,textvariable=self.province,values=['Tümü'],state='readonly',width=22)
        self.province_cb.grid(row=1,column=0,padx=(0,12),sticky='ew');self.province_cb.bind('<<ComboboxSelected>>',self.on_province)
        ttk.Label(flt,text='İlçe').grid(row=0,column=1,sticky='w')
        self.district_cb=ttk.Combobox(flt,textvariable=self.district,values=['Tümü'],state='readonly',width=24)
        self.district_cb.grid(row=1,column=1,padx=(0,12),sticky='ew');self.district_cb.bind('<<ComboboxSelected>>',self.on_district)
        ttk.Label(flt,text='Mahalle/Köy').grid(row=0,column=2,sticky='w')
        self.village_cb=ttk.Combobox(flt,textvariable=self.village,values=['Tümü'],state='readonly',width=28)
        self.village_cb.grid(row=1,column=2,padx=(0,12),sticky='ew');self.village_cb.bind('<<ComboboxSelected>>',lambda e:self.refresh_preview())
        ttk.Label(flt,text='Üretim yılı (rapor başlığı)').grid(row=0,column=3,sticky='w')
        ttk.Entry(flt,textvariable=self.year,width=12).grid(row=1,column=3,sticky='w')
        flt.columnconfigure(0,weight=1);flt.columnconfigure(1,weight=1);flt.columnconfigure(2,weight=1)
        statuses=ttk.LabelFrame(outer,text='Beyan/Tespit Durumu – askıya çıkarılacak durumları işaretleyin',padding=10)
        statuses.pack(fill='x',pady=10)
        buttons=ttk.Frame(statuses);buttons.pack(fill='x')
        ttk.Button(buttons,text='Tümünü Seç',command=lambda:self.toggle_all(True)).pack(side='left')
        ttk.Button(buttons,text='Seçimi Temizle',command=lambda:self.toggle_all(False)).pack(side='left',padx=8)
        self.status_frame=ttk.Frame(statuses);self.status_frame.pack(fill='x',pady=6)
        result=ttk.Frame(outer);result.pack(fill='x',pady=7)
        self.count_label=ttk.Label(result,text='Seçilen kayıt: 0');self.count_label.pack(side='left')
        ttk.Button(result,text='Önizlemeyi Yenile',command=self.refresh_preview).pack(side='right')
        self.print_btn=ttk.Button(result,text='PDF Kaydet',command=self.make_pdf);self.print_btn.pack(side='right',padx=(6,10))
        self.direct_print_btn=ttk.Button(result,text='Doğrudan Yazdır',command=self.direct_print);self.direct_print_btn.pack(side='right',padx=4)
        tbl=ttk.Frame(outer);tbl.pack(fill='both',expand=True)
        names=['Ad Soyad / İşletme','Mahalle/Köy','Ada','Parsel','Beyan Edilen Ürün','Durum']
        self.tree=ttk.Treeview(tbl,columns=names,show='headings',height=13)
        for n in names:self.tree.heading(n,text=n);self.tree.column(n,width=165 if n not in ('Ada','Parsel') else 75)
        self.tree.pack(side='left',fill='both',expand=True)
        scr=ttk.Scrollbar(tbl,orient='vertical',command=self.tree.yview);scr.pack(side='right',fill='y');self.tree.configure(yscrollcommand=scr.set)
        ttk.Label(outer,text='Önizleme ilk 200 kayıtla sınırlıdır. PDF oluşturulurken seçime uyan bütün kayıtlar alınır.',style='Sub.TLabel').pack(anchor='w',pady=(7,0))
        # Sponsorlu tanıtım yalnızca uygulama arayüzünde görünür; PDF rapora eklenmez.
        advert=tk.Frame(outer,bg='#102b48',padx=14,pady=9,cursor='hand2')
        advert.pack(fill='x',pady=(12,0))
        left=tk.Frame(advert,bg='#102b48',cursor='hand2')
        left.pack(side='left',fill='x',expand=True)
        brand=tk.Label(left,text='▶ HaberEngine  |  Profesyonel Haber Yayıncılık Platformu',
                       font=('Segoe UI',10,'bold'),fg='white',bg='#102b48',anchor='w',cursor='hand2')
        brand.pack(anchor='w')
        desc=tk.Label(left,text='Haber botu • Ajans/RSS havuzu • Modern haber sitesi yönetimi',
                      font=('Segoe UI',9),fg='#c7d9ed',bg='#102b48',anchor='w',cursor='hand2')
        desc.pack(anchor='w',pady=(2,0))
        visit=tk.Label(advert,text='Siteyi İncele  ↗',font=('Segoe UI',10,'bold'),
                       fg='#7ee0f0',bg='#102b48',cursor='hand2')
        visit.pack(side='right',padx=(12,0))
        for widget in (advert,left,brand,desc,visit):
            widget.bind('<Button-1>',lambda event:webbrowser.open('https://haberengine.site/'))


    def choose(self):
        path=filedialog.askopenfilename(title='ÇKS veri Excel dosyasını seçin',filetypes=[('Excel dosyası','*.xlsx')])
        if not path:return
        self.filename.set(os.path.basename(path));self.message.set('Excel okunuyor; büyük dosyalarda birkaç dakika sürebilir...')
        self.load_btn.config(state='disabled');self.print_btn.config(state='disabled');self.direct_print_btn.config(state='disabled');self.importing=True
        threading.Thread(target=self.import_excel,args=(path,),daemon=True).start()

    def import_excel(self,path):
        conn=None
        try:
            conn=sqlite3.connect(self.db_path)
            conn.execute('DELETE FROM records');conn.commit()
            wb=load_workbook(path,read_only=True,data_only=True,keep_links=False)
            total=0;skips=[];batch=[]
            for sheet in wb:
                header,mapping=find_headers(sheet)
                if not header or any(x not in mapping for x in REQUIRED):
                    skips.append(sheet.title+': uygun veri başlığı bulunamadı');continue
                for row in sheet.iter_rows(min_row=header+1,values_only=True):
                    vals=[sval(row[mapping[k]]) if k in mapping and mapping[k]<len(row) else '' for k in COLUMNS]
                    if not vals[0] or not vals[4] or not vals[12]:continue
                    batch.append(vals)
                    if len(batch)>=3000:
                        conn.executemany('INSERT INTO records ('+','.join(f'c{i}' for i in range(len(COLUMNS)))+') VALUES ('+','.join('?' for _ in COLUMNS)+')',batch)
                        total+=len(batch);batch=[];conn.commit()
                        if total%30000==0:self.events.put(('progress',f'{total:,} kayıt aktarıldı...'))
            if batch:
                conn.executemany('INSERT INTO records ('+','.join(f'c{i}' for i in range(len(COLUMNS)))+') VALUES ('+','.join('?' for _ in COLUMNS)+')',batch);total+=len(batch);conn.commit()
            wb.close()
            if total==0:raise ValueError('Parsel verisi bulunamadı. Sayfalar: '+'; '.join(skips))
            self.events.put(('done',total))
        except Exception as exc:self.events.put(('error',str(exc)))
        finally:
            if conn:conn.close()

    def poll(self):
        try:
            while True:
                typ,value=self.events.get_nowait()
                if typ=='progress':self.message.set(value)
                else:
                    self.importing=False;self.load_btn.config(state='normal');self.print_btn.config(state='normal');self.direct_print_btn.config(state='normal')
                    if typ=='done':self.message.set(f'{value:,} kayıt okundu. Askıya çıkarılacak durumları seçin.');self.update_filters()
                    if typ=='error':self.message.set('Excel okunamadı: '+value);messagebox.showerror('Excel yükleme hatası',value)
        except queue.Empty:pass
        self.after(200,self.poll)

    def update_filters(self):
        provinces=[x[0] for x in self.db.execute('SELECT DISTINCT c2 FROM records WHERE c2<>"" ORDER BY c2')]
        self.province_cb.config(values=['Tümü']+provinces)
        self.province.set('Tümü');self.on_province()
        for c in self.status_frame.winfo_children():c.destroy()
        self.status_vars={}
        for i,(val,count) in enumerate(self.db.execute('SELECT c12,COUNT(*) FROM records GROUP BY c12 ORDER BY c12')):
            v=tk.BooleanVar(value=False);self.status_vars[val]=v
            ttk.Checkbutton(self.status_frame,text=f'{val} ({count:,})',variable=v,command=self.refresh_preview).grid(row=i//3,column=i%3,sticky='w',padx=16,pady=3)
        self.refresh_preview()

    def on_province(self,event=None):
        province=self.province.get()
        if province=='Tümü':
            rows=self.db.execute('SELECT DISTINCT c3 FROM records WHERE c3<>"" ORDER BY c3')
        else:
            rows=self.db.execute('SELECT DISTINCT c3 FROM records WHERE c2=? AND c3<>"" ORDER BY c3',(province,))
        districts=[x[0] for x in rows]
        self.district_cb.config(values=['Tümü']+districts)
        self.district.set('Tümü');self.on_district()

    def on_district(self,event=None):
        clauses=['c4<>""'];args=[]
        if self.province.get()!='Tümü':clauses.append('c2=?');args.append(self.province.get())
        if self.district.get()!='Tümü':clauses.append('c3=?');args.append(self.district.get())
        villages=[x[0] for x in self.db.execute('SELECT DISTINCT c4 FROM records WHERE '+' AND '.join(clauses)+' ORDER BY c4',args)]
        self.village_cb.config(values=['Tümü']+villages)
        self.village.set('Tümü');self.refresh_preview()

    def toggle_all(self,flag):
        for v in self.status_vars.values():v.set(flag)
        self.refresh_preview()

    def selection(self):
        chosen=[name for name,var in self.status_vars.items() if var.get()]
        if not chosen:return None,None
        where=['c12 IN ('+','.join('?' for _ in chosen)+')'];args=list(chosen)
        if self.province.get()!='Tümü':where.append('c2=?');args.append(self.province.get())
        if self.district.get()!='Tümü':where.append('c3=?');args.append(self.district.get())
        if self.village.get()!='Tümü':where.append('c4=?');args.append(self.village.get())
        return ' AND '.join(where),args

    def refresh_preview(self):
        for x in self.tree.get_children():self.tree.delete(x)
        where,args=self.selection()
        if where is None:self.count_label.config(text='Seçilen kayıt: 0 (durum işaretleyin)');return
        count=self.db.execute('SELECT COUNT(*) FROM records WHERE '+where,args).fetchone()[0]
        self.count_label.config(text=f'Seçilen kayıt: {count:,}')
        for r in self.db.execute('SELECT c0,c4,c5,c6,c7,c12 FROM records WHERE '+where+' ORDER BY c2,c3,c4,c5,c6 LIMIT 200',args):
            self.tree.insert('', 'end',values=r)

    def make_pdf(self):
        where,args=self.selection()
        if where is None:messagebox.showwarning('Durum seçimi','En az bir Beyan/Tespit Durumu seçmelisiniz.');return
        count=self.db.execute('SELECT COUNT(*) FROM records WHERE '+where,args).fetchone()[0]
        if not count:messagebox.showwarning('Kayıt yok','Seçili filtrelere uygun kayıt bulunamadı.');return
        path=filedialog.asksaveasfilename(title='PDF kaydet',defaultextension='.pdf',filetypes=[('PDF','*.pdf')],initialfile='CKS_Aski_Listesi.pdf')
        if not path:return
        try:self.export_pdf(path,where,args,count)
        except Exception as exc:messagebox.showerror('PDF hatası',str(exc));return
        if messagebox.askyesno('PDF oluşturuldu',f'{count:,} kayıt içeren PDF kaydedildi.\n\nŞimdi PDF açılsın mı?'):
            if sys.platform=='win32':os.startfile(path)
            elif sys.platform=='darwin':subprocess.Popen(['open',path])
            else:subprocess.Popen(['xdg-open',path])

    @staticmethod
    def send_pdf_to_windows_printer(path):
        """PDF sayfalarını Windows GDI üzerinden varsayılan yazıcıya aktar.

        Windows'un PDF dosyası için 'print' dosya ilişkilendirmesine dayanmaz.
        PDF çok sayfalıysa sayfalar bellekte tek tek işlenir.
        """
        import fitz  # PyMuPDF
        import win32print
        import win32ui
        import win32con
        from PIL import Image, ImageWin

        printer_name = win32print.GetDefaultPrinter()
        if not printer_name:
            raise RuntimeError('Windows üzerinde varsayılan yazıcı ayarlanmamış.')

        printer_handle = win32print.OpenPrinter(printer_name)
        devmode = None
        try:
            details = win32print.GetPrinter(printer_handle, 2)
            devmode = details.get('pDevMode')
            if devmode is not None:
                devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                devmode.Fields |= win32con.DM_ORIENTATION
        finally:
            win32print.ClosePrinter(printer_handle)

        dc = win32ui.CreateDC()
        doc_started = False
        try:
            # Yazıcı sürücüsüne A4 yatay isteğini ve seçilen çıktı yönünü ilet.
            if devmode is not None:
                dc.CreateDC('WINSPOOL', printer_name, None, devmode)
            else:
                dc.CreatePrinterDC(printer_name)
            printable_w = dc.GetDeviceCaps(win32con.HORZRES)
            printable_h = dc.GetDeviceCaps(win32con.VERTRES)
            if printable_w <= 0 or printable_h <= 0:
                raise RuntimeError('Varsayılan yazıcının baskı ölçüleri alınamadı.')
            # Yazıcıdaki donanımsal marjlar HORZRES/VERTRES ile korunur.
            with fitz.open(path) as pdf:
                dc.StartDoc('ÇKS Askı Listesi')
                doc_started = True
                for page in pdf:
                    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), colorspace=fitz.csRGB, alpha=False)
                    page_image = Image.frombytes('RGB', (pix.width, pix.height), pix.samples)
                    del pix
                    scale = min(printable_w / page_image.width, printable_h / page_image.height)
                    width = max(1, int(page_image.width * scale))
                    height = max(1, int(page_image.height * scale))
                    x = (printable_w - width) // 2
                    y = (printable_h - height) // 2
                    dc.StartPage()
                    try:
                        ImageWin.Dib(page_image).draw(dc.GetHandleOutput(), (x, y, x + width, y + height))
                    finally:
                        dc.EndPage()
                dc.EndDoc()
                doc_started = False
            return printer_name
        except Exception:
            if doc_started:
                try:
                    dc.AbortDoc()
                except Exception:
                    pass
            raise
        finally:
            dc.DeleteDC()

    def direct_print(self):
        """PDF işleyicisine gerek kalmadan Windows yazıcısına doğrudan yazdır."""
        if sys.platform != 'win32':
            messagebox.showerror('Yazdırma', 'Doğrudan yazdırma yalnızca Windows üzerinde desteklenir.')
            return
        where, args = self.selection()
        if where is None:
            messagebox.showwarning('Durum seçimi', 'En az bir Beyan/Tespit Durumu seçmelisiniz.')
            return
        count = self.db.execute('SELECT COUNT(*) FROM records WHERE ' + where, args).fetchone()[0]
        if not count:
            messagebox.showwarning('Kayıt yok', 'Seçili filtrelere uygun kayıt bulunamadı.')
            return
        if not messagebox.askyesno('Doğrudan yazdır',
                f'{count:,} kayıt A4 yatay PDF düzeninde varsayılan yazıcıya gönderilecek.\n\nDevam edilsin mi?'):
            return
        folder = tempfile.mkdtemp(prefix='cks_yazdir_')
        path = os.path.join(folder, 'CKS_Aski_Listesi_Yazdir.pdf')
        try:
            self.export_pdf(path, where, args, count)
            printer = self.send_pdf_to_windows_printer(path)
        except Exception as exc:
            messagebox.showerror('Yazdırma hatası',
                'PDF hazırlandı ancak yazıcıya gönderilemedi.\n\n'
                f'Açıklama: {exc}\n\n'
                'Varsayılan yazıcıyı ve sürücüsünü kontrol edin. '
                'Alternatif olarak PDF Kaydet seçeneğini kullanabilirsiniz.')
            return
        finally:
            # Yazdırma GDI üzerinden senkron tamamlandığından geçici PDF silinebilir.
            try:
                import shutil
                shutil.rmtree(folder)
            except OSError:
                pass
        messagebox.showinfo('Yazdırma',
            f'Yazdırma görevi Windows yazıcı sistemine iletildi.\nYazıcı: {printer}\n'
            'Kağıt çıkışını yazıcının kuyruğundan kontrol edin.')

    def export_pdf(self,path,where,args,count):
        w,h=landscape(A4)
        font='Helvetica';bold='Helvetica-Bold'
        for candidate in ([r'C:\Windows\Fonts\arial.ttf',r'C:\Windows\Fonts\Arial.ttf','/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf']):
            if os.path.exists(candidate):
                pdfmetrics.registerFont(TTFont('CKSBody',candidate));font='CKSBody';break
        for candidate in ([r'C:\Windows\Fonts\arialbd.ttf','/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf']):
            if os.path.exists(candidate):
                pdfmetrics.registerFont(TTFont('CKSBold',candidate));bold='CKSBold';break
        c=canvas.Canvas(path,pagesize=(w,h));c.setTitle('ÇKS Beyan Kontrol Askı Listesi')
        left=16;right=w-16;tablew=right-left
        # PDF örneğinin 11 sütunlu düzeni; genişlikler toplam 100
        weights=[4,15,11,9,4,5,14,13,7,9,8,9,8,13]
        # 14 logical widths? Target actual 12 columns, align precisely:
        weights=[4,15,11,9,4,5,14,13,8,9,8,13]
        widths=[tablew*x/sum(weights) for x in weights]
        labels=['Sıra','Ad Soyad','ÇKS Beyan Sahibi\nT.C. No','Mahalle/Köy','Ada','Parsel',
                "ÇKS'ye Beyan\nEdilen Ürün", "ÇKS'ye Beyan\nEdilen Ürün Grubu",
                "ÇKS'ye Beyan\nEdilen Ürün Alanı (da)", 'Tespit Edilen\nÜrün Grubu',
                'Tespit Edilen\nÜrün Alanı (da)', 'Beyan/Tespit\nDurumu']
        date=datetime.date.today().strftime('%d.%m.%Y')
        year=self.year.get().strip() or '2026'
        province=self.province.get(); district=self.district.get(); village=self.village.get()
        # Page footer ensures rows never overlap notice
        y=0;page=0;row_num=0
        def truncate(text,maxwidth,size=6.15):
            s=str(text)
            if pdfmetrics.stringWidth(s,font,size)<=maxwidth:return s
            while s and pdfmetrics.stringWidth(s+'…',font,size)>maxwidth:s=s[:-1]
            return s+'…'
        def new_page():
            nonlocal y,page
            if page:c.showPage()
            page+=1
            c.setFont(bold,12);c.drawCentredString(w/2,h-30,f'{year} ÜRETİM YILI BEYAN KONTROL ASKI LİSTESİ')
            c.setFont(font,8);c.drawCentredString(w/2,h-45,f'İl: {province} — İlçe: {district} — Mahalle/Köy: {village} | Tarih: {date}')
            y=h-69
            c.setFillColorRGB(.92,.95,.98);c.rect(left,y-38,tablew,38,fill=1,stroke=0)
            c.setFillColorRGB(0,0,0);x=left
            for j,lbl in enumerate(labels):
                c.setFont(bold,6.1)
                parts=lbl.split('\n')
                for k,part in enumerate(parts):
                    c.drawCentredString(x+widths[j]/2,y-18+((len(parts)-1)/2-k)*9,truncate(part,widths[j]-3,6.1))
                x+=widths[j]
            c.setStrokeColorRGB(.67,.72,.77);c.line(left,y-38,right,y-38)
            y-=38
        new_page()
        sql='SELECT '+','.join('c'+str(i) for i in range(len(COLUMNS)))+' FROM records WHERE '+where+' ORDER BY c2,c3,c4,c5,c6,id'
        for r in self.db.execute(sql,args):
            if y-24<47:new_page()
            row_num+=1
            if row_num%2==0:c.setFillColorRGB(.975,.985,.995);c.rect(left,y-24,tablew,24,fill=1,stroke=0)
            c.setFillColorRGB(0,0,0)
            cells=[str(row_num),mask_name(r[0]),mask_id(r[1]),r[4],r[5],r[6],r[7],r[8],area(r[9]),r[10] or '-',area(r[11]),r[12]]
            x=left;c.setFont(font,6.1)
            for j,t in enumerate(cells):
                # middle vertically, clipping long strings to avoid overflow
                s=truncate(t,widths[j]-5)
                c.drawCentredString(x+widths[j]/2,y-14,s) if j in (0,4,5,8,10) else c.drawString(x+2,y-14,s)
                x+=widths[j]
            c.setStrokeColorRGB(.87,.89,.91);c.line(left,y-24,right,y-24)
            y-=24
        # footer on each page via callback isn't available in canvas; footer on last page only
        c.setFont(font,7);c.drawString(left,25,'Listede yer alan parsellerle ilgili itiraz ve düzeltme başvuruları ilgili İl/İlçe Tarım ve Orman Müdürlüğüne yapılır.')
        c.setFont(font,7);c.drawRightString(right,14,f'Sayfa {page} • Toplam {count:,} kayıt')
        c.save()

    def on_close(self):
        if self.importing and not messagebox.askyesno('Çıkış','Excel aktarımı sürüyor. Program kapatılsın mı?'):return
        try:self.db.close()
        except Exception:pass
        self.destroy()

if __name__=='__main__':Application().mainloop()
