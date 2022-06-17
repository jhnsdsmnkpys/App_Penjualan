#Boa:Frame:Frame1

import wx, MySQLdb, datetime, os
from xlwt import *
conn= MySQLdb.connect(host="localhost", user="root", passwd="",db="penjualan")
cur = conn.cursor()

def create(parent):
    return Frame1(parent)

[wxID_FRAME1, wxID_FRAME1LC, wxID_FRAME1PANEL1, wxID_FRAME1STATICTEXT1, 
 wxID_FRAME1STATICTEXT10, wxID_FRAME1STATICTEXT11, wxID_FRAME1STATICTEXT2, 
 wxID_FRAME1STATICTEXT3, wxID_FRAME1STATICTEXT4, wxID_FRAME1STATICTEXT5, 
 wxID_FRAME1STATICTEXT6, wxID_FRAME1STATICTEXT7, wxID_FRAME1STATICTEXT8, 
 wxID_FRAME1STATICTEXT9, wxID_FRAME1TMBSIMPAN, wxID_FRAME1TMBTAMBAH, 
 wxID_FRAME1TXT_BAYAR, wxID_FRAME1TXT_HRG, wxID_FRAME1TXT_JML, 
 wxID_FRAME1TXT_KDBRG, wxID_FRAME1TXT_KEMBALI, wxID_FRAME1TXT_NAMABRG, 
 wxID_FRAME1TXT_NOTA, wxID_FRAME1TXT_TGL, wxID_FRAME1TXT_TOTAL, 
 wxID_FRAME1TXT_TOTAL_SEMUA, 
] = [wx.NewId() for _init_ctrls in range(26)]

class Frame1(wx.Frame):
    def _init_coll_lc_Columns(self, parent):
        # generated method, don't edit

        parent.InsertColumn(col=0, format=wx.LIST_FORMAT_LEFT,
              heading='Kode Barang', width=-1)
        parent.InsertColumn(col=1, format=wx.LIST_FORMAT_LEFT,
              heading='Nama Barang', width=135)
        parent.InsertColumn(col=2, format=wx.LIST_FORMAT_LEFT, heading='Jumlah',
              width=-1)
        parent.InsertColumn(col=3, format=wx.LIST_FORMAT_LEFT, heading='Harga',
              width=-1)
        parent.InsertColumn(col=4, format=wx.LIST_FORMAT_LEFT, heading='Total',
              width=-1)

    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(359, 228), size=wx.Size(587, 500),
              style=wx.DEFAULT_FRAME_STYLE, title='Transaksi Penjualan')
        self.SetClientSize(wx.Size(571, 462))
        self.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.panel1 = wx.Panel(id=wxID_FRAME1PANEL1, name='panel1', parent=self,
              pos=wx.Point(0, 0), size=wx.Size(571, 462),
              style=wx.TAB_TRAVERSAL)

        self.staticText1 = wx.StaticText(id=wxID_FRAME1STATICTEXT1,
              label='No. Nota', name='staticText1', parent=self.panel1,
              pos=wx.Point(8, 16), size=wx.Size(43, 13), style=0)

        self.txt_nota = wx.TextCtrl(id=wxID_FRAME1TXT_NOTA, name='txt_nota',
              parent=self.panel1, pos=wx.Point(8, 40), size=wx.Size(100, 21),
              style=0, value='')
        self.txt_nota.SetEditable(False)

        self.staticText2 = wx.StaticText(id=wxID_FRAME1STATICTEXT2,
              label='Tanggal', name='staticText2', parent=self.panel1,
              pos=wx.Point(480, 17), size=wx.Size(38, 13), style=0)

        self.txt_tgl = wx.TextCtrl(id=wxID_FRAME1TXT_TGL, name='txt_tgl',
              parent=self.panel1, pos=wx.Point(464, 35), size=wx.Size(100, 21),
              style=0, value='')
        self.txt_tgl.SetEditable(False)

        self.staticText3 = wx.StaticText(id=wxID_FRAME1STATICTEXT3,
              label='Kode Barang', name='staticText3', parent=self.panel1,
              pos=wx.Point(16, 96), size=wx.Size(61, 13), style=0)

        self.txt_kdbrg = wx.TextCtrl(id=wxID_FRAME1TXT_KDBRG, name='txt_kdbrg',
              parent=self.panel1, pos=wx.Point(16, 120), size=wx.Size(100, 21),
              style=wx.TE_PROCESS_ENTER, value='')
        self.txt_kdbrg.Bind(wx.EVT_TEXT_ENTER, self.OnTxt_kdbrgTextEnter,
              id=wxID_FRAME1TXT_KDBRG)

        self.staticText4 = wx.StaticText(id=wxID_FRAME1STATICTEXT4,
              label='Nama Barang', name='staticText4', parent=self.panel1,
              pos=wx.Point(136, 96), size=wx.Size(64, 13), style=0)

        self.txt_namabrg = wx.TextCtrl(id=wxID_FRAME1TXT_NAMABRG,
              name='txt_namabrg', parent=self.panel1, pos=wx.Point(128, 120),
              size=wx.Size(104, 21), style=0, value='')
        self.txt_namabrg.SetEditable(False)

        self.staticText5 = wx.StaticText(id=wxID_FRAME1STATICTEXT5,
              label='Jumlah', name='staticText5', parent=self.panel1,
              pos=wx.Point(256, 96), size=wx.Size(33, 13), style=0)

        self.txt_jml = wx.TextCtrl(id=wxID_FRAME1TXT_JML, name='txt_jml',
              parent=self.panel1, pos=wx.Point(240, 120), size=wx.Size(64, 21),
              style=wx.TE_PROCESS_ENTER, value='')
        self.txt_jml.Bind(wx.EVT_TEXT_ENTER, self.OnTxt_jmlTextEnter,
              id=wxID_FRAME1TXT_JML)

        self.staticText6 = wx.StaticText(id=wxID_FRAME1STATICTEXT6,
              label='Harga', name='staticText6', parent=self.panel1,
              pos=wx.Point(328, 96), size=wx.Size(29, 13), style=0)

        self.txt_hrg = wx.TextCtrl(id=wxID_FRAME1TXT_HRG, name='txt_hrg',
              parent=self.panel1, pos=wx.Point(312, 120), size=wx.Size(72, 21),
              style=0, value='')
        self.txt_hrg.SetEditable(False)

        self.staticText7 = wx.StaticText(id=wxID_FRAME1STATICTEXT7,
              label='Total', name='staticText7', parent=self.panel1,
              pos=wx.Point(400, 96), size=wx.Size(24, 13), style=0)

        self.txt_total = wx.TextCtrl(id=wxID_FRAME1TXT_TOTAL, name='txt_total',
              parent=self.panel1, pos=wx.Point(392, 120), size=wx.Size(80, 21),
              style=0, value='')
        self.txt_total.SetEditable(False)

        self.tmbTambah = wx.Button(id=wxID_FRAME1TMBTAMBAH, label='Tambah',
              name='tmbTambah', parent=self.panel1, pos=wx.Point(480, 120),
              size=wx.Size(75, 23), style=0)
        self.tmbTambah.Bind(wx.EVT_BUTTON, self.OnTmbTambahButton,
              id=wxID_FRAME1TMBTAMBAH)

        self.lc = wx.ListCtrl(id=wxID_FRAME1LC, name='lc', parent=self.panel1,
              pos=wx.Point(16, 152), size=wx.Size(536, 152),
              style=wx.LC_REPORT)
        self._init_coll_lc_Columns(self.lc)
        self.lc.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnLcListItemSelected,
              id=wxID_FRAME1LC)

        self.staticText8 = wx.StaticText(id=wxID_FRAME1STATICTEXT8,
              label='Total', name='staticText8', parent=self.panel1,
              pos=wx.Point(40, 312), size=wx.Size(47, 25), style=0)
        self.staticText8.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, 'Tahoma'))

        self.txt_total_semua = wx.TextCtrl(id=wxID_FRAME1TXT_TOTAL_SEMUA,
              name='txt_total_semua', parent=self.panel1, pos=wx.Point(16, 336),
              size=wx.Size(176, 64), style=0, value='')
        self.txt_total_semua.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.BOLD,
              False, 'Tahoma'))
        self.txt_total_semua.SetEditable(False)

        self.staticText9 = wx.StaticText(id=wxID_FRAME1STATICTEXT9,
              label='Total', name='staticText9', parent=self.panel1,
              pos=wx.Point(40, 312), size=wx.Size(47, 25), style=0)
        self.staticText9.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, 'Tahoma'))

        self.staticText10 = wx.StaticText(id=wxID_FRAME1STATICTEXT10,
              label='Bayar', name='staticText10', parent=self.panel1,
              pos=wx.Point(240, 312), size=wx.Size(54, 23), style=0)
        self.staticText10.SetFont(wx.Font(14, wx.SWISS, wx.NORMAL, wx.BOLD,
              False, 'Tahoma'))

        self.txt_bayar = wx.TextCtrl(id=wxID_FRAME1TXT_BAYAR, name='txt_bayar',
              parent=self.panel1, pos=wx.Point(200, 336), size=wx.Size(168, 64),
              style=wx.TE_PROCESS_ENTER, value='')
        self.txt_bayar.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, 'Tahoma'))
        self.txt_bayar.Bind(wx.EVT_TEXT_ENTER, self.OnTxt_bayarTextEnter,
              id=wxID_FRAME1TXT_BAYAR)

        self.staticText11 = wx.StaticText(id=wxID_FRAME1STATICTEXT11,
              label='Kembali', name='staticText11', parent=self.panel1,
              pos=wx.Point(424, 312), size=wx.Size(77, 23), style=0)
        self.staticText11.SetFont(wx.Font(14, wx.SWISS, wx.NORMAL, wx.BOLD,
              False, 'Tahoma'))

        self.txt_kembali = wx.TextCtrl(id=wxID_FRAME1TXT_KEMBALI,
              name='txt_kembali', parent=self.panel1, pos=wx.Point(392, 336),
              size=wx.Size(160, 64), style=0, value='')
        self.txt_kembali.SetFont(wx.Font(16, wx.SWISS, wx.NORMAL, wx.NORMAL,
              False, 'Tahoma'))

        self.tmbSimpan = wx.Button(id=wxID_FRAME1TMBSIMPAN, label='Simpan',
              name='tmbSimpan', parent=self.panel1, pos=wx.Point(16, 416),
              size=wx.Size(88, 24), style=0)
        self.tmbSimpan.Bind(wx.EVT_BUTTON, self.OnTmbSimpanButton,
              id=wxID_FRAME1TMBSIMPAN)

    def __init__(self, parent):
        self._init_ctrls(parent)
        self.Awal()
    def Awal (self) :
        # Isi No Nota
        default_value = 0
        sql ="select COALESCE(max(nota), 0) from jual"

        #sql = "select max(nota) as nt from jual "
        cur.execute(sql)
        if cur.rowcount > 0 :
            hasil = cur.fetchone()
            jml = hasil[0] + 1
            self.txt_nota.SetValue(str(jml))
        else :
            self.txt_nota.SetValue("1")
            
        # Isi Tanggal Saat ini
        skrg = datetime.date.today()
        day = skrg.day
        month = skrg.month
        year = skrg.year
        self.txt_tgl.SetValue("%02d/%02d/%4d" % (day, month, year))
        # Bersihkan seluruh text di atas ListCtrl
        self.txt_kdbrg.SetValue("")
        self.txt_namabrg.SetValue("")
        self.txt_hrg.SetValue("")
        self.txt_jml.SetValue("")
        self.txt_total.SetValue("")
        self.txt_kdbrg.SetFocus()
        # Bersihkan ListCtrl lc  
        self.lc.DeleteAllItems()
        # Bersihkan textctrl di bawah ListCtrl
        self.txt_total_semua.SetValue("0")
        self.txt_bayar.SetValue("0")
        self.txt_kembali.SetValue("0")

    def OnTmbTambahButton(self, event):
        jumbar = self.lc.GetItemCount()
        self.lc.InsertStringItem(jumbar,self.txt_kdbrg.GetValue())
        self.lc.SetStringItem(jumbar,1,self.txt_namabrg.GetValue())
        self.lc.SetStringItem(jumbar,2,self.txt_jml.GetValue())
        self.lc.SetStringItem(jumbar,3,self.txt_hrg.GetValue())
        self.lc.SetStringItem(jumbar,4,self.txt_total.GetValue())
        # Tambahkan Total
        ta = int(self.txt_total.GetValue())
        tb = int(self.txt_total_semua.GetValue())
        tb = tb + ta
        self.txt_total_semua.SetValue(str(tb))
        

        # Bersihkan seluruh text di atas ListCtrl
        self.txt_kdbrg.SetValue("")
        self.txt_namabrg.SetValue("")
        self.txt_jml.SetValue("")
        self.txt_hrg.SetValue("")
        self.txt_total.SetValue("")
        self.txt_kdbrg.SetFocus()
 
          
    def OnTmbSimpanButton(self, event):
        # Simpan ke Tabel Jual
        nota1= int(self.txt_nota.GetValue())
        skrg = datetime.date.today()
        day = skrg.day
        month = skrg.month
        year = skrg.year
        tgl1 = datetime.date(year,month,day)
        total1 = int(self.txt_total_semua.GetValue())
        bayar1 = int(self.txt_bayar.GetValue())
        kembali1 = int(self.txt_kembali.GetValue())
                     
        sql = "insert into jual (nota,tgl,total,bayar,kembali) values \
        ('%d','%s','%d','%d','%d') "%(nota1,tgl1,total1,bayar1,kembali1)
        cur.execute(sql)
        conn.commit()
        # Simpan ke Tabel Detail_Jual
        i =0
        jumbar = self.lc.GetItemCount()
        while i <= jumbar-1 :
            kd_brg1 = self.lc.GetItem(i,0).GetText()
            jml1 = int(self.lc.GetItem(i,2).GetText())
            hrg_jual1 = int(self.lc.GetItem(i,3).GetText())
            total1 = int(self.lc.GetItem(i,4).GetText())
            sql = "insert into detail_jual (nota,kd_brg,jml,hrg_jual,total)\
            values ('%d','%s','%d','%d','%d') " % \
            (nota1,kd_brg1,jml1,hrg_jual1,total1)
            cur.execute(sql)
            conn.commit()
            i = i+1
        # Update stock barang
        i =0
        jumbar = self.lc.GetItemCount()
        while i <= jumbar-1 :
            kd_brg1 = self.lc.GetItem(i,0).GetText()
            jml1 = int(self.lc.GetItem(i,2).GetText())
            sql = "select stock from barang where kd_brg ='%s'"%(kd_brg1)
            cur.execute(sql)
            if cur.rowcount > 0 :
                hasil = cur.fetchone()
                st1 = hasil[0] - jml1
                sql = "update barang set stock ='%d' where \
                kd_brg = '%s'"%(st1,kd_brg1)
                cur.execute(sql)
                conn.commit()
            i = i +1
        tanya = wx.MessageDialog(self,message="Apakah Anda Hendak \
        Mencetak Nota"+self.txt_namabrg.GetValue()+" ?",style = wx.YES_NO)
        if tanya.ShowModal()==wx.ID_YES:
            self.Cetak()
        self.Awal()
    def Cetak(self) :
        #Buat Workbook book
        book = Workbook()
        #Buat Worksheet sheet1
        sheet1 = book.add_sheet('Sheet1')
        
        ## Setting Border
        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        ## Buat dan setting Style style0
        style0 = XFStyle()
        style0.borders = borders
        
        # Hitung Jumlah Baris
        jumbar =self.lc.GetItemCount()
        
        ## Cetak Header
        sheet1.write(0,1,"STRUK BELANJA TOKO SUKSES HUDAYA")
        sheet1.write(1,0,"No. Nota")
        sheet1.write(1,1,self.txt_nota.GetValue())
        sheet1.write(2,0,"Tanggal")
        sheet1.write(2,1,self.txt_tgl.GetValue())
        
        i = 3
        # Beri Judul
        sheet1.write(i,0,"Kode Barang",style0)
        sheet1.write(i,1,"Nama Barang",style0)
        sheet1.write(i,2,"Jumlah",style0)
        sheet1.write(i,3,"Harga",style0)
        sheet1.write(i,4,"Total",style0)
        j=0
        while j<=jumbar-1 :
            ## Isikan Item Data Barang
            kd_brg1 = self.lc.GetItem(j,0).GetText()
            sheet1.write(j+i+1,0,kd_brg1,style0)
            nama_brg1 = self.lc.GetItem(j,1).GetText()
            sheet1.write(j+i+1,1,nama_brg1,style0)
            hrg1= self.lc.GetItem(j,2).GetText()
            sheet1.write(j+i+1,2,int(hrg1),style0)
            jml1 = self.lc.GetItem(j,3).GetText()
            sheet1.write(j+i+1,3,int(jml1),style0)
            total1 = self.lc.GetItem(j,4).GetText()
            sheet1.write(j+i+1,4,int(total1),style0)
            j=j+1
        k = j+i+2
        sheet1.write(k,3,"Total Semua ",style0)
        sheet1.write(k,4,int(self.txt_total_semua.GetValue()),style0)
        sheet1.write(k+1,3,"Bayar ",style0)
        sheet1.write(k+1,4,int(self.txt_bayar.GetValue()),style0)
        sheet1.write(k+2,3,"Kembali ",style0)
        sheet1.write(k+2,4,int(self.txt_kembali.GetValue()),style0)
        sheet1.write(k+4,1,"TERIMA KASIH ATAS KUNJUNGAN ANDA")
        
        # Atur Lebar Kolom
        sheet1.col(0).width = 3500
        sheet1.col(1).width = 5000
        sheet1.col(3).width = 4000
        
        path1 = 'C:\\cetak.xls' 
        # Cek apakah ada file lama
        if os.path.exists(path1) :
            # Hapus file lama
            os.remove(path1)
        # Simpan file baru 
        book.save(path1)
        # Luncurkan file baru
        os.system("start excel.exe C:\\cetak.xls")
    def OnTxt_kdbrgTextEnter(self, event):
        sql = "select * from barang where kd_brg = '%s' "\
         %(self.txt_kdbrg.GetValue())
        cur.execute(sql)
        if cur.rowcount > 0 :
            hasil = cur.fetchone()
            self.txt_namabrg.SetValue(hasil[1])
            self.txt_hrg.SetValue(str(hasil[3]))
            self.txt_jml.SetFocus()

    def OnTxt_jmlTextEnter(self, event):
        hrg = int(self.txt_hrg.GetValue())
        jml = int(self.txt_jml.GetValue())
        tt = hrg * jml
        self.txt_total.SetValue(str(tt))
        self.tmbTambah.SetFocus()

    def OnTxt_bayarTextEnter(self, event):
        tt = int(self.txt_total_semua.GetValue())
        byr = int(self.txt_bayar.GetValue())
        kmb = byr - tt
        self.txt_kembali.SetValue(str(kmb))

    def OnLcListItemSelected(self, event):
        self.currentItem = event.m_itemIndex
        # mengambil no index baris yang dipilih
        b = self.lc.GetItem(self.currentItem).GetText()
        # no index baris dikonversi ke text/ string
        self.txt_kdbrg.SetValue(b)
        self.lc.DeleteItem(self.currentItem)
        

    
        
            
