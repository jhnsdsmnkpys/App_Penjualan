#Boa:Frame:Frame1

import wx, MySQLdb, datetime, os
from xlwt import *
conn= MySQLdb.connect(host="localhost", user="root", passwd="",db="penjualan")
cur = conn.cursor()

def create(parent):
    return Frame1(parent)

[wxID_FRAME1, wxID_FRAME1LC, wxID_FRAME1PANEL1, wxID_FRAME1STATICTEXT1, 
 wxID_FRAME1STATICTEXT2, wxID_FRAME1STATICTEXT3, wxID_FRAME1TGL_DARI, 
 wxID_FRAME1TGL_SAMPAI, wxID_FRAME1TMBCETAK, wxID_FRAME1TMBPROSES, 
 wxID_FRAME1TXT_TOTAL_SEMUA, 
] = [wx.NewId() for _init_ctrls in range(11)]

class Frame1(wx.Frame):
    def _init_coll_lc_Columns(self, parent):
        # generated method, don't edit

        parent.InsertColumn(col=0, format=wx.LIST_FORMAT_LEFT, heading='Nota',
              width=-1)
        parent.InsertColumn(col=1, format=wx.LIST_FORMAT_LEFT, heading='Total',
              width=-1)

    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(321, 204), size=wx.Size(463, 476),
              style=wx.DEFAULT_FRAME_STYLE, title='Laporan Penjualan')
        self.SetClientSize(wx.Size(447, 438))
        self.SetBackgroundColour(wx.Colour(255, 255, 255))

        self.panel1 = wx.Panel(id=wxID_FRAME1PANEL1, name='panel1', parent=self,
              pos=wx.Point(0, 0), size=wx.Size(447, 438),
              style=wx.TAB_TRAVERSAL)

        self.tgl_dari = wx.DatePickerCtrl(id=wxID_FRAME1TGL_DARI,
              name='tgl_dari', parent=self.panel1, pos=wx.Point(32, 64),
              size=wx.Size(96, 21), style=wx.DP_SHOWCENTURY)

        self.staticText1 = wx.StaticText(id=wxID_FRAME1STATICTEXT1,
              label='Dari', name='staticText1', parent=self.panel1,
              pos=wx.Point(40, 40), size=wx.Size(19, 13), style=0)

        self.staticText2 = wx.StaticText(id=wxID_FRAME1STATICTEXT2,
              label='Sampai', name='staticText2', parent=self.panel1,
              pos=wx.Point(312, 40), size=wx.Size(34, 13), style=0)

        self.tgl_sampai = wx.DatePickerCtrl(id=wxID_FRAME1TGL_SAMPAI,
              name='tgl_sampai', parent=self.panel1, pos=wx.Point(312, 64),
              size=wx.Size(96, 21), style=wx.DP_SHOWCENTURY)

        self.tmbProses = wx.Button(id=wxID_FRAME1TMBPROSES, label='Proses',
              name='tmbProses', parent=self.panel1, pos=wx.Point(336, 120),
              size=wx.Size(75, 23), style=0)
        self.tmbProses.Bind(wx.EVT_BUTTON, self.OnTmbProsesButton,
              id=wxID_FRAME1TMBPROSES)

        self.lc = wx.ListCtrl(id=wxID_FRAME1LC, name='lc', parent=self.panel1,
              pos=wx.Point(32, 160), size=wx.Size(376, 224),
              style=wx.LC_REPORT)
        self._init_coll_lc_Columns(self.lc)

        self.staticText3 = wx.StaticText(id=wxID_FRAME1STATICTEXT3,
              label='Total Keseluruhan', name='staticText3', parent=self.panel1,
              pos=wx.Point(32, 400), size=wx.Size(87, 13), style=0)

        self.txt_total_semua = wx.TextCtrl(id=wxID_FRAME1TXT_TOTAL_SEMUA,
              name='txt_total_semua', parent=self.panel1, pos=wx.Point(136,
              400), size=wx.Size(100, 21), style=0, value='')

        self.tmbCetak = wx.Button(id=wxID_FRAME1TMBCETAK, label='Cetak',
              name='tmbCetak', parent=self.panel1, pos=wx.Point(328, 400),
              size=wx.Size(75, 23), style=0)
        self.tmbCetak.Bind(wx.EVT_BUTTON, self.OnTmbCetakButton,
              id=wxID_FRAME1TMBCETAK)

    def __init__(self, parent):
        
        self._init_ctrls(parent)
        skrg = datetime.date.today()
        day = skrg.day
        month = skrg.month
        year = skrg.year
        displayed = wx.DateTimeFromDMY(day,month-1,year)
        displayed.Format("%d/%m/%Y")
        self.tgl_dari.SetValue(displayed)
        #self.tgl_sampai.SetValue(displayed)
    def OnTmbProsesButton(self, event):
        selected = self.tgl_dari.GetValue()
        month = selected.Month + 1
        day = selected.Day
        year = selected.Year
        dari1=datetime.date(year,month,day)
        selected = self.tgl_sampai.GetValue()
        month = selected.Month + 1
        day = selected.Day
        year = selected.Year
        sampai1=datetime.date(year,month,day)
        self.lc.DeleteAllItems()
        sql = "select * from jual where tgl <='%s' and tgl>='%s' "%(sampai1,dari1)
        
        cur.execute(sql)
        if cur.rowcount > 0 :
            hasil = cur.fetchall()
            jumbar = self.lc.GetItemCount()  
            total =0
            for i in hasil :
                self.lc.InsertStringItem(jumbar,str(i[0]))
                self.lc.SetStringItem(jumbar,1,str(i[2]))
                total=total+i[2]
                
                jumbar = jumbar + 1 
            self.txt_total_semua.SetValue(str(total))

    def OnTmbCetakButton(self, event):
        #Buat Workbook book
        book = Workbook()
        #Buat Worksheet sheet1
        sheet1 = book.add_sheet('Sheet1')
       
        ## Setting Font
        font0 = Font()
        font0.name = 'Arial'
        font0.height=200
        font0.bold = True

        
        ## Setting Border
        borders = Borders()
        borders.left = 1
        borders.right = 1
        borders.top = 1
        borders.bottom = 1
        
        ## Setting Pattern
        BkgPat = Pattern()
        BkgPat.pattern = Pattern.SOLID_PATTERN
        BkgPat.pattern_fore_colour = 22

        
        ## Buat dan setting Style style0 (Judul Kolom dan Rekap Total)
        style0 = XFStyle()
        style0.font = font0
        style0.borders = borders
        style0.pattern = BkgPat
        
        ## Buat dan setting Style style1
        style1 = XFStyle()
        style1.borders = borders
        
        ## Buat dan setting Style style_judul
        style_judul = XFStyle()
        style_judul.font = font0
        
        ## Format Tanggal
        fmt ='DD-MM-YYYY'
        style_tgl = XFStyle()
        style_tgl.num_format_str = fmt
        
        ## Format Nominal isi kolom
        fmt ='Rp #,##0.00'
        style_num1 = XFStyle()
        style_num1.borders = borders
        style_num1.num_format_str = fmt
        ## Format Nominal total omzet
        fmt ='Rp #,##0.00'
        style_num2 = XFStyle()
        style_num2.borders = borders
        style_num2.num_format_str = fmt
        style_num2.pattern = BkgPat
        style_num2.font = font0

        # Hitung Jumlah Baris
        jumbar =self.lc.GetItemCount()
        
        ## Ambil Data Tanggal dr wx.DatePickerCtrl
        selected = self.tgl_dari.GetValue()
        month = selected.Month + 1
        day = selected.Day
        year = selected.Year
        dari1=datetime.date(year,month,day)
        selected = self.tgl_sampai.GetValue()
        month = selected.Month + 1
        day = selected.Day
        year = selected.Year
        sampai1=datetime.date(year,month,day)
        
        ## Cetak Header
        sheet1.write(0,0,"  LAPORAN OMZET PENJUALAN",style_judul)
        sheet1.write(1,0,"      TOKO SUKSES HUDAYA ",style_judul)
        sheet1.write(2,0,"Dari Tgl", style_judul)
        sheet1.write(2,1,dari1,style_tgl)
        sheet1.write(3,0,"Sampai Tgl", style_judul)
        sheet1.write(3,1,sampai1,style_tgl)
        i = 4
        # Beri Judul Kolom
        sheet1.write(i,0,"Nota",style0)
        sheet1.write(i,1,"Total Harga",style0)
        
        j=0
        while j<=jumbar-1 :
            ## Isikan Item Data Barang
            nota1 = self.lc.GetItem(j,0).GetText()
            sheet1.write(j+i+1,0,nota1,style1)
            total1 = self.lc.GetItem(j,1).GetText()
            sheet1.write(j+i+1,1,int(total1),style_num1)
            j=j+1
        k=j+i+1
        sheet1.write(k,0,"Total Omzet",style0)
        sheet1.write(k,1,int(self.txt_total_semua.GetValue()),style_num2)
        
        # Atur Lebar Kolom
        sheet1.col(0).width = 3500
        sheet1.col(1).width = 5000
        
        path1 = 'C:\\cetak_lap_omzet.xls' 
        # Cek apakah ada file lama
        if os.path.exists(path1) :
            # Hapus file lama
            os.remove(path1)
        # Simpan file baru 
        book.save(path1)
        # Luncurkan file baru
        os.system("start excel.exe C:\\cetak_lap_omzet.xls")
