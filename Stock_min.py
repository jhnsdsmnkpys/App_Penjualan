#Boa:Frame:Frame1

import wx, MySQLdb, datetime, os
from xlwt import *
conn= MySQLdb.connect(host="localhost", user="root", passwd="",db="penjualan")
cur = conn.cursor()

def create(parent):
    return Frame1(parent)

[wxID_FRAME1, wxID_FRAME1BUTTON1, wxID_FRAME1LC, wxID_FRAME1PANEL1, 
 wxID_FRAME1TMBCETAK, 
] = [wx.NewId() for _init_ctrls in range(5)]

class Frame1(wx.Frame):
    def _init_coll_lc_Columns(self, parent):
        # generated method, don't edit

        parent.InsertColumn(col=0, format=wx.LIST_FORMAT_LEFT,
              heading='Kode Barang', width=105)
        parent.InsertColumn(col=1, format=wx.LIST_FORMAT_LEFT,
              heading='Nama Barang', width=160)
        parent.InsertColumn(col=2, format=wx.LIST_FORMAT_LEFT, heading='Stock',
              width=65)
        parent.InsertColumn(col=3, format=wx.LIST_FORMAT_LEFT,
              heading='Stock Min', width=77)
        parent.InsertColumn(col=4, format=wx.LIST_FORMAT_LEFT,
              heading='Harga Beli', width=100)

    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(379, 202), size=wx.Size(658, 485),
              style=wx.DEFAULT_FRAME_STYLE,
              title='Laporan Daftar Barang di Bawah Stock Min')
        self.SetClientSize(wx.Size(642, 447))

        self.panel1 = wx.Panel(id=wxID_FRAME1PANEL1, name='panel1', parent=self,
              pos=wx.Point(0, 0), size=wx.Size(642, 447),
              style=wx.TAB_TRAVERSAL)
        self.panel1.SetBackgroundColour(wx.Colour(192, 192, 192))

        self.lc = wx.ListCtrl(id=wxID_FRAME1LC, name='lc', parent=self.panel1,
              pos=wx.Point(16, 8), size=wx.Size(600, 376), style=wx.LC_REPORT)
        self._init_coll_lc_Columns(self.lc)

        self.button1 = wx.Button(id=wxID_FRAME1BUTTON1, label='button1',
              name='button1', parent=self.panel1, pos=wx.Point(192, 216),
              size=wx.Size(75, 23), style=0)

        self.tmbCetak = wx.Button(id=wxID_FRAME1TMBCETAK, label='Cetak',
              name='tmbCetak', parent=self.panel1, pos=wx.Point(512, 400),
              size=wx.Size(96, 23), style=0)
        self.tmbCetak.Bind(wx.EVT_BUTTON, self.OnTmbCetakButton,
              id=wxID_FRAME1TMBCETAK)

    def __init__(self, parent):
        self._init_ctrls(parent)
        sql = "select * from barang where stock < stock_min "
        cur.execute(sql)
        hasil= cur.fetchall()
        k =self.lc.GetItemCount()
        for i in hasil :
            self.lc.InsertStringItem(k,i[0])
            self.lc.SetStringItem(k,1,i[1])
            self.lc.SetStringItem(k,2,str(i[4]))
            self.lc.SetStringItem(k,3,str(i[5]))
            self.lc.SetStringItem(k,4,str(i[2]))
            k = k + 1

        
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

        
        ## Buat dan setting Style style0
        style0 = XFStyle()
        style0.font = font0
        style0.borders = borders
        style0.pattern = BkgPat
        
        ## Buat dan setting Style style1
        style1 = XFStyle()
        style1.borders = borders
        
        ## Buat dan setting Style style1
        style_judul = XFStyle()
        style_judul.font = font0


        # Hitung Jumlah Baris
        jumbar =self.lc.GetItemCount()
        
        ## Cetak Header
        sheet1.write(0,1,"DAFTAR BARANG DI BAWAH STOCK MINIMUM",style_judul)
        sheet1.write(1,1,"                  TOKO SUKSES HUDAYA        ",style_judul)
        
        i = 3
        # Beri Judul
        sheet1.write(i,0,"Kode Barang",style0)
        sheet1.write(i,1,"Nama Barang",style0)
        sheet1.write(i,2,"Stock",style0)
        sheet1.write(i,3,"Stock Min",style0)
        sheet1.write(i,4,"Harga Beli",style0)
        
        j=0
        while j<=jumbar-1 :
            ## Isikan Item Data Barang
            kd_brg1 = self.lc.GetItem(j,0).GetText()
            sheet1.write(j+i+1,0,kd_brg1,style1)
            nama_brg1 = self.lc.GetItem(j,1).GetText()
            sheet1.write(j+i+1,1,nama_brg1,style1)
            stock1= self.lc.GetItem(j,2).GetText()
            sheet1.write(j+i+1,2,int(stock1),style1)
            stock_min1 = self.lc.GetItem(j,3).GetText()
            sheet1.write(j+i+1,3,int(stock_min1),style1)
            hrg_beli1 = self.lc.GetItem(j,4).GetText()
            sheet1.write(j+i+1,4,int(hrg_beli1),style1)
            j=j+1
        
        # Atur Lebar Kolom
        sheet1.col(0).width = 3500
        sheet1.col(1).width = 5000
        sheet1.col(3).width = 4000
        sheet1.col(4).width = 3500
        path1 = 'C:\\cetak_stock_min.xls' 
        # Cek apakah ada file lama
        if os.path.exists(path1) :
            # Hapus file lama
            os.remove(path1)
        # Simpan file baru 
        book.save(path1)
        # Luncurkan file baru
        os.system("start excel.exe C:\\cetak_stock_min.xls")
    
        