#Boa:Frame:Dabar

import wx, MySQLdb
conn= MySQLdb.connect(host="localhost", user="root", passwd="",db="penjualan")
cur = conn.cursor()


def create(parent):
    return Dabar(parent)

[wxID_DABAR, wxID_DABARLC, wxID_DABARPANEL1, wxID_DABARSTATICTEXT1, 
 wxID_DABARSTATICTEXT2, wxID_DABARSTATICTEXT3, wxID_DABARSTATICTEXT4, 
 wxID_DABARSTATICTEXT5, wxID_DABARSTATICTEXT6, wxID_DABARTMBBERSIH, 
 wxID_DABARTMBHAPUS, wxID_DABARTMBSIMPAN, wxID_DABARTXT_HRGBELI, 
 wxID_DABARTXT_HRGJUAL, wxID_DABARTXT_KDBRG, wxID_DABARTXT_NAMABRG, 
 wxID_DABARTXT_STOCK, wxID_DABARTXT_STOCKMIN, 
] = [wx.NewId() for _init_ctrls in range(18)]

class Dabar(wx.Frame):
    def _init_coll_lc_Columns(self, parent):
        # generated method, don't edit

        parent.InsertColumn(col=0, format=wx.LIST_FORMAT_LEFT,
              heading='Kode Barang', width=-1)
        parent.InsertColumn(col=1, format=wx.LIST_FORMAT_LEFT,
              heading='Nama Barang', width=151)
        parent.InsertColumn(col=2, format=wx.LIST_FORMAT_LEFT, heading='Stock',
              width=-1)
        parent.InsertColumn(col=3, format=wx.LIST_FORMAT_LEFT,
              heading='Harga Jual', width=-1)

    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_DABAR, name='Dabar', parent=prnt,
              pos=wx.Point(397, 184), size=wx.Size(465, 521),
              style=wx.DEFAULT_FRAME_STYLE, title='Input Data Barang')
        self.SetClientSize(wx.Size(457, 494))
        self.SetBackgroundColour(wx.Colour(255, 255, 255))
        self.SetToolTipString('Input Data Barang')

        self.panel1 = wx.Panel(id=wxID_DABARPANEL1, name='panel1', parent=self,
              pos=wx.Point(0, 0), size=wx.Size(457, 494),
              style=wx.TAB_TRAVERSAL)

        self.staticText1 = wx.StaticText(id=wxID_DABARSTATICTEXT1,
              label='Kode Barang', name='staticText1', parent=self.panel1,
              pos=wx.Point(24, 13), size=wx.Size(61, 13), style=0)

        self.txt_kdbrg = wx.TextCtrl(id=wxID_DABARTXT_KDBRG, name='txt_kdbrg',
              parent=self.panel1, pos=wx.Point(128, 13), size=wx.Size(100, 21),
              style=wx.TE_PROCESS_ENTER, value='')
        self.txt_kdbrg.Bind(wx.EVT_TEXT_ENTER, self.OnTxt_kdbrgTextEnter,
              id=wxID_DABARTXT_KDBRG)

        self.staticText2 = wx.StaticText(id=wxID_DABARSTATICTEXT2,
              label='Nama Barang', name='staticText2', parent=self.panel1,
              pos=wx.Point(24, 53), size=wx.Size(64, 13), style=0)

        self.txt_namabrg = wx.TextCtrl(id=wxID_DABARTXT_NAMABRG,
              name='txt_namabrg', parent=self.panel1, pos=wx.Point(128, 53),
              size=wx.Size(152, 21), style=0, value='')

        self.staticText3 = wx.StaticText(id=wxID_DABARSTATICTEXT3,
              label='Harga Beli', name='staticText3', parent=self.panel1,
              pos=wx.Point(24, 96), size=wx.Size(48, 13), style=0)

        self.txt_hrgbeli = wx.TextCtrl(id=wxID_DABARTXT_HRGBELI,
              name='txt_hrgbeli', parent=self.panel1, pos=wx.Point(128, 96),
              size=wx.Size(100, 21), style=0, value='')

        self.staticText4 = wx.StaticText(id=wxID_DABARSTATICTEXT4,
              label='Harga Jual', name='staticText4', parent=self.panel1,
              pos=wx.Point(24, 144), size=wx.Size(51, 13), style=0)

        self.txt_hrgjual = wx.TextCtrl(id=wxID_DABARTXT_HRGJUAL,
              name='txt_hrgjual', parent=self.panel1, pos=wx.Point(128, 144),
              size=wx.Size(100, 21), style=0, value='')

        self.staticText5 = wx.StaticText(id=wxID_DABARSTATICTEXT5,
              label='Stock', name='staticText5', parent=self.panel1,
              pos=wx.Point(24, 192), size=wx.Size(26, 13), style=0)

        self.txt_stock = wx.TextCtrl(id=wxID_DABARTXT_STOCK, name='txt_stock',
              parent=self.panel1, pos=wx.Point(128, 192), size=wx.Size(100, 21),
              style=0, value='')

        self.staticText6 = wx.StaticText(id=wxID_DABARSTATICTEXT6,
              label='Stock Min', name='staticText6', parent=self.panel1,
              pos=wx.Point(24, 232), size=wx.Size(45, 13), style=0)

        self.txt_stockmin = wx.TextCtrl(id=wxID_DABARTXT_STOCKMIN,
              name='txt_stockmin', parent=self.panel1, pos=wx.Point(128, 232),
              size=wx.Size(100, 21), style=0, value='')

        self.tmbSimpan = wx.Button(id=wxID_DABARTMBSIMPAN, label='Simpan',
              name='tmbSimpan', parent=self.panel1, pos=wx.Point(24, 272),
              size=wx.Size(75, 23), style=0)
        self.tmbSimpan.Bind(wx.EVT_BUTTON, self.OnTmbSimpanButton,
              id=wxID_DABARTMBSIMPAN)

        self.tmbHapus = wx.Button(id=wxID_DABARTMBHAPUS, label='Hapus',
              name='tmbHapus', parent=self.panel1, pos=wx.Point(112, 272),
              size=wx.Size(75, 23), style=0)
        self.tmbHapus.Bind(wx.EVT_BUTTON, self.OnTmbHapusButton,
              id=wxID_DABARTMBHAPUS)

        self.tmbBersih = wx.Button(id=wxID_DABARTMBBERSIH, label='Bersih',
              name='tmbBersih', parent=self.panel1, pos=wx.Point(200, 272),
              size=wx.Size(75, 23), style=0)
        self.tmbBersih.Bind(wx.EVT_BUTTON, self.OnTmbBersihButton,
              id=wxID_DABARTMBBERSIH)

        self.lc = wx.ListCtrl(id=wxID_DABARLC, name='lc', parent=self.panel1,
              pos=wx.Point(24, 312), size=wx.Size(410, 160),
              style=wx.LC_REPORT)
        self._init_coll_lc_Columns(self.lc)
        self.lc.Bind(wx.EVT_LIST_ITEM_SELECTED, self.OnLcListItemSelected,
              id=wxID_DABARLC)

    def __init__(self, parent):
        self._init_ctrls(parent)
        self.Awal()
    def OnTxt_kdbrgTextEnter(self, event):
        self.Isi_Object()

    def OnTmbSimpanButton(self, event):
        if self.txt_kdbrg.GetValue()=="" :
            self.pesan = wx.MessageDialog(self,"Kode Barang Belum diisi","Konfirmasi",wx.OK)
            self.pesan.ShowModal()
            self.txt_kdbrg.SetFocus()
            event.Veto()
        if self.txt_namabrg.GetValue()=="" :
            self.pesan = wx.MessageDialog(self,"Nama Barang Belum diisi","Konfirmasi",wx.OK)
            self.pesan.ShowModal()
            self.txt_namabrg.SetFocus()
            event.Veto()
         
        sql = "select * from barang where kd_brg='%s' "%(self.txt_kdbrg.GetValue())
        cur.execute(sql)
        kd_brg1=self.txt_kdbrg.GetValue()
        nama_brg1=self.txt_namabrg.GetValue()
        hrg_beli1=int(self.txt_hrgbeli.GetValue())
        hrg_jual1=int(self.txt_hrgjual.GetValue())
        stock1=int(self.txt_stock.GetValue())
        stock_min1=int(self.txt_stockmin.GetValue())
        if cur.rowcount > 0 :
                         
            sql = "update barang set nama_brg ='%s', hrg_beli='%d', hrg_jual='%d', stock='%d', stock_min ='%d' where kd_brg ='%s' "%(nama_brg1,hrg_beli1,hrg_jual1,stock1,stock_min1,kd_brg1)
        else :
            sql = "insert into barang (kd_brg,nama_brg,hrg_beli,hrg_jual,stock,stock_min) values ('%s','%s','%d','%d','%d','%d') "%(kd_brg1,nama_brg1,hrg_beli1,hrg_jual1,stock1,stock_min1)     
        cur.execute(sql)
        conn.commit()
        self.Awal()

    def Awal(self) :
        self.Isi_List()
        self.txt_kdbrg.SetValue("")
        self.txt_namabrg.SetValue("")
        self.txt_hrgbeli.SetValue("")
        self.txt_hrgjual.SetValue("")
        self.txt_stock.SetValue("")
        self.txt_stockmin.SetValue("")
        self.txt_kdbrg.SetFocus()
    def Isi_Object(self) :
        sql = "select * from barang where kd_brg = '%s' " %(self.txt_kdbrg.GetValue())
        cur.execute(sql)
        if cur.rowcount > 0 :
            hasil = cur.fetchone()
            self.txt_namabrg.SetValue(hasil[1])
            self.txt_hrgbeli.SetValue(str(hasil[2]))
            self.txt_hrgjual.SetValue(str(hasil[3]))
            self.txt_stock.SetValue(str(hasil[4]))
            self.txt_stockmin.SetValue(str(hasil[5]))
        #else :
        #    self.pesan = wx.MessageDialog(self,"Data Tidak Ada","Konfirmasi",wx.OK)
        #    self.pesan.ShowModal()
        self.txt_namabrg.SetFocus()
    def Isi_List(self) :
        self.lc.DeleteAllItems()
        sql = "select * from barang order by kd_brg"
        cur.execute(sql)
        hasil = cur.fetchall()
        jumbar = self.lc.GetItemCount()
        for i in hasil :
            self.lc.InsertStringItem(jumbar,i[0])
            self.lc.SetStringItem(jumbar,1,i[1])
            self.lc.SetStringItem(jumbar,2,str(i[4]))
            self.lc.SetStringItem(jumbar,3,str(i[3]))
            jumbar = jumbar + 1 

    def OnTmbHapusButton(self, event):
        if self.txt_kdbrg.GetValue()=="" :
            self.pesan = wx.MessageDialog(self,"Kode Barang yang Akan Dihapus Belum diisi","Konfirmasi",wx.OK)
            self.pesan.ShowModal()
            self.txt_kdbrg.SetFocus()
            event.Veto()
        tanya = wx.MessageDialog(self,message="Anda yakin handak menghapus barang"+self.txt_namabrg.GetValue()+" ?",style = wx.YES_NO)
        if tanya.ShowModal()==wx.ID_YES:
            sql ="delete from barang where kd_brg ='%s' "%(self.txt_kdbrg.GetValue())
            cur.execute(sql)
            conn.commit()
        self.Awal()    
    def OnTmbBersihButton(self, event):
        self.Awal()

    def OnLcListItemSelected(self, event):
        self.currentItem = event.m_itemIndex
        # mengambil no index baris yang dipilih
        b = self.lc.GetItem(self.currentItem).GetText()
        # no index baris dikonversi ke text/ string
        self.txt_kdbrg.SetValue(b)
        self.Isi_Object()
