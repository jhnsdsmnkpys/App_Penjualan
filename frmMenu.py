#Boa:Frame:frmMenu

import wx, Dabar, Jual, Stock_min, LapPenjualan

def create(parent):
    return frmMenu(parent)

[wxID_FRMMENU] = [wx.NewId() for _init_ctrls in range(1)]

class frmMenu(wx.Frame):
    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRMMENU, name='frmMenu', parent=prnt,
              pos=wx.Point(-4, -4), size=wx.Size(1032, 748),
              style=wx.DEFAULT_FRAME_STYLE, title='Menu Utama Penjualan')
        self.SetClientSize(wx.Size(1024, 721))
        self.SetBackgroundColour(wx.Colour(255, 255, 255))

    def __init__(self, parent):
        self._init_ctrls(parent)
        menuMaster = wx.Menu()
        menuMaster.Append(1, "&Data Barang...")
        menuMaster.AppendSeparator()
        menuMaster.Append(2, "E&xit")
        menuTrans = wx.Menu()
        menuTrans.Append(3, "&Penjualan...")
        menuLap = wx.Menu()
        menuLap.Append(4, "&Laporan di Bawah Stock Min")
        menuLap.Append(5, "&Laporan Omzet")
        menuBar = wx.MenuBar()
        menuBar.Append(menuMaster, "&Master")
        menuBar.Append(menuTrans, "&Transaksi")
        menuBar.Append(menuLap, "&Laporan")
        self.SetMenuBar(menuBar)
        self.CreateStatusBar()
        self.SetStatusText("Selamat Datang di Python!")
        self.Bind(wx.EVT_MENU, self.OnDabar, id=1)
        self.Bind(wx.EVT_MENU, self.OnQuit, id=2)
        self.Bind(wx.EVT_MENU, self.OnJual, id=3)
        self.Bind(wx.EVT_MENU, self.OnLapStock_min, id=4)
        self.Bind(wx.EVT_MENU, self.OnLapPenjualan, id=5)
    def OnQuit(self, event):
        self.Close()
         
    def OnDabar(self, event):
        self.main = Dabar.create(None)
        self.main.Show()

    def OnJual(self, event):
        self.main = Jual.create(None)
        self.main.Show()
    def OnLapStock_min(self,event) :
        self.main = Stock_min.create(None)
        self.main.Show()    
    def OnLapPenjualan(self,event) :
        self.main = LapPenjualan.create(None)
        self.main.Show()
           
