#Boa:Frame:Frame1

import wx, xlrd, KoneksiDB
import MySQLdb

from time import sleep
import sys

conn = KoneksiDB.conn
cur = KoneksiDB.cur

def create(parent):
    return Frame1(parent)

[wxID_FRAME1, wxID_FRAME1BARIS_PROSES, wxID_FRAME1BUTTON2, 
 wxID_FRAME1JML_BARIS, wxID_FRAME1LOKASI_EXCEL, wxID_FRAME1PANEL1, 
 wxID_FRAME1STATICTEXT1, wxID_FRAME1STATICTEXT2, wxID_FRAME1STATICTEXT3, 
 wxID_FRAME1STATICTEXT4, wxID_FRAME1TMBBROWSE, wxID_FRAME1TMBIMPORT, 
] = [wx.NewId() for _init_ctrls in range(12)]

class Frame1(wx.Frame):
    def _init_ctrls(self, prnt):
        # generated method, don't edit
        wx.Frame.__init__(self, id=wxID_FRAME1, name='', parent=prnt,
              pos=wx.Point(461, 164), size=wx.Size(637, 179),
              style=wx.DEFAULT_FRAME_STYLE,
              title='Import Data RTSM dari Hasil Closing')
        self.SetClientSize(wx.Size(621, 141))

        self.panel1 = wx.Panel(id=wxID_FRAME1PANEL1, name='panel1', parent=self,
              pos=wx.Point(0, 0), size=wx.Size(621, 144),
              style=wx.TAB_TRAVERSAL)

        self.staticText1 = wx.StaticText(id=wxID_FRAME1STATICTEXT1,
              label='Lokasi File Closing', name='staticText1',
              parent=self.panel1, pos=wx.Point(24, 24), size=wx.Size(86, 13),
              style=0)

        self.lokasi_excel = wx.TextCtrl(id=wxID_FRAME1LOKASI_EXCEL,
              name=u'lokasi_excel', parent=self.panel1, pos=wx.Point(144, 24),
              size=wx.Size(416, 21), style=0, value=u'')

        self.tmbBrowse = wx.Button(id=wxID_FRAME1TMBBROWSE, label=u'...',
              name=u'tmbBrowse', parent=self.panel1, pos=wx.Point(568, 24),
              size=wx.Size(40, 23), style=0)
        self.tmbBrowse.Bind(wx.EVT_BUTTON, self.OnTmbBrowseButton,
              id=wxID_FRAME1TMBBROWSE)

        self.button2 = wx.Button(id=wxID_FRAME1BUTTON2, label='Hitung Baris',
              name='button2', parent=self.panel1, pos=wx.Point(24, 56),
              size=wx.Size(88, 23), style=0)
        self.button2.Bind(wx.EVT_BUTTON, self.OnButton2Button,
              id=wxID_FRAME1BUTTON2)

        self.staticText2 = wx.StaticText(id=wxID_FRAME1STATICTEXT2,
              label=u'Jumlah Baris', name='staticText2', parent=self.panel1,
              pos=wx.Point(464, 64), size=wx.Size(59, 13), style=0)

        self.jml_baris = wx.TextCtrl(id=wxID_FRAME1JML_BARIS, name=u'jml_baris',
              parent=self.panel1, pos=wx.Point(536, 61), size=wx.Size(72, 21),
              style=0, value=u'')

        self.staticText3 = wx.StaticText(id=wxID_FRAME1STATICTEXT3,
              label=u'Baris yang diproses', name='staticText3',
              parent=self.panel1, pos=wx.Point(144, 56), size=wx.Size(94, 13),
              style=0)

        self.baris_proses = wx.TextCtrl(id=wxID_FRAME1BARIS_PROSES,
              name=u'baris_proses', parent=self.panel1, pos=wx.Point(264, 56),
              size=wx.Size(100, 21), style=0, value=u'')

        self.tmbImport = wx.Button(id=wxID_FRAME1TMBIMPORT,
              label='Pindah ke MySQL', name='tmbImport', parent=self.panel1,
              pos=wx.Point(24, 96), size=wx.Size(104, 23), style=0)
        self.tmbImport.Bind(wx.EVT_BUTTON, self.OnTmbImportButton,
              id=wxID_FRAME1TMBIMPORT)

        self.staticText4 = wx.StaticText(id=wxID_FRAME1STATICTEXT4,
              label='staticText4', name='staticText4', parent=self.panel1,
              pos=wx.Point(152, 104), size=wx.Size(55, 13), style=0)

    def __init__(self, parent):
        self._init_ctrls(parent)

    def OnTmbBrowseButton(self, event):
        wildcard = "File Excel (*.xls)|*.xls"
        dialog = wx.FileDialog(None, "Pilih File",
                               wildcard=wildcard,
                               style=wx.OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            self.lokasi_excel.SetValue(dialog.GetPath())
        dialog.Destroy() 
        

    def OnButton2Button(self, event):
        wb = xlrd.open_workbook(self.lokasi_excel.GetValue())
        sh = wb.sheet_by_index(0)
        self.jml_baris.SetValue(str(sh.nrows))

    def OnTmbImportButton(self, event):
        wb = xlrd.open_workbook(self.lokasi_excel.GetValue())
        sh = wb.sheet_by_index(0)
        # Hapus Isi Data Sebelumnya
        sql = "delete from data_rtsm"
        cur.execute(sql)
        i = 1
        jml = int(self.jml_baris.GetValue())
        total = jml-1
        # iterate through ieach row
        while i <= jml-1 :
            no_rtsm1 = sh.cell(i,0).value
            pengurus1=sh.cell(i,12).value
            pengurus1 = pengurus1.replace("'","")
            sd1 = sh.cell(i,23).value
            smp1=sh.cell(i,24).value
            bumil1=sh.cell(i,25).value
            balita1=sh.cell(i,26).value
            alamat1 = sh.cell(i,22).value
            kec1 = sh.cell(i,33).value
            desa1= sh.cell(i,34).value
            query = "INSERT INTO data_rtsm (no_rtsm,pengurus,sd,smp,bumil,balita,alamat,desa,kec) VALUES ('%s', '%s', '%s', '%s', '%s', '%s','%s', '%s', '%s')"%(no_rtsm1,pengurus1,sd1,smp1,bumil1,balita1,alamat1,desa1,kec1)
            cur.execute(query)
            self.baris_proses.SetValue(str(i))
            i = i + 1
            
            
        # close cursor
        #cur.close()
 
        # We are using an InnoDB tables so we need to commit the transaction
        #conn.commit()
        
        