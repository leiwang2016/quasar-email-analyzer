# -*- coding: utf-8 -*-
"""
Created on Tue Nov 15 22:52:27 2016

@author: HB60983
"""

import os
import fnmatch
import sys
import time
import wx
from wx import xrc
from wx.lib.wordwrap import wordwrap
import wx.html2
import ExtractMsg as ExtractMsg

class quasarEmailAnalysis(wx.App):
    
    swVersion = "0.0.1 Beta"

    def OnInit(self):
        self.res = xrc.XmlResource('quasarEmailAnalyzerGui.xrc')
        self.init_frame()
        return True
        
    def init_frame(self):
        self.frame = self.res.LoadFrame(None, 'quasarEmailAnalysisFrame')
        self.m_panel1 = xrc.XRCCTRL(self.frame, 'm_panel1')
        self.statusbar = xrc.XRCCTRL(self.frame, 'm_statusBar1')
        self.emailDirectory = xrc.XRCCTRL(self.frame, 'm_dirPicker2') 
        self.fromKey = xrc.XRCCTRL(self.m_panel1, 'm_textCtrl13')
        self.toKey = xrc.XRCCTRL(self.m_panel1, 'm_textCtrl14')
        self.ccKey = xrc.XRCCTRL(self.m_panel1, 'm_textCtrl16')
        self.subjectKey = xrc.XRCCTRL(self.m_panel1, 'm_textCtrl18')
        self.contextKey = xrc.XRCCTRL(self.m_panel1, 'm_textCtrl21') 
        self.attachementKey = xrc.XRCCTRL(self.m_panel1, 'm_checkBox12')
        self.ccOption = xrc.XRCCTRL(self.m_panel1, 'm_choice4')
        self.contextOption = xrc.XRCCTRL(self.m_panel1, 'm_choice5')
        self.attachmentOption = xrc.XRCCTRL(self.m_panel1, 'm_choice6')
        self.attachmentText = xrc.XRCCTRL(self.m_panel1, 'm_textCtrl7')
        self.startDate = xrc.XRCCTRL(self.m_panel1, 'm_datePicker1')
        self.endDate = xrc.XRCCTRL(self.m_panel1, 'm_datePicker2') 
        self.logger = xrc.XRCCTRL(self.m_panel1, 'm_textCtrl9')
        self.frame.Bind( wx.EVT_BUTTON, self.OnRun, id=xrc.XRCID('m_button19'))
        self.frame.Bind( wx.EVT_CHECKBOX, self.OnAttachment, id=xrc.XRCID('m_checkBox12')) 
        self.frame.Bind(wx.EVT_MENU, self.OnOpen, id=xrc.XRCID('m_menu1'))
        self.frame.Bind(wx.EVT_MENU, self.OnExit, id=xrc.XRCID('m_menu3'))
        self.frame.Bind(wx.EVT_MENU, self.OnSave, id=xrc.XRCID('m_menuItem5'))
        self.frame.Bind(wx.EVT_MENU, self.OnTips, id=xrc.XRCID('m_menuItem6'))        
        self.frame.Bind(wx.EVT_MENU, self.OnAbout, id=xrc.XRCID('m_menuItem4'))
        self.logger.Bind( wx.EVT_LEFT_DCLICK, self.highlightText)
        sys.stdout = self.logger
        self.startDate.SetValue(wx.DateTime().Today() - wx.DateSpan(days=3650))
        self.statusbar.SetStatusWidths([-3, -2, 150])
        self.statusbar.SetStatusText("Version="+self.swVersion, 2)
        self.frame.Show()
        
    def OnOpen(self, event):
        dlg = wx.DirDialog(None, "Choose a search directory:", 
                           style=wx.DD_DEFAULT_STYLE | wx.DD_NEW_DIR_BUTTON)                              
        result = dlg.ShowModal()
        if result == wx.ID_OK:        
            self.emailDirectory.SetPath(dlg.GetPath()) 
            dlg.Destroy()
            
    def OnSave(self, event):
        dlg = wx.FileDialog(None, "Save log as...", os.getcwd(), "", "*.log", \
                    wx.SAVE|wx.OVERWRITE_PROMPT)
        result = dlg.ShowModal()
        inFile = dlg.GetPath()
        dlg.Destroy()

        if result == wx.ID_OK:          #Save button was pressed
            self.logger.SaveFile(inFile, 0)
            return True
        elif result == wx.ID_CANCEL:    #Either the cancel button was pressed or the window was closed
            return False
            
    def OnTips(self, event):
        fr=wx.Frame(None,-1, title="Tips", size=(600,800))
        browser = wx.html2.WebView.New(fr, size=(600, 800))
        browser.LoadURL("file://"+os.getcwd()+"\Tips.htm")
        fr.Show()
        
    def OnExit(self, event):
        dlg = wx.MessageDialog(None, 
            "Do you really want to close this application?",
            "Confirm Exit", wx.OK|wx.CANCEL|wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        if result == wx.ID_OK:
            self.frame.Destroy()
 
    def OnAbout(self, event):
        self.aboutDiag = wx.AboutDialogInfo()
        self.aboutDiag.Name = "Quasar Email Analyer"
        self.aboutDiag.Version = self.swVersion
        self.aboutDiag.Copyright = "(c) 2016 aimatee"
        self.aboutDiag.Description = wordwrap("This is an Email search app", 350,
                                            wx.ClientDC(self.frame))
        self.aboutDiag.WebSite = ("http://aimatee.blogspot.com/", "My Home Page")
        self.aboutDiag.Developers = ["Lei Wang"]
        self.aboutDiag.License = wordwrap("Open source and free.", 500,
                                            wx.ClientDC(self.frame))
        wx.AboutBox(self.aboutDiag)                    
        
    def OnAttachment(self, event):
        self.attachmentText.Enable(self.attachementKey.GetValue())
        
    def highlightText(self, event):
        position = event.GetPosition()
        (res,hitpos) = self.logger.HitTestPos(position)
        (col,line) = self.logger.PositionToXY(hitpos)
        the_line = self.logger.GetLineText(line)
        if the_line != "Searching..." and the_line != "Searching complete!": 
            os.startfile(the_line) 
            
    def matchKeywords(self, keywords, content):
        if keywords.lstrip() == "":
            return True
        flag = False
        wholekeys = keywords.lstrip().split('"')[1::2];
        for wholekey in wholekeys:
            if wholekey.lower() in content.lower():
                return True
        keys = keywords.lstrip().split()
        for key in keys:
            multikeys = key.split('+')
            if len(multikeys) > 0:
                flag = True
                for multikey in multikeys:
                    if multikey.lower() not in content.lower():
                        flag = False
                        break
            else:
                if key.lower() in content.lower():
                    return True
        return flag

    def OnRun(self, event):
        self.logger.Clear()
        print "Searching..."
        numResults = 0 
        self.statusbar.SetStatusText("Search in process...", 0)
        self.statusbar.SetStatusText("", 1)
        startTime = time.time()
        for root, dir, files in os.walk(self.emailDirectory.GetPath()):
            for items in fnmatch.filter(files, "*.msg"):
                items_path = os.path.join(root, items)
                msg = ExtractMsg.Message(items_path)
                x = wx.DateTime()
                if msg.date != None:
                    x.ParseDateTime(msg.date)
                if (x >= self.startDate.GetValue() and x <= self.endDate.GetValue()+wx.DateSpan(days=1)) or msg.date == None: 
                    if self.matchKeywords(self.fromKey.GetValue(), msg.sender):
                        if self.matchKeywords(self.toKey.GetValue(), msg.to) \
                        and (self.matchKeywords(self.ccKey.GetValue(), msg.cc) or self.ccOption.GetSelection() == 0)\
                        or (self.matchKeywords(self.ccKey.GetValue(), msg.cc) and self.ccOption.GetSelection() == 0):
                            if (self.matchKeywords(self.subjectKey.GetValue(), msg.subject) \
                            and (self.matchKeywords(self.contextKey.GetValue(), msg.body) or self.contextOption.GetSelection() == 0)\
                            or (self.matchKeywords(self.contextKey.GetValue(), msg.body) and self.contextOption.GetSelection() == 0))\
                            and (((len(msg.attachments) >= 1 and self.attachementKey.GetValue()) or not self.attachementKey.GetValue()) or self.attachmentOption.GetSelection() == 0)\
                            or (((len(msg.attachments) >= 1 and self.attachementKey.GetValue()) or not self.attachementKey.GetValue()) and self.attachmentOption.GetSelection() == 0):
                                if self.attachementKey.GetValue() and self.attachmentOption.GetSelection() == 1:
                                    for attachment in msg.attachments:
                                        if self.matchKeywords(self.attachmentText.GetValue(), attachment.longFilename): 
                                            numResults += 1                                             
                                            print items_path
                                            break
                                else:
                                    numResults += 1
                                    print items_path
        self.statusbar.SetStatusText("Search complete: "+str(numResults)+" item(s) found!", 0)
        self.statusbar.SetStatusText("Time elapsed: "+str(time.time()-startTime)+" seconds.", 1)                            
        print "Searching complete!"
        
if __name__ == "__main__": 
    app = quasarEmailAnalysis(False)
    app.MainLoop()