# -*- coding: cp1252 -*-

#pylint: disable=no-name-in-module
#pylint: disable=undefined-variable

REALSEEKER = False
#---------------------------------------------------------------------------------------
import gc
import os
import win32con, win32gui, win32event, win32api
import msvcrt
import math
import datetime
import time
import threading
import re
import unidecode
import traceback
from string import Template
import iso639

'''
import logging
logging.basicConfig(filename='.\\com.log',level=logging.DEBUG)
logging.getLogger('.').setLevel(logging.DEBUG)
_debug = logging.getLogger('.').debug
'''

#---------------------------------------------------------------------------------------
import sys
sys.coinit_flags = 0#comtypes.COINIT_MULTITHREADED

import comtypes.client
from comtypes import _ole32, IUnknown, IPersist, GUID, STDMETHOD,HRESULT
from ctypes import POINTER, byref, cast, c_void_p, c_double, windll, c_longlong ,CFUNCTYPE, create_string_buffer, pointer,c_long
from ctypes.wintypes import INT, DWORD, LPCOLESTR, LPWSTR, LPCWSTR, LPCSTR ,BOOL, HBITMAP, COLORREF, RECT, tagRECT

comtypes.client.GetModule('.\\tlb\\DirectShow.tlb')
from comtypes.gen.DirectShowLib import *

comtypes.client.GetModule('.\\tlb\\dvd.tlb')
from comtypes.gen.DvdLib import *

comtypes.client.GetModule('quartz.dll')
from comtypes.gen.QuartzTypeLib import IMediaControl, IMediaEventEx, IVideoWindow, IBasicVideo2, IBasicAudio

#---------------------------------------------------------------------------------------
LOG = False
HWACCEL = 'No'#'CopyBack'#'Native'
AUDIODELAY = 0
# UsePSYCO = False

#---------------------------------------------------------------------------------------
from osd import osdPlay, osdTimeline, osdVol, osdBattery ,osdLocalTime ,osdAudio ,osdSrt
from osd import stateFWD, stateRWD, statePAUSE, statePLAY, labelsSEPARATOR
seekRWD, seekFWD = stateRWD, stateFWD

#-------------------------------------------------------------------------------
from ctypes import c_ulong, c_int
class IPersistMemory(IPersist):
    _iid_ = GUID('{BD1AE5E0-A6AE-11CE-BD37-504200C10000}')
    _methods_ = [
        STDMETHOD(HRESULT, 'IsDirty',[]),
        STDMETHOD(HRESULT, 'Load', [c_void_p, c_ulong]),
        STDMETHOD(HRESULT, 'Save', [c_void_p, c_int, c_ulong]),
        STDMETHOD(HRESULT, 'GetSizeMax', [POINTER(c_ulong)])
    ]
#-------------------------------------------------------------------------------
NULL = c_void_p
SHADER_PROFILE = LPCSTR('ps_2_0')
ShaderStage_PreScale = 0
ShaderStage_PostScale = 1

# class IOsdServices(IUnknown):
#     _iid_ = GUID('{3AE03A88-F613-4BBA-AD3E-EE236976BF9A}')
#     _methods_ = [
#                     STDMETHOD(HRESULT, 'OsdSetBitmap', [LPCSTR, HBITMAP, NULL, COLORREF, INT, INT,BOOL, INT, DWORD, NULL, NULL, NULL, NULL])
#                     #STDMETHOD(OsdSetBitmap)(LPCSTR name,HBITMAP leftEye = NULL,HBITMAP rightEye = NULL,COLORREF colorKey = 0,int posX = 0,int posY = 0,bool posRelativeToVideoRect = false,int zOrder = 0,DWORD duration = 0,DWORD flags = 0,OSDMOUSECALLBACK callback = NULL,LPVOID callbackContext = NULL,LPVOID reserved = NULL)
#                 ]

class IMadVRExternalPixelShaders(IUnknown):
    _iid_ = GUID('{B6A6D5D4-9637-4C7D-AAAE-BC0B36F5E433}')
    _methods_ = [
                    STDMETHOD(HRESULT, 'ClearPixelShaders', [c_long]),#STDMETHOD(ClearPixelShaders)(int stage)
                    STDMETHOD(HRESULT, 'AddPixelShader', [LPCSTR, LPCSTR, c_long, c_long])#STDMETHOD(AddPixelShader)(LPCSTR sourceCode, LPCSTR compileProfile, int stage, LPVOID reserved)
                ]

class IMadVRExclusiveModeControl(IUnknown):
    _iid_ = GUID('{88A69329-3CD3-47D6-ADEF-89FA23AFC7F3}')
    _methods_ = [
                    STDMETHOD(HRESULT, 'DisableExclusiveMode', [BOOL])
                ]

from osd import IOsdRenderCallback

class IOsdServices(IUnknown):
    _iid_ = GUID('{3AE03A88-F613-4BBA-AD3E-EE236976BF9A}')
    _methods_ = [
                    STDMETHOD(HRESULT, 'OsdSetBitmap', [LPCSTR, HBITMAP, NULL, COLORREF, INT, INT ,BOOL, INT, DWORD, NULL, NULL ,NULL, NULL]),#STDMETHOD(OsdSetBitmap)(LPCSTR name,HBITMAP leftEye = NULL,HBITMAP rightEye = NULL,COLORREF colorKey = 0,int posX = 0,int posY = 0,bool posRelativeToVideoRect = false,int zOrder = 0,DWORD duration = 0,DWORD flags = 0,OSDMOUSECALLBACK callback = NULL,LPVOID callbackContext = NULL,LPVOID reserved = NULL)
                    STDMETHOD(HRESULT, 'OsdGetVideoRects', [POINTER(RECT), POINTER(RECT)]),
                    STDMETHOD(HRESULT, 'OsdSetRenderCallback', [LPCSTR, POINTER(IOsdRenderCallback), NULL]),
                    STDMETHOD(HRESULT, 'OsdRedrawFrame', [])
                ]

class IMadVRSettings(IUnknown):
    _iid_ = GUID('{6F8A566C-4E19-439E-8F07-20E46ED06DEE}')
    _methods_ = [
                  STDMETHOD(BOOL, 'SettingsGetRevision', [POINTER(c_longlong)]),#(LONGLONG* revision) = 0;
                  STDMETHOD(BOOL, 'SettingsExport', [POINTER(c_void_p), POINTER(INT)]),#(LPVOID* buf, int* size) = 0;
                  STDMETHOD(BOOL, 'SettingsImport', [c_void_p, INT]),#(LPVOID buf, int size) = 0;
                  STDMETHOD(BOOL, 'SettingsSetString', [LPCWSTR, LPCWSTR]),#(LPCWSTR path, LPCWSTR value) = 0;
                  STDMETHOD(BOOL, 'SettingsSetInteger', [LPCWSTR, INT]),#(LPCWSTR path, int     value) = 0;
                  STDMETHOD(BOOL, 'SettingsSetBoolean', [LPCWSTR, BOOL]),#(LPCWSTR path, BOOL    value) = 0;
                  STDMETHOD(BOOL, 'SettingsGetString', [LPCWSTR, LPCWSTR, POINTER(INT)]),#(LPCWSTR path, LPCWSTR value, int* bufLenInChars) = 0;
                  STDMETHOD(BOOL, 'SettingsGetInteger', [LPCWSTR, POINTER(INT)]),#(LPCWSTR path, int*    value) = 0;
                  STDMETHOD(BOOL, 'SettingsGetBoolean', [LPCWSTR, POINTER(BOOL)]),#(LPCWSTR path, BOOL*   value) = 0;
                  STDMETHOD(BOOL, 'SettingsGetBinary', [LPCWSTR, POINTER(c_void_p), POINTER(INT)])#(LPCWSTR path, LPVOID* value, int* bufLenInBytes) = 0;
                ]

#---------------------------------------------------------------------------------------
OATrue = -1
OAFalse = 0
ROTFLAGS_REGISTRATIONKEEPSALIVE = 0x1

EC_COMPLETE = 0x0001
EC_USERABORT = 0x0002
EC_ERRORABORT  = 0x0003
EC_DVD_DOMAIN_CHANGE = 0x0101
EC_PAUSED = 0x000e
EC_VIDEO_SIZE_CHANGED = 0x000A

DVD_CMD_FLAG_Flush = 0x00000001
DVD_CMD_FLAG_Block = 0x00000004

LOCALE_SLOCALIZEDLANGUAGENAME = 0x0000006f
getLocaleInfo = windll.kernel32.GetLocaleInfoW

#---------------------------------------------------------------------------------------
class MiniPlayer(object):
    filters = dict( DvdNavigator =  '{9B8C4620-2C1A-11D0-8493-00A02438AD48}',
                    LavSplitter =   '{B98D13E7-55DB-4385-A33D-09FD1BA26338}',
                    LavVideo =      '{EE30215D-164F-4A92-A4EB-9D4C13390F9F}',
                    LavAudio =      '{E8E73B6B-4CB3-44A4-BE99-4F7BCB96E491}',
                    madVR =         '{E1A8B82A-32CE-4B0D-BE0D-AA68C772E423}',
                    XYSubFilter =   '{2DFCB782-EC20-4A7C-B530-4577ADB33F21}',
                    # reclock =       '{9DC15360-914C-46B8-B9DF-BFE67FD36C6A}'
                    )
    
    videoFilters = ('LavVideo', 'LavAudio', 'XYSubFilter')#, 'reclock')
    filterGraphName = 'FilterGraph miniPlayer'
    TicksInSec = 1000*1000*10. #1 tick = 100 ns

    def __init__(self, window):
        self.hwnd = window.getHwnd()
        self.size = window.getSize()
        self.getBattery = None

        self.setLAVvideoHWAccel(HWACCEL) 
        self.setLAVaudioDelay(AUDIODELAY)

        self.fg  = comtypes.CoCreateInstance(FilterGraph._reg_clsid_, IGraphBuilder, comtypes.CLSCTX_INPROC_SERVER)
        
        self.control =      self.fg.QueryInterface(IMediaControl)
        self.event =        self.fg.QueryInterface(IMediaEventEx)
        self.seeking =      self.fg.QueryInterface(IMediaSeeking)
        self.videoWindow =  self.fg.QueryInterface(IVideoWindow)
        self.basicVideo =   self.fg.QueryInterface(IBasicVideo2)
        self.basicAudio =   self.fg.QueryInterface(IBasicAudio)

        self.initStates()
        self.initVolume()
        self.initOsd()
        
        self.buildingFiltergraph = threading.Lock()
        self.ChangeAspectRatioLock = threading.Lock()

        self.logFile = None
        if LOG:
            self.logFile = open('log miniPlayer.txt', 'w')
            self.fg.SetLogFile(msvcrt.get_osfhandle(self.logFile.fileno()))

    def initStates(self):
        self.playing = False
        self.ready = False
        self.chooseAudioSrt = False

    def setDvdOptions(self):
        self.dvdControl.SetOption(DVD_HMSF_TimeCodeEvents, 1)
        self.dvdControl.SetOption(DVD_ResetOnStop, 0)
        self.dvdControl.SetOption(DVD_DisableStillThrottle, 1)
        self.dvdControl.SetOption(DVD_EnableStreaming, 1)
        self.dvdControl.SetOption(DVD_EnablePortableBookmarks, 1)

    def createFilter(self, filterName):
        print filterName
        clsid = GUID(self.filters[filterName])
        f = comtypes.CoCreateInstance(clsid, IBaseFilter, comtypes.CLSCTX_INPROC_SERVER)
        self.fg.AddFilter(f, LPCWSTR(filterName))
        return f

    def getFilterName(self, f):
        i = f.QueryFilterInfo()
        name = cast(i.achName, LPCWSTR)
        return name.value

    def buildFiltergraph(self):
        self.initSeeker()
        self.initAUDIOSRT()

        filename = self.files[self.fileIndex]
        self.dvd = 'video_ts' in filename.lower()

        if self.dvd:
            sourceFilter = self.createFilter('DvdNavigator')
            self.dvdControl = sourceFilter.QueryInterface(IDvdControl2)
            self.dvdInfo =  sourceFilter.QueryInterface(IDvdInfo2)

            dirName = os.path.dirname(filename)
            self.dvdControl.SetDVDDirectory(LPCWSTR(dirName))
            self.setDvdOptions()
            self.setDefaultLanguage() #need dvdControl 
            
        else:
            self.setDefaultLanguage() #change registry before LAV is created
            sourceFilter = self.createFilter('LavSplitter')
            source = sourceFilter.QueryInterface(IFileSourceFilter)
            source.Load(LPCOLESTR(filename), None)

        madvr = self.createFilter('madVR')
        self.osdService = madvr.QueryInterface(IOsdServices)
        self.madSettings = madvr.QueryInterface(IMadVRSettings)
        self.exclusiveModeControl = madvr.QueryInterface(IMadVRExclusiveModeControl)
        self.pixelShaders = madvr.QueryInterface(IMadVRExternalPixelShaders)
        self.setRotation()

        self.videoWindow.Owner = self.hwnd #better for madvr

        for filterName in self.videoFilters:
            self.createFilter(filterName)

        for pin in sourceFilter.EnumPins():
            self.fg.Render(pin)

        self.addToROT()
        self.setWindow()
        self.runOsd()
        self.setVolume(self.vol) #set current volume, in case it has been changed by the user in a previous session
        print 'GRAPH build'
    
    def releaseFilterGraph(self):
        self.initStates()
        self.stop()
        self.stopOsd()
        self.stopSeeker()
        self.releaseAUDIOSRT()
        
        self.videoWindow.Visible = OAFalse
        #self.exclusiveModeControl.DisableExclusiveMode(True) #Ivideowindow would freeze when it was called from another thread than the one where we build the graph
        self.videoWindow.Owner = 0
        
        self.madSettings = None
        self.osdService = None
        self.pixelShaders = None
        self.exclusiveModeControl = None
        
        self.srt = None
        self.audio = None
        
        if self.dvd:
            self.dvdControl = None
            self.dvdInfo =  None

        for f in [f for f in self.fg.EnumFilters()]: #copy in a list before iterating, because filters are deleted
            print 'Remove Filter "%s"'%self.getFilterName(f)
            self.fg.RemoveFilter(f)

        self.delFromROT()

        if self.logFile is not None : self.logFile.close()
        gc.collect()
    
    def clean(self):
        print 'Release filterGraph'
        nb = self.fg.Release()
        for _ in range(nb):
            self.fg.Release()
        gc.collect()

    def setDefaultLanguage(self):
        if self.dvd:
            lcidFRE, lcidENG = 0x040c, 0x0409 
            self.dvdControl.SelectDefaultMenuLanguage(lcidFRE)
            if self.forceFrench:
                self.dvdControl.SelectDefaultAudioLanguage(lcidFRE, DVD_AUD_EXT_NotSpecified)
            else:
                self.dvdControl.SelectDefaultAudioLanguage(lcidENG, DVD_AUD_EXT_NotSpecified)
                self.dvdControl.SelectDefaultSubpictureLanguage(lcidFRE, DVD_SP_EXT_NotSpecified)
        else:
            self.setLAVsubtitlesPref(self.forceFrench)
    
    #---------------------------------------------------------------------------------------
    def handleGraphEvent(self):
        while True:
            try:
                evCode,evParam1,evParam2 = self.event.GetEvent(0)
                self.event.FreeEventParams(evCode, evParam1, evParam2)
                #print 'GRAPH event %s %s %s'%(hex(evCode),hex(evParam1),hex(evParam2))
                
                if evCode in (EC_USERABORT, EC_ERRORABORT):
                    print 'EC_ABORT'
                    self.movieEnded(False, None) #not viewed !

                elif evCode==EC_COMPLETE:
                    print 'EC_COMPLETE'
                    if self.needSeekInPlaylist(seekFWD):
                        self.movieEnded(None, None, seekFWD)
                    else:
                        self.movieEnded(True, None) #viewed
                
                elif evCode==EC_VIDEO_SIZE_CHANGED :
                    ratio = evParam1&0xffff , evParam1>>16
                    print 'change ivideowindow size %dx%d'%ratio
                    self.setAspectRatio(ratio=ratio, setWindowPos=False)
                    
                elif evCode==EC_DVD_DOMAIN_CHANGE:
                    print 'change DVD domain'
                    ratio = self.getAspectRatioDvd()
                    self.setAspectRatio(ratio=ratio, setWindowPos=False)
                   
            except: break

    def setWindow(self):
        self.videoWindow.WindowStyle = win32con.WS_CHILD | win32con.WS_CLIPSIBLINGS # win32con.WS_VISIBLE 
        self.videoWindow.WindowStyleEx = win32con.WS_EX_TOOLWINDOW #win32con.WS_EX_TOPMOST

        self.videoWindow.Visible = OATrue
        self.videoWindow.AutoShow = OATrue
        self.videoWindow.WindowState = win32con.SW_SHOW
        self.videoWindow.HideCursor(OATrue)
        self.videoWindow.SetWindowForeground = OATrue

        self.setAspectRatio()

    def setAspectRatio(self, ratio=None, setWindowPos=True):
        with self.ChangeAspectRatioLock:
            screenW,screenH = self.size
    
            if ratio is None:
                aspectW,aspectH = self.basicVideo.GetPreferredAspectRatio()
                sourceW,sourceH = self.basicVideo.GetVideoSize()
                if aspectW==0 or aspectH==0 or sourceW>0 or sourceH>0:
                    aspectW,aspectH = sourceW,sourceH
            else :
                aspectW,aspectH = ratio
            
            if self.angle in (90,270): #prevent other steps than 90°
                aspectW,aspectH = aspectH,aspectW
            
            adjustedH = int(round(float(aspectH)*screenW/aspectW)) if (aspectW > 0 and aspectH > 0) else 0
    
            if screenH >= adjustedH and adjustedH>0:
                topMargin = int(round((screenH - adjustedH)/2.))
                self.basicVideo.SetDestinationPosition(0, topMargin, screenW, adjustedH)
    
            elif adjustedH>0:
                adjustedW = int(round(float(aspectW)*screenH/aspectH))
                leftMargin  = int(round((screenW - adjustedW)/2.))
                self.basicVideo.SetDestinationPosition(leftMargin, 0, adjustedW, screenH)
    
            if setWindowPos:
                self.videoWindow.SetWindowPosition(0, 0, screenW, screenH)

    def getAspectRatioDvd(self):
        vAttr = self.dvdInfo.GetCurrentVideoAttributes()
        return vAttr.ulAspectX,vAttr.ulAspectY

    #---------------------------------------------------------------------------------------
    def playMovie(self, favorite, files, angle=0, forceFrench=False, isVideo=False):

        self.angle = angle%360
        self.isVideo = isVideo
        self.forceFrench = forceFrench
        
        with self.buildingFiltergraph:
            if favorite is not None:
                self.playFavorite(favorite)         
            else:
                self.playFiles(files)

        if self.isVideo : self.exclusiveModeControl.DisableExclusiveMode(True)
        self.ready = True
        print 'MiniPlayer ready'

    def playFiles(self, files):
        self.files = files if isinstance(files, (list, tuple)) else [files]
        self.fileIndex = 0

        self.buildFiltergraph()
        if not self.dvd: self.doSeek(0) # if we don't do that, ImediaSeeking.setpositions seeks to the position when the graph was last stopped
        self.play()
            
        if self.dvd:
            try:
                self.dvdControl.ShowMenu(2,0) #main menu
            except comtypes.COMError,e:
                #pylint: disable=no-member
                print "can't acces menu, hresult %s"%hex(0xffffffff+e.hresult) 

    #---------------------------------------------------------------------------------------
    def seekInPlaylist(self, seek): #won't work with DVD playlist!
        with self.buildingFiltergraph:
            self.fileIndex += seek
            print '-----------------------------\nPlay File index %s in playlist'%self.fileIndex
            self.buildFiltergraph()
            pos = self.getDuration()- 1*self.TicksInSec if seek==seekRWD else 0
            self.doSeek(pos)
            self.play()
            if self.isVideo : self.exclusiveModeControl.DisableExclusiveMode(True)
            self.ready = True
            print 'seek in Playlist done'

    def needSeekInPlaylist(self, seek):
        return 0<= self.fileIndex + seek <len(self.files)

    #---------------------------------------------------------------------------------------
    def playFavorite(self, favorite):
        if len(favorite)==2:
            files,dvdstateData = favorite

            self.files = files
            self.fileIndex = 0
            self.buildFiltergraph()

            self.pause()

            dvdState = self.dvdInfo.GetState()
            persist = dvdState.QueryInterface(IPersistMemory)
            loadBuffer = create_string_buffer(dvdstateData, len(dvdstateData))
            persist.Load(cast(loadBuffer, c_void_p), len(loadBuffer))

            self.dvdControl.SetState(dvdState, DVD_CMD_FLAG_Block | DVD_CMD_FLAG_Flush)

        else:
            files, fileIndex, pos, audioIndex, srtIndex = favorite

            self.files = files
            self.fileIndex = fileIndex
            self.buildFiltergraph()

            self.pause()

            self.doSeek(pos)
            self.getAUDIOSRT()
            self.audioIndex, self.srtIndex = audioIndex, srtIndex
            self.selectSRT()
            self.selectAUDIO()

        self.play()

    def getFavorite(self):
        if self.dvd:
            dvdState = self.dvdInfo.GetState()
            persist = dvdState.QueryInterface(IPersistMemory)
            
            size = c_ulong(0) 
            persist.GetSizeMax(byref(size))
            size = size.value
            saveBuffer = create_string_buffer(size)
            persist.Save(cast(saveBuffer, c_void_p),1,size)
            dvdstateData = saveBuffer.raw

            return self.files,dvdstateData
            
        else:
            self.getAUDIOSRT()
            pos = self.getPos()
            #title = self.dvdInfo.GetCurrentLocation().TitleNum if self.dvd else None
            return self.files, self.fileIndex, pos, self.audioIndex, self.srtIndex

    def isViewed(self):
        BEFOREendOfMOvie = 15*60
        PERCENTviewed = .90
        isViewed = False
        if self.fileIndex == len(self.files)-1:
            try:
                pos,dur = self.getPosAndDuration()
                posSec = pos/self.TicksInSec
                totalSec = dur/self.TicksInSec
                viewSec = max(totalSec-BEFOREendOfMOvie, totalSec*PERCENTviewed)
                isViewed = posSec>viewSec
                print 'Viewed %s (pos:%is,view:%s,total:%ss)' %(isViewed, posSec, viewSec, totalSec)
            except : pass
        return isViewed

    def close(self):
        isViewed =  False
        favorite = None

        if self.ready and not self.isInDVDMenu():
            self.pause()
            isViewed = self.isViewed()
            favorite = self.getFavorite()
       
        self.movieEnded(isViewed, favorite)

    def movieEnded(self, viewed, favorite, seek=None):
        angle = self.angle if self.isVideo else 0
        self.stopMsgLoop(viewed, favorite, angle, seek) #so that a possible call to seekInPLaylist would be in the msg loop thread (buildgraph thread too!)

    def forceClose(self):
        print 'Force CLOSE miniplayer'
        with self.buildingFiltergraph:
            self.close()

    #---------------------------------------------------------------------------------------
    def Pos2HMS(self, seconds):
        seconds = int(round(seconds/self.TicksInSec))
        m, s = divmod(seconds, 60)
        h, m = divmod(m, 60)
        return tagDVD_HMSF_TIMECODE(h ,m, s, 0)

    def HMS2Pos(self,hms):
        return int(round((hms.bHours*3600 + hms.bMinutes*60 + hms.bSeconds) * self.TicksInSec))

    def getPosAndDuration(self):
        return (self.getPos(), self.getDuration()) if self.dvd else self.seeking.GetPositions()
            
    def getDuration(self):
        return self.HMS2Pos(self.dvdInfo.GetTotalTitleTime()[0]) if self.dvd else self.seeking.GetDuration()

    def getPos(self):
        return self.HMS2Pos(self.dvdInfo.GetCurrentLocation().TimeCode) if self.dvd else self.seeking.GetCurrentPosition()

    def initSeeker(self):
        self.lastState = statePLAY
        self.lastStopTime = 0

    def stopSeeker(self):
        self.lastState = statePAUSE
       
        
    DVDminTimeBeforeNextSeek = .7 # to prevent too close DVD seeks that would hangs

    def seek(self, state):
        if self.dvd and time.time()-self.lastStopTime < self.DVDminTimeBeforeNextSeek:
            return
        
        if state in (stateFWD, stateRWD):
            self.pause()

            if state != self.lastState: # start seeking
                self.duration = self.showState(state) #set BEGIN SEEK
                print 'SEEK BEGIN duration %s'%self.duration
            
            pos = self.renderCallback.getCurrentPos()
            pos = max(0, min(pos, self.duration))

            if pos in (0, self.duration):
                if self.needSeekInPlaylist(state):
                    self.doSeek(pos)
                    self.movieEnded(None, None, seek=state) #-1 RWD, 1 FWD
                    return

            if REALSEEKER:
                self.pos = pos
                #print 'seek pos%s'%self.pos
                win32event.SetEvent(self.seekEvent) #self.doSeek(pos) 
                time.sleep(0.1)
        
        else: 
            self.showState(state) #set END SEEK
            pos = self.renderCallback.getCurrentPos()
            pos = max(0,min(pos,self.duration))
            print 'SEEK END PLAY %s'%pos
            
            if pos == self.duration and not self.needSeekInPlaylist(seekFWD) and not self.dvd:
                #exit right after the user release the seek button, except for DVD where we're not sure it's the main VOB
                self.movieEnded(True, None)
                return
            else:
                self.doSeek(pos)
                self.play() 
                self.lastStopTime = time.time()

        self.lastState = state

    def doSeek(self, pos):
        if self.dvd:
            hms = self.Pos2HMS(pos)
            #print hms.bHours,hms.bMinutes,hms.bSeconds
            flags = DVD_CMD_FLAG_Block | DVD_CMD_FLAG_Flush
            self.dvdControl.PlayAtTime(hms,flags)
        else:
            flags = AM_SEEKING_AbsolutePositioning | AM_SEEKING_SeekToKeyFrame #|AM_SEEKING_NoFlush
            self.seeking.SetPositions(int(round(pos)), flags, 0, AM_SEEKING_NoPositioning)
    
    #---------------------------------------------------------------------------------------
    def showState(self, state):
        p,d = self.getPosAndDuration()
        self.renderCallback.setState(state, p, d)
        
        if state == statePAUSE:
            self.showOSD(osdPlay, 'Pause', delay=self.delayINFINITE)
            self.showBattery()
            self.showOSD(osdLocalTime, '', delay=self.delayINFINITE)
            self.showOSD(osdTimeline, delay=self.delayINFINITE) 

        elif state == statePLAY:
            self.showOSD(osdPlay, 'Play', delay=1)
            self.hideOSD(osdBattery, delay=0)
            self.hideOSD(osdLocalTime, delay=0)
            self.hideOSD(osdTimeline, delay=1)
        
        elif state in (stateFWD, stateRWD):
            label = {stateRWD:'<<' ,stateFWD:'>>'}[state]

            self.showOSD(osdPlay, label, delay=self.delayINFINITE)
            self.hideOSD(osdBattery)
            self.hideOSD(osdLocalTime)
            self.showOSD(osdTimeline, delay=self.delayINFINITE) 
        
        return d

    #---------------------------------------------------------------------------------------
    def showVol(self):
        #vol = int(round(self.getVolume()))
        self.showOSD(osdVol, 'Vol: %d%%'%self.vol, self.vol, delay=1)

    #---------------------------------------------------------------------------------------
    def setBatteryFunc(self, GetBatteryFunc):
        self.getBattery = GetBatteryFunc
    
    def showBattery(self):
        label = None
        value = 100
        if self.getBattery is not None:
            bat = self.getBattery()
            if bat is not None:
                value = bat
                label = 'POW:%d%%'%bat
        self.showOSD(osdBattery, label, value, delay=self.delayINFINITE)
        
    #---------------------------------------------------------------------------------------
    delayINFINITE = -1
    defaultDelay = 1 #s
    
    def initOsd(self):
        self.renderCallback = comtypes.CoCreateInstance(IOsdRenderCallback._iid_, IOsdRenderCallback, comtypes.CLSCTX_INPROC_SERVER)

    def runOsd(self):
        w,h = self.size
        self.renderCallback.init(w, h, self.osdService,len(self.files), self.fileIndex)
        self.renderCallback.run()
        self.osdService.OsdSetRenderCallback('miniPlayer.osd', self.renderCallback,None)

    def isOsdInitialized(self):
        return self.renderCallback.isInitialized()

    def stopOsd(self):
        self.renderCallback.stop()
        self.osdService.OsdSetRenderCallback('miniPlayer.osd', None, None)

    def showOSD(self, name,label='', value=-1, delay=defaultDelay):
        self.renderCallback.showButton(name, label, value, delay)

    def hideOSD(self, name, delay=0):
        self.renderCallback.hideButton(name, delay)

    def hideAllOSDExceptAudioSrt(self):
        for name in  (osdPlay, osdTimeline, osdVol, osdBattery, osdLocalTime):
            self.hideOSD(name, delay=0)

    #---------------------------------------------------------------------------------------
    def showAudioSrt(self, name, label, labels, index, selected):
        labels = labelsSEPARATOR.join(labels)
        self.renderCallback.showCombo(name, label, labels, index, selected)

    def hideAudioSrt(self, name):
        self.renderCallback.hideCombo(name)

    def changeSelectionAudioSrt(self, name, index,selected):
        self.renderCallback.selectCombo(name, index, selected)

    def beginSelectAudioSrt(self):
        self.chooseAudioSrt = True
        self.getAUDIOSRT()
        self.col = 0

        self.hideAllOSDExceptAudioSrt()
        self.showAudioSrt(osdAudio, 'Audio', [s.name for s in self.audio], self.audioIndex, self.col==0)
        self.showAudioSrt(osdSrt, 'Sous-titres', [s.name for s in self.srt], self.srtIndex, self.col==1)
        self.pause()

    def endSelectAudioSrt(self, select):
        self.hideAudioSrt(osdAudio)
        self.hideAudioSrt(osdSrt)

        if select:
            self.selectSRT()
            self.selectAUDIO()
        self.chooseAudioSrt = False
        self.showState(statePLAY)
        self.play()
    
    def selectAudioSrt(self, cmd):
        dx,dy = dict(up=(0,-1), down=(0,1), left=(-1,0), right=(1,0))[cmd]

        self.col = (self.col+dx)%2
        if self.col ==0:
            self.audioIndex = self.keepInRange(0, self.audioIndex+dy, len(self.audio)-1)
        else:
            self.srtIndex = self.keepInRange(0, self.srtIndex+dy, len(self.srt)-1)

        self.changeSelectionAudioSrt(osdAudio, self.audioIndex, self.col==0)
        self.changeSelectionAudioSrt(osdSrt, self.srtIndex, self.col==1)

    def keepInRange(self, amin, v, amax):
        return max(amin, min(v, amax))
            
    #---------------------------------------------------------------------------------------
    NoSubtitleTxt = 'Pas de Sous-Titres'

    class sourceStream():
        types = {1:'A', 2:'S', 6590018:'S', 6590016:'embedded'} 

        def __init__(self, index, iss, filterName, dummyNoSubtitle=False, betterName=True):
            self.index = index
            self.iss = iss
            self.showStream = None
            self.hideStream = None
            self.lav = 'Lav' in filterName

            if dummyNoSubtitle:
                self.type = 'S'
                self.name = MiniPlayer.NoSubtitleTxt
                self.dummyNoSubtitle = True
            
            else:
                info = iss.Info(index)
                    
                self.selected = info[1]!=0
                #name = unidecode.unidecode(cast(info[4],LPCWSTR).value)
                name = cast(info[4],LPCWSTR).value
                
                self.type = self.types.get(info[3],None)
                #print '%d: "%s"'%(info[3],self.name)

                if   name == 'Show Subtitles' : self.type = 'show'
                elif name == 'Hide Subtitles' : self.type = 'hide'
                #elif name == 'Subtitle (embeded)' : self.type = 'dummy'

                self.name = re.sub('^[AS]:','',name).strip()
                if self.name.lower() == 'no subtitles' : 
                    self.name = MiniPlayer.NoSubtitleTxt
                elif betterName:
                    self.getBetterName()

                self.dummyNoSubtitle = False
        
        def getBetterName(self):
            forced = ''
            name = unidecode.unidecode(self.name.lower())
            
            if self.isSrt():
                for srch in ('force', 'forzado'): # "forced" in french, english, spanish
                    if srch in name:
                        forced = u'(Forcé)'
                        break
                    
            words = re.findall(r'\b\S+\b', name)
            for abr in (w for w in words if len(w)==3): #3 letters words are potentially an ISO 639-2 language code
                lang = iso639.langs.get(abr, None)
                if lang is not None: #assume that the first one is a valid language code
                    self.name = '%s %s'%(lang.capitalize(), forced)
                    return 

        def isAudio(self):  return self.type =='A'
        def isSrt(self):    return self.type == 'S'
        def isHideSrt(self) : return self.type == 'hide' #XY stream to hide subtitles
        def isShowSrt(self) : return self.type == 'show' #XY stream to show subtitles
        def isEmbeddedSrt(self) : return self.type == 'embedded' #XY stream to select embedded (LAV) subtitles
        
        def _select(self):
            self.iss.Enable(self.index, 1)
        
        def select(self, player):
            if self.dummyNoSubtitle:
                player.hideSrtStream._select() #hide srt with XYsubfilter
            else:
                self._select()
                if self.isSrt():
                    player.showSrtStream._select() #show srt with XYsubfilter
                    if self.lav and player.embeddedSrtStream is not None:
                        player.embeddedSrtStream._select() #embeddedSrtStream is the selected LAV subtitle in XYsubfilter

        def release(self):
            del self.iss

    class dvdTrack(object):
        def getLang(self, lcid):
            name = LPWSTR(' '*255)
            getLocaleInfo(lcid, LOCALE_SLOCALIZEDLANGUAGENAME, name, 255)
            return name.value

        def release(self): pass

    class dvdSRT(dvdTrack):
        def __init__(self, index, player, noSrt=False):
            self.index = index
            self.noSrt = noSrt
            if self.noSrt :
                self.name = MiniPlayer.NoSubtitleTxt
            else:
                attr = player.dvdInfo.GetSubpictureAttributes(index)
                lang = self.getLang(attr.Language) if attr.Language else 'Inconnu'
                ext = 'Commentaires ' if attr.LanguageExtension in (DVD_SP_EXT_DirectorComments_Normal, DVD_SP_EXT_DirectorComments_Big, DVD_SP_EXT_DirectorComments_Children) else ''
                self.name = '%s%s'%(ext, lang)

        def select(self, player):
            if self.noSrt :
                player.dvdControl.SetSubpictureState(0, DVD_CMD_FLAG_Block)
            else:
                player.dvdControl.SelectSubpictureStream(self.index, DVD_CMD_FLAG_Block)
                player.dvdControl.SetSubpictureState(1, DVD_CMD_FLAG_Block)

    class dvdAUDIO(dvdTrack):
        def __init__(self, index, player):
            self.index = index

            attr = player.dvdInfo.GetAudioAttributes(index)
            lang = self.getLang(attr.Language) if attr.Language else 'Inconnu' 
            ext = 'Commentaires ' if attr.LanguageExtension in (DVD_AUD_EXT_DirectorComments1, DVD_AUD_EXT_DirectorComments2) else ''

            freq = '%dkHz'%(attr.dwFrequency/1000)

            aformat = {DVD_AudioFormat_AC3:'AC3', DVD_AudioFormat_MPEG1:'MPEG1', DVD_AudioFormat_MPEG1_DRC:'MPEG1', DVD_AudioFormat_MPEG2:'MPEG2', DVD_AudioFormat_MPEG2_DRC:'MPEG2', DVD_AudioFormat_LPCM:'LPCM', DVD_AudioFormat_DTS:'DTS',DVD_AudioFormat_SDDS:'SDDS'}
            aformat = aformat.get(attr.AudioFormat, '')

            channels = {1:'Mono', 2:u'Stéréo', 3:'2.1', 4:'Quadriphonie', 5:'5.0', 6:'5.1', 7:'7.0', 8:'7.1'}
            channels = channels.get(attr.bNumberOfChannels, '')
            
            self.name = '%s%s (%s)'%(ext, lang, ' '.join((aformat, channels, freq)) )
        
        def select(self,player):
            player.dvdControl.SelectAudioStream(self.index, DVD_CMD_FLAG_Block)

    def getStreamFromFilter(self, filtername, betterName=True):
        f = self.fg.FindFilterByName(LPCWSTR(filtername))
        iss = f.QueryInterface(IAMStreamSelect)

        for i in range(iss.Count()):
            yield self.sourceStream(i, iss, filtername, betterName=betterName)

    def initAUDIOSRT(self):
        self.srt, self.audio = [], []
        self.embeddedSrtStream = None
        self.showSrtStream = None
        self.hideSrtStream = None

    def releaseAUDIOSRT(self):
        for stream in self.srt+self.audio: stream.release()
        self.initAUDIOSRT()

    def getAUDIOSRT(self):
        self.releaseAUDIOSRT()
        
        if self.dvd:
            nb,index = self.dvdInfo.GetCurrentAudio()
            self.audio = [self.dvdAUDIO(i,self) for i in range(nb)]
            self.audioIndex = index

            nb,index,disabled = self.dvdInfo.GetCurrentSubpicture()
            nb = min(nb,16)
            self.srt = [self.dvdSRT(0,self,True)] + [self.dvdSRT(i, self) for i in range(nb)]
            self.srtIndex = 0 if disabled else index+1

        else :
            self.srt, self.audio = [], []
            self.srtIndex, self.audioIndex = 0, 0

            for stream in self.getStreamFromFilter('LavSplitter'):
                if stream.isAudio():
                    if stream.selected : self.audioIndex = len(self.audio)
                    self.audio.append(stream)
                elif stream.isSrt():
                    if stream.selected: self.srtIndex = len(self.srt)
                    self.srt.append(stream)
            nbLAVsrt = len(self.srt)

            for stream in self.getStreamFromFilter('XYSubFilter', betterName=False):
                if stream.isSrt():
                    if stream.selected : self.srtIndex = len(self.srt)
                    self.srt.append(stream)
                elif stream.isEmbeddedSrt():
                    self.embeddedSrtStream = stream
                elif stream.isShowSrt():
                    self.showSrtStream = stream
                elif stream.isHideSrt():
                    self.hideSrtStream = stream
                    hideStreamSelected = stream.selected
            nbXYsrt = len(self.srt)-nbLAVsrt

            if nbLAVsrt==0:
                #XYsrt only, add dummyNoSubtitle to let the user hide subtitles
                if hideStreamSelected : self.srtIndex = len(self.srt) #select dummyNoSubtitle since hideSrtStream is selected
                self.srt.append(self.sourceStream(0, None, '', dummyNoSubtitle=True))

            elif  self.embeddedSrtStream is None and nbXYsrt>0:  
                # We got some LAVsrt but embeddedSrtStream is missing (needed to switch from XYsrt to LAVsrt)
                self.srt[0].select(self) #select first LAVsrt (we're sure it's not Nosubtitle)
                self.getAUDIOSRT() # do again to get embeddedSrtStream

            else:
                #put LAVsrt in front of XYsrt so that LAV noSubtitle stream is last
                self.srt = self.srt[nbLAVsrt:] + self.srt[:nbLAVsrt] #XY srt + LAv srt
                self.srtIndex = (self.srtIndex-nbLAVsrt) %len(self.srt)

    def selectSRT(self):
        self.srt[self.srtIndex].select(self)

    def selectAUDIO(self):
        self.audio[self.audioIndex].select(self)

    #---------------------------------------------------------------------------------------
    def initVolume(self):
        self.dampVol = 50.
        self.cVol = self.percentTodB(100)
        self.kVol = 10000./(self.percentTodB(0)-self.cVol)

        self.vol = 100

    def percentTodB(self,p):
        return math.pow(10, -math.log(1+(p/self.dampVol)))

    '''
    def getVolume(self):
        vol = self.basicAudio.Volume
        vol = self.dampVol * (math.exp(-math.log(self.cVol-vol/(self.kVol),10))-1)
        self.vol = round(vol, 1)
        return self.vol
    '''

    def setVolume(self,vol):
        self.vol = vol = min(100,max(0,vol))
        vol = -self.kVol * (self.percentTodB(vol)-self.cVol)
        self.basicAudio.Volume = int(vol)

    incVOL = 2
    def updateVolume(self, dvol):
        dvol = self.incVOL if dvol is 'up' else -self.incVOL
        self.vol = self.keepInRange(0, self.vol+dvol, 100)
        self.showVol()
        self.setVolume(self.vol)

    def getBalance(self):
        bal = self.basicAudio.Balance
        return 50 + bal.value/200

    def setBalance(self,vol):
        bal = ((min(100, max(0, vol))-50) * 200)
        self.basicAudio.Balance = int(bal)

    #---------------------------------------------------------------------------------------
    def computeRot(self, angle):
        rad = -math.pi * angle/180.
        cos,sin = math.cos(rad), math.sin(rad)
        scale =  abs(sin) + abs(cos)
        return  scale, cos, sin

    def setRotation(self):
        self.pixelShaders.ClearPixelShaders(ShaderStage_PreScale)
        
        if self.angle in (90, 180, 270):

            shaderCode  = '''
                        sampler s0 : register(s0);                         
                        float4 main(float2 tex: TEXCOORD0) : COLOR       
                        {                                                
                            float2 t;                                       
                            tex.x -= 0.5;                                   
                            tex.y -= 0.5;                                   
                            t.x = $SCALE*(tex.x*$COS - tex.y*$SIN) + 0.5;           
                            t.y = $SCALE*(tex.y*$COS + tex.x*$SIN) + 0.5;           
                            if(t.x >= 0 && t.x <= 1 && t.y >= 0 && t.y <= 1) 
                                return tex2D(s0, t);                          
                            return (0,0,0,0);                               
                        }                                                
                        ''' 
            scale,cos,sin = self.computeRot(self.angle)
            shaderCode = Template(shaderCode).substitute(SCALE=scale, COS=cos, SIN=sin)

            self.pixelShaders.AddPixelShader(shaderCode, SHADER_PROFILE, ShaderStage_PreScale, 0)

    def changeRotation(self, inc):
        self.angle = (self.angle+inc) % 360
        self.setRotation()
        self.setAspectRatio(setWindowPos=False)
            
    #---------------------------------------------------------------------------------------
    def isInDVDMenu(self):
        if self.dvd:
            return self.dvdInfo.GetCurrentDomain() != DVD_DOMAIN_Title
        return False

    def dvdButton(self,cmd):
        try:
            if   cmd == 'select' :  self.dvdControl.ActivateButton()
            elif cmd == 'left' :    self.dvdControl.SelectRelativeButton(DVD_Relative_Left)
            elif cmd == 'right' :   self.dvdControl.SelectRelativeButton(DVD_Relative_Right)
            elif cmd == 'up' :      self.dvdControl.SelectRelativeButton(DVD_Relative_Upper)
            elif cmd == 'down' :    self.dvdControl.SelectRelativeButton(DVD_Relative_Lower)
        except : pass

    def playPause(self):
        self.playing = self.control.GetState(1000) == State_Running
        if self.playing :
            self.showState(statePAUSE)
            self.pause()
        else:
            self.showState(statePLAY)
            self.play()
    
    def play(self):
        self.renderCallback.forceRedraw(False)
        #self.madSettings.SettingsSetInteger(LPCWSTR('exclusiveSettings\\preRenderFrames'),8)
        self.playing = True
        self.control.Run()
        #print 'PLAY'

    def pause(self):
        self.playing = False
        self.control.Pause()
        #self.madSettings.SettingsSetInteger(LPCWSTR('exclusiveSettings\\preRenderFrames'),1)
        self.renderCallback.forceRedraw(True)
        #print 'PAUSE'
    
    def stop(self):
        self.playing = False
        self.control.Stop()

    #---------------------------------------------------------------------------------------
    ROOTKEY = win32con.HKEY_CURRENT_USER
    LAVPATH = r'Software\LAV'
    
    def setRegKey(self, root, path, key, rtype, value):
        try:
            hKey = win32api.RegOpenKey(root, path, 0, win32con.KEY_ALL_ACCESS)
            win32api.RegSetValueEx(hKey, key, 0, rtype, value)        
        except:
            print u"ERROR, can't access REG %s\\%s" %(path, key)

    def setLAVvideoHWAccel(self, accel):
        path = self.LAVPATH + r"\Video\HWAccel"
        key = r"HWAccel"
        rtype = win32con.REG_DWORD
        value = dict(No=0, CopyBack=3, Native=4).get(accel, 0)
        self.setRegKey(self.ROOTKEY, path, key, rtype, value)
        print 'Lav Video - HWAccel : %s'%accel

    def setLAVaudioDelay(self, delay):
        path = self.LAVPATH + r"\Audio"
        rtype = win32con.REG_DWORD

        if delay != 0:
            keyDelay = r"AudioDelay"
            self.setRegKey(self.ROOTKEY, path, keyDelay, rtype, delay)

        enable = 1 if delay !=0 else 0
        keyEnable = r"AudioDelayEnabled"
        self.setRegKey(self.ROOTKEY, path, keyEnable, rtype, enable)
            
        print 'Lav Audio - Delay : %d'%delay

    def setLAVsubtitlesPref(self, forceFrench=False):
        path = self.LAVPATH + r"\Splitter"
        rtype = win32con.REG_SZ

        if forceFrench :
            langPrefs = 'fre'
            subPrefs = '*:off'
        else:
            langs = ('eng', 'jap', 'dan', 'spa', 'kor', 'ger', 'swe', 'por', 'ita')
            langPrefs = ','.join(langs)
            subPrefs = ','.join(['%s:fre|n!f*'%l for l in langs] + ['*:fre|n!f*','*:*'])
        
        print 'LAV Audio langs : %s'%langPrefs

        key = r"prefAudioLangs"
        self.setRegKey(self.ROOTKEY, path, key, rtype, langPrefs)
        
        key = r"subtitleAdvanced"
        self.setRegKey(self.ROOTKEY, path, key, rtype, subPrefs)

    #---------------------------------------------------------------------------------------
    def handleCommands(self, cmd, pressed):
        if self.ready and self.isOsdInitialized():
            #print 'CMD %s %s'%(cmd,pressed)
            try:
                inDVDMenu = self.isInDVDMenu()
                
                if pressed :
                    if cmd == 'close':          self.close()
    
                    elif cmd == 'lang':
                        if not self.chooseAudioSrt:
                            if not inDVDMenu:   self.beginSelectAudioSrt()
    
                    elif cmd == 'select':
                        if self.chooseAudioSrt: self.endSelectAudioSrt(True)
                        elif inDVDMenu:         self.dvdButton('select')
                        else:                   self.playPause()
                            
                    elif cmd == 'back':
                        if self.chooseAudioSrt: self.endSelectAudioSrt(False)
                            
                    elif cmd == 'up':
                        if self.chooseAudioSrt: self.selectAudioSrt('up')
                        elif inDVDMenu:         self.dvdButton('up')
                        else:                   self.updateVolume('up')
    
                    elif cmd == 'down':
                        if self.chooseAudioSrt: self.selectAudioSrt('down')
                        elif inDVDMenu:         self.dvdButton('down')
                        else:                   self.updateVolume('down')
                            
                    elif cmd == 'left':
                        if self.chooseAudioSrt: self.selectAudioSrt('left')
                        elif inDVDMenu:         self.dvdButton('left')
                        else:                   self.seek(stateRWD)
    
                    elif cmd == 'right':
                        if self.chooseAudioSrt: self.selectAudioSrt('right')
                        elif inDVDMenu:         self.dvdButton('right')
                        else:                   self.seek(stateFWD)

                    elif cmd == 'rotateInc':
                        if not self.chooseAudioSrt and self.isVideo: self.changeRotation(90)

                    elif cmd == 'rotateDec':
                        if not self.chooseAudioSrt and self.isVideo: self.changeRotation(-90)
                        
                else:
                    if cmd in ('right','left'):
                        if not self.chooseAudioSrt:
                            if not inDVDMenu:   self.seek(statePLAY)
                
            except :
                print 'ERROR in miniPlayer CMD %s'%cmd
                print traceback.format_exc() 

    #---------------------------------------------------------------------------------------
    stopEvent = win32event.CreateEvent(None, 0, 0, None)
    seekEvent = win32event.CreateEvent(None, 0, 0, None)

    WAIT_stopEvent = win32event.WAIT_OBJECT_0
    WAIT_graphEvent = win32event.WAIT_OBJECT_0 +1
    WAIT_seekEvent = win32event.WAIT_OBJECT_0 +2
    WAIT_msgEvent =  win32event.WAIT_OBJECT_0 +3

    def pumpMessages(self):
        print 'Msg Pump begin'
        while True:
            rc = win32event.MsgWaitForMultipleObjects((self.stopEvent, self.graphEvent,self.seekEvent),0,-1,win32event.QS_ALLINPUT)
            if rc == self.WAIT_stopEvent:
                break 

            elif rc == self.WAIT_graphEvent:  
                self.handleGraphEvent() 

            elif rc == self.WAIT_seekEvent:
                self.doSeek(self.pos) 

            elif rc == self.WAIT_msgEvent:
                if win32gui.PumpWaitingMessages() :
                    break

            else: raise RuntimeError("unexpected MsgWaitForMultipleObjects return value")
        print 'Msg Pump end'

    def runMsgLoop(self, movieEndedCallback): # Need to be called from the same thread than the one that buildFiltergraph()
        self.graphEvent = self.event.GetEventHandle()
        
        while True:
            self.pumpMessages() #win32gui.PumpMessages()

            self.releaseFilterGraph() # Need to be called from the same thread than the one that buildFiltergraph()

            viewed, favorite, angle, seek = self.endArgs
            if seek is not None:
                self.seekInPlaylist(seek) # Need to be called from the same thread than the one that buildFiltergraph()
            else:
                movieEndedCallback(viewed, favorite, angle)
                break

    def stopMsgLoop(self, viewed, favorite, angle, seek):
        self.endArgs = viewed, favorite, angle, seek
        win32event.SetEvent(self.stopEvent) #exit pumpMessages

    #---------------------------------------------------------------------------------------
    def addToROT(self):
        context = POINTER(IBindCtx)()
        _ole32.CreateBindCtx(0, byref(context))
        self.rot = context.GetRunningObjectTable()

        moniker = POINTER(IMoniker)()
        _ole32.CreateItemMoniker(LPCOLESTR('!'), LPCOLESTR(self.filterGraphName), byref(moniker))
        self.regID = self.rot.Register(ROTFLAGS_REGISTRATIONKEEPSALIVE, self.fg, moniker)

    def delFromROT(self):
        self.rot.Revoke(self.regID)
        del self.rot

# if UsePSYCO:
#     import psyco
#     psyco.bind(MiniPlayer)
       
#---------------------------------------------------------------------------------------
#---------------------------------------------------------------------------------------
if __name__ == "__main__":

    import cPickle as pickle
    import win32api, win32con, win32gui

    #---------------------------------------------------------------------------------------
    class MiniWindow:
        def __init__(self, x,y, w,h):
            self.player = None
            self.size = w,h
            
            win32gui.InitCommonControls()
            self.hinst = win32api.GetModuleHandle(None)
            className = 'miniPlayer'
            
            wc = win32gui.WNDCLASS()
            wc.style = 0#win32con.CS_HREDRAW | win32con.CS_VREDRAW 
            wc.lpfnWndProc = { win32con.WM_DESTROY: self.onDestroy, win32con.WM_KEYDOWN : self.onkeyDown, win32con.WM_KEYUP : self.onkeyUp}
            wc.lpszClassName = className
            try : win32gui.RegisterClass(wc)
            except: pass

            style =  win32con.WS_POPUP 
            styleEx = win32con.WS_EX_TOOLWINDOW 
            self.hwnd = win32gui.CreateWindowEx(styleEx,className,"miniPlayer", style, x,y, w,h, None, None, self.hinst, None)

            win32gui.UpdateWindow(self.hwnd)
            win32gui.SetWindowPos(self.hwnd, win32con.HWND_TOPMOST, 0,0,0,0, win32con.SWP_NOSIZE|win32con.SWP_NOMOVE|win32con.SWP_SHOWWINDOW)
            win32gui.ShowWindow(self.hwnd, win32con.SW_SHOW)

        def movieEnded(self,viewed,favorite,angle):
            print '-----------------'
            print viewed
            print favorite
            print angle
            if favorite is not None:
                with open(favoriteFile,'wb') as f:
                    pickle.dump(favorite, f, -1)

        def onDestroy(self, hwnd, message, wparam, lparam):
            win32gui.PostQuitMessage(0)
            return True

        def onkeyDown(self,hwnd, msg, wparam, lparam):
            self.handleKeys(wparam, True)
            return True

        def onkeyUp(self,hwnd, msg, wparam, lparam):
            self.handleKeys(wparam, False)
            return True
        
        def getHwnd(self):
            return self.hwnd

        def getSize(self):
            return self.size

        VK_LEFT, VK_UP, VK_RIGHT, VK_DOWN, VK_SPACE, VK_ESCAPE, VK_BACK, VK_NUMPAD0, VK_DELETE = 0x25, 0x26, 0x27, 0x28, 0x20, 0x1B, 0x08, 0x60, 0x2e

        cmds = {VK_SPACE:'select', VK_BACK:'back', VK_ESCAPE : 'close' , VK_DELETE : 'close', ord('S'):'lang', VK_NUMPAD0:'lang',
                VK_UP:'up', VK_DOWN:'down', VK_LEFT: 'left', VK_RIGHT :'right'}

        def handleKeys(self, key, pressed):
            if self.player is not None:
                cmd = self.cmds.get(key, None)
                if key is not None:
                    self.player.handleCommands(cmd, pressed)

        def setPlayer(self, player):
            self.player = player

    def mainLoop():
        win32gui.PumpMessages()

    #---------------------------------------------------------------------------------------        

    PLAYFAVORITE = False
    favoriteFile = 'favorite.miniplayer'

    X,Y = 0,0
    W,H = 1920,1200

    window = MiniWindow(X, Y, W, H)
    
    player = MiniPlayer(window)
    window.setPlayer(player)
    
    favorite = None
    if PLAYFAVORITE :
        with open(favoriteFile, 'rb') as f:
            favorite = pickle.load(f)

    movieFile = 'f:/movie/automatic/28 Jours plus tard 2002 1080p FR EN x264 AC3-mHDgz.mkv' #['movie1.avi', 'movie2.mkv'] # or a string if only one file in the playlist

    player.playMovie(favorite, movieFile,0)
    player.runMsgLoop(window.movieEnded) #need to be called in the same thread as playMovie !
    print 'EXIT'
    
