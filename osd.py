import comtypes.client
import comtypes.server.localserver
from comtypes import IUnknown,GUID,STDMETHOD,HRESULT,c_longlong,c_void_p

import traceback
import time
import math
import threading
import datetime
from string import Template

#-------------------------------------------------------------------------------------------
from directx.types import *
from directx.d3d import IDirect3DVertexBuffer9,IDirect3DIndexBuffer9,IDirect3DTexture9,IDirect3DPixelShader9,IDirect3DStateBlock9
from directx.d3dx import d3dxdll, IDirect3DDevice9, ID3DXFont,ID3DXBuffer,ID3DXSprite,ID3DXConstantTable

class Vertex(Structure):
    _fields_ = [
        ('x', c_float),
        ('y', c_float),
        ('z', c_float),
        ('rhw', c_float),     
        ('diffuse', DWORD),
    ]

VERTEXFVF = D3DFVF.XYZRHW | D3DFVF.DIFFUSE

#-------------------------------------------------------------------------------------------
class Color():
    def __init__(self,a,r,g,b):
        self.a,self.r,self.g,self.b = a,r,g,b

    def getDWORD(self):
        return (int(self.a)<<24)+(int(self.r)<<16)+(int(self.g)<<8)+int(self.b)

    def getFloat4String(self):
        return '%.3ff,%.3ff,%.3ff,%.3ff'%(self.r/255.,self.g/255.,self.b/255.,self.a/255.)

    def interColor(self,c1,c2,t):
        self.r = t*c1.r+(1-t)*c2.r
        self.g = t*c1.g+(1-t)*c2.g
        self.b = t*c1.b+(1-t)*c2.b

#-------------------------------------------------------------------------------------------
ColorWhite = Color(0,0xf8,0xf8,0xff)
ColorFGd = Color(0,47,120,173)
ColorBGd = Color(0,24,60,86)
ColorCursor = Color(0,34,70,96)
ColorTemp = Color(0,0,0,0)

#-------------------------------------------------------------------------------------------
REFERENCE_TIME = c_longlong

class IOsdServices(IUnknown):
    _iid_ = GUID('{3AE03A88-F613-4BBA-AD3E-EE236976BF9A}')
    _methods_ =     [ STDMETHOD(HRESULT, 'dummy',[]) ]*3 +   [ STDMETHOD(HRESULT, 'OsdRedrawFrame',[]) ]

class IOsdRenderCallback(IUnknown):
    _iid_ = GUID('{57FBF6DC-3E5F-4641-935A-CB62F00C9958}')
    _methods_ = [
                    STDMETHOD(HRESULT, 'SetDevice',[POINTER(IDirect3DDevice9)]),
                    STDMETHOD(HRESULT, 'ClearBackground',[LPCSTR,REFERENCE_TIME,POINTER(RECT),POINTER(RECT)]),
                    STDMETHOD(HRESULT, 'RenderOsd',[LPCSTR,REFERENCE_TIME,POINTER(RECT),POINTER(RECT)]),
                    STDMETHOD(HRESULT, 'showButton',[LPCSTR,LPCSTR,c_float,INT]),
                    STDMETHOD(HRESULT, 'hideButton',[LPCSTR,INT]),
                    STDMETHOD(HRESULT, 'showCombo',[LPCSTR,LPCSTR,LPCSTR,INT,INT]),
                    STDMETHOD(HRESULT, 'hideCombo',[LPCSTR]),
                    STDMETHOD(HRESULT, 'selectCombo',[LPCSTR,INT,INT]),
                    STDMETHOD(HRESULT, 'init',[UINT,UINT,POINTER(IOsdServices),UINT,UINT]),
                    STDMETHOD(HRESULT, 'forceRedraw',[BOOL]),
                    STDMETHOD(HRESULT, 'run',[]),
                    STDMETHOD(HRESULT, 'stop',[]),
                    STDMETHOD(HRESULT, 'setState',[INT,FLOAT,FLOAT]),
                    STDMETHOD(FLOAT,   'getCurrentPos',[]),
                    STDMETHOD(BOOL,    'isInitialized',[])                    
                ]

#-------------------------------------------------------------------------------------------
osdPlay = 'osd.play'
osdTimeline = 'osd.timeline'
osdVol = 'osd.vol'
osdBattery = 'osd.battery'
osdLocalTime = 'osd.localTime'
osdAudio = 'osd.audio'
osdSrt = 'osd.srt'

#-------------------------------------------------------------------------------------------
FONTHEIGHT = 55
FONTHEIGHT2 = 45

#-------------------------------------------------------------------------------------------
alphaMINtoShow = .05
SMOOTH = .9

class Button(object):
    def __init__(self,osd,name,pos,size,initVbuffer=True):
        self.device = osd.device
        self.font = osd.font
        self.shaderOsd = osd.shaderOsd

        self.name = name
        self.pos = pos
        self.size = size
        self.label = ''
        self.value = 100

        x,y = self.pos
        w,h = self.size
        self.rectLabel = RECT(x,y,x+w,y+h)

        self.radius = 12

        self.vbuffer = None
        if initVbuffer:
            quad = self.getQuad(x,y,w,h,0)
            self.vbuffer = createVBuffer(self.device,self.getVertex(quad))

        self.alpha = self.talpha = 0
        self.timeToFade = None
        self.drawing = False

    def getQuad(self,x,y,w,h,color):
        return ((x, y, color),(x+w, y, color), (x, y+h, color), (x+w, y+h, color))

    def getVertex(self,points2d):
        vertexes = (Vertex * len(points2d))()
        for i,(x,y,c) in enumerate(points2d):
            vertexes[i] = Vertex(x,y,0,1,c)

        return vertexes

    def release(self):
        if self.vbuffer is not None:
            del self.vbuffer
        
        del self.device
        del self.font
        del self.shaderOsd
    
    def fade(self,smooth):
        if self.timeToFade is not None:
            if time.clock()>= self.timeToFade:
                self.talpha = 0
                self.timeToFade = None

        self.alpha = (1-smooth)*self.talpha+smooth*self.alpha
        if self.alpha<alphaMINtoShow :
            self.alpha = 0
            self.drawing = False
            #print 'stop Drawing %s'%self.name
    
    def draw(self):
        self.fade(SMOOTH)

        (x,y),(w,h) = self.pos,self.size

        shader,var = self.shaderOsd
        var.SetVector(self.device,'RECT',D3DXVECTOR4(x,y,w,h))
        var.SetVector(self.device,'constants',D3DXVECTOR4(self.alpha,self.radius,self.value,0))
        
        self.device.SetPixelShader(shader)

        self.device.SetStreamSource(0, self.vbuffer, 0, sizeof(Vertex))
        self.device.DrawPrimitive(D3DPT.TRIANGLESTRIP, 0, 2)

        ColorWhite.a = self.alpha*255
        colorTxt = ColorWhite.getDWORD()

        self.font.DrawTextW(None, self.label, -1, self.rectLabel, D3DXFONT.CENTER | D3DXFONT.VCENTER | D3DXFONT.WORDBREAK,colorTxt)

        #print 'render %s'%self.name

    def show(self,label,value,waitBeforeFade):
        if label != '': self.label = label
        if value != -1: self.value = value

        self.talpha = 1.0
        self.drawing = True
        self.timeToFade = time.clock()+waitBeforeFade if waitBeforeFade != -1 else None

    def hide(self,waitBeforeFade):
        self.timeToFade = time.clock()+waitBeforeFade

#-------------------------------------------------------------------------------------------
stateFWD = 1
stateRWD = -1
statePAUSE = 0
statePLAY = 2

class Timeline(Button):
    def __init__(self,osd,name,pos,size):
        Button.__init__(self,osd,name,pos,size)
        self.font2 = osd.font2
        self.seekPos = None

    def release(self):
        Button.release(self)
        del self.font2

    TicksInSec = 1000*1000*10. #1 tick = 100 ns
    coefL = 10. #linear ramp up
    coefP =  2.5 #exp ramp up

    def setFileInfo(self,nbFiles,fileIndex):
        self.boundaryLabels = []

        if nbFiles > 1:
            if fileIndex>=1:
                self.boundaryLabels.append(('< CD%d'%(fileIndex),D3DXFONT.LEFT))
            if fileIndex< nbFiles-1:
                self.boundaryLabels.append(('CD%d >'%(fileIndex+2),D3DXFONT.RIGHT))
        
    def setState(self,state,pos,duration):
        #print 'set Timeline State %s %s %s'%(state,pos,duration) 
        self.state = state
        self.statePos = pos if self.seekPos is None else self.seekPos 
        self.duration = duration
        self.stateTime = time.clock()
       
        self.direction = self.state

    def getCurrentPos(self):
        pos = self.statePos
        dt = time.clock()-self.stateTime
        self.seekPos = None

        if self.state in (stateFWD,stateRWD):
            step = self.direction*self.coefL*pow(dt,self.coefP)
            pos += step*self.TicksInSec
            self.seekPos = max(0,pos) #si on revient a 0
        elif self.state == statePLAY:
            pos += dt*self.TicksInSec 
        
        #print 'timeline state %s, pos %s'%(self.state,str(datetime.timedelta(seconds=round(pos/self.TicksInSec))))
        return pos

    def draw(self):
        pos = self.getCurrentPos()
        pos = max(0,min(pos,self.duration))

        self.value = pos/self.duration*100
        
        pos = str(datetime.timedelta(seconds=round(pos/self.TicksInSec)))
        dur = str(datetime.timedelta(seconds=round(self.duration/self.TicksInSec)))
        self.label = '%s/%s'%(pos,dur)
        
        Button.draw(self)

        ColorWhite.a = self.alpha*64
        colorTxt = ColorWhite.getDWORD()
        for label,flag in self.boundaryLabels:
            self.font2.DrawTextW(None, label, -1, self.rectLabel, D3DXFONT.VCENTER | D3DXFONT.WORDBREAK |flag,colorTxt)

#-------------------------------------------------------------------------------------------
class LocalTime(Button):
    def __init__(self,osd,name,pos,size):
        Button.__init__(self,osd,name,pos,size)

    def draw(self):
        #self.label = time.strftime('%Hh%M:%S',time.localtime())
        self.label = '%sh%s:%s'%tuple(time.ctime()[11:19].split(':')) #localtime marche pas !
        Button.draw(self)
        
#-------------------------------------------------------------------------------------------
labelsSEPARATOR = '\x01'
class Combo(Button):
    def __init__(self,osd,name,pos,size):
        Button.__init__(self,osd,name,pos,size,initVbuffer=False)
        self.font2 = osd.font2
        self.vertexes = self.getVertex(((0,0,0),)*4) 

    def release(self):
        Button.release(self)
        del self.font2

    def show(self,label,labels,index,selected):
        self.labels = labels.split(labelsSEPARATOR)
        self.label = label

        (x,y),(w,h) = self.pos,self.size

        self.hb = FONTHEIGHT
        self.h = FONTHEIGHT2+6
        h = len(self.labels)*self.h+self.hb+15
        self.size = w,h

        quad = self.getQuad(x,y,w,self.hb,ColorFGd.getDWORD())
        quad += self.getQuad(x,y+self.hb,w,h-self.hb,ColorBGd.getDWORD())
        self.vbuffer = createVBuffer(self.device,self.getVertex(quad))
        
        self.rectLabel = RECT(x,y,x+w,y+self.hb)
        self.cR = self.h/4
        self.dd = round((self.h-self.cR*2)/2.)

        x += self.h
        y += self.hb+10
        self.rectLabels = [ RECT(x,y+i*self.h,x+w-(self.h+10),y+(i+1)*self.h) for i,label in enumerate(self.labels)]
        self.labels = zip(self.labels,self.rectLabels)

        self.setIndex(index)
        self.posCursor = self.tposCursor

        self.select(selected)
        self.talpha = 1.
        self.drawing = True

    def hide(self):
        self.talpha = 0

    def select(self,selected):
        self.selected = selected
        self.refTime = time.clock()

    def setIndex(self,index):
        self.tposCursor = self.h*index

    def draw(self):
        self.fade(SMOOTH)

        self.posCursor = (1-SMOOTH)*self.tposCursor+SMOOTH*self.posCursor

        (x,y),(w,h) = self.pos,self.size

        shader,var = self.shaderOsd
        self.device.SetPixelShader(shader)

        var.SetVector(self.device,'RECT',D3DXVECTOR4(x,y,w,h))
        var.SetVector(self.device,'constants',D3DXVECTOR4(self.alpha,self.radius,-1,0))

        self.device.SetStreamSource(0, self.vbuffer, 0, sizeof(Vertex))
        self.device.DrawPrimitive(D3DPT.TRIANGLESTRIP, 0, 6)

        #pulsing cursor
        yCur = round(y+self.hb+10+self.posCursor)
        
        if self.selected:
            t = .5*(1+math.sin(2*math.pi*(time.clock()-self.refTime)))
            ColorTemp.interColor(ColorFGd,ColorCursor,t)
            self.setQuadVertex(self.vertexes,x,yCur+5,w,self.h-10,ColorTemp.getDWORD())
            self.device.DrawPrimitiveUP(D3DPT.TRIANGLESTRIP, 2, self.vertexes, sizeof(Vertex))

        #circle
        ColorWhite.a = self.alpha*255
        colorTxt = ColorWhite.getDWORD()

        xx,yy,ww,hh = x+self.dd,yCur+self.dd,self.cR*2,self.cR*2
        
        var.SetVector(self.device,'RECT',D3DXVECTOR4(xx,yy,ww,hh))
        var.SetVector(self.device,'constants',D3DXVECTOR4(self.alpha,self.cR,-1,0))

        self.setQuadVertex(self.vertexes,xx,yy,ww,hh,colorTxt)
        self.device.DrawPrimitiveUP(D3DPT.TRIANGLESTRIP, 2, self.vertexes, sizeof(Vertex))

        #texts
        self.font.DrawTextW(None, self.label, -1, self.rectLabel, D3DXFONT.CENTER | D3DXFONT.VCENTER | D3DXFONT.WORDBREAK,colorTxt)

        for label,rect in self.labels:
            self.font2.DrawTextW(None, label, -1, rect, D3DXFONT.LEFT | D3DXFONT.VCENTER ,colorTxt)

    def setQuadVertex(self,v,x,y,w,h,color):
        v[0].x,v[0].y,v[0].diffuse = x,y,color
        v[1].x,v[1].y,v[1].diffuse = x+w,y,color
        v[2].x,v[2].y,v[2].diffuse = x,y+h,color
        v[3].x,v[3].y,v[3].diffuse = x+w,y+h,color


#-------------------------------------------------------------------------------------------

def createFont(device,h):
    ANTIALIASED_QUALITY = 4
    MWLF_TYPE_TRUETYPE = 7
    weight = 0#MWLF_WEIGHT_BOLD = 700
    font = POINTER(ID3DXFont)()
    
    d3dxdll.D3DXCreateFontW(device, h, 0, weight, 0, BOOL(0), DWORD(0), MWLF_TYPE_TRUETYPE, ANTIALIASED_QUALITY, DWORD(0), LPCWSTR(u'GeosansLight'), byref(font))
    font.PreloadCharacters(32, 127)

    print 'init Font'
    return font


def createVBuffer(device,vertex,typeVertex=Vertex,fvf=VERTEXFVF):
    nb = len(vertex)

    vbuffer = POINTER(IDirect3DVertexBuffer9)()
    device.CreateVertexBuffer(sizeof(typeVertex) * nb, D3DUSAGE.WRITEONLY, fvf, D3DPOOL.DEFAULT, byref(vbuffer), None)

    ptr = c_void_p()
    vbuffer.Lock(0, 0, byref(ptr), 0)
    cdll.msvcrt.memcpy(ptr, vertex, sizeof(typeVertex) * nb)
    vbuffer.Unlock()

    return vbuffer

def createShader(device,shaderDef):
    mainFunc, shader = shaderDef
    #print shader
    
    try:
        code = POINTER(ID3DXBuffer)()
        error = POINTER(ID3DXBuffer)()
        pixelShader = POINTER(IDirect3DPixelShader9)()
        constants =  POINTER(ID3DXConstantTable)()
        
        hr = d3dxdll.D3DXCompileShader(shader,len(shader),None,None,mainFunc,'ps_3_0',0,byref(code),byref(error),byref(constants))
        if hr != 0 : print LPCSTR(error.GetBufferPointer()).value

        device.CreatePixelShader(cast(code.GetBufferPointer(),POINTER(DWORD)),byref(pixelShader))
    except:
        print traceback.format_exc()

    print 'init Shader'
    return pixelShader,constants


def createStateBlock(device):
    stateBlock = POINTER(IDirect3DStateBlock9)()

    device.BeginStateBlock()
    
    device.SetRenderState(D3DRS.ALPHABLENDENABLE,True)
    device.SetRenderState(D3DRS.SRCBLEND,D3DBLEND.SRCALPHA)
    device.SetRenderState(D3DRS.DESTBLEND,D3DBLEND.INVSRCALPHA)
    device.SetFVF(VERTEXFVF)

    device.EndStateBlock(byref(stateBlock))

    print 'init StateBlock'
    return stateBlock

#-------------------------------------------------------------------------------------------
shaderOSD = 'psOSD',Template('''

static float4 ColorFGd = float4($COLORFGD);
static float4 ColorBGd = float4($COLORBGD);

float4 RECT;
float4 constants;

float smoothCorner(float2 pos, float2 c, float radius)
{
    float dist = length(pos-c);
    if (dist>radius)
        return 1.0f-smoothstep(radius,radius+1.5,dist);
    else 
        return 1.0f;
}

float4 psOSD( float4 Color : COLOR0, float2 pos  : VPOS) : COLOR
{

    float4 Out; 
    float2 c;

    static const float X = RECT.x;
    static const float Y = RECT.y;
    static const float W = RECT.z;
    static const float H = RECT.w;
	
    static const float alpha = constants.x;
    static const float radius = constants.y;
    static const float ratio = constants.z;

    if (ratio == -1)
	Out.rgb = Color.rgb;
    else if ( (pos.x-X) < (W*ratio/100))
	Out.rgb = ColorFGd.rgb;

    else 
	Out.rgb = ColorBGd.rgb;

    Out.a = alpha;

    c = float2(X+radius,Y+radius);
    if ((pos.x < c.x) && (pos.y < c.y))
	    Out.a *= smoothCorner(pos,c,radius);

    c = float2(X+W-radius-1,Y+radius);
    if ((pos.x > c.x) && (pos.y < c.y))
	    Out.a *= smoothCorner(pos,c,radius);
    
    c = float2(X+W-radius-1,Y+H-radius-1);
    if ((pos.x > c.x) && (pos.y > c.y))
	    Out.a *= smoothCorner(pos,c,radius);
    
    c = float2(X+radius,Y+H-radius-1);
    if ((pos.x < c.x) && (pos.y > c.y))
	    Out.a *= smoothCorner(pos,c,radius);
    

    return Out;                                
}

''').substitute(COLORFGD=ColorFGd.getFloat4String(),COLORBGD=ColorBGd.getFloat4String())


#-------------------------------------------------------------------------------------------
class OSD(comtypes.COMObject):
    _com_interfaces_ = [IOsdRenderCallback]
    _reg_clsid_ = IOsdRenderCallback._iid_

    _reg_threading_ = "Both"
    _reg_clsctx_ = comtypes.CLSCTX_INPROC_SERVER | comtypes.CLSCTX_LOCAL_SERVER
    _regcls_ = comtypes.server.localserver.REGCLS_MULTIPLEUSE #REGCLS_MULTI_SEPARATE

    def __init__(self):
        self.initlock = threading.Lock()
        self.deviceInitialized = False
        print 'init OSD'

    def initButtons(self):
        dx,dy = 10,10
        w,h = 200,50
        wb = 225
        W,H = self.size
        wl,hl = (W-3*dx)/2,800

        buttons =   (#class        name          pos                 size          
                    (Button,       osdPlay,      (dx,dy),            (w,h)),
                    (Button,       osdVol,       (dx,dy),            (w,h)),
                    (Timeline,     osdTimeline,  (w+dx+2,dy),        (W-2*dx-2-w,h)),
                    (LocalTime,    osdLocalTime, (dx,H-h-dy),        (wb,h)),
                    (Button,       osdBattery,   (W-wb-dx,H-h-dy),   (wb,h)),
                    (Combo,        osdAudio,     (dx,dy),            (wl,hl)),
                    (Combo,        osdSrt,       (wl+2*dx,dy),       (wl,hl))
                    )
        
        #for cls,name,pos,size in buttons: print name
        buttons = [ (name,cls(self,name,pos,size)) for cls,name,pos,size in buttons]
        self.buttons = zip(*buttons)[1] #dans l'ordre de la declaration (pour algo du peintre dans draw car pas de zbuffer)
        self.buttonsByName = dict(buttons)
        
        self.timeline = self.buttonsByName[osdTimeline]
        self.timeline.setFileInfo(self.nbFiles,self.fileIndex)
        print 'init Buttons Done'

    #d3d
    def SetDevice(self,this,device):
        #print 'init d3d Device %s'%device
        with self.initlock:
            if device:
                if not self.deviceInitialized:
                    self.device = device
    
                    self.font = createFont(device,FONTHEIGHT)
                    self.font2 = createFont(device,FONTHEIGHT2)
                    self.shaderOsd = createShader(device,shaderOSD)
                    self.stateBlock = createStateBlock(device)
    
                    self.initButtons()
    
                    self.deviceInitialized = True
                    print 'd3d init done'
            else:
                if self.deviceInitialized:
                    self.deviceInitialized = False
                    del self.font
                    del self.font2
                    del self.shaderOsd
                    del self.stateBlock
                    del self.device
                    
                    for button in self.buttons:
                        button.release()
    
                print 'd3d Ressources Released'

    def ClearBackground(self,this,name,frameStart,fullOutputRect,activeVideoRect):
        if self.forceRedrawFlag:
            return ERROR_EMPTY #madvr low latency mode
    
    def RenderOsd(self,this,name,frameStart,fullOutputRect,activeVideoRect):
        self.stateBlock.Apply()

        for button in self.buttons:
            if button.drawing : 
                button.draw()

        if self.forceRedrawFlag:
            return ERROR_EMPTY #madvr low latency mode

    #initdone ?
    def isInitialized(self,this):
        return self.deviceInitialized
    
    #button
    def showButton(self,this,name,label,value,waitBeforeFade):
        if self.deviceInitialized:
            button = self.buttonsByName[name]
            button.show(label,value,waitBeforeFade)

    def hideButton(self,this,name,waitBeforeFade):
        if self.deviceInitialized:
            button = self.buttonsByName[name]
            button.hide(waitBeforeFade)

    #combo
    def showCombo(self,this,name,label,labels,index,selected):
        if self.deviceInitialized:
            combo = self.buttonsByName[name]
            combo.show(label,labels,index,selected)

    def hideCombo(self,this,name):
        if self.deviceInitialized:
            combo = self.buttonsByName[name]
            combo.hide()

    def selectCombo(self,this,name,index,selected):
        if self.deviceInitialized:
            combo = self.buttonsByName[name]
            combo.select(selected)
            combo.setIndex(index)

    #timeline
    def getCurrentPos(self,this):
        if self.deviceInitialized:
            return self.timeline.getCurrentPos()

    def setState(self,this,state,currentPos,duration):
        if self.deviceInitialized:
            self.timeline.setState(state,currentPos,duration)
    
    #osd loop
    def init(self,this,w,h,osdService,nbFiles,fileIndex):
        self.size = w,h
        self.nbFiles,self.fileIndex = nbFiles,fileIndex
        self.osdService = osdService
        self.forceRedrawFlag = False

    def forceRedraw(self,this,forceRedrawFlag):
        if self.deviceInitialized:
            self.forceRedrawFlag = forceRedrawFlag

    def run(self,this):
        pass #no need with madvr low latency mode since version 0.88.17
        #self.running = True
        #self.runThread = threading.Thread(target=self._run)
        #self.runThread.start()
        
    ''' #no need with madvr low latency mode since version 0.88.17
    def _run(self):
        print '----------\nOSD loop running\n----------'
        while self.running:
            if self.forceRedrawFlag and self.deviceInitialized:
                self.osdService.OsdRedrawFrame()

            time.sleep(0.02)
        print '----------\nOSD loop stopped\n----------'
    '''

    def stop(self,this):
        self.forceRedrawFlag = False
        #self.running = False #no need with madvr low latency mode since version 0.88.17
        #self.runThread.join()
        self.osdService = None

if __name__ == "__main__": #osd.py /regserver
    from comtypes.server.register import UseCommandLine
    UseCommandLine(OSD)
