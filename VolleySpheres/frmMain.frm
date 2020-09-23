VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "3D DirectX Engine - George Papadopoulos"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   1680
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DX As DirectX8 'The master Object, everything comes from here
Dim D3D As Direct3D8 'This controls all things 3D
Dim D3DDevice As Direct3DDevice8 'This actually represents the hardware doing the rendering
Dim bRunning As Boolean 'Controls whether the program is running or not...

'This is the Flexible-Vertex-Format description for a 2D vertex (Transformed and Lit)
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'This structure describes a transformed and lit vertex - it's identical to the DirectX7 type "D3DTLVERTEX"
Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Dim p2V(0 To 3) As TLVERTEX '//We're going to have two squares - one with colour, the other without
Dim p1V(0 To 3) As TLVERTEX
Dim ballV(0 To 3) As TLVERTEX '//This is going to be our transparent part - it'll follow the mouse...
Dim netV(0 To 3) As TLVERTEX
Dim lostmsgV(0 To 3) As TLVERTEX
Dim backV(0 To 3) As TLVERTEX
Dim ballposV(0 To 3) As TLVERTEX
Dim p1shadV(0 To 3) As TLVERTEX
Dim p2shadV(0 To 3) As TLVERTEX
Dim ballshadV(0 To 3) As TLVERTEX
Dim ballMaskV(0 To 3) As TLVERTEX

'//NEW TEXTURING STUFF
Dim D3DX As D3DX8 '//A helper library

'##### TEXTURES
Dim P1Texture As Direct3DTexture8
Dim P2Texture As Direct3DTexture8 '//This texture will have transparency information encoded into it....
Dim BallTexture As Direct3DTexture8
Dim NetTexture As Direct3DTexture8
Dim LostP1MsgTexture As Direct3DTexture8
Dim BackTexture As Direct3DTexture8
Dim BallposTexture As Direct3DTexture8
Dim ShadTexture As Direct3DTexture8
Dim BallMaskTexture As Direct3DTexture8

Dim MainFont As D3DXFont
Dim MainFontDesc As IFont
Dim fnt As New StdFont
Dim Text1Rect As RECT
Dim Text2Rect As RECT

'// Initialise : This procedure kick starts the whole process.
'// It'll return true for success, false if there was an error.
Public Function Initialise() As Boolean
On Error GoTo ErrHandler:

Dim DispMode As D3DDISPLAYMODE '//Describes our Display Mode
Dim D3DWindow As D3DPRESENT_PARAMETERS '//Describes our Viewport
Dim ColorKeyVal As Long '//What colour becomes transparent...

Set DX = New DirectX8  '//Create our Master Object
Set D3D = DX.Direct3DCreate() '//Make our Master Object create the Direct3D Interface
Set D3DX = New D3DX8 '//Create our helper library...

'//We're going to use Fullscreen mode because I prefer it to windowed mode :)
Dim maxx, maxy
maxx = 640
maxy = 480
'DispMode.Format = D3DFMT_X8R8G8B8
DispMode.Format = D3DFMT_R5G6B5 'If this mode doesn't work try the commented one above...
DispMode.Width = maxx
DispMode.Height = maxy

refs = (Val(Form2.Check3.Value))
If refs = 1 Then refs = 0 Else refs = 1
D3DWindow.Windowed = refs
D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
D3DWindow.BackBufferCount = 1 '//1 backbuffer only
D3DWindow.BackBufferFormat = DispMode.Format 'What we specified earlier
D3DWindow.BackBufferHeight = maxy
D3DWindow.BackBufferWidth = maxx
D3DWindow.hDeviceWindow = frmmain.hWnd

'//This line creates a device that uses a hardware device if possible; software vertex processing and uses the form as it's target
'//See the lesson text for more information on this line...
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmmain.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                        D3DWindow)

'//Set the vertex shader to use our vertex format
D3DDevice.SetVertexShader FVF

'//Transformed and lit vertices dont need lighting
'   so we disable it...
D3DDevice.SetRenderState D3DRS_LIGHTING, False

D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True


'ColorKeyVal = &HFF000000 '//Black
'ColorKeyVal = &HFFFF0000 '//Red
ColorKeyVal = &HFF00FF00 '//Green
'ColorKeyVal = &HFF0000FF '//Blue
'ColorKeyVal = &HFFFF00FF '//Magenta
'ColorKeyVal = &HFFFFFF00 '//Yellow
'ColorKeyVal = &HFF00FFFF '//Cyan
'ColorKeyVal = &HFFFFFFFF '//White


'//We now want to load our texture;
Set P1Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat6.x", 276, 276, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
Set P2Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat7.x", 210, 211, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
Set BallTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat2.x", 128, 128, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
Set NetTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat5.x", 36, 225, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
Set LostP1MsgTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat3.x", 130, 333, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
Set BackTexture = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\Data\dat9.x")
Set BallposTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat1.x", 25, 46, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
Set ShadTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat8.x", 229, 114, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)
Set BallMaskTexture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Data\dat10.x", 128, 128, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, ColorKeyVal, ByVal 0, ByVal 0)


fnt.Name = "Verdana"
fnt.Size = 18
fnt.Bold = True
Set MainFontDesc = fnt
Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)



'//We can only continue if Initialise Geometry succeeds;
'   If it doesn't we'll fail this call as well...
If InitialiseGeometry() = True Then
    Initialise = True '//We succeeded
    Exit Function
End If


ErrHandler:
'//We failed; for now we wont worry about why.
Debug.Print "Error Number Returned: " & err.Number
Initialise = False
End Function

Public Sub Render()
On Error Resume Next
'//1. We need to clear the render device before we can draw anything
'       This must always happen before you start rendering stuff...
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0 '//Clear the screen black

'//2. Rendering the graphics...

D3DDevice.BeginScene
    'All rendering calls go between these two lines
    
    'Backround
    D3DDevice.SetTexture 0, BackTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, backV(0), Len(backV(0))
    
    'Shadows
    D3DDevice.SetTexture 0, ShadTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ballshadV(0), Len(ballshadV(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, p1shadV(0), Len(p1shadV(0))
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, p2shadV(0), Len(p2shadV(0))
            
    
    D3DDevice.SetTexture 0, P1Texture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, p1V(0), Len(p1V(0))
    
    D3DDevice.SetTexture 0, P2Texture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, p2V(0), Len(p2V(0))
    
    D3DDevice.SetTexture 0, BallTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ballV(0), Len(ballV(0))
    
    D3DDevice.SetTexture 0, BallMaskTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ballMaskV(0), Len(ballMaskV(0))
        
    D3DDevice.SetTexture 0, NetTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, netV(0), Len(netV(0))
    
    D3DDevice.SetTexture 0, BallposTexture
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, ballposV(0), Len(ballposV(0))
        

    If Form1.lostmsg.Visible = True And Form1.lostmsg.Tag = "show" Then
        D3DDevice.SetTexture 0, LostP1MsgTexture 'Else D3DDevice.SetTexture 0, LostP2MsgTexture
        
        D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, lostmsgV(0), Len(netV(0))
    End If
        

    If Form2.gameType = "sp" Then
        D3DX.DrawText MainFont, &HFFFFFF00, Form2.Text10 & " : " & Form1.Score1.Caption, Text1Rect, DT_TOP Or DT_CENTER
        D3DX.DrawText MainFont, &HFF00FF00, Form2.Text11 & " : " & Form1.score2.Caption, Text2Rect, DT_TOP Or DT_CENTER
    Else
        If Form2.Atri = "s" Then
            D3DX.DrawText MainFont, &HFFFFFF00, MyNick & " : " & Form1.Score1.Caption, Text1Rect, DT_TOP Or DT_CENTER
            D3DX.DrawText MainFont, &HFF00FF00, OtherNick & " : " & Form1.score2.Caption, Text2Rect, DT_TOP Or DT_CENTER
        Else
            D3DX.DrawText MainFont, &HFFFFFF00, OtherNick & " : " & Form1.Score1.Caption, Text1Rect, DT_TOP Or DT_CENTER
            D3DX.DrawText MainFont, &HFF00FF00, MyNick & " : " & Form1.score2.Caption, Text2Rect, DT_TOP Or DT_CENTER
       End If
    End If


    
D3DDevice.EndScene

'//3. Update the frame to the screen...
'       This is the same as the Primary.Flip method as used in DirectX 7
'       These values below should work for almost all cases...
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Sub Form_Click()
bRunning = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
DoEvents
If KeyCode = vbKeyEscape Then End
Call Form1.Form_KeyDown(KeyCode, Shift)
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
DoEvents
Call Form1.Form_KeyUp(KeyCode, Shift)
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
End Sub





Public Function InitialiseGeometry() As Boolean

    
On Error GoTo BailOut: '//Setup our Error handler



screenx = (Screen.TwipsPerPixelX + 1.1)
screeny = (Screen.TwipsPerPixelY + 1.1)


Dim p1left As Single
Dim p1top As Single
Dim p1w As Single
Dim p1h As Single

p1left = Form1.p1.Left / screenx
p1top = Form1.p1.Top / screeny
p1w = Form1.p1.Width / screenx
p1h = Form1.p1.Height / screeny

p1V(0) = CreateTLVertex(p1left, p1top, 0, 1, RGB(255, 255, 255), 0, 0, 0)
p1V(1) = CreateTLVertex(p1left + p1w, p1top, 0, 1, RGB(255, 255, 255), 0, 1, 0)
p1V(2) = CreateTLVertex(p1left, p1top + p1h, 0, 1, RGB(255, 255, 255), 0, 0, 1)
p1V(3) = CreateTLVertex(p1left + p1w, p1top + p1h, 0, 1, RGB(255, 255, 255), 0, 1, 1)



Dim p2left As Single
Dim p2top As Single
Dim p2w As Single
Dim p2h As Single

p2left = Form1.p2.Left / screenx
p2top = Form1.p2.Top / screeny
p2w = Form1.p2.Width / screenx
p2h = Form1.p2.Height / screeny

p2V(0) = CreateTLVertex(p2left, p2top, 0, 1, RGB(255, 255, 255), 0, 0, 0)
p2V(1) = CreateTLVertex(p2left + p2w, p2top, 0, 1, RGB(255, 255, 255), 0, 1, 0)
p2V(2) = CreateTLVertex(p2left, p2top + p2h, 0, 1, RGB(255, 255, 255), 0, 0, 1)
p2V(3) = CreateTLVertex(p2left + p2w, p2top + p2h, 0, 1, RGB(255, 255, 255), 0, 1, 1)




Dim ballleft As Single
Dim balltop As Single
Dim ballw As Single
Dim ballh As Single

ballleft = Form1.ball.Left / screenx
balltop = Form1.ball.Top / screeny
ballw = Form1.ball.Width / screenx
ballh = Form1.ball.Height / screeny

Dim pi As Double

pi = 3.14159265358979

Dim x1 As Single, y1 As Single
y1 = (balltop + (ballh / 2)) - (Cos(BallData.Angle) * (ballh / 2))
x1 = (ballleft + (ballw / 2)) - (Sin(BallData.Angle) * (ballw / 2))

Dim x2 As Single, y2 As Single
y2 = (balltop + (ballh / 2)) - (Cos(BallData.Angle - pi / 2) * (ballh / 2))
x2 = (ballleft + (ballw / 2)) - (Sin(BallData.Angle - pi / 2) * (ballw / 2))

Dim x3 As Single, y3 As Single
y3 = (balltop + (ballh / 2)) - (Cos(BallData.Angle - pi) * (ballh / 2))
x3 = (ballleft + (ballw / 2)) - (Sin(BallData.Angle - pi) * (ballw / 2))


Dim x4 As Single, y4 As Single
y4 = (balltop + (ballh / 2)) - (Cos(BallData.Angle + pi / 2) * (ballh / 2))
x4 = (ballleft + (ballw / 2)) - (Sin(BallData.Angle + pi / 2) * (ballw / 2))





'*** FAKES

Dim Fx1 As Single, Fy1 As Single
Fy1 = (balltop + (ballh / 2)) - (Cos(tuy) * (ballh / 2))
Fx1 = (ballleft + (ballw / 2)) - (Sin(tuy) * (ballw / 2))

Dim Fx2 As Single, Fy2 As Single
Fy2 = (balltop + (ballh / 2)) - (Cos(tuy - pi / 2) * (ballh / 2))
Fx2 = (ballleft + (ballw / 2)) - (Sin(tuy - pi / 2) * (ballw / 2))

Dim Fx3 As Single, Fy3 As Single
Fy3 = (balltop + (ballh / 2)) - (Cos(tuy - pi) * (ballh / 2))
Fx3 = (ballleft + (ballw / 2)) - (Sin(tuy - pi) * (ballw / 2))


Dim Fx4 As Single, Fy4 As Single
Fy4 = (balltop + (ballh / 2)) - (Cos(tuy + pi / 2) * (ballh / 2))
Fx4 = (ballleft + (ballw / 2)) - (Sin(tuy + pi / 2) * (ballw / 2))





ballV(0) = CreateTLVertex(Fx1, Fy1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
ballV(1) = CreateTLVertex(Fx2, Fy2, 0, 1, RGB(255, 255, 255), 0, 1, 0)
ballV(2) = CreateTLVertex(Fx4, Fy4, 0, 1, RGB(255, 255, 255), 0, 0, 1)
ballV(3) = CreateTLVertex(Fx3, Fy3, 0, 1, RGB(255, 255, 255), 0, 1, 1)

ballMaskV(0) = CreateTLVertex(x1, y1, 0, 1, RGB(255, 255, 255), 0, 0, 0)
ballMaskV(1) = CreateTLVertex(x2, y2, 0, 1, RGB(255, 255, 255), 0, 1, 0)
ballMaskV(2) = CreateTLVertex(x4, y4, 0, 1, RGB(255, 255, 255), 0, 0, 1)
ballMaskV(3) = CreateTLVertex(x3, y3, 0, 1, RGB(255, 255, 255), 0, 1, 1)


BallData.Angle = BallData.Angle + BallData.RotSpeed
If BallData.Angle > pi * 2 Then BallData.Angle = 0



Dim netleft As Single
Dim nettop As Single
Dim netw As Single
Dim neth As Single

netleft = Form1.net.Left / screenx
nettop = Form1.net.Top / screeny
netw = Form1.net.Width / screenx
neth = Form1.net.Height / screeny

netV(0) = CreateTLVertex(netleft, nettop, 0, 1, RGB(255, 255, 255), 0, 0, 0)
netV(1) = CreateTLVertex(netleft + netw, nettop, 0, 1, RGB(255, 255, 255), 0, 1, 0)
netV(2) = CreateTLVertex(netleft, nettop + neth, 0, 1, RGB(255, 255, 255), 0, 0, 1)
netV(3) = CreateTLVertex(netleft + netw, nettop + neth, 0, 1, RGB(255, 255, 255), 0, 1, 1)




Dim lostmsgleft As Single
Dim lostmsgtop As Single
Dim lostmsgw As Single
Dim lostmsgh As Single

lostmsgleft = Form1.lostmsg.Left / screenx
lostmsgtop = Form1.lostmsg.Top / screeny
lostmsgw = Form1.lostmsg.Width / screenx
lostmsgh = Form1.lostmsg.Height / screeny

lostmsgV(0) = CreateTLVertex(lostmsgleft, lostmsgtop, 0, 1, RGB(255, 255, 255), 0, 0, 0)
lostmsgV(1) = CreateTLVertex(lostmsgleft + lostmsgw, lostmsgtop, 0, 1, RGB(255, 255, 255), 0, 1, 0)
lostmsgV(2) = CreateTLVertex(lostmsgleft, lostmsgtop + lostmsgh, 0, 1, RGB(255, 255, 255), 0, 0, 1)
lostmsgV(3) = CreateTLVertex(lostmsgleft + lostmsgw, lostmsgtop + lostmsgh, 0, 1, RGB(255, 255, 255), 0, 1, 1)




Dim backleft As Single
Dim backtop As Single
Dim backw As Single
Dim backh As Single

backleft = Form1.back.Left / screenx
backtop = Form1.back.Top / screeny
backw = Form1.back.Width / screenx
backh = 7695 / screeny

backV(0) = CreateTLVertex(backleft, backtop, 0, 1, RGB(255, 255, 255), 0, 0, 0)
backV(1) = CreateTLVertex(backleft + backw, backtop, 0, 1, RGB(255, 255, 255), 0, 1, 0)
backV(2) = CreateTLVertex(backleft, backtop + backh, 0, 1, RGB(255, 255, 255), 0, 0, 1)
backV(3) = CreateTLVertex(backleft + backw, backtop + backh, 0, 1, RGB(255, 255, 255), 0, 1, 1)







Dim ballposleft As Single
Dim ballpostop As Single
Dim ballposw As Single
Dim ballposh As Single

ballposleft = Form1.ballpos.Left / screenx
ballpostop = Form1.ballpos.Top / screeny
ballposw = Form1.ballpos.Width / screenx
ballposh = Form1.ballpos.Height / screeny

ballposV(0) = CreateTLVertex(ballposleft, ballpostop, 0, 1, RGB(255, 255, 255), 0, 0, 0)
ballposV(1) = CreateTLVertex(ballposleft + ballposw, ballpostop, 0, 1, RGB(255, 255, 255), 0, 1, 0)
ballposV(2) = CreateTLVertex(ballposleft, ballpostop + ballposh, 0, 1, RGB(255, 255, 255), 0, 0, 1)
ballposV(3) = CreateTLVertex(ballposleft + ballposw, ballpostop + ballposh, 0, 1, RGB(255, 255, 255), 0, 1, 1)





Dim ballshadleft As Single
Dim ballshadtop As Single
Dim ballshadw As Single
Dim ballshadh As Single

ballshadleft = Form1.ballshad.Left / screenx
ballshadtop = Form1.ballshad.Top / screeny
ballshadw = Form1.ballshad.Width / screenx
ballshadh = Form1.ballshad.Height / screeny

ballshadV(0) = CreateTLVertex(ballshadleft, ballshadtop, 0, 1, RGB(255, 255, 255), 0, 0, 0)
ballshadV(1) = CreateTLVertex(ballshadleft + ballshadw, ballshadtop, 0, 1, RGB(255, 255, 255), 0, 1, 0)
ballshadV(2) = CreateTLVertex(ballshadleft, ballshadtop + ballshadh, 0, 1, RGB(255, 255, 255), 0, 0, 1)
ballshadV(3) = CreateTLVertex(ballshadleft + ballshadw, ballshadtop + ballshadh, 0, 1, RGB(255, 255, 255), 0, 1, 1)





Dim p1shadleft As Single
Dim p1shadtop As Single
Dim p1shadw As Single
Dim p1shadh As Single

p1shadleft = Form1.p1shad.Left / screenx
p1shadtop = Form1.p1shad.Top / screeny
p1shadw = Form1.p1shad.Width / screenx
p1shadh = Form1.p1shad.Height / screeny

p1shadV(0) = CreateTLVertex(p1shadleft, p1shadtop, 0, 1, RGB(255, 255, 255), 0, 0, 0)
p1shadV(1) = CreateTLVertex(p1shadleft + p1shadw, p1shadtop, 0, 1, RGB(255, 255, 255), 0, 1, 0)
p1shadV(2) = CreateTLVertex(p1shadleft, p1shadtop + p1shadh, 0, 1, RGB(255, 255, 255), 0, 0, 1)
p1shadV(3) = CreateTLVertex(p1shadleft + p1shadw, p1shadtop + p1shadh, 0, 1, RGB(255, 255, 255), 0, 1, 1)






Dim p2shadleft As Single
Dim p2shadtop As Single
Dim p2shadw As Single
Dim p2shadh As Single

p2shadleft = Form1.p2shad.Left / screenx
p2shadtop = Form1.p2shad.Top / screeny
p2shadw = Form1.p2shad.Width / screenx
p2shadh = Form1.p2shad.Height / screeny

p2shadV(0) = CreateTLVertex(p2shadleft, p2shadtop, 100, 1, RGB(255, 255, 255), 0, 0, 0)
p2shadV(1) = CreateTLVertex(p2shadleft + p2shadw, p2shadtop, 0, 1, RGB(255, 255, 255), 0, 1, 0)
p2shadV(2) = CreateTLVertex(p2shadleft, p2shadtop + p2shadh, 0, 1, RGB(255, 255, 255), 0, 0, 1)
p2shadV(3) = CreateTLVertex(p2shadleft + p2shadw, p2shadtop + p2shadh, 0, 1, RGB(255, 255, 255), 0, 1, 1)





Text1Rect.Left = Form1.Score1.Left / screenx
Text1Rect.Top = Form1.Score1.Top / screeny
Text1Rect.Right = Text1Rect.Left + (Form1.Score1.Width / screenx)
Text1Rect.bottom = Text1Rect.Top + (Form1.Score1.Height / screeny)

Text2Rect.Left = Form1.score2.Left / screenx
Text2Rect.Top = Form1.score2.Top / screeny
Text2Rect.Right = Text2Rect.Left + (Form1.score2.Width / screenx)
Text2Rect.bottom = Text2Rect.Top + (Form1.score2.Height / screeny)




InitialiseGeometry = True
Exit Function
BailOut:
InitialiseGeometry = False
End Function

'//This is just a simple wrapper function that makes filling the structures much much easier...
Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, specular As Long, tu As Single, tv As Single) As TLVERTEX
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
CreateTLVertex.X = X
CreateTLVertex.Y = Y
CreateTLVertex.Z = Z
CreateTLVertex.rhw = rhw
CreateTLVertex.color = color
CreateTLVertex.specular = specular
CreateTLVertex.tu = tu
CreateTLVertex.tv = tv
End Function




Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub



Private Sub Cleanall()
On Error Resume Next 'If the objects were never created;
'                               (the initialisation failed) we might get an
'                               error when freeing them... which we need to
'                               handle, but as we're closing anyway...
Set D3DDevice = Nothing
Set D3D = Nothing
Set DX = Nothing
'Debug.Print "All Objects Destroyed"

End Sub

Private Sub Form_Unload(Cancel As Integer)
Cleanall
Form2.Show
Form2.Timer2.Enabled = True
DoEvents
DoEvents
DoEvents
Unload Form2
DoEvents
DoEvents
Unload Form1
End
End Sub

Private Sub Timer1_Timer()
Me.Show '//Make sure our window is visible

bRunning = Initialise()
Form1.Loadup
Timer1.Enabled = False
End Sub
