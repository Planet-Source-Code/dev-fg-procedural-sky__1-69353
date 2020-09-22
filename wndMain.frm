VERSION 5.00
Begin VB.Form wndMain 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15360
   Icon            =   "wndMain.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "wndMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'require variable declaration
Option Explicit

'general directx8 objects
Private objDX As DirectX8
Private objD3D As Direct3D8


'last mouse position
Private msXl As Long
Private msYl As Long

'accurate system timer function
Private Declare Function GetTickCount Lib "kernel32" () As Long
'returns active window handle
Private Declare Function GetForegroundWindow Lib "user32" () As Long
'accurate sleep function
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)



'controls
Private Sub Form_KeyPress(KeyAscii As Integer)

  Select Case UCase(Chr(KeyAscii))
  
    'loop thru different render modes
    Case " "
      swZenith = Not swZenith
  
    Case "W"
      sun_angle = sun_angle - 0.01
      If sun_angle < -2.25 Then sun_angle = -2.25
    
    Case "S"
      sun_angle = sun_angle + 0.01
      If sun_angle > 2.25 Then sun_angle = 2.25
    
    Case "A"
      sun_turbidity = sun_turbidity - 0.01
      If sun_turbidity < 1.7 Then sun_turbidity = 1.7
  
    Case "D"
      sun_turbidity = sun_turbidity + 0.01
      If sun_turbidity > 15 Then sun_turbidity = 15
  
    'exit demo
    Case Chr(27)
      Unload Me
  
  End Select
  
  'recompute sky
  mhSky.skyCompute sun_turbidity, Sin(sun_angle), 0, Cos(sun_angle)

End Sub


'startup
Private Sub Form_Load()

  'show window
  Show
  DoEvents

  'reset mouse coords
  msXl = 0
  msYl = 0

  'variables for fps counter
  Dim fps As Long
  Dim tick As Long
  Dim slow As Boolean
  fps = 0
  tick = 0
  slow = False
  
  'initialize direct3d
  Set objDX = New DirectX8
  Set objD3D = objDX.Direct3DCreate
  
  'prepare configuration for rendering device
  With typConfig
    .AutoDepthStencilFormat = D3DFMT_D16   '16 bit z-buffer with no stencil
    .BackBufferCount = 1                   'single backbuffer only
    .BackBufferFormat = D3DFMT_A8R8G8B8    '32 bit backbuffer with alpha channel
    .BackBufferHeight = ScaleHeight        'current window height
    .BackBufferWidth = ScaleWidth          'current window width
    .EnableAutoDepthStencil = 1            'enable automatic depth testing
    .flags = 0                             'not required
    .FullScreen_PresentationInterval = 0   'not required (fullscreen only)
    .FullScreen_RefreshRateInHz = 0        'not required (fullscreen only)
    .hDeviceWindow = hWnd                  'target window handle
    .MultiSampleType = D3DMULTISAMPLE_NONE 'no anti-aliasing needed
    .SwapEffect = D3DSWAPEFFECT_FLIP       'fastest mode (not compatible with fsaa)
    .Windowed = 1                          'run in window
  End With
  
  'create rendering device
  Set objDev = objD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, typConfig)
  
  'load all resources
  appInitialize

  'start rendering loop
  Do
  
    'process scene frame
    appDraw
    'show final result
    objDev.Present ByVal 0, ByVal 0, 0, ByVal 0
  
    'increase fps counter
    fps = fps + 1
    'update fps & doevents (every 1000ms)
    If tick < GetTickCount - 1000 Then
      'when window is inactive - slow down rendering process
      If GetForegroundWindow = hWnd Then
        slow = False
      Else
        slow = True
      End If
      'update window title
      Caption = "Procedural Sky: " & ScaleWidth & "x" & ScaleHeight & "@" & fps & "Fps.   Sun Angle=" & Format(sun_angle * 180 / Pi, "0.00") & "; Turbidity=" & Format(sun_turbidity, "0.00")
      DoEvents
      'reset fps & renew timing
      fps = 0
      tick = GetTickCount
    End If
    
    'slow down rendering process, if required
    If slow Then Sleep 100
    
    'we need to control the demo with mouse at any time
    DoEvents
    
  Loop

End Sub


'mouse controls
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'only when left mb pressed
  If Button = vbLeftButton Then
  
    'add mouse coords
    msX = msX + (X - msXl)
    msY = msY + (Y - msYl)
    
    'range check
    If msY < -200 Then msY = -200
    If msY > 300 Then msY = 300
    
  End If

  'remember last mouse coords
  msXl = X
  msYl = Y

End Sub


'shutdown
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  'release all resources
  appRelease
  
  'destroy directx objects
  Set objDev = Nothing
  Set objD3D = Nothing
  Set objDX = Nothing

  'exit process
  End

End Sub
