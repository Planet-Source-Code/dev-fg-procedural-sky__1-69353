Attribute VB_Name = "mdlMain"


'require variable declaration
Option Explicit


'device settings (required for device creation, camera projection & rt initialization)
Public typConfig As D3DPRESENT_PARAMETERS
'rendering device itself
Public objDev As Direct3DDevice8

'mouse position (with no clamping)
Public msX As Long
Public msY As Long

'camera matrices
Private matProjection As D3DMATRIX
Private matWorld As D3DMATRIX
Private matView As D3DMATRIX


'the sky mesh & vertex program itself
Public mhSky As clsSkyHemisphere

'render zenith viewport
Private rViewportOrig As D3DVIEWPORT8
Private rViewport As D3DVIEWPORT8
Public swZenith As Boolean


'sky settings
Public sun_angle As Single
Public sun_turbidity As Single



'scene initialization
Public Sub appInitialize()
  
  'reset world matrix
  D3DXMatrixIdentity matWorld
  objDev.SetTransform D3DTS_WORLD, matWorld
  
  'defaults
  sun_turbidity = 2
  sun_angle = 1
  swZenith = True
  msY = -100
  msX = 200
  
  'create a sky
  Set mhSky = New clsSkyHemisphere
  mhSky.skyBuild 35, 70, 100

  'compute sky
  mhSky.skyCompute sun_turbidity, Sin(sun_angle), 0, Cos(sun_angle)
  

End Sub


'scene render
Public Sub appDraw()

  With objDev
    
    .BeginScene
    
    'prepare for rendering sky
    .SetRenderState D3DRS_LIGHTING, 0
    .SetRenderState D3DRS_CULLMODE, D3DCULL_CW
    .SetRenderState D3DRS_ZENABLE, 0
    .SetRenderState D3DRS_ZWRITEENABLE, 0
    
    'adjust camera
    D3DXMatrixPerspectiveFovLH matProjection, 1, typConfig.BackBufferHeight / typConfig.BackBufferWidth, 1, 1000
    D3DXMatrixLookAtLH matView, vec3x(0, 0, 0), vec3x(Sin(msX / 200), Cos(msX / 200), Exp(msY / 100)), vec3x(0, 0, 1)
    .SetTransform D3DTS_VIEW, matView
    .SetTransform D3DTS_PROJECTION, matProjection
    
    'clear backbuffer with black color (alpha=0)
    .Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 0, 0
    
    
    'render sky
    mhSky.skyRender
  
    'render sky zenith viewport
    If swZenith Then
      'modify viewport a little
      .GetViewport rViewportOrig
      rViewport = rViewportOrig
      rViewport.Width = 256
      rViewport.Height = 256
      .SetViewport rViewport
      'render with alpha (this will smooth edges of the mini-map a little)
      .SetRenderState D3DRS_ALPHABLENDENABLE, 1
      .SetRenderState D3DRS_SRCBLEND, D3DBLEND_BOTHSRCALPHA
      .SetRenderState D3DRS_DESTBLEND, D3DBLEND_BOTHSRCALPHA
      'setup camera (to render zenith viewport)
      D3DXMatrixPerspectiveFovLH matProjection, 1, 1, 1, 1000
      D3DXMatrixLookAtLH matView, vec3x(0, 0, -190), vec3x(0, 0, 0), vec3x(Sin(msX / 200), -Cos(msX / 200), 0)
      .SetTransform D3DTS_VIEW, matView
      .SetTransform D3DTS_PROJECTION, matProjection
      'draw sky
      mhSky.skyRender
      'set back original viewport
      .SetViewport rViewportOrig
      .SetRenderState D3DRS_ALPHABLENDENABLE, 0
    End If
     
    .EndScene
  
  End With

End Sub


'scene destruction
Public Sub appRelease()

  'destroy the sky
  mhSky.skyDestroy
  Set mhSky = Nothing

End Sub
