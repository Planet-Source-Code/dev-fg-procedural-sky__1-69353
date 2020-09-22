VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSkyHemisphere"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'require variable declaration
Option Explicit
Option Base 0


'vertex structure
Private Type structVertex
  pos As vec3f  'position
  col As Long   'color
End Type

'vertex shader & format
Private Const vFormat As Long = D3DFVF_XYZ Or D3DFVF_DIFFUSE
Private Const vLength As Long = 16 '16 bytes per each vertex

'sky hemisphere data
Private objIndex As Direct3DIndexBuffer8   'index buffer
Private objBuffer As Direct3DVertexBuffer8 'vertex buffer
Private arrNormalized() As vec3f           'normalized vertex position array
Private arrStream() As structVertex        'vertex data array
Private arrIndex() As Integer              'index array
Private dwVertices As Long                 'number of vertices
Private dwIndexes As Long                  'number of indexes


'temp vars (for perezFunc)
Private gamma As Single
Private cosGammaSq As Single

'temp vars (for skyBuild)
Private ring As Long
Private seg As Long
Private vert As Long
Private ind As Long
Private vInd As Long
Private r0 As Single
Private y0 As Single
Private x0 As Single
Private z0 As Single
Private fDeltaRingAngle As Single
Private fDeltaSegAngle As Single

'temp vars (for skyCompute)
Private sunPos As vec3f
Private v As vec3f
Private i As Long
Private t2 As Single
Private chi As Single
Private theta As vec4f
Private sunquad As Single
Private clrYxy As vec3f
Private ratio As Single
Private xyz As vec3f
Private zenith As vec3f
Private l As vec3f

'const variables (for skyCompute)
Private ch As Single
Private cx1 As vec4f
Private cx2 As vec4f
Private cx3 As vec4f
Private cy1 As vec4f
Private cy2 As vec4f
Private cy3 As vec4f
Private rCoeffs As vec3f
Private gCoeffs As vec3f
Private bCoeffs As vec3f

'precomputed vars (for perezFunc)
Private ay0 As Single
Private by0 As Single
Private cy0 As Single
Private dy0 As Single
Private ey0 As Single
Private ax As Single
Private bx As Single
Private cx As Single
Private dx As Single
Private ex As Single
Private ay As Single
Private by As Single
Private cy As Single
Private dy As Single
Private ey As Single

'general share
Private t As Single
Private thetaSun As Single


'perez sky model: allweather function
Private Function perezFunc(cosTheta As Single, cosGamma As Single, preCalc As Boolean) As vec3f
  
  'can accelerate calculation, because thetaSun=acos(l.z) - is precomputed in "skyCompure" procedure
  If preCalc Then
    gamma = thetaSun
  Else
    gamma = acos(cosGamma)
  End If
  
  'calc squared cosGamma
  cosGammaSq = cosGamma * cosGamma
  
  'do perez calculations
  With perezFunc
    .X = (1 + ay0 * Exp(by0 / cosTheta)) * (1 + cy0 * Exp(dy0 * gamma) + ey0 * cosGammaSq)
    .Y = (1 + ax * Exp(bx / cosTheta)) * (1 + cx * Exp(dx * gamma) + ex * cosGammaSq)
    .z = (1 + ay * Exp(by / cosTheta)) * (1 + cy * Exp(dy * gamma) + ey * cosGammaSq)
  End With

End Function


'rendering procedure
Public Sub skyRender()

  With objDev
  
    'apply a vertex shader
    .SetVertexShader vFormat
    
    'setup data source
    .SetStreamSource 0, objBuffer, vLength
    .SetIndices objIndex, 0
    
    'draw sphere
    .DrawIndexedPrimitive D3DPT_TRIANGLESTRIP, 0, dwVertices, 0, dwIndexes - 2
  
  End With

End Sub


'computes sky hemisphere lighting according to the perez sky model
Public Sub skyCompute(turbidity As Single, sunx As Single, suny As Single, sunz As Single)
  
  'set general turabidity
  t = turbidity
  
  'prepare data for perez function
  ay0 = 0.17872 * t - 1.46303
  by0 = -0.3554 * t + 0.42749
  cy0 = -0.02266 * t + 5.32505
  dy0 = 0.12064 * t - 2.57705
  ey0 = -0.06696 * t + 0.37027
  ax = -0.01925 * t - 0.25922
  bx = -0.06651 * t + 0.00081
  cx = -0.00041 * t + 0.21247
  dx = -0.06409 * t - 0.89887
  ex = -0.00325 * t + 0.04517
  ay = -0.01669 * t - 0.26078
  by = -0.09495 * t + 0.00921
  cy = -0.00792 * t + 0.21023
  dy = -0.04405 * t - 1.65369
  ey = -0.01092 * t + 0.05291
  
  'some precalculations to make process faster
  l = norm3(vec3(sunx, suny, sunz)) 'normalized sun position
  thetaSun = acos(l.z)              'sun theta angle
  
  'calculate perezZenith
  t2 = t * t
  chi = (ch - t / 120) * (Pi - 2 * thetaSun)
  sunquad = thetaSun * thetaSun
  theta = vec4(1, thetaSun, sunquad, sunquad * thetaSun)
  With zenith
    .X = (4.0453 * t - 4.971) * Tan(chi) - 0.2155 * t + 2.4192
    .Y = t2 * dot4(cx1, theta) + t * dot4(cx2, theta) + dot4(cx3, theta)
    .z = t2 * dot4(cy1, theta) + t * dot4(cy2, theta) + dot4(cy3, theta)
  End With
  
  'reset index & start processing
  i = 0
  Do
    
    'get normalized vertex position
    v = arrNormalized(i)
    
    'process only hemisphere, not full sphere
    '(all vertices below the horizon are last in the array, so just stop processing, or we'll get overflow somewhere)
    If v.z < 0 Then Exit Do
    
    'calculate vertex data
    clrYxy = div3(mul3(zenith, perezFunc(v.z, dot3(l, v), False)), perezFunc(1, l.z, True))
    'convert vertex data into color
    clrYxy.X = 1 - Exp(-clrYxy.X / 25)
    ratio = clrYxy.X / clrYxy.z
    xyz.X = clrYxy.Y * ratio
    xyz.Y = clrYxy.X
    xyz.z = ratio - xyz.X - xyz.Y
    'setup vertex color
    arrStream(i).col = D3DColorARGB(255, dot3clmp(rCoeffs, xyz) * 255, dot3clmp(gCoeffs, xyz) * 255, dot3clmp(bCoeffs, xyz) * 255)
    
    'next vertex
    i = i + 1
  Loop While i < dwVertices
  
  'upload data into vertex buffer
  D3DVertexBuffer8SetData objBuffer, 0, vLength * dwVertices, 0, arrStream(0)

End Sub


'setup shader constants; generate hemisphere data & build vertex+index buffers
Public Sub skyBuild(numRings As Long, numSegments As Long, mhRadius As Single)

  'setup constants for periz-sky shading
  cx1 = vec4(0, 0.00209, -0.00375, 0.00165)
  cx2 = vec4(0.00394, -0.03202, 0.06377, -0.02903)
  cx3 = vec4(0.25886, 0.06052, -0.21196, 0.11693)
  cy1 = vec4(0, 0.00317, -0.0061, 0.00275)
  cy2 = vec4(0.00516, -0.04153, 0.0897, -0.04214)
  cy3 = vec4(0.26688, 0.0667, -0.26756, 0.15346)
  rCoeffs = vec3(3.240479, -1.53715, -0.49853)
  gCoeffs = vec3(-0.969256, 1.875991, 0.041556)
  bCoeffs = vec3(0.055684, -0.204043, 1.057311)
  ch = 4 / 9

  'start building the hemisphere, first reset indexes
  vInd = 0
  vert = 0
  ind = 0
  
  'how many vertices & indexes we will have
  dwVertices = (numRings + 1) * (numSegments + 1)
  dwIndexes = 2 * numRings * (numSegments + 1)
  
  'allocate memory
  ReDim arrStream(dwVertices - 1) As structVertex
  ReDim arrNormalized(dwVertices - 1) As vec3f
  ReDim arrIndex(dwIndexes - 1) As Integer
  
  'we need a hemisphere, but not full sphere, so we use: pi/2 instead of pi
  fDeltaRingAngle = Pi / 2 / numRings
  fDeltaSegAngle = 2 * Pi / numSegments
    
  For ring = 0 To numRings Step 1
  
    z0 = Cos(ring * fDeltaRingAngle) * mhRadius
    r0 = Sin(ring * fDeltaRingAngle)
  
    For seg = 0 To numSegments Step 1
    
      y0 = r0 * Sin(seg * fDeltaSegAngle) * mhRadius
      x0 = r0 * Cos(seg * fDeltaSegAngle) * mhRadius
    
      With arrStream(vert)
        'set vertex position
        With .pos
          .X = x0
          .Y = y0
          .z = z0
        End With
        'black by default
        .col = 0
        'set normalized position (required for faster perez calculations)
        arrNormalized(vert) = norm3(.pos)
      End With
      vert = vert + 1

      'generate index array on the fly
      If ring <> numRings Then
        arrIndex(ind) = vInd
        ind = ind + 1
        arrIndex(ind) = vInd + numSegments + 1
        ind = ind + 1
        vInd = vInd + 1
      End If
      
    Next seg
  
  Next ring
  
  'create & load index buffer
  Set objIndex = objDev.CreateIndexBuffer(dwIndexes * 2, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
  D3DIndexBuffer8SetData objIndex, 0, dwIndexes * 2, 0, arrIndex(0)
  
  'create & load vertex buffer
  Set objBuffer = objDev.CreateVertexBuffer(vLength * dwVertices, D3DUSAGE_WRITEONLY, vFormat, D3DPOOL_DEFAULT)
  D3DVertexBuffer8SetData objBuffer, 0, vLength * dwVertices, 0, arrStream(0)

End Sub


'sky-sphere destruction
Public Sub skyDestroy()

  'destroy vertex & index buffers
  Set objIndex = Nothing
  Set objBuffer = Nothing

  'erase data arrays
  Erase arrIndex()
  Erase arrStream()
  Erase arrNormalized()

End Sub

