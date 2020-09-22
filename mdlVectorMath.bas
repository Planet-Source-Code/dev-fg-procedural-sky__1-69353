Attribute VB_Name = "mdlVectorMath"

'require variable declaration
Option Explicit


'pi constants
Public Const Pi As Single = 3.14159265358979
Public Const PiHalf As Single = 1.5707963267949


'xyz vector structure
Public Type vec3f
  X As Single
  Y As Single
  z As Single
End Type

'wxyz vector structure
Public Type vec4f
  w As Single
  X As Single
  Y As Single
  z As Single
End Type



'project x coord into screenspace
Public Function prjU(u As Long) As Single
  
  With typConfig
    prjU = (u - .BackBufferWidth / 2) / (.BackBufferWidth / 2)
  End With

End Function


'project y coord into screenspace
Public Function prjV(v As Long) As Single
  
  With typConfig
    prjV = -(v - .BackBufferHeight / 2) / (.BackBufferHeight / 2)
  End With

End Function


'creates d3dvector construction
Public Function vec3x(X As Single, Y As Single, z As Single) As D3DVECTOR
  
  With vec3x
    .X = X
    .Y = Y
    .z = z
  End With

End Function


'creates vec3f construction
Public Function vec3(X As Single, Y As Single, z As Single) As vec3f
  
  With vec3
    .X = X
    .Y = Y
    .z = z
  End With

End Function


'creates vec4f construction
Public Function vec4(w As Single, X As Single, Y As Single, z As Single) As vec4f
  
  With vec4
    .w = w
    .X = X
    .Y = Y
    .z = z
  End With

End Function


'mupliply vec3f
Public Function mul3(a As vec3f, b As vec3f) As vec3f
  
  With mul3
    .X = a.X * b.X
    .Y = a.Y * b.Y
    .z = a.z * b.z
  End With

End Function


'divide vec3f
Public Function div3(a As vec3f, b As vec3f) As vec3f
  
  With div3
    .X = a.X / b.X
    .Y = a.Y / b.Y
    .z = a.z / b.z
  End With

End Function


'arccosine calculation
Public Function acos(f As Single) As Single
  
  If Abs(f) < 1 Then
     
     'calculate arccosine
     acos = Atn(-f / Sqr(-f * f + 1)) + PiHalf
  
  ElseIf f = 1 Then
     
     'return 0 (when input is 1)
     acos = 0
  
  ElseIf f = -1 Then
     
     'return pi (when input is -1)
     acos = Pi
  
  End If

End Function


'dot product 3 calculation
Public Function dot3(a As vec3f, b As vec3f) As Single
  
  dot3 = a.X * b.X + a.Y * b.Y + a.z * b.z

End Function


'dot product 3 calculation with 0-1 clamping
Public Function dot3clmp(a As vec3f, b As vec3f) As Single
  
  'dot3 product
  dot3clmp = dot3(a, b)
  
  'clamp
  If dot3clmp < 0 Then dot3clmp = 0
  If dot3clmp > 1 Then dot3clmp = 1

End Function


'dot product 4 calculation
Public Function dot4(a As vec4f, b As vec4f) As Single
  
  dot4 = a.w * b.w + a.X * b.X + a.Y * b.Y + a.z * b.z

End Function


'subtract vec3f
Public Function sub3(a As vec3f, b As vec3f) As vec3f
  
  With sub3
    .X = a.X - b.X
    .Y = a.Y - b.Y
    .z = a.z - b.z
  End With

End Function


'vector length (vec3f)
Public Function len3(a As vec3f) As Single
  
  With a
    len3 = Sqr(.X ^ 2 + .Y ^ 2 + .z ^ 2)
  End With

End Function


'normalize vector (vec3f)
Public Function norm3(a As vec3f) As vec3f
  
  'temp variables
  Dim vectorLength As Single
  Dim vectorTemp As vec3f
  
  'store vector
  vectorTemp = a
  
  'calculate it's length
  vectorLength = len3(a)
  
  'normalize
  If vectorLength > 0 Then
    With vectorTemp
      .X = .X / vectorLength
      .Y = .Y / vectorLength
      .z = .z / vectorLength
    End With
  End If
  
  'return result
  norm3 = vectorTemp

End Function

