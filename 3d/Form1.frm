VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   3165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Thank you for downloading my code! If you have any suggestions
'or find any bugs, email me at Seriously05@aol.com. Other than
'that; have fun, please don't steal my code, and don't forget
'to vote!

Option Explicit

Dim DX8 As New DirectX8 'declare a new directx8 project
Dim D3D As Direct3D8 'used to create the rendering device
Dim DDEV As Direct3DDevice8 'the rendering device
Dim VB As Direct3DVertexBuffer8 'the vertex buffer (obviously)

'used to display a colored vertex on the screen
Private Type CUSTOMVERTEX
X As Single 'defines a place on the x-axis
Y As Single 'y-axis
Z As Single 'z-axis
COLOR As Long 'the color
End Type

'describes the custom vertex structure
Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZ Or D3DFVF_DIFFUSE)

'what kind of world would it be without our friend pi???
'it is used for so many things...but in this project,
'it sets up the projection matrix
Const PI = 3.1415

'vertex-setup helper function
Function Vecs(X As Single, Y As Single, Z As Single) As D3DVECTOR

Vecs.X = X
Vecs.Y = Y
Vecs.Z = Z

End Function

Sub Matrices()

Dim MATWORLD As D3DMATRIX 'positions and orients everything drawn into the world
D3DXMatrixRotationY MATWORLD, Timer 'rotate everything drawn around the y-axis
DDEV.SetTransform D3DTS_WORLD, MATWORLD 'set it up!

Dim MATVIEW As D3DMATRIX 'how we view the objects (camera)

'places the camera 3 spots up on the y-axis, 5 spots
'back on the z-axis, tells it to focus on the origin,
'and tells it that the y-axis is up
'also notice that we are using a left handed view
'of the coordinate system:
'     y  z
'     | /
'     |/
'----------- x
'    /|
'   / |

D3DXMatrixLookAtLH MATVIEW, Vecs(0#, 3#, -5#), _
                            Vecs(0#, 0#, 0#), _
                            Vecs(0#, 1#, 0#)

DDEV.SetTransform D3DTS_VIEW, MATVIEW 'set it up!

Dim MATPROJ As D3DMATRIX 'describes the camera's lenses

'perspective view space: (makes things smaller in the distance)
'sets up the field of view (1/4 pi), aspect ratio, and clipping
'planes (how far away an object has to be before its no longer
'rendered
'notice that, like above, it too is left handed
D3DXMatrixPerspectiveFovLH MATPROJ, PI / 4, 1, 0, 1000
DDEV.SetTransform D3DTS_PROJECTION, MATPROJ 'set it up!

End Sub
Function Init(hwnd As Long) As Boolean

'if a bug is present...move on
On Local Error Resume Next

'create the direct3d object; if something goes wrong...ABORT!
Set D3D = DX8.Direct3DCreate
If D3D Is Nothing Then Exit Function

'gets the default display mode
Dim MODE As D3DDISPLAYMODE
D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, MODE

'fill in the type structure used to create the device
Dim DPP As D3DPRESENT_PARAMETERS
DPP.Windowed = True 'the rendering will be windowed
DPP.BackBufferFormat = MODE.Format 'default back buffer format
DPP.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC 'describes how it wants the object to be rendered

'create the direct3ddevice; if something goes wrong...ABORT!!
'uses hardware acceleration with software vertex processing
Set DDEV = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hwnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, DPP)
If DDEV Is Nothing Then Exit Function

DDEV.SetRenderState D3DRS_CULLMODE, D3DCULL_CW 'tells the rendering device not to draw backfaces with clockwise vertices
DDEV.SetRenderState D3DRS_LIGHTING, 0 'turn the lighting off (because our vertices are colorful)

Dim VERT(23) As CUSTOMVERTEX 'says we'll be dealing with 24 vertices
Dim VERTSIZE As Long 'how many bytes will this be???

VERTSIZE = Len(VERT(0)) 'this many bytes * 3...

'defines every (colorful) vertice of the object
With VERT(0): .X = 0: .Y = 2: .Z = 0: .COLOR = &HFF00FFFF: End With
With VERT(1): .X = 1: .Y = 0: .Z = 1: .COLOR = &HFF00FF00: End With
With VERT(2): .X = -1: .Y = 0: .Z = 1: .COLOR = &HFFFF0000: End With

With VERT(3): .X = 0: .Y = 2: .Z = 0: .COLOR = &HFF00FFFF: End With
With VERT(4): .X = 1: .Y = 0: .Z = -1: .COLOR = &HFFFF0000: End With
With VERT(5): .X = 1: .Y = 0: .Z = 1: .COLOR = &HFF00FF00: End With

With VERT(6): .X = 0: .Y = 2: .Z = 0: .COLOR = &HFF00FFFF: End With
With VERT(7): .X = -1: .Y = 0: .Z = 1: .COLOR = &HFFFF0000: End With
With VERT(8): .X = -1: .Y = 0: .Z = -1: .COLOR = &HFF00FF00: End With

With VERT(9): .X = -1: .Y = 0: .Z = -1: .COLOR = &HFF00FF00: End With
With VERT(10): .X = 1: .Y = 0: .Z = -1: .COLOR = &HFFFF0000: End With
With VERT(11): .X = 0: .Y = 2: .Z = 0: .COLOR = &HFF00FFFF: End With

With VERT(12): .X = -1: .Y = 0: .Z = 1: .COLOR = &HFFFF0000: End With
With VERT(13): .X = 1: .Y = 0: .Z = 1: .COLOR = &HFF00FF00: End With
With VERT(14): .X = 0: .Y = -2: .Z = 0: .COLOR = &HFF00FFFF: End With

With VERT(15): .X = 1: .Y = 0: .Z = 1: .COLOR = &HFF00FF00: End With
With VERT(16): .X = 1: .Y = 0: .Z = -1: .COLOR = &HFFFF0000: End With
With VERT(17): .X = 0: .Y = -2: .Z = 0: .COLOR = &HFF00FFFF: End With

With VERT(18): .X = -1: .Y = 0: .Z = -1: .COLOR = &HFF00FF00: End With
With VERT(19): .X = -1: .Y = 0: .Z = 1: .COLOR = &HFFFF0000: End With
With VERT(20): .X = 0: .Y = -2: .Z = 0: .COLOR = &HFF00FFFF: End With

With VERT(21): .X = 0: .Y = -2: .Z = 0: .COLOR = &HFF00FFFF: End With
With VERT(22): .X = 1: .Y = 0: .Z = -1: .COLOR = &HFFFF0000: End With
With VERT(23): .X = -1: .Y = 0: .Z = -1: .COLOR = &HFF00FF00: End With

'create the vertex buffer; if something goes wrong...ABORT!!
Set VB = DDEV.CreateVertexBuffer(VERTSIZE * 24, 0, ByVal 0, D3DPOOL_DEFAULT)
If VB Is Nothing Then Exit Function

'fill the vertex buffer from the erray
D3DVertexBuffer8SetData VB, 0, VERTSIZE * 24, 0, VERT(0)

'yay, we made it!
Init = True

End Function
Sub Render()

Dim V As CUSTOMVERTEX 'all of the declared points
Dim VSIZE As Long 'size (bytes) of the declared points

'if the rendering device isn't functioning then ABORT!
If DDEV Is Nothing Then Exit Sub

'clear the backbuffer to a black color
DDEV.Clear 0, ByVal 0, D3DCLEAR_TARGET, &H0&, 1#, 0

'begin the scene
DDEV.BeginScene

'setup the world, view, and projection matrices
Matrices

VSIZE = Len(V) 'size (bytes) of the declared points

'draw the declared points (triangles) in the vertex buffer
DDEV.SetStreamSource 0, VB, VSIZE
DDEV.SetVertexShader D3DFVF_CUSTOMVERTEX
DDEV.DrawPrimitive D3DPT_TRIANGLELIST, 0, 8

'end the scene
DDEV.EndScene

'present everything in the backbuffer to the front buffer
DDEV.Present ByVal 0, ByVal 0, ByVal 0, ByVal 0

End Sub

'it's a dirty job, but somebody has to do it
'clears all initialze objects
Sub Clean()
Set D3D = Nothing
Set DDEV = Nothing
Set VB = Nothing
End Sub
Private Sub Form_Load()
DoEvents 'lets windows breath

Form1.WindowState = vbMaximized 'this is obvious...make the form maximized

Me.Show 'present the form, itself

Dim B As Boolean 'temporary boolean

'fill our temporary boolean with a new boolean (init)
'and tell the rendering device to display our object
'in the form
B = Init(Form1.hwnd)

'if something went wrong...tell the user, then unload
If Not B Then
MsgBox "Unable to Initialize DirectX.", vbOKOnly, "..."
Form_Unload (0)
End If

'let's begin the rendering!
Timer1.Enabled = True
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if someone clicks the form...unload
Form_Unload (0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
'self destruct!
Timer1.Enabled = False
Clean
End
End Sub

Private Sub Timer1_Timer()
'let the show begin
Render
End Sub
