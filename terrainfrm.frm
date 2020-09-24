VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   4605
   ClientTop       =   2295
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7995
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------
' TrueVision 8 Cloud sample + Splash screen
'-------------------------------------------

Public TV As TrueVision8
Public Scene As Scene8
Public Land As Landscape8
Public Inp As InputEngine8
Public Scr As Screen8
Public TexFactory As New TextureFactory8
Public PosX As Single, PosY As Single, PosZ As Single, Ang As Single, AngY As Single, MX As Long, MY As Long
Public Camera As New Camera8
Public mesh As New Mesh8
Dim GCamera(2) As CameraGravity
Public Effects As New GraphicEffect8
Dim Tx As Long, Ty As Long
Private Sub Form_Load()
  Dim Point1 As D3DVECTOR, Point2 As D3DVECTOR
  'Create the objects
  Set TV = New TrueVision8
  Set Scene = New Scene8
  Set Land = New Landscape8
  Set Scr = New Screen8
 
  'Initialize TrueVision WITHOUT console (no loading info)
  TV.Console = False
  If TV.ShowDriverDialog = False Then Exit Sub
  TV.Initialize
  Set Inp = New InputEngine8
  TV.DisplayFPS = True
  
  Set mesh = Scene.CreateMeshBuilder
  mesh.SetVertexType TV_NORMAL_VERTEX
  
  'Change the default path
  TV.SetSearchDirectory App.Path

  'Load the different textures
  TexFactory.LoadTexture "cloudx.dds", "Cloud1", , , TV_COLORKEY_BLACK
  TexFactory.LoadTexture "cloudx.dds", "Cloudx", , , TV_COLORKEY_BLACK
  TexFactory.LoadTexture "map_farview.jpg", "Terrain"
  TexFactory.LoadTexture "sun.jpg", "sun"
  TexFactory.LoadTexture "water.bmp", "water"
  TexFactory.LoadTexture "detail.jpg", "detail"
  TexFactory.LoadTexture "Back.jpg", "Front"
  TexFactory.LoadTexture "Front.jpg", "Back"
  TexFactory.LoadTexture "Left.jpg", "Right"
  TexFactory.LoadTexture "Right.jpg", "Left"
  TexFactory.LoadTexture "top.jpg", "Down"
  TexFactory.LoadTexture "Down.jpg", "Top"
  TexFactory.LoadTexture "Wall.Bmp", "Wall"
  TexFactory.LoadTexture "Floor.Bmp", "Floor"
  TexFactory.LoadTexture "flare1.jpg", "Flare1"
  TexFactory.LoadTexture "flare2.jpg", "Flare2"
  TexFactory.LoadTexture "flare3.jpg", "Flare3"
  TexFactory.LoadTexture "flare4.jpg", "Flare4"
  
  'Load the logo and the loading text
  TexFactory.LoadTexture "loading.jpg", "Loading"
  TexFactory.LoadTexture "logo.jpg", "Logo"
 

  Scene.SetCollisionPrecision 10
  'In fact you can display a splash screen like in the game loop
  'with the Screen class
  'Get the selected video mode to center the title
  
  Dim width As Long, height As Long
  TV.GetVideoMode width, height, 0
  
  
  'Clear the screen
  TV.Clear
    
  Scr.DrawTexture GetTex("Loading"), width / 2 - 150, height / 2 - 50, width / 2 + 250, height / 2 + 50
  Scr.DrawTexture GetTex("Logo"), width / 2 - 320, height / 2 - 100, width / 2 - 175, height / 2 + 100
  
  
  'Render the memory buffer to the screen
  TV.RenderToScreen
  Scene.EnableMipMapping True
  
  Scene.SetSceneBackGround 0.4, 0.6, 0.9
  Land.GenerateHugeTerrain "land_map.jpg", TV_PRECISION_HIGH, 10, 10, 0, 0, True
  Land.EnableLOD True
  Land.SetWaterTexture GetTex("water")
  Land.SetWaterEffect True, 0.9, True, D3DBLEND_DESTCOLOR, D3DBLEND_SRCCOLOR, False
  Land.SetWaterTextureScale 5
  Land.SetWaterAltitude 71
  Land.SetWaterEnable True
  
  'Init the sun
  Land.InitSun GetTex("Sun"), Vector(-1000, 500, 0), 2
  
  'A lot of work to make a simple mesh with Tx and Ty to move it very fast
  Tx = 1200
  Ty = 950
  mesh.AddWall GetTex("Wall"), 0 + Tx, 0 + Ty, 100 + Tx, 0 + Ty, 100, 0, 5, 5
  mesh.AddWall GetTex("Wall"), 0 + Tx, 0 + Ty, 0 + Tx, 100 + Ty, 100, 0, 5, 5
  mesh.AddWall GetTex("Wall"), 100 + Tx, 0 + Ty, 100 + Tx, 100 + Ty, 100, 0, 5, 5
  mesh.AddWall GetTex("Wall"), 0 + Tx, 100 + Ty, 100 + Tx, 100 + Ty, 100, 0, 5, 5
  mesh.AddFloor GetTex("Floor"), 0 + Tx, 0 + Ty, 100 + Tx, 100 + Ty, 100, 5, 5
  mesh.AddFloor GetTex("Floor"), 0 + Tx, 0 + Ty, 100 + Tx, 100 + Ty, 150, 5, 5
  mesh.AddWall GetTex("Wall"), 100 + Tx, 100 + Ty, 110 + Tx, 100 + Ty, 80, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 90 + Ty, 110 + Tx, 90 + Ty, 82, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 80 + Ty, 110 + Tx, 80 + Ty, 84, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 70 + Ty, 110 + Tx, 70 + Ty, 86, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 60 + Ty, 110 + Tx, 60 + Ty, 88, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 50 + Ty, 110 + Tx, 50 + Ty, 90, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 40 + Ty, 110 + Tx, 40 + Ty, 92, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 30 + Ty, 110 + Tx, 30 + Ty, 94, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 20 + Ty, 110 + Tx, 20 + Ty, 96, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 10 + Ty, 110 + Tx, 10 + Ty, 100, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 0 + Ty, 110 + Tx, 0 + Ty, 100, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 100 + Ty, 110 + Tx, 90 + Ty, 80, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 90 + Ty, 110 + Tx, 80 + Ty, 82, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 80 + Ty, 110 + Tx, 70 + Ty, 84, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 70 + Ty, 110 + Tx, 60 + Ty, 86, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 60 + Ty, 110 + Tx, 50 + Ty, 88, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 50 + Ty, 110 + Tx, 40 + Ty, 90, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 40 + Ty, 110 + Tx, 30 + Ty, 92, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 30 + Ty, 110 + Tx, 20 + Ty, 94, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 20 + Ty, 110 + Tx, 10 + Ty, 96, 0, 2, 1
  mesh.AddWall GetTex("Wall"), 110 + Tx, 10 + Ty, 110 + Tx, 0 + Ty, 100, 0, 2, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 100 + Ty, 100 + Tx, 90 + Ty, 80, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 90 + Ty, 100 + Tx, 80 + Ty, 82, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 80 + Ty, 100 + Tx, 70 + Ty, 84, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 70 + Ty, 100 + Tx, 60 + Ty, 86, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 60 + Ty, 100 + Tx, 50 + Ty, 88, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 50 + Ty, 100 + Tx, 40 + Ty, 90, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 40 + Ty, 100 + Tx, 30 + Ty, 92, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 30 + Ty, 100 + Tx, 20 + Ty, 94, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 20 + Ty, 100 + Tx, 10 + Ty, 96, 1, 1
  mesh.AddFloor GetTex("Floor"), 110 + Tx, 10 + Ty, 100 + Tx, 0 + Ty, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 0 + Tx, 0 + Ty, 10 + Tx, 0 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 0 + Tx, 0 + Ty, 0 + Tx, 10 + Ty, 50, 100, 1
  mesh.AddWall GetTex("Wall"), 10 + Tx, 10 + Ty, 10 + Tx, 0 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 0 + Tx, 10 + Ty, 10 + Tx, 10 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 0 + Tx, 90 + Ty, 10 + Tx, 90 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 0 + Tx, 90 + Ty, 0 + Tx, 100 + Ty, 50, 100, 1
  mesh.AddWall GetTex("Wall"), 10 + Tx, 100 + Ty, 10 + Tx, 90 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 0 + Tx, 100 + Ty, 10 + Tx, 100 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 90 + Tx, 90 + Ty, 100 + Tx, 90 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 90 + Tx, 90 + Ty, 90 + Tx, 100 + Ty, 50, 100, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 100 + Ty, 100 + Tx, 90 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 90 + Tx, 100 + Ty, 100 + Tx, 100 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 90 + Tx, 10 + Ty, 100 + Tx, 10 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 90 + Tx, 10 + Ty, 90 + Tx, 20 + Ty, 50, 100, 1
  mesh.AddWall GetTex("Wall"), 100 + Tx, 20 + Ty, 100 + Tx, 10 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 90 + Tx, 20 + Ty, 100 + Tx, 20 + Ty, 50, 100, 1, 1
  mesh.AddWall GetTex("Wall"), 0 + Tx, 0 + Ty, 100 + Tx, 0 + Ty, 10, 150, 5, 5
  mesh.AddWall GetTex("Wall"), 0 + Tx, 0 + Ty, 0 + Tx, 100 + Ty, 10, 150, 5, 5
  mesh.AddWall GetTex("Wall"), 100 + Tx, 0 + Ty, 100 + Tx, 100 + Ty, 10, 150, 5, 5
  mesh.AddWall GetTex("Wall"), 0 + Tx, 100 + Ty, 100 + Tx, 100 + Ty, 10, 150, 5, 5
  mesh.AddFloor GetTex("Floor"), 0 + Tx, 0 + Ty, 100 + Tx, 100 + Ty, 160, 5, 5
   
  Land.SetTexture GetTex("Terrain")
  Land.SetTextureScale 1, 1
  Land.ExpandTexture GetTex("Terrain"), 0, 0, 10, 10
  
  Land.SetDetailTexture GetTex("detail")
  Land.SetDetailTextureScale 15, 15
   

  
  Land.InitLensFlare 4
  Land.SetLensFlare 1, 2 * 5, GetTex("Flare1"), 40, RGBA(1, 1, 1, 0.5), RGBA(1, 1, 1, 0.5)
  Land.SetLensFlare 2, 2 * 1, GetTex("Flare2"), 18, RGBA(1, 1, 1, 0.5), RGBA(1, 1, 1, 0.5)
  Land.SetLensFlare 3, 2 * 1.8, GetTex("Flare3"), 15, RGBA(1, 1, 1, 0.5), RGBA(0.7, 1, 1, 0.5)
  Land.SetLensFlare 4, 2 * 1, GetTex("Flare4"), 6, RGBA(1, 0.1, 0, 0.5), RGBA(0.5, 1, 1, 0.5)
  
  Scene.SetSkyTexture GetTex("Front"), GetTex("Back"), GetTex("Left"), GetTex("Right"), GetTex("Top"), GetTex("Down"), 1500
  
  GCamera(0).x = 700
  GCamera(0).z = 900

  Do
    ChangePos
    TV.Clear
    Scene.RenderSky
    
    
    
    
    If GCamera(0).y > 70 Then
      Effects.SetFog False, 1, 1, 1, 0.1
      Land.RenderLensFlare
      Effects.SetFog True, 0.2, 0.4, 0.6, 0.1, 1, 2000
    Else
      Effects.SetFog True, 0, 0.4, 0.5, 100, 1, 100
    End If
    Land.Render False
    Effects.SetFog False, 1, 1, 1, 1
    Scene.RenderAllMeshes
    TV.RenderToScreen
   
  Loop Until Inp.IsKeyPressed(TV_KEY_ESCAPE)
  
  Set TV = Nothing
  End

End Sub
Sub ChangePos()
  GCamera(1) = GCamera(0)
   If Inp.IsKeyPressed(TV_KEY_UP) = True Then
       GCamera(1).x = GCamera(0).x + Cos(GCamera(0).Ang) * TV.TimeElapsed * 0.08
       GCamera(1).z = GCamera(0).z + Sin(GCamera(0).Ang) * TV.TimeElapsed * 0.08
   End If
   If Inp.IsKeyPressed(TV_KEY_DOWN) = True Then
       GCamera(1).x = GCamera(0).x - Cos(GCamera(0).Ang) * TV.TimeElapsed * 0.05
       GCamera(1).z = GCamera(0).z - Sin(GCamera(0).Ang) * TV.TimeElapsed * 0.05
   End If
   If Inp.IsKeyPressed(TV_KEY_LEFT) = True Then
       GCamera(1).Ang = GCamera(0).Ang + TV.TimeElapsed * 0.005
   End If
   If Inp.IsKeyPressed(TV_KEY_RIGHT) = True Then
       GCamera(1).Ang = GCamera(0).Ang - TV.TimeElapsed * 0.005
   End If
   If Inp.IsKeyPressed(TV_KEY_SUBTRACT) = True Then
       GCamera(1).y = GCamera(0).y + TV.TimeElapsed * 0.05
   End If
   If Inp.IsKeyPressed(TV_KEY_ADD) = True Then
       GCamera(1).y = GCamera(0).y - TV.TimeElapsed * 0.5 + 5
   End If
   If Inp.IsKeyPressed(TV_KEY_S) = True Then
       TV.ScreenShot "c:\Screenshot.bmp"
   End If
   Inp.GetMouseState MX, MY
     AngY = AngY - (MY / 100)
   If AngY > ((22 / 7) * 0.1) Then
     AngY = ((22 / 7) * 0.1)
   End If
   If AngY < -((22 / 7) * 0.8) Then
     AngY = -((22 / 7) * 0.8)
   End If
   GCamera(1).Ang = GCamera(1).Ang - (MX / 100)
   If GCamera(1).x > 1190 And GCamera(1).x < 1320 And GCamera(1).z > 940 And GCamera(1).z < 1060 Then
      GCamera(0) = Scene.CheckGravity(GCamera(), TV.TimeElapsed, 0.04, 0.04)
      GCamera(1) = GCamera(0)
      If GCamera(0).y < 89 Then
        GCamera(0).y = 89
      End If
   Else
      GCamera(1).y = Land.GetHeight(GCamera(1).x, GCamera(1).z) + 10
      GCamera(0) = GCamera(1)
   End If
   Scene.SetCamera GCamera(0).x, GCamera(0).y, GCamera(0).z, GCamera(0).x + Cos(GCamera(0).Ang), GCamera(0).y + Cos(AngY), GCamera(0).z + Sin(GCamera(0).Ang)
End Sub


