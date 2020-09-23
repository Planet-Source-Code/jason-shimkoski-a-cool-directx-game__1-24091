Attribute VB_Name = "modSurfaces"
'these are all of the surfaces used by direct draw.
Public ddsBG As DirectDrawSurface7
Public ddsdBG As DDSURFACEDESC2
Public rectBG As RECT

Public ddsGreeny As DirectDrawSurface7
Public ddsdGreeny As DDSURFACEDESC2
Public rectGreeny As RECT

Public ddsYellow As DirectDrawSurface7
Public ddsdYellow As DDSURFACEDESC2
Public rectYellow As RECT

Public ddsLogo As DirectDrawSurface7
Public ddsdLogo As DDSURFACEDESC2
Public rectLogo As RECT

Public ddsMainMenu As DirectDrawSurface7
Public ddsdMainMenu As DDSURFACEDESC2
Public rectMainMenu As RECT

Public ddsGums As DirectDrawSurface7
Public ddsdGums As DDSURFACEDESC2
Public rectGums As RECT

'this initializes the graphics
Sub InitGraphics()
    CreateGraphicsFromFile "bg", ddsBG, ddsdBG, 16, 16
    CreateGraphicsFromFile "greeny", ddsGreeny, ddsdGreeny, 256, 128
    CreateGraphicsFromFile "yellow", ddsYellow, ddsdYellow, 256, 128
    CreateGraphicsFromFile "logo", ddsLogo, ddsdLogo, 400, 128
    CreateGraphicsFromFile "menu-main", ddsMainMenu, ddsdMainMenu, 357, 177
    CreateGraphicsFromFile "gums", ddsGums, ddsdGums, 63, 72
End Sub
