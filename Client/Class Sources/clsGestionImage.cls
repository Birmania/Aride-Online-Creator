VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGestionImage"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================================================================
'Nom du contrôle : DKS_GESTION_IMAGE
'Version : 1.0.2
'Auteur : Teillet nicolas
'Environnement de développement : Visual basic 6.0
'Résumé : Librairie de gestion d'image pour un accès rapide
'====================================================================================================
'====================================================================================================
'Nom du fichier : CLS_Image.cls
'Crée le : 27/03/2004
'Rôle : classe principale de la librairie contenant l'objet d'image
'====================================================================================================

'déclaration des types privées pour les API :
'structure stockant les informations des en-têtes bitmap
Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'structure stockant les informations des bits d'une image
Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

'stucture permettant de définir un bitmap
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

'fonction permettant de récupèrer les bits d'une image
Private Declare Function GetDIBits Lib "gdi32" ( _
ByVal hdc As Long, _
ByVal hBitmap As Long, _
ByVal nStartScan As Long, _
ByVal nScanLines As Long, _
lpBits As Any, _
lpBitmapInfo As BITMAPINFO, _
ByVal wUsage As Long) As Long

'fonction permettant de définir les bits d'un bitmap dans un DC
Private Declare Function SetDIBits Lib "gdi32.dll" ( _
ByVal hdc As Long, _
ByVal hBitmap As Long, _
ByVal nStartScan As Long, _
ByVal nNumScans As Long, _
ByRef lpBits As Any, _
ByRef lpBI As BITMAPINFO, _
ByVal wUsage As Long) As Long
     
'fonction permettant de détruire un Device Context
Private Declare Function DeleteDC Lib "gdi32" ( _
ByVal hdc As Long) As Long

'fonction permettant de détruire l'objet sélectionné pour le tracé
Private Declare Function DeleteObject Lib "gdi32" ( _
ByVal hObject As Long) As Long

'déclaration des constantes privées pour les API :
Const BI_RGB = 0 'définit le type de compression d'un bitmap
Const DIB_RGB_COLORS = 0 'table de couleur en RGB
Private Const pixR As Integer = 3 'index du pixel de couleur rouge
Private Const pixG As Integer = 2 'index du pixel de couleur vert
Private Const pixB As Integer = 1 'index du pixel de couleur bleu

'déclaration des variables membres privées à l'objet
Private BYT_ImageByts_Image() As Byte 'stocke le tableau de bits associés à l'image
Private LNG_Handle_Bitmap As Long 'stocke le handle du bitmap
Private LNG_Handle_DC As Long 'stocke le handle du device context
Private PCT_PictureBox_Bitmap As PictureBox 'stocke le hDC de l'image
Private TYP_Info_Bitmap As BITMAPINFO 'stocke les info du bitmap associé à l'image

Public Sub Refresh()

'===========================================================================
'Permet de coller l'image stockée dans l'objet
'===========================================================================

'on active la routine de traitement d'erreur
'On Error Resume Next
'l'erreur etant minime, on continue l'éxécution normalement
                
    'Affiche le résultat
    Call SetDIBits(PCT_PictureBox_Bitmap.hdc, PCT_PictureBox_Bitmap.Image, 0, PCT_PictureBox_Bitmap.ScaleHeight, BYT_ImageByts_Image(1, 1, 1), TYP_Info_Bitmap, DIB_RGB_COLORS)
    Let PCT_PictureBox_Bitmap.Picture = PCT_PictureBox_Bitmap.Image

End Sub

Private Function BlueValue(ByRef LNG_Couleur As Long) As Integer

'==================================================================
'Permet d'obtenir la valeur bleu d'une couleur RGB spécifiée
'
'COL_Couleur : Couleur de type RGB
'
'Renvoie la valeur bleu de la couleur envoyé sous forme d'un entier
'Renvoie "-1" en cas d'erreur
'===================================================================

'on active la routine de traitement d'erreur
'On Error GoTo erreur0

    'on renvoie les bits correspondants à la valeur rouge de la couleur : ceux du milieu
    Let BlueValue = ((LNG_Couleur \ 65536) And &HFF)
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu retourner la valeur de bleu
'Solution : on renvoie une valeur d'erreur : "-1"

Let BlueValue = -1

End Function

Private Sub PRO_Extraire_Bits()

'===========================================================================
'Permet d'extraire le tableau de bits de l'image
'===========================================================================

'on active la routine de traitement d'erreur
'On Error Resume Next
'l'erreur etant minime, on continue l'éxécution normalement

    'on initialise la structure du bitmap
    Let TYP_Info_Bitmap.bmiHeader.biSize = Len(TYP_Info_Bitmap.bmiHeader)
    Let TYP_Info_Bitmap.bmiHeader.biWidth = PCT_PictureBox_Bitmap.ScaleWidth
    Let TYP_Info_Bitmap.bmiHeader.biHeight = -PCT_PictureBox_Bitmap.ScaleHeight
    Let TYP_Info_Bitmap.bmiHeader.biPlanes = 1
    Let TYP_Info_Bitmap.bmiHeader.biBitCount = 32
    Let TYP_Info_Bitmap.bmiHeader.biCompression = BI_RGB
    Let TYP_Info_Bitmap.bmiHeader.biSizeImage = ((((TYP_Info_Bitmap.bmiHeader.biWidth * TYP_Info_Bitmap.bmiHeader.biBitCount) + 31) \ 32) * 4) - (((TYP_Info_Bitmap.bmiHeader.biWidth * TYP_Info_Bitmap.bmiHeader.biBitCount) + 7) \ 8) * Abs(TYP_Info_Bitmap.bmiHeader.biHeight)

    ' Load the bitmap's data.
    ReDim BYT_ImageByts_Image(1 To 4, 1 To PCT_PictureBox_Bitmap.ScaleWidth, 1 To PCT_PictureBox_Bitmap.ScaleHeight)
    Call GetDIBits(PCT_PictureBox_Bitmap.hdc, PCT_PictureBox_Bitmap.Image, 0, PCT_PictureBox_Bitmap.ScaleHeight, BYT_ImageByts_Image(1, 1, 1), TYP_Info_Bitmap, DIB_RGB_COLORS)
    
End Sub

Public Function GetPixel(ByRef LNG_X As Long, ByRef LNG_Y As Long) As Long

'=============================================================================================
'Permet de récupèrer la couleur d'un pixel de l'image
'
'LNG_X : Coordonnée X en pixel du point de l'image dont on veut récupèrer la couleur
'LNG_Y : Coordonnée Y en pixel du point de l'image dont on veut récupèrer la couleur
'
'Renvoie la couleur du point en RGB
'Renvoie -1 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
'On Error GoTo erreur0

    'on renvoie la couleur correspondante
    Let GetPixel = RGB(BYT_ImageByts_Image(pixR, LNG_X + 1, LNG_Y + 1), BYT_ImageByts_Image(pixG, LNG_X + 1, LNG_Y + 1), BYT_ImageByts_Image(pixB, LNG_X + 1, LNG_Y + 1))
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On ne connaît pas la couleur du point
'Solution : On renvoie une valeur d'erreur : -1

Let GetPixel = -1

End Function

Public Function GetPixelRGB(ByRef LNG_X As Long, ByRef LNG_Y As Long, ByRef BYT_Red As Byte, ByRef BYT_Green As Byte, ByRef BYT_Blue As Byte) As Long

'=============================================================================================
'Permet de récupèrer la couleur d'un pixel de l'image
'
'LNG_X : Coordonnée X en pixel du point de l'image dont on veut récupèrer la couleur
'LNG_Y : Coordonnée Y en pixel du point de l'image dont on veut récupèrer la couleur
'BYT_Red : Composante rouge de la couleur
'BYT_Green : Composante verte de la couleur
'BYT_Blue : Composante bleue de la couleur
'
'Renvoie 1 si la fonction réussie
'Renvoie 0 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
'On Error GoTo erreur0

    'on renvoie la couleur correspondante
    Let BYT_Red = BYT_ImageByts_Image(pixR, LNG_X + 1, LNG_Y + 1)
    Let BYT_Green = BYT_ImageByts_Image(pixG, LNG_X + 1, LNG_Y + 1)
    Let BYT_Blue = BYT_ImageByts_Image(pixB, LNG_X + 1, LNG_Y + 1)
    Let GetPixelRGB = 1
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On ne connaît pas la couleur du point
'Solution : On renvoie une valeur d'erreur : 0

Let GetPixelRGB = 0

End Function

Private Function GreenValue(ByRef LNG_Couleur As Long) As Integer

'===================================================================
'Permet d'obtenir la valeur vert d'une couleur RGB
'
'COL_Couleur : Couleur de type RGB
'
'Renvoie la valeur vert de la couleur envoyé sous forme d'un entier
'Renvoie "-1" en cas d'erreur
'===================================================================

'on active la routine de traitement d'erreur
'On Error GoTo erreur0

    'on renvoie les bits correspondants à la valeur rouge de la couleur : ceux du milieu à côté de ceux du bleu
    Let GreenValue = ((LNG_Couleur \ 256) And &HFF)
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu retourner la valeur de vert
'Solution : on renvoie une valeur d'erreur : "-1"

Let GreenValue = -1

End Function

Public Property Set PictureBox(ByVal PCT_PictureBox As Object)

'==============================================================================================
'Permet de définir le hDC d'une image stockée dans l'objet ainsi que les dimensions de l'image
'==============================================================================================

    Set PCT_PictureBox_Bitmap = PCT_PictureBox
    
    'on extrait les bits de l'image
    Call PRO_Extraire_Bits
    
End Property

Private Function RedValue(ByRef LNG_Couleur As Long) As Integer

'===================================================================
'Permet d'obtenir la valeur rouge d'une couleur RGB spécifiée
'
'COL_Couleur : Couleur de type RGB
'
'Revoie la valeur rouge de la couleur envoyé sous forme d'un entier
'Renvoie -1 en cas d'erreur
'====================================================================

'on active la routine de traitement d'erreur
'On Error GoTo erreur0

    'on renvoie les bits correspondants à la valeur rouge de la couleur : ceux de droite
    Let RedValue = (LNG_Couleur And &HFF)
    
'la fonction est finie
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu retourner la valeur de rouge
'Solution : on renvoie une valeur d'erreur : "-1"

Let RedValue = -1

End Function

Public Function SetPixel(ByRef LNG_X As Long, ByRef LNG_Y As Long, ByRef LNG_Color As OLE_COLOR) As Long

'=============================================================================================
'Permet de définir la couleur d'un pixel de l'image
'
'LNG_X : Coordonnée X en pixel du point de l'image dont on veut définir la couleur
'LNG_Y : Coordonnée Y en pixel du point de l'image dont on veut définir la couleur
'LNG_Color : Couleur à affecter au pixel
'
'Renvoie 1 si la fonction réussie
'Renvoie 0 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
'On Error GoTo erreur0

    'on applique la couleur du pixel
    Let BYT_ImageByts_Image(pixR, LNG_X + 1, LNG_Y + 1) = RedValue(LNG_Color)
    Let BYT_ImageByts_Image(pixG, LNG_X + 1, LNG_Y + 1) = GreenValue(LNG_Color)
    Let BYT_ImageByts_Image(pixB, LNG_X + 1, LNG_Y + 1) = BlueValue(LNG_Color)

'la fonction a réussie
Let SetPixel = 1
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu définir la couleur du pixel
'Solution : On renvoie une valeur d'erreur : 0

Let SetPixel = 0

End Function

Public Function SetPixelRGB(ByRef LNG_X As Long, ByRef LNG_Y As Long, ByRef BYT_Red As Byte, ByRef BYT_Green As Byte, ByRef BYT_Blue As Byte) As Long

'=============================================================================================
'Permet de définir la couleur d'un pixel de l'image
'
'LNG_X : Coordonnée X en pixel du point de l'image dont on veut définir la couleur
'LNG_Y : Coordonnée Y en pixel du point de l'image dont on veut définir la couleur
'BYT_Red : composante rouge du pixel à appliquer
'BYT_Green : composante verte du pixel à appliquer
'BYT_Blue : composante bleue du pixel à appliquer
'
'Renvoie 1 si la fonction réussie
'Renvoie 0 si la fonction échoue
'=============================================================================================

'on active la routine de traitement d'erreur
'On Error GoTo erreur0

    'on applique la couleur du pixel
    Let BYT_ImageByts_Image(pixR, LNG_X + 1, LNG_Y + 1) = BYT_Red
    Let BYT_ImageByts_Image(pixG, LNG_X + 1, LNG_Y + 1) = BYT_Green
    Let BYT_ImageByts_Image(pixB, LNG_X + 1, LNG_Y + 1) = BYT_Blue

'la fonction a réussie
Let SetPixelRGB = 1
Exit Function

'routine de traitement d'erreur
erreur0:
'Problème : On n'a pas pu définir la couleur du pixel
'Solution : On renvoie une valeur d'erreur : 0

Let SetPixelRGB = 0

End Function

Private Sub Class_Initialize()

    'on redimensionne le tableau de bits de l'image
    ReDim BYT_ImageByts_Image(0)

End Sub

Private Sub Class_Terminate()

    'on détruit les device contexts et les bitmaps crées
    Call DeleteDC(LNG_Handle_DC)
    Call DeleteObject(LNG_Handle_Bitmap)

End Sub
