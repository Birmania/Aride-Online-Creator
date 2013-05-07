VERSION 5.00
Begin VB.Form frmCraftEditor 
   Caption         =   "Editer un patron"
   ClientHeight    =   7125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleWidth      =   11325
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Ajouts"
      Height          =   2535
      Left            =   5640
      TabIndex        =   11
      Top             =   1920
      Width           =   5535
      Begin VB.CommandButton cmdAddProduct 
         Caption         =   "Ajouter un &produit"
         Height          =   375
         Left            =   3120
         TabIndex        =   18
         Top             =   1800
         Width           =   1815
      End
      Begin VB.CommandButton cmdAddMaterial 
         Caption         =   "Ajouter un &composant"
         Height          =   435
         Left            =   3120
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
      Begin VB.TextBox txtCount 
         Height          =   285
         Left            =   3120
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.ListBox lstItem 
         Height          =   1425
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre requis :"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Objets :"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdDeleteProduct 
      Caption         =   "S&upprimer un produit"
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton cmdEditProduct 
      Caption         =   "E&diter un produit"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   4200
      Width           =   1935
   End
   Begin VB.ListBox lstProduct 
      Height          =   1620
      Left            =   360
      TabIndex        =   8
      Top             =   3960
      Width           =   2775
   End
   Begin VB.CommandButton cmdEditMaterial 
      Caption         =   "&Editer un composant"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   1680
      Width           =   1935
   End
   Begin VB.CommandButton cmdDeleteMaterial 
      Caption         =   "&Supprimer un composant"
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Annuler"
      Height          =   255
      Left            =   5880
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.ListBox lstMaterial 
      Height          =   1620
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   2895
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label5 
      Caption         =   "Produits :"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Composants :"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nom du patron :"
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmCraftEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'    Copyright (C) 2013  BRULTET Antoine
'
'    This file is part of Aride Online Creator.
'
'    Aride Online Creator is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.
'
'    Aride Online Creator is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Aride Online Creator.  If not, see <http://www.gnu.org/licenses/>.

Private Sub cmdAddMaterial_Click()
    Dim I As Integer
    
    If lstItem.ListIndex >= 0 Then
        If IsNumeric(txtCount.Text) Then
            If txtCount.Text > 0 Then
                ' Look if item not already in the list and find an empty slot
                For I = 0 To MAX_MATERIALS
                    If lstMaterial.ItemData(I) = lstItem.ListIndex Then
                        MsgBox "Cet objet est déjà un composant de la schématique."
                        
                        Exit For
                    Else
                        If lstMaterial.ItemData(I) = -1 Then
                            ' Empty slot find
                            lstMaterial.RemoveItem (I)
                            lstMaterial.AddItem I & " : " & Trim$(Item(lstItem.ListIndex).name) & " ( " & txtCount.Text & " )", I
                            lstMaterial.ItemData(I) = lstItem.ListIndex
                            
                            Exit For
                        End If
                    End If
                Next I
            Else
                MsgBox "Le nombre d'objet doit être supérieur à 0."
            End If
        Else
            MsgBox "Le nombre d'objet doit être un chiffre."
        End If
    End If
End Sub

Private Sub cmdAddProduct_Click()
    Dim I As Integer
    
    If lstItem.ListIndex >= 0 Then
        If IsNumeric(txtCount.Text) Then
            If txtCount.Text > 0 Then
                ' Look if item not already in the list of products and find an empty slot
                For I = 0 To MAX_MATERIALS
                    If lstProduct.ItemData(I) = lstItem.ListIndex Then
                        MsgBox "Cet objet est déjà un produit de la schématique."
                        
                        Exit For
                    Else
                        If lstProduct.ItemData(I) = -1 Then
                            ' Empty slot find
                            lstProduct.RemoveItem (I)
                            lstProduct.AddItem I & " : " & Trim$(Item(lstItem.ListIndex).name) & " ( " & txtCount.Text & " )", I
                            lstProduct.ItemData(I) = lstItem.ListIndex
                            
                            Exit For
                        End If
                    End If
                Next I
            Else
                MsgBox "Le nombre d'objet doit être supérieur à 0."
            End If
        Else
            MsgBox "Le nombre d'objet doit être un chiffre."
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
    frmCraftEditor.Visible = False
End Sub

Private Sub cmdDeleteMaterial_Click()
    Dim MaterialPosition As Integer
    
    MaterialPosition = lstMaterial.ListIndex

    If MaterialPosition >= 0 Then
        Call ClearMaterial(MaterialPosition)
    End If
End Sub

Public Sub ClearMaterial(position As Integer)
    Call lstMaterial.RemoveItem(position)
    lstMaterial.AddItem position & " : Aucun", position
    lstMaterial.ItemData(position) = -1
End Sub

Private Sub cmdDeleteProduct_Click()
    Dim ProductPosition As Integer
    
    ProductPosition = lstProduct.ListIndex

    If MaterialPosition >= 0 Then
        Call ClearProduct(ProductPosition)
    End If
End Sub

Public Sub ClearProduct(position As Integer)
    Call lstProduct.RemoveItem(position)
    lstProduct.AddItem position & " : Aucun", position
    lstProduct.ItemData(position) = -1
End Sub

Private Sub cmdEditMaterial_Click()
    Dim MaterialPosition As Integer
    Dim nbMaterial As Integer
    Dim ItemNum As Integer
    
    MaterialPosition = lstMaterial.ListIndex

    If MaterialPosition >= 0 Then
        If lstMaterial.ItemData(MaterialPosition) >= 0 Then
            nbMaterial = InputBox("Combien d'unités de ce composant sont nécessaire ?", "Saisie du nombre d'un composant")
            
            If IsNumeric(nbMaterial) And nbMaterial > 0 Then
                ItemNum = lstMaterial.ItemData(MaterialPosition)
                lstMaterial.RemoveItem (MaterialPosition)
                lstMaterial.AddItem MaterialPosition & " : " & Trim$(Item(ItemNum).name) & " ( " & nbMaterial & " )", MaterialPosition
                lstMaterial.ItemData(MaterialPosition) = ItemNum
            Else
                MsgBox "Erreur : La saisie doit être un chiffre.", vbCritical, "Erreur"
            End If
        End If
    End If
End Sub

Private Sub cmdEditProduct_Click()
    Dim ProductPosition As Integer
    Dim nbProduct As Integer
    Dim ItemNum As Integer
    
    ProductPosition = lstProduct.ListIndex

    If ProductPosition >= 0 Then
        If lstProduct.ItemData(ProductPosition) >= 0 Then
            nbProduct = InputBox("Combien d'unités de ce produit seront produites ?", "Saisie du nombre d'un produit")
            
            If IsNumeric(nbProduct) And nbProduct > 0 Then
                ItemNum = lstProduct.ItemData(ProductPosition)
                lstProduct.RemoveItem (ProductPosition)
                lstProduct.AddItem ProductPosition & " : " & Trim$(Item(ItemNum).name) & " ( " & nbProduct & " )", ProductPosition
                lstProduct.ItemData(ProductPosition) = ItemNum
            Else
                MsgBox "Erreur : La saisie doit être un chiffre.", vbCritical, "Erreur"
            End If
        End If
    End If
End Sub

Private Sub cmdOk_Click()
    Dim I As Integer
    Dim itemCount As String
    Dim regEx As RegExp
    Dim occurences As MatchCollection
    
    Set regEx = New RegExp
    regEx.Pattern = "\(.*\)"
    regEx.Global = False

    Call ClearCraft(txtName.DataField)

    Crafts(txtName.DataField).name = txtName.Text
    
    ' Get the materials
    For I = 0 To MAX_MATERIALS
        If lstMaterial.ItemData(I) > -1 Then
            ReDim Crafts(txtName.DataField).Materials(0 To GetNbMaterials(txtName.DataField)) As MaterialRec
        Else
            Exit For
        End If
        
        Crafts(txtName.DataField).Materials(I).ItemNum = lstMaterial.ItemData(I)
        Set occurences = regEx.Execute(lstMaterial.List(I))
        If occurences.Count = 1 Then
            itemCount = occurences.Item(0)
            
            ' Delete the parenthesis
            itemCount = Right(Left(itemCount, Len(itemCount) - 1), Len(itemCount) - 2)
            Crafts(txtName.DataField).Materials(I).Count = Val(itemCount)
        End If
    Next I
    
    ' Get the products
    For I = 0 To MAX_MATERIALS
        If lstProduct.ItemData(I) > -1 Then
            ReDim Crafts(txtName.DataField).Products(0 To GetNbProducts(txtName.DataField)) As MaterialRec
        Else
            Exit For
        End If
        
        Crafts(txtName.DataField).Products(I).ItemNum = lstProduct.ItemData(I)
        Set occurences = regEx.Execute(lstProduct.List(I))
        If occurences.Count = 1 Then
            itemCount = occurences.Item(0)
            
            ' Delete the parenthesis
            itemCount = Right(Left(itemCount, Len(itemCount) - 1), Len(itemCount) - 2)
            Crafts(txtName.DataField).Products(I).Count = Val(itemCount)
        End If
    Next I
    
    SendSaveCraft txtName.DataField
    frmMirage.Editeurcraft_Click
    frmCraftEditor.Visible = False
End Sub
