VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form fRMmANI 
   Caption         =   "KILL - KLON"
   ClientHeight    =   9750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   9750
   ScaleWidth      =   11130
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1e-4
      Max             =   100
   End
   Begin VB.TextBox pathy 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Text            =   "Enter a path here for a fast navigation"
      Top             =   5040
      Width           =   6135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "KILL!!!"
      Height          =   495
      Left            =   9840
      TabIndex        =   5
      Top             =   6120
      Width           =   1215
   End
   Begin VB.ListBox ListaNegra 
      Height          =   840
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   9615
   End
   Begin VB.CommandButton CmdComprobar 
      Caption         =   "START !"
      Height          =   495
      Left            =   9840
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
   Begin VB.DirListBox Dir1 
      Height          =   4365
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   6135
   End
   Begin VB.FileListBox File1 
      Height          =   5160
      Left            =   6360
      TabIndex        =   2
      Top             =   240
      Width           =   4695
   End
   Begin VB.Image Muestra2 
      Height          =   2175
      Left            =   4320
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   2655
   End
   Begin VB.Image Muestra1 
      Height          =   2175
      Left            =   1560
      Stretch         =   -1  'True
      Top             =   7440
      Width           =   2655
   End
End
Attribute VB_Name = "fRMmANI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdComprobar_Click()
On Error Resume Next
Dim I As Integer, j As Integer, Original As String
Dim Size As Long
Dim Orig1(255) As Byte, Comp1(255) As Byte
Dim Orig2(255) As Byte, Comp2(255) As Byte
Dim Orig3(255) As Byte, Comp3(255) As Byte
CmdComprobar.Enabled = False
'Recorre todos los archivos.

PB.Max = File1.ListCount - 1
PB.Value = 0

For I = 0 To File1.ListCount - 1

    ' Open the file to compare
    Size = FileLen(Dir1.Path & "\" & File1.List(I))
    Original = File1.List(I)
    Size = 256
    
    Open Dir1.Path & "\" & File1.List(I) For Binary Access Read As #1
    
        Get #1, Size, Orig1()
        Get #1, (Size * 2), Orig2()
        Get #1, (100 + (Size * Size)), Orig3()
    
    Close #1


    ' open the Comparizon file
    For j = I To File1.ListCount - 1
    
        If Original <> File1.List(j) Then
        
            Open Dir1.Path & "\" & File1.List(j) For Binary Access Read As #1
            
                Get #1, Size, Comp1()
                Get #1, (Size * 2), Comp2()
                Get #1, (100 + (Size * Size)), Comp3()
            
            Close #1
        
        
            'Se realiza la comparacion
            'It compares
            
            Dim k As Integer
            Dim CompResult As Integer
            'Si la comparacion da como resultado 257, entonces son iguales.
            'If the comparison is equal to 257, then is equal.
            
            For k = 0 To 255
            
                If Comp1(k) = Orig1(k) Then
                    CompResult = CompResult + 1
                End If
                
                If Comp2(k) = Orig2(k) Then
                    CompResult = CompResult + 1
                End If
                If Comp3(k) = Orig3(k) Then
                    CompResult = CompResult + 1
                End If
                    
            Next k
        
            If CompResult = 768 And FileLen(Dir1.Path & "\" & File1.List(I)) = FileLen(Dir1.Path & "\" & File1.List(j)) Then
                
                Muestra1.Picture = LoadPicture(Dir1.Path & "\" & File1.List(I))
                Muestra2.Picture = LoadPicture(Dir1.Path & "\" & File1.List(j))
                If MsgBox("The File is the same?:" & vbNewLine & Original & vbNewLine & File1.List(j), vbInformation + vbYesNo, "Put File 2 to Deletion List?") = vbYes Then
                    ListaNegra.AddItem Dir1.Path & "\" & File1.List(j)
                End If
                
                
            End If
            
        End If
    CompResult = 0
    
    Next j
    PB.Value = I
Next I


MsgBox "Check Done!", vbInformation
File1.SetFocus
CmdComprobar.Enabled = True

End Sub

Private Sub Command1_Click()
On Error Resume Next
If MsgBox("Shure?", vbQuestion + vbYesNo) = vbYes Then

    Dim I As Integer
    For I = 0 To ListaNegra.ListCount - 1
        Kill ListaNegra.List(I)
    Next I

End If

MsgBox "Done!"
ListaNegra.Clear

End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1
    
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
    
End Sub

Private Sub pathy_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

    Dir1.Path = pathy.Text

End If
End Sub

