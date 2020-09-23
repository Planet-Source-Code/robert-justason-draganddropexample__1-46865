VERSION 5.00
Begin VB.Form frmdragdropexample 
   AutoRedraw      =   -1  'True
   Caption         =   "Drag and drop example"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtdropped 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      OLEDropMode     =   1  'Manual
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   120
      Width           =   7575
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "Clear"
      Height          =   255
      Left            =   5160
      TabIndex        =   3
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   255
      Left            =   6600
      TabIndex        =   2
      Top             =   4440
      Width           =   1095
   End
   Begin VB.TextBox txtfile 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4440
      Width           =   1095
   End
End
Attribute VB_Name = "frmdragdropexample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclear_Click()
'clears the textboxes
txtdropped = ""
txtfile = ""
End Sub

Private Sub Cmdexit_Click()
End 'exit the program
End Sub

Private Sub txtdropped_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim dragfilename, strcompare As String

On Error Resume Next

strcompare = "\" 'character used to find the filename in the full path
dragfilename = Data.Files(1) 'put the filename and path into dragfilename string

txtfile.Text = Right(dragfilename, (Len(dragfilename) - (InStrRev(dragfilename, strcompare))))
'this statement finds the last occurence of "\" in the full path, extracts the filename after it and puts it into the "txtfile" textbox

Open dragfilename For Input As #1 'open the dragged file
    txtdropped.Text = Input(LOF(1), 1) 'place contents of file into the textbox called "textdropped"
    Close #1 'close the file
    
End Sub

