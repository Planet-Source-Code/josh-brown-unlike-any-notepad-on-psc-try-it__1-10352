Attribute VB_Name = "Module1"
'This is used by most of the Constants below
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'This is used for testing if Pasting is enabled or not
Public Const WM_USER = &H400

'This is the Undo command from the Windows API
Public Const EM_UNDO = &HC7
'This tests if you can Undo or Not
Public Const EM_CANUNDO = &HC6

'This is the Cut command from the Windows API
Public Const WM_CUT = &H300
'This is the Copy command from the Windows API
Public Const WM_COPY = &H301
'This is the Paste command from the Windows API
Public Const WM_PASTE = &H302
'This tests if you can Paste or Not
Public Const EM_CANPASTE = (WM_USER + 50)

'This is the Clear/Delete command from the Windows API
Public Const WM_CLEAR = &H303

'This is where the Undo Magic Takes Place
Sub EditUndo()
    'This is the main command for Undoing
    SendMessage Form1.rtb1.hwnd, EM_UNDO, 0, 0
    'Set the Focus the the forms RichTextBox Control
    Form1.rtb1.SetFocus
End Sub

'This is where we test to see if we are capable of Undoing
Public Function CanUndo() As Boolean
    CanUndo = SendMessage(Form1.rtb1.hwnd, EM_CANUNDO, 0, 0)
End Function

'This is where the Cut Magic Takes Place
Sub EditCut()
    'This is the main command for Cuting
    SendMessage Form1.rtb1.hwnd, WM_CUT, 0, 0
    'Set the Focus the the forms RichTextBox Control
    Form1.rtb1.SetFocus
End Sub

'This is where the Copy Magic Takes Place
Sub EditCopy()
    'T is used for setting the cursors position
    Dim T As Long
        T = Form1.rtb1.SelLength

    'This is the main command for Copying
    SendMessage Form1.rtb1.hwnd, WM_COPY, 0, 0
    
    'Set the cursors position
    Form1.rtb1.SelStart = T + Form1.rtb1.SelStart
    'Set the Focus the the forms RichTextBox Control
    Form1.rtb1.SetFocus
End Sub

'This is where the Paste Magic Takes Place
Sub EditPaste()
    'This is the main command for Pasting
    SendMessage Form1.rtb1.hwnd, WM_PASTE, 0, 0
    'Set the Focus the the forms RichTextBox Control
    Form1.rtb1.SetFocus
End Sub

'This is where we test to see if we are capable of Pasting
Public Function CanPaste() As Boolean
    CanPaste = SendMessage(Form1.rtb1.hwnd, EM_CANPASTE, 0, 0)
End Function

'This is where the Clear/Delete Magic Takes Place
Sub EditClear()
    'This is the main command for Clearing/Deleting
    SendMessage Form1.rtb1.hwnd, WM_CLEAR, 0, 0
    'Set the Focus the the forms RichTextBox Control
    Form1.rtb1.SetFocus
End Sub

Sub EditSelectAll()
    With Form1.rtb1
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

