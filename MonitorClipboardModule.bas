Attribute VB_Name = "Module1"
Option Explicit

Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Private Declare PtrSafe Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare PtrSafe Function AddClipboardFormatListener Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function RemoveClipboardFormatListener Lib "user32" (ByVal hwnd As LongPtr) As Long
Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrW" (ByVal hwnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcW" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function GetLastError Lib "kernel32" () As Long

Private Const WM_CLIPBOARDUPDATE As Long = &H31D
Private Const CF_BITMAP As Long = 2
Private Const GWL_WNDPROC As Long = -4

Private hwndNextViewer As LongPtr
Private prevWndProc As LongPtr
Public isMonitoring As Boolean
Private status As Integer
Private formStatus As Boolean

Public Sub showForm()
    UserForm1.Show
End Sub

Public Sub StartMonitor()
    AddClipboardFormatListener (Application.hwnd)
    ' Excel窗口句柄
    Debug.Print "Application.hwnd:" & Application.hwnd
    Debug.Print "WindowProc:" & Hex(AddressOf WindowProc)
    prevWndProc = SetWindowLongPtr(Application.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    Debug.Print "SetWindowLongPtr: " & GetLastError()
    isMonitoring = True
    formStatus = UserForm1.OptionButton1.Value
    UserForm1.Hide
    MsgBox "Clipboard Monitor Start", vbInformation
    If formStatus = True Then
        UpdateStatus "(標準モード)"
    Else
        UpdateStatus "(選択モード)"
    End If
End Sub

Public Sub EndMonitor()
    SetWindowLongPtr Application.hwnd, GWL_WNDPROC, prevWndProc
    RemoveClipboardFormatListener (Application.hwnd)
    isMonitoring = False
    ClearStatus
End Sub

Private Function WindowProc(ByVal hwnd As LongPtr, ByVal Msg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    If Msg = WM_CLIPBOARDUPDATE Then
        If IsClipboardImage() Then
            status = status + 1
            If status = 2 Then
                If formStatus = True Then
                    PasteStandardMode
                Else
                    PasteSelectMode
                End If
                status = 0
            End If
        End If
    End If

    WindowProc = CallWindowProc(prevWndProc, hwnd, Msg, wParam, lParam)
End Function

Private Function IsClipboardImage() As Boolean
    If OpenClipboard(0) <> 0 Then
        IsClipboardImage = (IsClipboardFormatAvailable(CF_BITMAP) <> 0)
        CloseClipboard
    End If
End Function

Private Sub PasteStandardMode()
    On Error GoTo ErrorHandler
        Dim nextPicRow As Integer
        nextPicRow = GetNextPicRow()
        If nextPicRow = 3 Then
            ActiveSheet.Paste Destination:=ActiveSheet.Range("B4")
        Else
            ActiveSheet.Paste Destination:=ActiveSheet.Range("B" & nextPicRow)
        End If
        ActiveSheet.Range("B" & nextPicRow).Select
        AdjustPic
        Debug.Print "PasteStandardMode End"
ErrorHandler:
    Debug.Print Err.Number & ":" & Err.Description
End Sub

Private Function GetNextPicRow() As Integer
    Dim dataRowA As Integer
    Dim dataRowB As Integer
    Dim shp As Shape
    Dim picRow As Integer
    dataRowA = ActiveSheet.Cells(ActiveSheet.Rows.Count, "A").End(xlUp).row
    dataRowB = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).row
    For Each shp In ActiveSheet.shapes
        If shp.Type = msoPicture And shp.BottomRightCell.row > picRow Then
            picRow = shp.BottomRightCell.row
        End If
    Next shp
    GetNextPicRow = Application.WorksheetFunction.Max(dataRowA, dataRowB, picRow) + 2
End Function

Private Sub PasteSelectMode()
    On Error GoTo ErrorHandler
        If TypeName(Selection) <> "Range" Then
            Exit Sub
        End If
        Dim selectedCell As Range
        Dim topLeftCell As Range
        Set selectedCell = Selection
        Set topLeftCell = Selection.Cells(1, 1)
        topLeftCell.Select
        ActiveSheet.Paste
        AdjustPic
        Debug.Print "PasteSelectMode End"
ErrorHandler:
    Debug.Print Err.Number & ":" & Err.Description
End Sub

Private Sub AdjustPic()
    Dim pastedImage As Shape
    Set pastedImage = ActiveSheet.shapes(ActiveSheet.shapes.Count)
    pastedImage.LockAspectRatio = msoCTrue
    pastedImage.Width = 1050
End Sub

Private Sub UpdateStatus(message As String)
    Application.StatusBar = "Clipboard Monitor Start" & message & "-" & ActiveWorkbook.Name & " | " & Format(Now, "hh:mm:ss")
End Sub


Private Sub ClearStatus()
    Application.StatusBar = False
End Sub

'Public Sub ExitClean()
    'If isMonitoring Then
        'SetWindowLongPtr Application.hwnd, GWL_WNDPROC, prevWndProc
        'RemoveClipboardFormatListener (Application.hwnd)
        'isMonitoring = False
        'ClearStatus
    'End If
'End Sub
