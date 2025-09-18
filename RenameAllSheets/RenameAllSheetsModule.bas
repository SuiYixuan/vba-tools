Attribute VB_Name = "Module1"
Option Explicit

Private undoStack As Object
Private stackCounter As Integer

Public Sub showForm()
    'loadUndoStackFromStorage
    UserForm1.Show
End Sub

Public Sub renameAllSheets(newStr As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim previousSheetNames As Object
    
    Set wb = ActiveWorkbook
    Set previousSheetNames = CreateObject("Scripting.Dictionary")
    If undoStack Is Nothing Then
        Set undoStack = CreateObject("Scripting.Dictionary")
        stackCounter = 1
    End If

    If newStr <> "" Then
        If InStr(newStr, "?") > 0 Then
            For Each ws In wb.Worksheets
                previousSheetNames.Add ws.codeName, ws.Name
                Dim holderStr
                holderStr = Replace(newStr, "?", ws.Name)
                If Len(holderStr) < 31 Then
                    ws.Name = Replace(newStr, "?", ws.Name)
                Else
                    MsgBox "最大長を超える"
                    Exit Sub
                End If
            Next ws
            undoStack.Add stackCounter, previousSheetNames
            stackCounter = stackCounter + 1
        Else
            MsgBox "元の名前のプレースホルダ「?」を含めてください"
        End If
    End If
End Sub

Public Sub undoRename()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim currentKey As Variant
    
    Dim maxKey As Integer
    Dim lastTimeDictionary As Object
    Set wb = ActiveWorkbook
    maxKey = 0
    
    If Not undoStack Is Nothing Then
        For Each currentKey In undoStack.Keys
            If currentKey > maxKey Then
                maxKey = currentKey
            End If
        Next currentKey
        
        If maxKey <> 0 Then
            Set lastTimeDictionary = undoStack(maxKey)
        
            For Each ws In wb.Worksheets
                If lastTimeDictionary.Exists(ws.codeName) Then
                    ws.Name = lastTimeDictionary(ws.codeName)
                End If
            Next ws
            undoStack.Remove maxKey
        End If
    End If
End Sub

Public Sub saveUndoStackToStorage()
    Dim innerDict As Variant
    Dim hiddenSheet As Worksheet
    Dim codeName As Variant
    Dim originalName As String
    Dim cursor As Integer
    cursor = 2
    Set hiddenSheet = ThisWorkbook.Sheets("UndoStackStorage")
    If hiddenSheet Is Nothing Then
        Set hiddenSheet = ThisWorkbook.Sheets.Add
        hiddenSheet.Name = "UndoStackStorage"
        hiddenSheet.Visible = xlSheetVeryHidden
    Else
        hiddenSheet.Cells.Clear
    End If
    hiddenSheet.Range("A1").Value = "cursor"
    hiddenSheet.Range("B1").Value = "codeName"
    hiddenSheet.Range("C1").Value = "originalName"
    
    If Not undoStack Is Nothing Then
        For Each innerDict In undoStack.Items
            For Each codeName In innerDict.Keys
                hiddenSheet.Range("A" & cursor) = cursor - 1
                hiddenSheet.Range("B" & cursor) = codeName
                hiddenSheet.Range("C" & cursor) = innerDict(codeName)
                cursor = cursor + 1
            Next codeName
        Next innerDict
    End If
End Sub

Private Sub loadUndoStackFromStorage()
    Dim hiddenSheet As Worksheet
    Dim columnA As Range
    Dim innerDict As Object
    Dim cursor As Integer
    Dim lastTimeCursor As Integer
    Dim codeName As Variant
    Dim originalName As String
    Dim index As Integer
    index = 2
    
    If Not hiddenSheet Is Nothing Then
        Set hiddenSheet = ThisWorkbook.Sheets("UndoStackStorage")
        Set undoStack = CreateObject("Scripting.Dictionary")
        Set innerDict = CreateObject("Scripting.Dictionary")
        stackCounter = 1
    
        For Each columnA In hiddenSheet.Range("A2:A" & Application.CountA(hiddenSheet.Range("A:A")) - 1)
            cursor = hiddenSheet.Range("A" & index).Value
            If index > 2 Then
                lastTimeCursor = hiddenSheet.Range("A" & index - 1).Value
            Else
                lastTimeCursor = 1
            End If
            codeName = hiddenSheet.Range("B" & index).Value
            originalName = hiddenSheet.Range("C" & index).Value
        
            If cursor = lastTimeCursor Then
                innerDict.Add codeName, originalName
                index = index + 1
            Else
                undoStack.Add stackCounter, innerDict
                stackCounter = stackCounter + 1
                Set innerDict = CreateObject("Scripting.Dictionary")
            End If
        Next columnA
        undoStack.Add stackCounter, innerDict
        stackCounter = stackCounter + 1
    End If
End Sub
