Attribute VB_Name = "SheetNameAdd"
Option Explicit

Public Sub addSheetScopedName()
Attribute addSheetScopedName.VB_ProcData.VB_Invoke_Func = "N\n14"
    ' Adds Name to *sheet* scope for the selected cell,
    ' based on the contents of the cell to the immediate left.
    ' Will CLOBBER any existing name!
    
    Dim wkSht As Worksheet
    Dim cel As Range
    Dim newName As String
    Dim wkName As Name
    
    Dim errNum As Long
    
    Set wkSht = ActiveSheet
    
    For Each cel In Selection
        If IsEmpty(cel.Offset(0, -1)) Then
            MsgBox "No text in " & cel.Offset(0, -1).Address & "!" & _
                    Chr(10) & Chr(10) & _
                    "Skipping.", vbOKOnly + vbInformation, _
                    "Skipping Empty Name Cell"
            GoTo Skip_Cell
        End If
    
        ' Define the new name
        newName = cleanNameName(cel.Offset(0, -1))
    
        ' Remove any old Names assigned to the cell
        ' May turn out this is undesirable;
        ' usage will see
        For Each wkName In Names
            If isNameSheetScopeAndOnCell(wkName, cel) Then
            wkName.Delete
            End If
        Next wkName
    
        ' Apply the new name, revising if needed
        On Error Resume Next
            wkSht.Names.Add newName, cel
        errNum = Err.Number: Err.Clear: On Error GoTo 0
        If errNum = 1004 Then
            ' Invalid name, append an underscore. This handles
            ' (3) from the 'docstring' of cleanNameName.
            wkSht.Names.Add newName & "_", cel
        ElseIf errNum > 0 And errNum <> 40040 Then
            ' 40040 is apparently set as Err.Number when Name
            ' creation succeeds...?
            ' Anyways, this re-raises any unexpected errors.
            Err.Raise errNum
        End If
        
Skip_Cell:
    Next cel
    
End Sub

Function isNameSheetScopeAndOnCell(n As Name, c As Range) As Boolean
    ' Return True if:
    '  1) n is scoped to the worksheet, not workbook; and
    '  2) refers exactly to just the cell c
    
    If InStr(n, "!") < 1 Then
        isNameSheetScopeAndOnCell = False
        Exit Function
    End If
    
    isNameSheetScopeAndOnCell = (c.Address = Mid(n.RefersTo, 1 + InStr(n.RefersTo, "!")))
    
End Function

Function cleanNameName(ByVal n As String) As String
    ' The names of Names must:
    '   1. Start with underscore or letter
    '   2. Must not have spaces or other invalid characters
    '   3. Must not conflict with Excel names
    '   4. Must not start with a valid RC-style reference
    '
    ' This function takes care of (1) and (2) and (4)
    '
    
    Dim iter As Long, c As String
    Dim rxRC As New RegExp
    
    With rxRC
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^([RC]|RC)0*[1-9]"
    End With
    
    cleanNameName = n
    
    ' Fix (1) if needed by prepending underscore
    c = Mid(cleanNameName, 1, 1)
    If Not (isCharLetter(c) Or isCharUnderscore(c)) Then
        cleanNameName = "_" & cleanNameName
    End If
    
    ' Fix (2) if needed by substituting underscores
    For iter = 1 To Len(cleanNameName)
        c = Mid(cleanNameName, iter, 1)
        
        If Not ( _
            isCharLetter(c) _
            Or isCharUnderscore(c) _
            Or isCharNumeral(c) _
        ) Then
            cleanNameName = swapChar(cleanNameName, iter, "_")
        End If
            
    Next iter
    
    ' Fix (4) if needed by prepending underscore
    If rxRC.Test(cleanNameName) Then
        cleanNameName = "_" & cleanNameName
    End If

End Function

Function isCharLetter(ByVal c As String) As Boolean

    isCharLetter = False
    
    If Asc(c) >= 97 And Asc(c) <= 122 Then isCharLetter = True
    If Asc(c) >= 65 And Asc(c) <= 90 Then isCharLetter = True

End Function

Function isCharUnderscore(ByVal c As String) As Boolean

    isCharUnderscore = False
    
    If Asc(c) = 95 Then isCharUnderscore = True

End Function

Function isCharNumeral(ByVal c As String) As Boolean

    isCharNumeral = False
    
    If Asc(c) >= 48 And Asc(c) <= 57 Then isCharNumeral = True

End Function

Function swapChar(ByVal s As String, pos As Long, c As String) As String
    ' Swap the single character in 's' at position 'pos' with
    ' character (or string...) 'c'.
    
    swapChar = Left(s, pos - 1) & c & Right(s, Len(s) - pos)
    
End Function
