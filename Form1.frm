VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3840
   LinkTopic       =   "Form1"
   ScaleHeight     =   2385
   ScaleWidth      =   3840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Create Despatch Note"
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' If the same variable name is used more than once in the template, this
' array saves the application performing the same work again to get that
' data.  It simply lifts it from this array.
Private UsedVariables() As String


Private Sub Command1_Click()

    FillTemplates
    
End Sub

Private Sub FillTemplates()

    Dim WordApp As Word.Application
    Dim WordDoc As Word.Document
    Dim i As Integer, j As Integer
    Dim NewResult As String
    
    
    On Error GoTo ErrHandler
    
    ReDim UsedVariables(0)
    
    Set WordApp = CreateObject("Word.Application")
    Set WordDoc = WordApp.Documents.Open(App.Path & "\template.doc")
    
    
    ' For each section (header and footer)
    For i = 1 To WordDoc.Sections.Count
    
        ' Headers
        Debug.Print "Fields in Header:" & WordDoc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Fields.Count
        For j = 1 To WordDoc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Fields.Count
        
            If WordDoc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Fields(j).Type = wdFieldDocVariable Then
            
                ' Get the text for the field from the user
                NewResult = GetNewResult(WordDoc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Fields(j), WordDoc)
                'Insert New Text into the field
                If NewResult <> "" Then
                    WordDoc.Sections(i).Headers(wdHeaderFooterPrimary).Range.Fields(j).Result.Text = NewResult
                End If
                
            End If
        
        Next
        
        ' Footers
        Debug.Print "Fields in Footer:" & WordDoc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Fields.Count
        For j = 1 To WordDoc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Fields.Count
        
            If WordDoc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Fields(j).Type = wdFieldDocVariable Then
        
                ' Get the text for the field from the user
                NewResult = GetNewResult(WordDoc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Fields(j), WordDoc)
                'Insert New Text into the field
                If NewResult <> "" Then
                    WordDoc.Sections(i).Footers(wdHeaderFooterPrimary).Range.Fields(j).Result.Text = NewResult
                End If
            
            End If
        
        
        Next
    
    Next
                
    ' In main body
    Debug.Print "Fields in main body: " & WordDoc.Fields.Count
    For i = 1 To WordDoc.Fields.Count
            
        If WordDoc.Fields(i).Type = wdFieldDocVariable Then
    
            ' Get the text for the field from the user
            NewResult = GetNewResult(WordDoc.Fields(i), WordDoc)
            'Insert New Text into the field
            If NewResult <> "" Then
                WordDoc.Fields(i).Result.Text = NewResult
            End If
                
        End If
                
    Next
        
    ' lock the document to stop changes
    WordDoc.Protect wdAllowOnlyComments, , "jd837djh82"
    WordDoc.SaveAs App.Path & "\despatchnote.doc"
    
    WordDoc.Close
    
    WordApp.Quit
    Set WordDoc = Nothing
    Set WordApp = Nothing

    MsgBox "Finished!"

Exit Sub
ErrHandler:
    
    MsgBox "Unhanled Error: " & Err.Description

End Sub

Private Function GetNewResult(wField As Word.Field, WordDoc As Word.Document) As String

    Dim StopPos As Long
    Dim Variable As String
    Dim UsedVariable As String
    Dim VariableValue As String
    Dim wRange As Word.Range
    
    Debug.Print wField.Code
    
    ' These three lines strip down the field code to find
    ' out it's name
    StopPos = InStrRev(wField.Code, "\*")
    Variable = Left(wField.Code, StopPos - 3)
    Variable = Right(Variable, Len(Variable) - 14)
    
    ' Check this field hasn't already appeared in this
    ' document.
    If CheckUsedVariable(Variable) Then
                  
        VariableValue = GetVariableValue(Variable)
        
    Else
        
        Select Case UCase(Variable)
        
            ' I don't simply want to insert a string -
            ' I wish to insert a table at the Product Field.
            Case "PRODUCT"
                                            
                ' Get the range (location) of the product field
                Set wRange = wField.Code
                ' Delete the field, as any text will be inserted into the
                ' {} of the existing field.
                wField.Delete
                
                ' Enter our table information including headers.
                ' Ideally, I would get this data from an ADO recordset
                ' using GetString().
                With wRange
                
                    .Text = "PRODUCT" & vbTab & "CTSBATCHNO" & vbTab & "SUPP REF" & vbTab & "PACKNO" & vbTab & "STORAGE" & vbTab & "QTY UNITS" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3" & vbCrLf & _
                                "989797897" & vbTab & "hjhkhk" & vbTab & "kjhkjhkh" & vbTab & "kjhkjh" & vbTab & "Frozen" & vbTab & "3"
                                
                    .FormattedText.Font.Name = "Arial"
                    .FormattedText.Font.Size = "8"
                
                    ' Once the data is there, we can convert it to a table
                    ' structure and format it to look pretty!
                    .ConvertToTable vbTab, , , , wdTableFormatColorful2
                
                End With
                
                ' Send back blank string as field does not exist anymore
                VariableValue = ""
        
            Case Else
                
                ' Get the value of the field from the user
                VariableValue = InputBox("Enter value for: " & Variable, "Value not recognised for Despatch Note!")
                AddNewVariable Variable, VariableValue
        
        End Select
        
    End If
    
    GetNewResult = VariableValue
        
End Function

Private Function GetVariableValue(Variable As String) As String
Dim i As Integer

    For i = 0 To UBound(UsedVariables)
        If Left(UsedVariables(i), Len(Variable)) = Variable Then
            GetVariableValue = Right(UsedVariables(i), Len(UsedVariables(i)) - Len(Variable))
            Exit For
        End If
    Next
    
End Function

Private Sub AddNewVariable(Variable As String, TheValue As String)
Dim ArraySize As Integer

    ArraySize = UBound(UsedVariables)
    ReDim Preserve UsedVariables(ArraySize + 1)
    UsedVariables(ArraySize) = Variable & TheValue

End Sub

Private Function CheckUsedVariable(Variable As String) As Boolean
Dim i As Integer

    For i = 0 To UBound(UsedVariables)
        If Left(UsedVariables(i), Len(Variable)) = Variable Then
            CheckUsedVariable = True
            Exit For
        End If
    Next
    
End Function

