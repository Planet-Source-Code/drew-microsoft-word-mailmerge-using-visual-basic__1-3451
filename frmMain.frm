VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Mail Merge Sample"
   ClientHeight    =   855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   855
   ScaleWidth      =   2610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdMailMerge 
      Caption         =   "&Mail Merge"
      Default         =   -1  'True
      Height          =   495
      Left            =   930
      TabIndex        =   0
      Top             =   180
      Width           =   1500
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmMain.frx":0442
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**(MODULE HEADER)*************************************************
'*
'*   Author: Microsoft Corporation
'*  Purpose: This VB Project was created using sample code from
'*           Microsoft's Knowledgebase.
'*
'******************************************************************

Dim wrdApp      As Word.Application
Dim wrdDoc      As Word.Document

Private Sub cmdMailMerge_Click()
    Dim wrdSelection    As Word.Selection
    Dim wrdMailMerge    As Word.MailMerge
    Dim wrdMergeFields  As Word.MailMergeFields
    Dim StrToAdd        As String
    
    On Error GoTo Error_Handler
    
    Screen.MousePointer = vbHourglass
    
    
    ' Create an instance of Word  and make it visible
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = True
    
    ' Add a new document
    Set wrdDoc = wrdApp.Documents.Add
    wrdDoc.Select
    Set wrdSelection = wrdApp.Selection
    Set wrdMailMerge = wrdDoc.MailMerge
    
    ' Create MailMerge Data file
    CreateMailMergeDataFile
    
    ' Create a string and insert it into the document
    StrToAdd = "State University" & vbCr & "Electrical Engineering Department"
    wrdSelection.ParagraphFormat.Alignment = wdAlignParagraphCenter
    wrdSelection.TypeText StrToAdd
    InsertLines 4   ' Insert merge data
    wrdSelection.ParagraphFormat.Alignment = wdAlignParagraphLeft
    Set wrdMergeFields = wrdMailMerge.Fields
    wrdMergeFields.Add wrdSelection.Range, "FirstName"
    wrdSelection.TypeText " "
    wrdMergeFields.Add wrdSelection.Range, "LastName"
    wrdSelection.TypeParagraph
    wrdMergeFields.Add wrdSelection.Range, "Address"
    wrdSelection.TypeParagraph
    wrdMergeFields.Add wrdSelection.Range, "CityStateZip"
    InsertLines 2
    
    ' Right justify the line and insert a date field' with the current date
    wrdSelection.ParagraphFormat.Alignment = wdAlignParagraphRight
    wrdSelection.InsertDateTime _
    DateTimeFormat:="dddd, MMMM dd, yyyy", InsertAsField:=False
    InsertLines 2
    
    ' Justify the rest of the document
    wrdSelection.ParagraphFormat.Alignment = wdAlignParagraphJustify
    wrdSelection.TypeText "Dear "
    wrdMergeFields.Add wrdSelection.Range, "FirstName"
    wrdSelection.TypeText ","
    InsertLines 2
    
    ' Create a string and insert it into the document
    StrToAdd = "Thank you for your recent request for next " & _
                "semester's class schedule for the Electrical " & _
                "Engineering Department. Enclosed with this " & _
                "letter is a booklet containing all the classes " & _
                "offered next semester at State University.  " & _
                "Several new classes will be offered in the " & _
                "Electrical Engineering Department next semester.  " & _
                "These classes are listed below."
    wrdSelection.TypeText StrToAdd
    InsertLines 2    ' Insert a new table with 9 rows and 4 columns
    wrdDoc.Tables.Add wrdSelection.Range, NumRows:=9, _
    NumColumns:=4
    With wrdDoc.Tables(1)    ' Set the column widths
        .Columns(1).SetWidth 51, wdAdjustNone
        .Columns(2).SetWidth 170, wdAdjustNone
        .Columns(3).SetWidth 100, wdAdjustNone
        .Columns(4).SetWidth 111, wdAdjustNone
        
        ' Set the shading on the first row to light gray
        .Rows(1).Cells.Shading.BackgroundPatternColorIndex = wdGray25
        
        ' Bold the first row
        .Rows(1).Range.Bold = True
        
        ' Center the text in Cell (1,1)
        .Cell(1, 1).Range.Paragraphs.Alignment = wdAlignParagraphCenter
        
        ' Fill each row of the table with data
        FillRow wrdDoc, 1, "Class Number", "Class Name", "Class Time", "Instructor"
        FillRow wrdDoc, 2, "EE220", "Introduction to Electronics II", "1:00-2:00 M,W,F", "Dr. Jensen"
        FillRow wrdDoc, 3, "EE230", "Electromagnetic Field Theory I", "10:00-11:30 T,T", "Dr. Crump"
        FillRow wrdDoc, 4, "EE300", "Feedback Control Systems", "9:00-10:00 M,W,F", "Dr. Murdy"
        FillRow wrdDoc, 5, "EE325", "Advanced Digital Design", "9:00-10:30 T,T", "Dr. Alley"
        FillRow wrdDoc, 6, "EE350", "Advanced Communication Systems", "9:00-10:30 T,T", "Dr. Taylor"
        FillRow wrdDoc, 7, "EE400", "Advanced Microwave Theory", "1:00-2:30 T,T", "Dr. Lee"
        FillRow wrdDoc, 8, "EE450", "Plasma Theory", "1:00-2:00 M,W,F", "Dr. Davis"
        FillRow wrdDoc, 9, "EE500", "Principles of VLSI Design", "3:00-4:00 M,W,F", "Dr. Ellison"
    End With
  
    ' Go to the end of the document
    wrdApp.Selection.GoTo wdGoToLine, wdGoToLast
    InsertLines 2
    
    ' Create a string and insert it into the document
    StrToAdd = "For additional information regarding the " & _
                "Department of Electrical Engineering, " & _
                "you can visit our Web site at "
    wrdSelection.TypeText StrToAdd
    
    ' Insert a hyperlink to the Web page
    wrdSelection.Hyperlinks.Add Anchor:=wrdSelection.Range, Address:="http://www.ee.stateu.tld"
    
    ' Create a string and insert it into the document
    StrToAdd = ".  Thank you for your interest in the classes " & _
                "offered in the Department of Electrical " & _
                "Engineering.  If you have any other questions, " & _
                "please feel free to give us a call at " & _
                "555-1212." & vbCr & vbCr & _
                "Sincerely," & vbCr & vbCr & _
                "Kathryn M. Hinsch" & vbCr & _
                "Department of Electrical Engineering" & vbCr
    wrdSelection.TypeText StrToAdd
    
    ' Where to send the document?'
    wrdMailMerge.Destination = wdSendToNewDocument
'    wrdMailMerge.Destination = wdSendToEmail
'    wrdMailMerge.Destination = wdSendToFax
'    wrdMailMerge.Destination = wdSendToPrinter
    
    ' --- Perform MAIL MERGE --- '
    wrdMailMerge.Execute False
    
    wrdDoc.PrintPreview
    
    
    ' Close the original form document
    wrdDoc.Saved = True
'    wrdDoc.Close False
    
    ' Notify user we are done.
    MsgBox "Mail Merge Complete.", vbMsgBoxSetForeground
    
    ' Release References
    Set wrdSelection = Nothing
    Set wrdMailMerge = Nothing
    Set wrdMergeFields = Nothing
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    
    ' Cleanup temp file
'    Kill "C:\DataDoc.doc"
    Screen.MousePointer = vbDefault
Exit Sub

Error_Handler:
    Screen.MousePointer = vbDefault
    MsgBox "Error: " & Err.Number & vbLf & vbLf & Err.Description, vbExclamation, "Mail Merge Error!"
End Sub



Public Sub InsertLines(LineNum As Integer)
    Dim iCount As Integer
    'INSERT BLANK LINES IN MS WORD
    For iCount = 1 To LineNum
        wrdApp.Selection.TypeParagraph
    Next iCount
End Sub
    
Public Sub FillRow(Doc As Word.Document, Row As Integer, _
                   Text1 As String, Text2 As String, _
                   Text3 As String, Text4 As String)
                   
    With Doc.Tables(1)    ' Insert the data into the specific cell
        .Cell(Row, 1).Range.InsertAfter Text1
        .Cell(Row, 2).Range.InsertAfter Text2
        .Cell(Row, 3).Range.InsertAfter Text3
        .Cell(Row, 4).Range.InsertAfter Text4
    End With
End Sub

Public Sub CreateMailMergeDataFile()
    Dim wrdDataDoc  As Word.Document
    Dim X           As Integer
    
    ' Create a data source at C:\DataDoc.doc containing the field data
    wrdDoc.MailMerge.CreateDataSource Name:="C:\DataDoc.doc", HeaderRecord:="FirstName, LastName, Address, CityStateZip"
    
    ' Open the file to insert data
    Set wrdDataDoc = wrdApp.Documents.Open("C:\DataDoc.doc")
    For X = 1 To 2
        wrdDataDoc.Tables(1).Rows.Add
    Next X
    
    ' Fill in the data
    FillRow wrdDataDoc, 2, "Steve", "DeBroux", "4567 Main Street", "Buffalo, NY  98052"
    FillRow wrdDataDoc, 3, "Jan", "Miksovsky", "1234 5th Street", "Charlotte, NC  98765"
    FillRow wrdDataDoc, 4, "Brian", "Valentine", "12348 78th Street  Apt. 214", "Lubbock, TX  25874"
    
    ' Save and close the file
    wrdDataDoc.Save
    wrdDataDoc.Close False
End Sub








