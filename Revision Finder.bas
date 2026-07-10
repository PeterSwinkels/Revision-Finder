Attribute VB_Name = "RevisionFinderModule"
'This module contains this program's core procedures.
Option Explicit

'The Microsoft Windows API functions used by this program:
Private Declare Function GetFullPathNameA Lib "Kernel32.dll" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long

'The constants used by this program:
Private Const MAX_PATH As Long = 260    'Defines the maximum length for a path.
'This procedure converts any relative path specified to a complete path.
Private Function CompletePath(Path As String) As String
On Error GoTo ErrorTrap
Dim CompletedPath As String
Dim Length As Long

   CompletedPath = String$(MAX_PATH, vbNullChar)
   Length = GetFullPathNameA(Path, Len(CompletedPath), CompletedPath, vbNullString)
   CompletedPath = Left$(CompletedPath, Length)
   
EndRoutine:
   CompletePath = CompletedPath
   Exit Function
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Function


'This procedure checks the specified document for revisions.
Private Sub GetRevisions(DocumentPath As String)
On Error GoTo ErrorTrap
Dim DocumentO As Word.Document
Dim FileH As Integer
Dim RevisionO As Word.Revision
Dim WordO As New Word.Application

   WordO.Visible = False
   WordO.Documents.Open DocumentPath
   Set DocumentO = WordO.Documents.Item(WordO.Documents.Count)
   
   If DocumentO.TrackRevisions Then
      If DocumentO.Revisions.Count > 0 Then
         FileH = FreeFile()
         Open DocumentPath & ".txt" For Output As FileH
            Print #FileH, "Track Revisions: "; DocumentO.TrackRevisions
            Print #FileH, "Revision Count: "; CStr(DocumentO.Revisions.Count)
            For Each RevisionO In DocumentO.Revisions
               Print #FileH, "Revision: "; CStr(RevisionO.Index)
               Print #FileH, "Type: "; CStr(RevisionO.Type)
               Print #FileH, "Contents: "
               Print #FileH, RevisionO.Range.Text
               Print #FileH,
            Next RevisionO
         Close FileH
         
         Shell "Notepad.exe " & DocumentPath & ".txt", vbNormalFocus
      Else
         MsgBox "The specified document does not contain any revisions.", vbInformation
      End If
   Else
      MsgBox "The specified document does not have revision tracking enabled.", vbInformation
   End If
   
EndRoutine:
   WordO.Quit SaveChanges:=False
   
   Set DocumentO = Nothing
   Set RevisionO = Nothing
   Set WordO = Nothing
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure handles any errors that occur.
Private Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.LastDllError
   On Error Resume Next
   MsgBox Description & vbCr & "Error code: " & CStr(ErrorCode), vbInformation
End Sub

'This procecure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
Dim DocumentPath As String

   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path
   
   DocumentPath = Command$()
   If DocumentPath = vbNullString Then DocumentPath = InputBox$("Specify a Microsoft Word document:")
   If Not DocumentPath = vbNullString Then GetRevisions CompletePath(DocumentPath)
   
EndRoutine:
   Exit Sub
   
ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

