Public Sub DownloadAttachment(Mail As Outlook.MailItem)

  Dim SaveFolder As String
  Dim SubjectLine As String
  Dim Attach As Outlook.Attachment
  Dim FileName As String
  Dim FilePath As String
  Dim bad_char As Variant
  Dim env As String

  ' Change to preferred location'
  env = CStr(Environ("USERPROFILE")')
  SaveFolder = env & "\Documents\outlook-emails\"
  SubjectLine = Mail.Subject

  ' Remove bad characters from subject line'
  Const BadCharacters As String = "!,@,#,$,%,^,&,*,(,),{,[,],},?,/,:"
  For Each bad_char In Split(BadCharacters, ",")
    SubjectLine = Replace(SubjectLine, bad_char, "_")
  Next

  For Each Attach In Mail.Attachments
    FileName = SubjectLine & "___" & Attach.DisplayName
    FilePath = SaveFolder & FileName
    Attach.SaveAsFile FilePath
  Next

End Sub
