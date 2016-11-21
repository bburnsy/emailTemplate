Attribute VB_Name = "Module1"
Sub insertHtml()


Dim insp As Inspector
Set insp = ActiveInspector
If insp.IsWordMail Then
    Dim wordDoc As Word.Document
    Set wordDoc = insp.WordEditor
    wordDoc.Application.Selection.InsertFile "C:\Users\JJenkins\Desktop\htmlEmail\emailHtmlTemplate.html", , False, False, False
End If


End Sub
