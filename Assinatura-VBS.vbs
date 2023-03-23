'On Error Resume Next
Const end_table = 7
Set objSysInfo = CreateObject("ADSystemInfo")
strUser = objSysInfo.UserName
Set objUser = GetObject("LDAP://" & strUser)
strnome = objUser.Fullname
strTitle = objUser.Title
strPhone = objUser.TelephoneNumber
strMobile = objUser.mobile
stremail = objUser.mail
strnotice = "Seu Texto aqui"
Set objword = CreateObject("Word.Application")
With objword 
  Set objDoc = .Documents.Add()
  Set objSelection = .Selection
  Set objEmailOptions = .EmailOptions
  Set objRange = objDoc.Range()
  objDoc.Tables.Add objRange,7,1
  Set objTable = objDoc.Tables(1)

End With

Set objSignatureObject = objEmailOptions.EmailSignature
Set objSignatureEntries = objSignatureObject.EmailSignatureEntries
objdoc.Paragraphs.SpaceAfter = 0

With objSelection
	objTable.Columns.Width = 800	
	
	objTable.Cell(1, 1).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(1, 1).Range.Font.Bold = True
	objTable.Cell(1, 1).Range.Font.Size = "10"
	objTable.Cell(1, 1).Range.Font.Name = "Verdana"
	objTable.Cell(1, 1).Range.Font.Color = RGB(0, 0, 0)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)
	objTable.Cell(1, 1).Range.Text = strnome
	
	objTable.Cell(2, 1).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(2, 1).Range.Font.Bold = False
	objTable.Cell(2, 1).Range.Font.Size = "10"
	objTable.Cell(2, 1).Range.Font.Name = "Verdana"
	objTable.Cell(2, 1).Range.Font.Color = RGB(0, 0, 0)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)	
	objTable.Cell(2, 1).Range.Text = strTitle

'verifica se tem mais de 1 telefone
	If strMobile <> "NULL" then
		objTable.Cell(3, 1).Range.ParagraphFormat.SpaceAfter = 0
		objTable.Cell(3, 1).Range.Font.Bold = False
		objTable.Cell(3, 1).Range.Font.Size = "10"
		objTable.Cell(3, 1).Range.Font.Name = "Verdana"
		objTable.Cell(3, 1).Range.Font.Color = RGB(0, 0, 0)
		objTable.Columns(1).Width = objWord.InchesToPoints(1)
		objTable.Cell(3, 1).Range.Text = strPhone & " / " & strMobile
	Else
		objTable.Cell(3, 1).Range.ParagraphFormat.SpaceAfter = 0
		objTable.Cell(3, 1).Range.Font.Bold = False
		objTable.Cell(3, 1).Range.Font.Size = "10"
		objTable.Cell(3, 1).Range.Font.Name = "Verdana"
		objTable.Cell(3, 1).Range.Font.Color = RGB(0, 0, 0)
		objTable.Columns(1).Width = objWord.InchesToPoints(1)
		objTable.Cell(3, 1).Range.Text = strPhone
	End If
	
	objTable.Cell(4, 1).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(4, 1).Range.Font.Bold = False
	objTable.Cell(4, 1).Range.Font.Size = "10"
	objTable.Cell(4, 1).Range.Font.Name = "Verdana"
	objTable.Cell(4, 1).Range.Font.Color = RGB(0, 0, 0)
	objTable.Columns(1).Width = objWord.InchesToPoints(1)
	objTable.Cell(4, 1).Range.Text = stremail
	
	objTable.Cell(5, 1).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(5, 1).Range.Font.Bold = False
	objTable.Cell(5, 1).Range.Font.Underline = True
	objTable.Cell(5, 1).Range.Font.Size = "10"
	objTable.Cell(5, 1).Range.Font.Name = "Verdana"
	objTable.Cell(5, 1).Range.Font.Color = RGB(0, 0, 0)
	objTable.Columns(1).Width = objWord.InchesToPoints(1) 
	objTable.Cell(5, 1).Range.Text = strweb
	
	objtable.Cell(6,1).width = 750
	objtable.Cell(6,1).height = 140
	objTable.Cell(6,1).Range.InlineShapes.AddPicture("\\caminho-da-imagem-logo.png")
	
	objTable.Cell(7, 1).Range.ParagraphFormat.SpaceAfter = 0
	objTable.Cell(7, 1).Range.Font.Bold = False
	objTable.Cell(7, 1).Range.Font.Underline = True
	objTable.Cell(7, 1).Range.Font.Size = "10"
	objTable.Cell(7, 1).Range.Font.Name = "Verdana"
	objTable.Cell(7, 1).Range.Font.Color = RGB(0, 0, 0)
	objtable.Cell(7, 1).width = 750
	objtable.Cell(7, 1).height = 140	
	objTable.Cell(7, 1).Range.Text = strnotice
	
	.TypeText Chr(1)
	
	'.EndKey end_table
End With

Set objSelection = objDoc.Range()
objSignatureEntries.Add "Assinatura", objSelection
objSignatureObject.NewMessageSignature = "Assinatura"
objSignatureObject.ReplyMessageSignature = "Assinatura"

objDoc.Saved = True
objword.Quit