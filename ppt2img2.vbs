'//-----------------------------------------------------------------------------
'// PowerPoint�ļ�תͼ��ű�(ppt2img)
'// ����: www.51windows.net������
'// ʹ�÷���: �����ļ�����sendto�ļ���, Ȼ����ppt�ļ��ϵ��Ҽ�, ���͵�,
'//           ppt2img.vbs��,����Ҫ���ͼ��ĸ�ʽ,Ȼ������ͼ��Ŀ����,
'//           �ű�������һ��ͬ�����ļ�,����Ϊ���ɵ�ͼ���ļ�.
'// ע��: ������Ҫ��װPowerpoint����
'// �޸�:
'//      2011/04/02 �˽� (eight.ring@gmail.com) �޸���֧�����ɳ���ͼƬ
'//-----------------------------------------------------------------------------
on error resume next
Set ArgObj = WScript.Arguments
ArgCnt = WScript.Arguments.Count

pptfilepath = ""
imgType = "png"
imgW = 600
imgH = imgW * 0.75

If ArgCnt = 4 Then
	pptfilepath = ArgObj.Item(0)
	imgType = ArgObj.Item(1)
	imgW = ArgObj.Item(2)
	imgH = ArgObj.Item(3)
ElseIf ArgCnt = 3 Then
	pptfilepath = ArgObj.Item(0)
	imgType = ArgObj.Item(1)
	imgW = ArgObj.Item(2)
	imgH = imgW * 0.75
ElseIf ArgCnt = 2 Then
	pptfilepath = ArgObj.Item(0)
	imgType = ArgObj.Item(1)
ElseIf ArgCnt = 1 Then
	pptfilepath = ArgObj.Item(0)
Else
	Wscript.Echo "Usage: pp2img filePath imgType imgWidth imgHeight"
	Wscript.Quit
End If

'msgbox pptfilepath & imgType & imgW & imgH

call Form_Load(pptfilepath,imgType)

Private Sub Form_Load(Filepath, format)
	if format = "" then
		format = "png"
	end if

	ext = GetExtName(Filepath)

	Folderpath = left(Filepath, len(Filepath) - len(ext) - 1)
	CreateFolder(Folderpath)
    Set ppApp = CreateObject("PowerPoint.Application")
    Set ppPresentations = ppApp.Presentations
    Set ppPres = ppPresentations.Open(Filepath, -1, 0, 0)
    Set ppSlides = ppPres.Slides

	For i = 1 To ppSlides.Count
		iname = "000000"&i
		iname = right(iname, 4) 'ȡ��λ��
		Call ppSlides.Item(i).Export(Folderpath&"\"&iname&"."&format, format, imgW, imgH)
	Next

	Set ppApp = Nothing
	Set ppPres = Nothing
End Sub

Function GetBaseName(DriveSpec)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   GetBaseName = fso.GetFileName(DriveSpec)
   Set fso = Nothing
End Function

Function GetDirName(DriveSpec)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   GetDirName = fso.GetParentFolderName(Drivespec)
   Set fso = Nothing
End Function

Function GetExtName(DriveSpec)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   GetExtName = fso.GetExtensionName(Drivespec)
   Set fso = Nothing
End Function

Function CreateFolder(Filepath)
	Dim fso, f
	on error resume next
	Set fso = CreateObject("Scripting.FileSystemObject")
	if not fso.FolderExists(Filepath) then
		Set f = fso.CreateFolder(Filepath)
	end if
	CreateFolder = f.Path
	set fso = Nothing
	set f = Nothing
End Function
