'//-----------------------------------------------------------------------------
'// PowerPoint文件转图像脚本(ppt2img)
'// 作者: www.51windows.net，海娃
'// 使用方法: 将此文件放在sendto文件中, 然后在ppt文件上点右键, 发送到,
'//           ppt2img.vbs中,输入要输出图像的格式,然后输入图像的宽与高,
'//           脚本会生成一个同名的文件,里面为生成的图像文件.
'// 注意: 机器上要安装Powerpoint程序
'// 修改:
'//      2011/04/02 八戒 (eight.ring@gmail.com) 修改以支持生成长条图片
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
		iname = right(iname, 4) '取四位数
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
