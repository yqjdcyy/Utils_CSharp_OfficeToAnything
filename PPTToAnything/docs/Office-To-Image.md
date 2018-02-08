

[TOC]


# PPT

## EMPTY
>
	Void SaveAs(System.String, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): Presentation (unknown member) : Failed.
	   在 Microsoft.Office.Interop.PowerPoint.PresentationClass.SaveAs(String FileName, PpSaveAsFileType FileFormat, MsoTriState EmbedTrueTypeFonts)
	   在 PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) 位置 D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:行号 52

## READONLY
>
	Void SaveAs(System.String, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80048240): Presentation (unknown member) : Invalid request.  Presentation cannot be modified.
	   在 Microsoft.Office.Interop.PowerPoint.PresentationClass.SaveAs(String FileName, PpSaveAsFileType FileFormat, MsoTriState EmbedTrueTypeFonts)
	   在 PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) 位置 D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:行号 52

## ERROR
>
	Microsoft.Office.Interop.PowerPoint.Presentation Open2007(System.String, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): 对 COM 组件的调用返回了错误 HRESULT E_FAIL。
	   在 Microsoft.Office.Interop.PowerPoint.Presentations.Open2007(String FileName, MsoTriState ReadOnly, MsoTriState Untitled, MsoTriState WithWindow, MsoTriState OpenAndRepair)
	   在 PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) 位置 D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:行号 47

## PASSWORD
>
	Microsoft.Office.Interop.PowerPoint.Presentation Open2007(System.String, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): 对 COM 组件的调用返回了错误 HRESULT E_FAIL。
	   在 Microsoft.Office.Interop.PowerPoint.Presentations.Open2007(String FileName, MsoTriState ReadOnly, MsoTriState Untitled, MsoTriState WithWindow, MsoTriState OpenAndRepair)
	   在 PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) 位置 D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:行号 47

## FIX
>
	Microsoft.Office.Interop.PowerPoint.Presentation Open2007(System.String, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): 对 COM 组件的调用返回了错误 HRESULT E_FAIL。
	   在 Microsoft.Office.Interop.PowerPoint.Presentations.Open2007(String FileName, MsoTriState ReadOnly, MsoTriState Untitled, MsoTriState WithWindow, MsoTriState OpenAndRepair)
	   在 PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) 位置 D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:行号 47

## NORMAL
>
	![1.jpg](http://otzm88f21.bkt.clouddn.com/13423c56-df84-4ece-913b-2f2f0bd4f614.jpg)
	![3.jpg](http://otzm88f21.bkt.clouddn.com/eccbdd4c-3a9f-45f2-b31b-3aec3d74dc7f.jpg)
	![4.jpg](http://otzm88f21.bkt.clouddn.com/2f06020f-fac5-44c2-b61c-03086bed2e95.jpg)
	![2.jpg](http://otzm88f21.bkt.clouddn.com/ff2341d7-a204-4395-9a4e-e4169f0128e4.jpg)


# Excel

## EMPTY
> 0 == book.Sheets.HPageBreaks.Count && 0 == book.Sheets.VPageBreaks.Count

## ERROR
> Excel 无法打开文件“error.xlsx”，因为文件格式或文件扩展名无效。请确定文件未损坏，并且文件扩展名与文 件的格式匹配。

## PASSWORD
> 您所提供的密码不正确。请检查 CAPS LOCK 键的状态，并确认使用了正确的大小写。

## NORMAL
> [51911633-24c2-4f67-9f2e-36531fb8d03d.pdf](http://otzm88f21.bkt.clouddn.com/8b81f591-62a6-4c9f-80e1-198f5ac6b126.pdf)


# Word

## EMPTY
> [63c60ae5-7e8f-46ab-a0e3-efe3db00e305.pdf](http://otzm88f21.bkt.clouddn.com/699be5f6-178f-4bb3-85be-8f05db8f676f.pdf)


## ERROR
> 
	Microsoft.Office.Interop.Word.Document Open(System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef)
	文件可能已经损坏。

## PASSWORD
- 未设置密码或`PasswordDocument`值为空时，自动弹出密码输入框
	- 未设置密码时，错误密码不影响打开文件

> 
	Microsoft.Office.Interop.Word.Document Open(System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef)
	命令失败


## NORMAL
> [37d13071-4a9e-44b7-833d-0ceacfef2915.pdf](http://otzm88f21.bkt.clouddn.com/2c230f5f-faf5-4154-b2e5-6d84180ae592.pdf)

