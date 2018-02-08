

[TOC]


# PPT

## EMPTY
>
	Void SaveAs(System.String, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): Presentation (unknown member) : Failed.
	   �� Microsoft.Office.Interop.PowerPoint.PresentationClass.SaveAs(String FileName, PpSaveAsFileType FileFormat, MsoTriState EmbedTrueTypeFonts)
	   �� PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) λ�� D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:�к� 52

## READONLY
>
	Void SaveAs(System.String, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80048240): Presentation (unknown member) : Invalid request.  Presentation cannot be modified.
	   �� Microsoft.Office.Interop.PowerPoint.PresentationClass.SaveAs(String FileName, PpSaveAsFileType FileFormat, MsoTriState EmbedTrueTypeFonts)
	   �� PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) λ�� D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:�к� 52

## ERROR
>
	Microsoft.Office.Interop.PowerPoint.Presentation Open2007(System.String, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): �� COM ����ĵ��÷����˴��� HRESULT E_FAIL��
	   �� Microsoft.Office.Interop.PowerPoint.Presentations.Open2007(String FileName, MsoTriState ReadOnly, MsoTriState Untitled, MsoTriState WithWindow, MsoTriState OpenAndRepair)
	   �� PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) λ�� D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:�к� 47

## PASSWORD
>
	Microsoft.Office.Interop.PowerPoint.Presentation Open2007(System.String, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): �� COM ����ĵ��÷����˴��� HRESULT E_FAIL��
	   �� Microsoft.Office.Interop.PowerPoint.Presentations.Open2007(String FileName, MsoTriState ReadOnly, MsoTriState Untitled, MsoTriState WithWindow, MsoTriState OpenAndRepair)
	   �� PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) λ�� D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:�к� 47

## FIX
>
	Microsoft.Office.Interop.PowerPoint.Presentation Open2007(System.String, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState, Microsoft.Office.Core.MsoTriState)
	System.Runtime.InteropServices.COMException (0x80004005): �� COM ����ĵ��÷����˴��� HRESULT E_FAIL��
	   �� Microsoft.Office.Interop.PowerPoint.Presentations.Open2007(String FileName, MsoTriState ReadOnly, MsoTriState Untitled, MsoTriState WithWindow, MsoTriState OpenAndRepair)
	   �� PPT2HTML5.Expand.Service.Utils.PPTUtils.ConvertToIMAGE(String filePath, String destPath) λ�� D:\work\git\yk\csharp\yk_pptconvertor\code\powerpoint2html5\PPT2HTML5.ConversionEngine\PPT2HTML5.Expand.Service\PPT2HTML5.Expand.Service\Utils\PPTUtils.cs:�к� 47

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
> Excel �޷����ļ���error.xlsx������Ϊ�ļ���ʽ���ļ���չ����Ч����ȷ���ļ�δ�𻵣������ļ���չ������ ���ĸ�ʽƥ�䡣

## PASSWORD
> �����ṩ�����벻��ȷ������ CAPS LOCK ����״̬����ȷ��ʹ������ȷ�Ĵ�Сд��

## NORMAL
> [51911633-24c2-4f67-9f2e-36531fb8d03d.pdf](http://otzm88f21.bkt.clouddn.com/8b81f591-62a6-4c9f-80e1-198f5ac6b126.pdf)


# Word

## EMPTY
> [63c60ae5-7e8f-46ab-a0e3-efe3db00e305.pdf](http://otzm88f21.bkt.clouddn.com/699be5f6-178f-4bb3-85be-8f05db8f676f.pdf)


## ERROR
> 
	Microsoft.Office.Interop.Word.Document Open(System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef)
	�ļ������Ѿ��𻵡�

## PASSWORD
- δ���������`PasswordDocument`ֵΪ��ʱ���Զ��������������
	- δ��������ʱ���������벻Ӱ����ļ�

> 
	Microsoft.Office.Interop.Word.Document Open(System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef, System.Object ByRef)
	����ʧ��


## NORMAL
> [37d13071-4a9e-44b7-833d-0ceacfef2915.pdf](http://otzm88f21.bkt.clouddn.com/2c230f5f-faf5-4154-b2e5-6d84180ae592.pdf)

