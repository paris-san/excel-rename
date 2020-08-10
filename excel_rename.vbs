Dim WorkBookName, SearchKeyWord , Replace, dir, objDlg, Fil, fso, FLD, r
abc = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"

Set objDlg = WScript.CreateObject("Shell.Application")			
'Set dir = objDlg.BrowseForFolder(&H0, "Dialekse fakelo", &H10, "")			
Set fso = CreateObject("Scripting.FileSystemObject") 			
dir=fso.GetParentFolderName(wscript.ScriptFullName) 			
set FLD = FSO.GetFolder(dir)			
			
SearchKeyWord= InputBox("ti leksi psaxneis")	
If SearchKeyWord= "" Then	WScript.Quit
Replace=InputBox("me poia thes na antikatastiseis (tha allaksei olo to keli oxi mono to arxiko simeio)")	
If Replace = "" Then WScript.Quit
Set appExcel = CreateObject("Excel.Application")		
appExcel.Visible = True			
			
For Each Fil In FLD.Files

	If Instr(1, Fil.Name, "xls") > 0 Then	
		Set objWorkBook = appExcel.Workbooks.Open(Fil.Path) 'opens the sheet
		For j = 1 to objWorkBook.Worksheets.Count
            For k=1	to 22
				For l=1 to 100
				    set r=objWorkBook.Worksheets(j).Cells(l,k)
					If r.Value=SearchKeyWord Then r.Value= Replace
					'set r=appExcel.Worksheets(j).Range("A1:Z100").Find(SearchKeyWord)	
					'i=0		
					'If Not r Is Nothing Then i=1
					'temp="A1:Z100"			
					'Do While i>0	
				
					'	If Not r Is Nothing Then
					'		If r.Value = SearchKeyWord Then 
					'			r.Value=Replace	
					'		Else
					'			MsgBox r.Value
					'			temp1 = r.column+1
					'			temp2 = r.row
					'			temp1 = convert(temp1)
					'			temp=temp1&temp2&":Z100"			
					'		End if
					'		set r=appExcel.Worksheets(j).Range(temp).Find(SearchKeyWord)
					'	Else	
					'		i=0
					'	End If	
						
					'Loop		
				Next
			Next
		Next	
		objWorkBook.Save		
		objWorkBook.Close	
		
	End If
Next			
			
MsgBox "done"			
Set appExcel = Nothing			


function convert(n)
  convert = Mid(abc, n, 1)
end function