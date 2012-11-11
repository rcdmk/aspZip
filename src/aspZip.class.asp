<%
' Classic ASP CSV creator
' By RCDMK <rcdmk@rcdmk.com>
'
' The MIT License (MIT) - http://opensource.org/licenses/MIT
' Copyright (c) 2012 RCDMK <rcdmk@rcdmk.com>
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
' associated documentation files (the "Software"), to deal in the Software without restriction,
' including without limitation the rights to use, copy, modify, merge, publish, distribute,
' sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial
' portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
' NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
' OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

class aspZip
	dim BlankZip, NoInterfaceYesToAll
	dim fso, curArquieve, created, saved
	dim files, m_path, zipApp, zipFile	
	
	public property get Count()
		Count = files.Count
	end property
	
	public property get Path
		Path = m_path
	end property
	
	
	private sub class_initialize()
		BlankZip = Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0) 	' Create the blank file structure
		NoInterfaceYesToAll = 4 or 16 or 1024 ' http://msdn.microsoft.com/en-us/library/windows/desktop/bb787866(v=vs.85).aspx
		
		' initialize components
		set fso = createObject("scripting.filesystemobject")
		set files = createObject("Scripting.Dictionary")
		
		Set zipApp = CreateObject("Shell.Application")
	end sub
	
	private sub class_terminate()
		' some cleanup
		set curArquieve = nothing
		set zipApp = nothing
		set files = nothing
		
		' If we created the file but did not saved it, delete it
		' since its empty
		if created and not saved then
			on error resume next
			fso.deleteFile m_path
			on error goto 0
		end if
		
		set fso = nothing
	end sub
	
	
	' Opens or creates the arquieve
	public sub OpenArquieve(byval path)
		dim file
		' Make sure the path is complete and in a correct format
		path = replace(path, "/", "\")
		m_path = Server.MapPath(path)
		
		' Create an empty file if it already doesn't exists
		if not fso.fileexists(m_path) then
			set file = fso.createTextFile(m_path)
			file.write BlankZip
			file.close()
			set file = nothing
			
			set curArquieve = zipApp.NameSpace(m_path)
			created = true
		else
			' Open the existing file and load its contents
			
			dim cnt
			set curArquieve = zipApp.NameSpace(m_path)
			
			cnt = 0
			for each file in curArquieve.Items
				cnt = cnt + 1
				files.add file.path, cnt
			next
		end if

		saved = false
	end sub
	
	
	' Add a file or folder to the list
	public sub Add(byval path)
		path = replace(path, "/", "\")		
		if instr(path, ":") = 0 then path = Server.mappath(path)
		
		if not fso.fileExists(path) and not fso.folderExists(path) then
			err.raise 1, "File not exists", "The input file name doen't correspond to an existing file"
			
		elseif not files.exists(path) Then
			files.add path, files.Count + 1
		end if
	end sub
	
	' Remove a file or folder from the to be added list (currently it only works for new files)
	public sub Remove(byval path)
		if files.exists(path) then files.Remove(path)
	end sub
	
	' Clear the to be added list
	public sub RemoveAll()
		files.RemoveAll()
	end sub
	
	' Writes the to the arquieve
	public sub CloseArquieve()
		dim filepath, file, initTime, fileCount
		dim cnt
		cnt = 0

		For Each filepath In files.keys
			' do not try add the contents that are already in the arquieve
			if instr(filepath, m_path) = 0 then
				curArquieve.Copyhere filepath, NoInterfaceYesToAll
				fileCount = curArquieve.items.Count
				
				'Keep script waiting until Compressing is done
				On Error Resume Next
				'Do Until fileCount < curArquieve.Items.Count
					wscript.sleep(10)
					cn = cnt + 1
				'Loop
				On Error GoTo 0
			end if
		next
		
		saved = true
	end sub
	
	
	public sub ExtractTo(byval path)
		if typeName(curArquieve) = "Folder3" Then
			path = Server.MapPath(path)
			
			if not fso.folderExists(path) then
				fso.createFolder(path)
			end if
			
			zipApp.NameSpace(path).CopyHere curArquieve.Items, NoInterfaceYesToAll
		end if
	end sub
end class
%>