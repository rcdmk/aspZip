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
	dim BlankZip, fso, curArquieve
	dim files, m_path, zipApp, zipFile	
	
	public property get Count()
		Count = files.Count
	end property
	
	public property get Path
		Path = m_path
	end property
	
	
	private sub class_initialize()
		dim x
		
		' Create the blank file structure
		BlankZip = Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, 0)
		
		set fso = createObject("scripting.filesystemobject")
		set files = createObject("Scripting.Dictionary")
		
		Set zipApp = CreateObject("Shell.Application")
	end sub
	
	private sub class_terminate()
		if typeName(curArquieve) = "TextStream" then
			on error resume next
			curArquieve.close
			err.clear
			on error goto 0
		end if
		
		set curArquieve = nothing
		set fso = nothing
		set files = nothing
		set zipApp = nothing
	end sub
	
	
	public sub OpenArquieve(byval path)
		dim file
		m_path = Server.MapPath(path)
		
		if not fso.fileexists(m_path) then
			set file = fso.createTextFile(m_path)
			file.write BlankZip
			file.close()
			set file = nothing
			
			set curArquieve = zipApp.NameSpace(m_path)
		else
			dim cnt
			set curArquieve = zipApp.NameSpace(m_path)
			
			cnt = 0
			for each file in curArquieve.Items
				cnt = cnt + 1
				files.add file, cnt
			next
		end if
	end sub
	
	public sub AddFile(byval path)
		path = replace(Server.mappath(path), "/", "\")
		if not fso.fileExists(path) and not fso.folderExists(path) then
			err.raise 1, "File not exists", "The input file name doen't correspond to an existing file"
		else
			if not files.exists(path) Then
				files.add path, files.Count + 1
			end if
		end if
	end sub
	
	public sub RemoveFile(byval path)
		if files.exists(path) then files.Remove(path)
	end sub
	
	public sub RemoveAll()
		files.RemoveAll()
	end sub
	
	public sub CloseArquieve()
		dim filepath, file, initTime, fileCount
		
		For Each filepath In files.keys
			curArquieve.Copyhere filepath
			fileCount = curArquieve.items.Count
			
			'Keep script waiting until Compressing is done
			On Error Resume Next
			Do Until fileCount < curArquieve.Items.Count
				initTime = now
				while (now - initTime) < 1
					' Wait for the file to be added
				wend
			Loop
			On Error GoTo 0
		next
	end sub
	
	
	public sub Extract(byval path)
		path = Server.MapPath(path)
		
		if not fso.folderExists(path) then
			fso.createFolder(path)
		end if
		
		zipApp.NameSpace(path).Copyhere curArquieve.Items
	end sub
end class
%>