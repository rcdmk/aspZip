#Classic ASP ZIP file creator 0.3

##The MIT License (MIT) - http://opensource.org/licenses/MIT

Copyright (c) 2012 RCDMK &lt;rcdmk@rcdmk.com&gt;

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject
to the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.


##Usage

Include the class file in the desired page and instantiate the class

    dim zip
    set zip = new aspZip

Open a ZIP file (create or open an existing file in disk)
	
    zip.OpenArquieve("path\to\file.zip") ' this creates the arquieve if it doesn't exists

Add some files or folders

    zip.Add("..\src")
    zip.Add(".\default.asp")
    
Write the files to disk

    zip.CloseArquieve()

If you want to extract the contents of a ZIP file, use the `ExtractTo(DestinationPath)` method

    zip.ExtractTo(".\test")

If the archieve contains no files, it will be deleted when the object is destroyed.
	
>**Note:**  
>In the current release (0.3) it only extracts the folder strutucture from the file. There are
some "mysterious" things happening when extracting that I could not bypass. Some kind of
access restriction imposed by Windows that even giving "Full control" access to the IIS users
it still gives the error and hangs, so I just ignored the erros for it to work without crashing
the server.
