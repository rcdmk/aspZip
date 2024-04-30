# Classic ASP ZIP file creator 0.4

## The MIT License (MIT) - http://opensource.org/licenses/MIT

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


## Usage

Include the class file in the desired page and instantiate the class

    dim zip
    set zip = new aspZip

Open a ZIP file (create or open an existing file in disk)
	
    zip.OpenArchieve("path\to\file.zip") ' this creates the archieve if it doesn't exists

Add some files or folders

    zip.Add("..\src")
    zip.Add(".\default.asp")
    
Write the files to disk

    zip.CloseArchieve()

If you want to extract the contents of a ZIP file, use the `ExtractTo(DestinationPath)` method

    zip.ExtractTo(".\test")

If the archieve contains no files, it will be deleted when the object is destroyed.
	
>**Note:**  
>In the current release (0.4) the issue with extracting only directory structures should be solved, but processing times are a lot longer now, due to waiting on directory creation. Minimum extra waiting time is 500 ms.