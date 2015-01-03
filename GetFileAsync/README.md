This example demostrates how to use the [Document.getFileAsync method](http://msdn.microsoft.com/en-us/library/office/jj715284%28v=office.1501401%29.aspx) to obtain the editing document from Word.

When the user click the button, the Agave will obtain the file, and send it to the server. The server-side code is a C# Http Handler, which extract the docx file and return the data in document.xml to the Agave client.

You may implement your customized opeartion with the docx file.
