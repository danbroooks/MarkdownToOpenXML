# MarkdownToOpenXML

Converts Markdown to a DocX file in C#.

Very rough beginnings, currently only supports header 1, bold, italics and underline. Will eventually include support for customizing styles (normal, header 1 etc) aswell as inline fonts etc. There was nothing like this already online so this could be a potential, very rough starting point.

Usage:

```c#
using MarkdownToOpenXML;

string markdown = @"# This is a header";
MD2OXML.CreateDocX(markdown, @"C:\MyMarkdown.docx");
```