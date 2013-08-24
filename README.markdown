# MarkdownToOpenXML

Converts Markdown to a DocX file, allowing you to output Word files from your C# application via Markdown.

Very rough beginnings, currently only supports header1 & normal styles & bold, italics and underline inline styles. Will eventually include support for customizing styles (normal, header 1 etc) aswell as inline fonts etc. There was nothing like this already online so this could be a potential, very rough starting point.

[Getting started](https://github.com/dangerdan/MarkdownToOpenXML/wiki/Getting-started)

More to come.


## Building project

In order to build the project you will need to install the [Open XML SDK 2.5](http://www.microsoft.com/en-ca/download/details.aspx?id=30425)
and add a reference to **DocumentFormat.OpenXML** if it does not exist
