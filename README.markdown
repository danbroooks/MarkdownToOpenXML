# MarkdownToOpenXML

Converts Markdown to a DocX file, allowing you to output Word files from your C# application via Markdown.

## Building project

In order to build the project you will need to install the [Open XML SDK 2.5](http://www.microsoft.com/en-ca/download/details.aspx?id=30425)
and add a reference to **DocumentFormat.OpenXML** if it does not exist

## Getting started ##

```c#
using MarkdownToOpenXML;

string markdown = @"# This is a header";
MD2OXML.CreateDocX(markdown, @"C:\MyMarkdown.docx");
```

By default the MD2OXML will use the standard markdown syntax:

```
# Header1 text

Two newlines represents a new paragraph, without the # at the start this will output normal text

*Italic* out puts italic text and **bold** outputs bold text
```

However there is a custom syntax mode which adds in support for underline, before running the main MD2OXML method, set this:

```c#
MarkdownToOpenXML.MD2OXML.customMode = true;
```

This will then support this syntax:

```
**Bold text**

`Italics`

_Underlined text_
```

In both cases, syntax markers can be mixed and matched as you please.