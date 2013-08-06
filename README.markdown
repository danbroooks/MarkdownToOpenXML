# MarkdownToOpenXML

Converts Markdown to a DocX file, allowing you to output Word files from your C# application via Markdown.

Very rough beginnings, currently only supports header1 & normal styles & bold, italics and underline inline styles. Will eventually include support for customizing styles (normal, header 1 etc) aswell as inline fonts etc. There was nothing like this already online so this could be a potential, very rough starting point.

## Usage:

```c#
using MarkdownToOpenXML;

string markdown = @"# This is a header";
MD2OXML.CreateDocX(markdown, @"C:\MyMarkdown.docx");
```

By default the class will use the standard markdown syntax:

```
# Header1 text

Two newlines represents a new paragraph, without the # at the start this will output normal text

*Italic* out puts italic text and **bold** outputs bold text
```

However there is a custom syntax mode which adds in support for underline - before running the main MD2OXML method, set this:

```c#
MarkdownToOpenXML.MD2OXML.customMode = true;
```

This will then support this syntax:

```
**Bold text**

`Italics`

_Underlined text_
```

In both modes, syntax markers can be mixed and matched as you please.

More to come.