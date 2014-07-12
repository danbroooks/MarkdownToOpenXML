# MarkdownToOpenXML

Converts Markdown to a DocX file, allowing you to output Word files from your C# application via Markdown.

## Building project

Clone the repository, open in Visual Studio, and build. All dependencies (such as `DocumentFormat.OpenXML`) are Nuget packages.

## Getting started ##

### Usage ###

The main method in MD2OXML is CreateDocX, which takes two string values - the markdown, and the path where you want to store the file:

```c#
using MarkdownToOpenXML;

string markdown = @"# This is a header";
MD2OXML.CreateDocX(markdown, @"C:\MyMarkdown.docx");
```

### Current syntax support ###

#### Headers ####

Headers can be produced either the atx style or setext style.

##### ATX #####

An ATX style header has a number of hashes at the beginning of the line to signify a header:

```
# This is an H1

## This is an H2

###### This is an H6
```

You can optionally close atx headers:

```
# This is an H1 #

## This is an H2 ##

### This is an H3 ###
```

##### Setext #####

A setext header is underlined by either equal signs or dashes on the next line:

```
This is an H1
=============

This is an H2
-------------
```

#### Formatting ####

By default MD2OXML is set to use it's own extended formatting syntax, adding in support for underline.

```
**Bold text**

`Italics`

_Underlined text_
```

However you can set the syntax back to the original markdown rules by setting the ExtendedMode flag to false:

```c#
MarkdownToOpenXML.MD2OXML.ExtendedMode = false;
```

This will then support this syntax:

```
*Italic* out puts italic text and **bold** outputs bold text
```

#### Alignment ####

This is something that isn't seen in markdown parsers for the web as placement would normally be dealt with by CSS styles. However MD2OXML supports inline alignment statements that should be placed at the start of a line:

```
>< This centers the text
<> This fully justifies the text
>> Right alignment
<< Left alignment
```