
# Markdown To Powerpoint User Guide

This document describes the md2pptx Markdown preprocessor, which turns Markdown text into a Powerpoint pptx presentation.

In this document we'll refer to it as "md2pptx", pronounced "em dee to pee pee tee ex".

### Table Of Contents

* [Why md2pptx?](#why-md2pptx)
	* [A Real World Use Case](#a-real-world-use-case)
* [How Do You Use md2pptx?](#how-do-you-use-md2pptx)
* [python-pptx license](#pythonpptx-license)
* [Change Log](#change-log)
* [Creating Slides](#creating-slides)
	* [Presentation Title Slides](#presentation-title-slides)
	* [Presentation Section Slides](#presentation-section-slides)
	* [Bullet Slides](#bullet-slides)
	* [Graphics Slides](#graphics-slides)
	* [Table Slides](#table-slides)
		* [Special Case: Two Graphics Side By Side](#special-case-two-graphics-side-by-side)
		* [Special Case: Two By Two Grid Of Graphics](#special-case-two-by-two-grid-of-graphics)
		* [Special Case: Three Graphics On A Slide](#special-case-three-graphics-on-a-slide)
		* [Special Case: One Graphic Above Another](#special-case-one-graphic-above-another)
	* [Code Slides](#code-slides)
	* [Task List Slides](#task-list-slides)
* [Hyperlinks](#hyperlinks)
* [HTML Comments](#html-comments)
* [Special Text Formatting](#special-text-formatting)
	* [Using HTML `<style>` Elements To Specify Text Colours And Underlining](#using-html-<style>-elements-to-specify-text-colours-and-underlining)
	* [HTML Entity References](#html-entity-references)
	* [Numeric Character References](#numeric-character-references)
	* [Escaped Characters](#escaped-characters)
	* [CriticMarkup](#criticmarkup)
* [Creating A Glossary Of Terms](#creating-a-glossary-of-terms)
* [Creating Footnotes](#creating-footnotes)
	* [Creating A Footnote](#creating-a-footnote)
	* [Referring To A Footnote](#referring-to-a-footnote)
* [Controlling The Presentation With Metadata](#controlling-the-presentation-with-metadata)
	* [Specifying Metadata](#specifying-metadata)
	* [Metadata Keys](#metadata-keys)
		* [Slide Numbers - `numbers`](#slide-numbers-numbers)
		* [Page Title Size - `pageTitleSize`](#page-title-size-pagetitlesize)
		* [Section Title Size - `sectionTitleSize`](#section-title-size-sectiontitlesize)
		* [Section Subtitle Size - `sectionSubtitleSize`](#section-subtitle-size-sectionsubtitlesize)
		* [Monospace Font - `monoFont`](#monospace-font-monofont)
		* [Margin size - `marginBase` and `tableMargin`](#margin-size-marginbase-and-tablemargin)
		* [Associating A Class Name with A Background Colour With `style.bgcolor`](#associating-a-class-name-with-a-background-colour-with-stylebgcolor)
		* [Associating A Class Name with A Foreground Colour With `style.fgcolor`](#associating-a-class-name-with-a-foreground-colour-with-stylefgcolor)
		* [Associating A Class Name With Text Emphasis With `style.emphasis`](#associating-a-class-name-with-text-emphasis-with-styleemphasis)
		* [Template Presentation - `template`](#template-presentation-template)
		* ["Chevron Style" Table Of Contents - `tocStyle` And `tocTitle`](#chevron-style-table-of-contents-tocstyle-and-toctitle)
		* [Specifying An Abstract Slide With `abstractTitle`](#specifying-an-abstract-slide-with-abstracttitle)
		* [Specifying Text Size With `baseTextSize` And `baseTextDecrement`](#specifying-text-size-with-basetextsize-and-basetextdecrement)
		* [Specifying Bold And Italic Text Colour With `BoldColour` And `ItalicColour`](#specifying-bold-and-italic-text-colour-with-boldcolour-and-italiccolour)
		* [Specifying Bold And Italic Text Effects With `BoldBold` And `ItalicItalic`](#specifying-bold-and-italic-text-effects-with-boldbold-and-italicitalic)
		* [Shrinking Tables With `compactTables`](#shrinking-tables-with-compacttables)
		* [Controlling Task Slide Production With `taskSlides` and `tasksPerSlide`](#controlling-task-slide-production-with-taskslides-and-tasksperslide)
		* [Controlling Glossary Slide Production With `glossaryTitle`, `glossaryTerm`, `glossaryMeaning`,`glossaryMeaningWidth`, and `glossaryPerPage`](#controlling-glossary-slide-production-with-glossarytitle-glossaryterm-glossarymeaningglossarymeaningwidth-and-glossaryperpage)
* [Modifying The Slide Template](#modifying-the-slide-template)
	* [Basics](#basics)
	* [Slide Template Sequence](#slide-template-sequence)

## Why md2pptx?

There are advantages in creating presentations using a flat file format. Some of these are:

* You can use any text editor on any platform to create the file.
* Other tools can generate the file.

	For example, the author uses iThoughtsX on Mac, with its counterpart (iThoughts) on iOS, to generate presentations from outlines.

* Text editing tools are far quicker and more flexible that the Powerpoint presentation editor.
* Versioning and collaboration tools - such as git - are much easier to use with a text file than a Powerpoint presentation file.
* Other flat file formats can be embedded.

	With md2pptx you can use a simple Task Management format called [Taskpaper](https://support.omnigroup.com/omnifocus-taskpaper-reference/) to embed tasks. md2pptx will extract such tasks and generated a "Tasks" slide at the end of the presentation.

The flat file format that md2pptx uses is Markdown. Using Markdown has further advantages:

* The same text could be used to start, or even complete, a document of a different kind. Perhaps a long-form document.
* You can render the material in a web browser. Builds of this very documentation are checked that way.
* Markdown is easy to write.
* Markdown is compact; The files are tiny.
* Markdown is used in popular sites, such as [Github](https://github.com).

Every piece of text you use to create a Powerpoint presentation with md2pptx is valid Markdown. While it might not render exactly the same way put through another Markdown processor, it is generally equivalent. This is one of the key aims of md2pptx.

### A Real World Use Case

The author developed a presentation over 10 years in Powerpoint and OpenOffice and LibreOffice. It became very inconsistent in formatting - fonts, colours, indentations, bullets, etc.. It was a horrible mess.

He took the trouble to convert it to Markdown and regenerated it with a very early version of md2pptx. The presentation looks nice again, with consistent formatting.

It was relatively little trouble to convert to Markdown. In fact it took about an hour to convert the 40 page presentation. The consistency gain was automatic.

## How Do You Use md2pptx?

You write Markdown in exactly the same way as normal, with some understanding of how Markdown is converted to slides (using the information in [Creating Slides](#creating-slides)).

To use md2pptx you need to

1. Download it.
1. Have Python 3 installed - at a reasonably high level.
1. Install python-pptx using the command `pip3 install python-pptx`. (You might have to install pip firsst.)
1. Invoke it.

The following instructions are for Unix-like systems. (It's developed and used by the developer on Mac OS but should also have identical syntax on Linux.) Windows users will need a slightly different form, but the principle is the same.

Here is a sample invocation:

	md2pptx powerpoint-filename < markdown-filename

An alternative is to have the Markdown be in-stream. md2pptx reads from stdin. You can, of course, use stdin in a pipeline. Indeed the developer uses this to pipe from another program.
Alternatively, you can specify both an input file and an output file:

	md2pptx markdown-filename powerpoint-filename

If the input file doesn't exist md2pptx will terminate with a message. If the input file is empty the same thing will happen.

If you don't specify an input filename and don't redirect stdin md2pptx will await terminal input. This works but is probably only useful when experimenting with syntax with md2pptx.

Messages are written to stderr.

## python-pptx license

While [python-pptx](http://python-pptx.readthedocs.io/en/latest/) is not included in md2pptx it is used by it.

To quote from the python-pptx license statement:

	The MIT License (MIT)
	Copyright (c) 2013 Steve Canny, https://github.com/scanny

	Permission is hereby granted, free of charge, to any person obtaining a copy
	of this software and associated documentation files (the "Software"), to deal
	in the Software without restriction, including without limitation the rights
	to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
	copies of the Software, and to permit persons to whom the Software is
	furnished to do so, subject to the following conditions:

	The above copyright notice and this permission notice shall be included in
	all copies or substantial portions of the Software.

	THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
	IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
	FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
	AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
	LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
	THE SOFTWARE.

## Change Log

|Level|Date|What|
|:--|:---|:-----|
|1.1|15 October 2020| Introduce Template as a better replacement for Master - which still works. Add German characters.|
|1.0|13 October 2020| Python 3&comma; Support input filename as first command line parameter.|
|0.9|4 September 2020| Footnote slide support|
|0.8|14 June 2020|`bgcolor` is now `style.bgcolor`. Added `style.fgcolor` and `style.emphasis`.|
|0.7.3|24 May 2020|Allow background colouring via `span` elements|
|0.7.2|14 April 2020| Support three graphics on a slide. Added `&equals;` entity reference. Added `tableMargin`.|
|0.7.1|14 November 2019| Make slide titles longer. Fixed formatting issue with one-graphic-above-another table slide.|
|0.7|3 November 2019|Support `abbr` element as a glossary item. Each distinct term leads to a glossary slide entry at the back of the presentation.|
|0.6|8 October 2019|Support vertical pair of graphics in a table<br/>Fixed some issues with Markdown-syntax hyperlinks<br/>Support escaped square brackets `\[` and `\]`&comma;`&lsqb;` and `&rsqb;` being newly-supported alternatives|
|0.5|12 May 2019|CriticMarkup support|
|0.4.5|5 May 2019|Some numeric character references|
|0.4.4|6 March 2019|Processing summary slide shows build date and time|
|0.4.3|20 January 2019|Support a few HTML entity references - punctuation and arrows.<br/>Support split task slide sets - completed and incomplete.<br/>Task tags are sorted.|
|0.4.2|13 January 2019|Tasks slide set controllable with metadata `taskSlides` and `tasksPerSlide`|
|0.4.1|9 January 2019|Enhanced Taskpaper support with `@due`&comma; `@tags`&comma; and `@done`&comma; and reworked as a series of table slides.|
|0.4|7 January 2019|Support shrinking of table cell font and margins.<br/>Added two-to-by-two grid of graphics on a slide.|
|0.3.2|3 January 2019|Support \\# as a literal octothorpe/hash/pound.<br/>Tidied up reporting.<br/>Added superscript, subscript, strikethrough, and underline text effects.|
|0.3.1|3 November 2018|Fixed support for `<br/>` so it won't create a bullet on the new line.|
|0.3|22 October 2018|Added customisation for bold and italic text|
|0.2|3 September 2018|Added ways of controlling bullet sizes|
|0.1|1 April 2018|Initial Prototype|

## Creating Slides

Let's start with a simple example. Consider the following text.

	template: Martin Template.pptx
	pageTitleSize: 24
	sectionTitleSize: 30

	# This Is the Presentation Title Page

	## This Is A Section

	### This Is A Bullet Slide

	* Bullet One
		* Sub-bullet A
		* Sub-bullet B
	* Bullet Two
	* Bullet Three

You can try it if you like. Just cut it and paste it into a file. Call it something like Example.markdown.

It will render something like this:


![](Simple1.png)
![](Simple2.png)
![](Simple3.png)
![](Simple4.png)

The first slide is special, and an inevitable feature of using the python-pptx library. You will probably want to remove it before publishing.

Because the first slide has to be there md2pptx uses it to create a processing summary. The processing summary slide shows processing options, the time and date the presentation was created by md2pptx, and metadata.

Metadata is specified in the first three lines of this sample. In general metadata is the lines before the first blank line. It consists of key/value pairs, with the key separated from the value by a colon.

In this case the metadata specifies three things:

1. The Powerpoint file the presentation is based on is "Martin Template.pptx" - which is provided with md2pptx.
1. Each page with a title has a title font 24 pixels high.
1. Each presentation section slide has a title font 30 pixels high.

All of the above are optional but you will almost certainly want to specify a template. Feel free to copy Martin Template.pptx and make stylistic changes.

For more on metadata see [Controlling The Presentation With Metadata](#controlling-the-presentation-with-metadata).

As you can see the format of each slide is fairly straightforward. How to code slides is described in the following sections.

### Presentation Title Slides

You code a presentation title slide with a Markdown Heading Level 1:

	# This Is the Presentation Title Page

If you type anything in subsequent lines - before a blank line - the text will appear as extra lines in the presentation title. You might use this, for example, to add the presentation authors' details.

### Presentation Section Slides

You code a presentation section slide with a Markdown Heading Level 2:

	## This Is A Section

You can code multiple lines, as with [Presentation Title slides](#presentation-title-slides).

### Bullet Slides

Bullet slides use Markdown bulleted lists, which can be nested. This example shows two levels of nesting.

The title of the slide is defined by coding a Markdown Heading Level 3.


	### This Is A Bullet Slide

	* Bullet One
		* Sub-bullet A
		* Sub-bullet B
	* Bullet Two
	* Bullet Three

Bulleted list items are introduced by an asterisk.

**NOTE:** Some dialects of Markdown allow other bullet markers but md2pptx doesn't. You can be sure by coding `*` you have valid Markdown that md2pptx can also process correctly. For an explanation of why you have to stick to `*` see [here](#task-list-slides).

To nest bullets use a tab character or 4 spaces to indent the sub-bullets. md2pptx doesn't have a limit on the level of nesting but Powerpoint probably does.

Terminate the bulleted list slide with a blank line.

### Graphics Slides

As with [bullet slides](#bullet-slides), code the slide title as a heading level 3. Specify the graphic to embed with the standard Markdown image reference:

	### A Graphic Slide

	![](graphics/my-graphic.png)

The graphic will be scaled to sensibly fill the slide area.

Don't code anything inside the square brackets.

**HINT:** If you want two graphics side by side use a single-row table, described [here](#special-case-two-graphics-side-by-side). If you want two graphics one above the other use a two-row, single-column table, described [here](#special-case-one-graphic-above-another).

### Table Slides

You can create a table slide using Markdown's table format.

Code a title with a heading level 3. Then code a table. Here is a simple example:

	|Left Heading|Centre Heading|Right Heading|
	|:----|:-:|--:|
	|Alpha|Bravo|1|
	|Charlie|Delta|2|

In this example there are three columns and three rows. The first row is the heading row. The third and fourth rows are data rows.

The second row controls the alignment of each column and their width:

* In the first column the leading colon denotes the text is to be left-justified.
* In the second column the colons either end denotes the text is to be centred.
* In the third column the trailing colon denotes the text is to be right-justified.
* According to the number of dashes the columns have widths in the ratio of 4 to 1 to 2.

In other Markdown processors the widths of the columns can't be specified in this way; The relative width specifications will be ignored.

Each cell can consist of text, which will wrap as necessary. You can't embed images in a table slide. But see [here](#special-case-two-graphics-side-by-side) and [here](#special-case-two-by-two-grid-of-graphics).

#### Special Case: Two Graphics Side By Side

The best Markdown fit for two graphics side by side is a single row table with two cells. md2pptx will "special case" such a table.

If you code something like this the two graphics will be placed next to each other:

	|![](left-graphic.png)|![](right-graphic.png)|

A table won't be created in this case.

Don't code any headings or more than one row.

#### Special Case: Two By Two Grid Of Graphics

The best Markdown fit for four graphics on a slide is a two row table with two pairs of cells. md2pptx will "special case" such a table.

If you code something like this the four graphics will be placed in two rows of two:

	|![](top-left-graphic.png)|![](top-right-graphic.png)|
	|![](bottom-left-graphic.png)|![](bottom-right-graphic.png)|

A table won't be created in this case.

Don't code any headings or more than two rows.

To achieve the best result some margins around the graphics are reduced.

#### Special Case: Three Graphics On A Slide

The best Markdown fit for three graphics on a slide is a two row table&colon;

* The first row has two graphics.
* The second row has one graphic, centred in the row.

md2pptx will "special case" such a table.

If you code something like this the three graphics will be placed appropriately:

	|![](top-left-graphic.png)|![](top-right-graphic.png)|
	|![](bottom-graphic.png)|

A table won't be created in this case.

Don't code any headings or more than two rows.

To achieve the best result some margins around the graphics are reduced.

#### Special Case: One Graphic Above Another

The best Markdown fit for two graphics, on above the other, on a slide is a two row table with a single cell in each row. md2pptx will "special case" such a table.

If you code something like this the two graphics will be placed in two rows of one:

	|![](top-graphic.png)|
	|![](bottom-graphic.png)|

A table won't be created in this case.

Don't code any headings or more than two rows.

### Code Slides

You can create a slide where the body is in a monospace font, without bullets.

The heading for the slide is introduced with heading level 3 - `### `.

Each line of the code fragment - to be displayed in a monospace font - is indented with 4 spaces:

	### This Is A Code Slide

	    for i in range(10):
	        print(i)


### Task List Slides

You can create tasks in a subset of the [Taskpaper](https://support.omnigroup.com/omnifocus-taskpaper-reference/) format by coding a line that starts with a `-`:

	- MARTIN: Complete The User Guide

If md2pptx detects any such tasks it removes them from the body of the presentation and adds them to a special set of "Tasks" slides at the end of the presentation. If no tasks are detected these slides are not created.

Taskpaper is a very flexible and simple text-based task management system. md2pptx parses anything after the `-` simplistically but doesn't invalidate the Taskpaper format:

* Anything after the `-` leading character and before the first `@` symbol, if any, is the task title.
* Anything bracketed by `@due(` and `)` is treated as a due date - but the date isn't actively parsed.
* Anything bracketed by `@tags(` and `)` is treated as a set of tags. Tags are separated by a space or a comma and they are sorted.
* Anything bracketed by `@done(` and `)` is treated as a completion date - but it isn't actively parsed. (An uncompleted task need not have anything in inside the bracket - or the `@done` could be missing.)

The task title, any due date, any tags, and any completion information, are added as a table row to the set of tasks.

Because of Taskpaper support you can't start a bullet with a `-`. So always start bulleted list items with a `*`.

Tasks on the Tasks slides are shown with the slide number they were coded on.

Here's a more comprehensive example. Coding

	- Complete abstract @due(2019-01-11) @tags(Anna,Martin)

will cause a task to appear with title "Complete abstract", a due date of "2019-01-11", and tags "Anna,Martin". In this case the task has implicitly not been completed. (It would be possible to achieve the same effect by coding `@done()`.)

Task slides are paginated: Multiple task slides are created, each with the task slide number appended to the title, if there are more than a certain number of tasks.

You can control task slide production by specifying `taskSlides` and `tasksPerSlide`. See [Controlling Task Slide Production With `taskSlides` and `tasksPerSlide`](#controlling-task-slide-production-with-taskslides-and-tasksperslide).

## Hyperlinks

To code a hyperlink in a slide code something like:

	[IBM Website](http://www.ibm.com)

It will be rendered with the text "IBM Website" displayed: [IBM Website](http://www.ibm.com)

## HTML Comments

You can use HTML-style comments, ranging over multiple lines.

Start the first line with `<!--`.

End the last line with `-->`.

md2pptx will throw away HTML comments, rather than adding them to the output file.

**NOTE:** Other Markdown processors will copy the comment into the output file. Put nothing in the comments that is sensitive.

## Special Text Formatting

Markdown and md2pptx allow additional ways of formatting text. The syntax md2pptx supports is a subset of what many Markdown processors allow.

To specify **bold** surround the text with pairs of asterisks - `**bold**`.

To specify *italics* surround the text with single asterisks - `*italics*`.

If you actually want an asterisk code either `\*` or the asterisk surrounded by spaces. (An asterisk at the end of a line need only have a preceding space.) Alternatively you can code an HTML entity reference - `&lowast;`.

If you actually want an octothorpe/hash/pound symbol (rendered "\#") code `\#`.

You can use bold and italics syntax to change the colour of highlighted text. See [here](#specifying-bold-and-italic-text-colour-with-boldcolour-and-italiccolour) for more.

To specify a `monospace font` use the back tick character - `` ` `` - at the start and end of the text run.

To force a line break code `<br/>`. This, being HTML, is legitimate in Markdown and will be treated as a line break. I don't want one here so I won't code one here.

Some other HTML-originated text effects work - as Markdown allows you to embed HTML (elements and attributes):

|Effect|HTML Element|Example|Produces|
|:--|:---|:-----|:-|
|Superscript|`sup`|`x<sup>2</sup>`|x<sup>2</sup>|
|Subscript|`sub`|`C<sub>6</sub>H<sub>12</sub>O<sub>6</sub>`|C<sub>6</sub>H<sub>12</sub>O<sub>6</sub>|
|Underline|`ins`|`this is <ins>important</ins>`|this is <ins>important</ins>|
|Strikethrough|`del`|`this is <del>obsolete>/del>`|this is <del>obsolete</del>|

### Using HTML `<style>` Elements To Specify Text Colours And Underlining

You can set the background or foreground colour of a piece of text. To do this use the `<span>` HTML element. Here is an example:

    I would like to highlight <span class="yellow">this bit</span> but not **this** bit.

In this example the `span` element specifies a `class` attribute. The class name must match one specified in the metadata using one of

* `style.bgcolor.` - described in <a href="#associating-a-class-name-with-a-background-colour-with-stylebgcolor">Associating A Class Name With A Background Colour With <code>style.bgcolor</code></a>.
* `style.fgcolor` - described in <a href="#associating-a-class-name-with-a-foreground-colour-with-stylefgcolor">Associating A Class Name With A Foreground Colour With <code>style.fgcolor</code></a>.
* `style.emphasis` - described in <a href="#associating-a-class-name-with-text-emphasis-with-styleemphasis">Associating A Class Name With Text Emphasis With <code>style.emphasis</code></a>.


**Note:** A fragment of text in a span can't use any other text effect, such as bolding or italics.

If you want to be able to process the text using a normal Markdown processor you can code Cascading Style Sheet (CSS) using the HTML `<style>` element. md2pptx will ignore any HTML after the metadata and before the first real Markdown text. For example:

	<style>
	.mytest{
	    text-decoration: underline;
	    color: #FF0000;
	    background-color: #FFFF00;
	    font-weight: bold;
	    font-style: italic;

	}
	</style>

The above uses only the style elements that md2pptx supports with `style.` metadata. I relies on you coding

	<span class="mytest">Here is some text</span>

for example - as it uses the class `mytest`.

### HTML Entity References

md2pptx supports a few [HTML entity references](https://en.wikipedia.org/wiki/List_of_XML_and_HTML_character_entity_references)&colon;


|Entity Reference|Character|Entity Reference|Character|Entity Reference|Character|
|:--|:---|:-----|:-|:-|:-|
|`&lt;`|&lt;|`&larr;`|&larr;|`&auml;`|&auml;|
|`&gt;`|&gt;|`&rarr;`|&rarr;|`&Auml;`|&Auml;|
|`&ge;`|&ge;|`&uarr;`|&uarr;|`&uuml;`|&uuml;|
|`&le;`|&le;|`&darr;`|&darr;|`&Uuml;`|&Uuml;|
|`&asymp;`|&asymp;|`&harr;`|&harr;|`&ouml;`|&ouml;|
|`&Delta;`|&Delta;|`&varr;`|&varr;|`&Ouml;`|&Ouml;|
|`&delta;`|&delta;|`&nearr;`|&nearr;|`&szlig;`|&szlig;|
|`&sim;`|&sim;|`&nwarr;`|&nwarr;|`&euro;`|&euro;|
|`&lowast;`|&lowast;|`&searr;`|&searr;|
|`&semi;`|&semi;|`&swarr;`|&swarr;|
|`&colon;`|&colon;|`&lsqb;`|&lsqb;|
|`&amp;`|&amp;|`&rsqb;`|&rsqb;|
|`&comma;`|&comma;|`&infin;`|&infin;|

### Numeric Character References

md2pptx supports a few [HTML numeric character references](https://en.wikipedia.org/wiki/Numeric_character_reference)&colon;

* Some like `&#223;` - with 3 or 4 numeric digits. (This produces the character '&#223;').
* Some like `&#x03A3;` - with 4 hexadecimal digits. (This produces the character '&#x03A3;').

(Some emoji work. For example you can paste 🦖 in to the source.)

###  Escaped Characters

md2pptx supports a few escaped characters. Of most interest are the two square bracket characters:

* `\[`
* `\]`

### CriticMarkup

md2pptx supports [CriticMarkup](http://criticmarkup.com/) for text. To quote from their home page:

>CriticMarkup is a way for authors and editors to track changes to documents in plain text. As with Markdown, small groups of distinctive characters allow you to highlight insertions, deletions, substitutions and comments, all without the overhead of heavy, proprietary office suites.

md2pptx supports all five markup elements. In common with other CriticMarkup processors, md2pptx shows the markup and merely colours the markup and marked up text appropriately:

* Insertion - `{++` and  `++}` - rendered in green.

    For example <span style="color:#00C300">`{++ This text was inserted ++}`</span>

* Deletion - `{--` and `--}` - rendered in red.

    For example <span style="color:#C30000">`{-- This text was deleted --}`</span>

* Comment - `{>>` and `<<}` - rendered in blue.

    For example <span style="color:#0000C3">`{>> This text is a comment <<}`</span>

* Replacement - `{~~`, `~>` `~~}` - rendered in orange.

    For example <span style="color:#FF8C00">`{~~old text~>replacement text~~}`</span>

* Highlight - `{==` and `==}` - rendered in purple.

    For example <span style="color:#C300C3">`{== This is a highlight ==}`</span>

In the above examples the deletions, insertions, replacements etc don't actually happen; They are just marked. Editing tools are needed to actually perform these actions, once the reviewer's comments have been accepted.

## Creating A Glossary Of Terms

You can use the HTML `abbr` element to generate a glossary entry. For example,

	<abbr title='British Broadcasting Corporation'>BBC</abbr>

In this example the glossary term is "BBC" and its definition is "British Broadcasting Corporation".

One or more glossary slides will appear at the end of the presentation if any such terms are defined. When you create a glossary entry two things will happen:

* Both the term and its definition will appear in the glossary.
* Only the term will appear in the slide with the `abbr` element.

If you define the term more than once only the first use will be included in the glossary. All uses of the term will appear in the normal slides.

A Glossary Table slide comprises two columns: A narrow column with the terms (or acronyms), and a wider column with their definitions (or meanings).

If you use the `abbr` HTML element most markdown processors will treat it as HTML and hovering over the term will reveal the definition.

You can control various aspects of the glossary's appearance using metadata. How to is described [here](#controlling-glossary-slide-production-with-glossarytitle-glossaryterm-glossarymeaningglossarymeaningwidth-and-glossaryperpage).

## Creating Footnotes

You can create and reference footnotes.

### Creating A Footnote

To define a footnote code `[^name]: ` on a new line. The remainder of the line will be the footnote text.

If you have defined footnotes one or more Footnotes pages will be added to the end of the presentation.

Footnotes are automatically numbered, starting with 1.

### Referring To A Footnote

To refer to a footnote code `[^name]`. The footnote's number (automatically generated) will appear like so:

&nbsp;&nbsp;&nbsp;&nbsp;This is a footnote reference<sup>[4]</sup>.

If the name doesn't match a footnote a question mark will be printed instead of the footnote number.

## Controlling The Presentation With Metadata

You can control some aspects of md2pptx's processing using metadata.

### Specifying Metadata

You specify metadata in the lines before the first blank line. It consists of key/value pairs, with the key separated from the value by a colon.

While some Markdown processors handle metadata, most ignore it. Conversely, while md2pptx will print **all** the metadata it encounters on the first slide, it will practically ignore metadata it doesn't understand.

### Metadata Keys

The following sections describe each of the metadata keys.

#### Slide Numbers - `numbers`

md2pptx can add slide numbers. These are generated by md2pptx itself (or hardcoded) and are not the same as ones you can turn on in a footer.

The default value is `no`. You can turn them on for all slides with `yes` or non-title slides with `content`.

Example:

	numbers: yes

#### Page Title Size - `pageTitleSize`

You can specify the point size of each page that isn't a section divider or title slide. The size is specified in points.

Example:

	pageTitleSize: 24

The default is 30 points.

#### Section Title Size - `sectionTitleSize`

You can specify the point size of the title text for each page that's a section divider or title slide. The size is specified in points.

Example:

	sectionTitleSize: 42

The default is 40 points.

#### Section Subtitle Size - `sectionSubtitleSize`

You can specify the point size of the subtitle text for each page that's a section divider or title slide. The size is specified in points.

The subtitle text is the second and subsequent lines of the title - generally a separate text shape on the slide.

Example:

	sectionSubtitleSize: 24

The default is 28 points.

#### Monospace Font - `monoFont`

You can specify which font to use for monospaced text - such as on [code slides](#code-slides).

Example:

	monoFont: Arial

The default is Courier.

#### Margin size - `marginBase` and `tableMargin`

You can increase or decrease the margin around things - in decimal fractions of an inch.

For a table you can specify the left and right margins using `tableMargin`.
For everything else use `marginBase`.

Example:

	marginBase: 0.5

The default is 0.2 (inches).

#### Associating A Class Name with A Background Colour With `style.bgcolor`

You can use HTML `<span>` elements to set the background colour, as described in <a href="#using-html-ltstylegt-elements-to-specify-text-colours-and-underlining">Using HTML &lt;style&gt; Elements To Specify Text Colours And Underlining</a>.

Here is an example:

    style.bgcolor.yellow: FFFF00

In this example the class "yellow" is associated with a background colour, defined in RGB terms as hexadecimal FFFF00, which is:

* 255 of red
* 255 of green
* 0 of blue

which is in fact yellow.

#### Associating A Class Name with A Foreground Colour With `style.fgcolor`

You can use HTML `<span>` elements to set the foreground colour, as described in <a href="#using-html-ltstylegt-elements-to-specify-text-colours-and-underlining">Using HTML &lt;style&gt; Elements To Specify Text Colours And Underlining</a>.

Here is an example:

    style.fgcolor.red: FF0000

In this example the class "red" is associated with a foreground colour, defined in RGB terms as hexadecimal FF0000, which is:

* 255 of red
* 0 of green
* 0 of blue

which is in fact red.

#### Associating A Class Name With Text Emphasis With `style.emphasis`

You can use HTML `<span>` elements to bold text, make it italic, or underline it - as described in <a href="#using-html-ltstylegt-elements-to-specify-text-colours-and-underlining">Using HTML &lt;style&gt; Elements To Specify Text Colours And Underlining</a>.

Here is an example:

    style.emphasis.important: bold underline

In this example the class "important" is associated with bolding the text and underlining it.

You can also use `italic`.

As the example shows, separate the emphasis attributes with a space.

#### Template Presentation - `template`

You can specify a different template file to create the presentation from than the one supplied with python-pptx. The one supplied with md2pptx is a very good one to work from:

	template: Martin Template.pptx

If you want to create your own template you probably want to take Martin Template.pptx and modify it. See [Modifying The Slide Template](modifying-the-slide-template) for more information on how to do so.

(For compatibility purposes, you can continue to use `master` instead of `template`. It's probably better practice, though, to use `template`.)

#### "Chevron Style" Table Of Contents - `tocStyle` And `tocTitle`

If you have a Table Of Contents slide - with each section title listed as a top level bullet you can create a "Chevron Style" Table Of Contents slide. It will look something like this:

![](chevronTOC.png)

If your Table Of Contents slide's title is "Topics" you need only code

	tocStyle: chevron

If your Table Of Contents slide's title is something else you need to additionally code something like

	tocTitle: Agenda

If you specify `tocStyle: chevron` the section headings will be rendered something like this:

![](chevronSection.png)

Here the section is highlighted by removing the background.

**NOTES:**

* Ensure only one slide has the same title as the Table Of Contents slide. Otherwise md2pptx will attempt to render the other slides as if they were a Table Of Contents slide.
* Ensure the section slides' titles are unique. Otherwise more than one chevron will be highlighted on the relevant section slide.

#### Specifying An Abstract Slide With `abstractTitle`

You can arrange for a single-level bulleted list slide to be formatted specially - as an abstract.

Instead of each list item having a bullet, it is treated as a paragraph. This is more appropriate for an abstract slide. (An extra blank paragraph is added between each paragraph - to space it out.)

To indicate an abstract slide code

	abstractTitle: Abstract

Any slide with the title matching the value of abstractTitle will be rendered as an abstract slide.

#### Specifying Text Size With `baseTextSize` And `baseTextDecrement`

You can control the size of text - in table slides, code slides, and bulleted list slides - with two metadata tags: `baseTextSize` And `baseTextDecrement`.
If you don't specify `baseTextSize` the base presentation's font sizes are used.

If you specify a `baseTextSize` value code, tables, and the top-level bullet use this size, which is specified in points.
Further, if you specify `baseTextDecrement` each successive level of bullets' font size is decremented by this number of points.
The default for `baseTextDecrement` is 2 points.

For example, if you code

	baseTextSize: 20
	baseTextDecrement: 1

the top-level bullet uses a 20 point font, the next level down a 19 point font, and so on.

If you just coded

	baseTextSize: 20

the top-level bullet uses a 20 point font, the next level down a 18 point font, and so on.

#### Specifying Bold And Italic Text Colour With `BoldColour` And `ItalicColour`

You can modify how md2pptx formats bold and italic colours:

* To specify the colour of bold text use `BoldColour` (or `BoldColor`).
* To specify the colour of italic text use `ItalicColour` (or `ItalicColor`).

For example:

	BoldColour: ACCENT 1

Will cause text marked like so `**I am bold**` to be rendered in the presentation's smartmaster's "Accent 1" colour.

The values you can use for `BoldColour` and `ItalicColour` are:

* NONE
* ACCENT 1
* ACCENT 2
* ACCENT 3
* ACCENT 4
* ACCENT 5
* ACCENT 6
* BACKGROUND 1
* BACKGROUND 2
* DARK 1
* DARK 2
* FOLLOWED HYPERLINK
* HYPERLINK
* LIGHT 1
* LIGHT 2
* TEXT 1
* TEXT 2
* MIXED

As you can probably guess, these are standard values for python-pptx and, ultimately, PowerPoint.

**Note:** For the values you can use any capitalisation you like (or none). e.g. `ItalicColor: dark 1`.

#### Specifying Bold And Italic Text Effects With `BoldBold` And `ItalicItalic`

You can modify how md2pptx formats bold and italic text.

If you don't want bold text to actually be bold code

	BoldBold: no

If you don't want italic text to actually be italic code

	ItalicItalic: no

The default for both of these is, of course, `yes` so that bold text is bold and italic text is italic.

These options were added so that `BoldColour` and `ItalicColour` could just become colour effects. See [here](#specifying-bold-and-italic-text-colour-with-boldcolour-and-italiccolour).

#### Shrinking Tables With `compactTables`

You can reduce the size of a table on the page with `compactTables`. If you specify a value larger than 0 two things will happen:

* The font will use whatever point size you specify.
* The margins around the text in a cell will be reduced to 0.

For example, to remove the margins and reduce the font size to 16pt code

    compactTables: 16

#### Controlling Task Slide Production With `taskSlides` and `tasksPerSlide`

Before unleashing your presentation on the world you probably want to remove the Task List slides from it. You can control what tasks are shown, if any, with `taskSlides`. It can take four different values:

* `all` (the default) which shows both complete and incomplete tasks.
* `none` which hides the task list slides altogether.
* `done` which shows completed tasks only.
* `remaining` which shows only tasks that haven't been completed.
* `separate` which separates the tasks into slides with completed tasks and slides with incomplete tasks.

A completed task is one where the `@done` attribute has been coded with something inside the brackets. An incomplete task is one where either the `@done` attribute wasn't coded or there is nothing inside the brackets.

For example, to suppress all task slides code

	taskSlides: none

Though you wouldn't normally need to do this, you can control how many tasks appear on a slide with `tasksPerSlide`. For example, coding

	tasksPerSlide: 10

will limit the numberof tasks on a slide to 10. The default is 20 tasks per slide.

#### Controlling Glossary Slide Production With `glossaryTitle`, `glossaryTerm`, `glossaryMeaning`,`glossaryMeaningWidth`, and `glossaryPerPage`

[Creating A Glossary Of Terms](#creating-a-glossary-of-terms) describes how you can use the `abbr` element to generate a glossary of terms.

You can control various aspects of the appearance of the Glossary Slide(s) using the following metadata items.

For example, coding

	glossaryTitle: Definitions

will cause the Glossary Slide title to be "Definitions". The default is "Glossary".

Coding

	glossaryTerm: Acronym

will cause the first heading of the table in the Glossary Slide to be "Acronym". The default is "Term".

Coding

	glossaryMeaning: Definition

will cause the second heading of the table in the Glossary Slide to be "Definition". The default is "Meaning".

Coding

	glossaryMeaningWidth: 3

will cause the width of the second column in the table to be three times that of the first column.

Coding

	glossaryPerPage: 10

will cause the maximum number of glossary items on a Glossary Slide to be 10. If there are more terms, a second page will be created. And so on. The default is 20.


## Modifying The Slide Template

The included template presentation - Martin Template.pptx - is what the author tested with and gives good results. However, you probably want to develop your own template from it.

This section is a basic introduction to the rules of the game for doing so.

### Basics

Don't change the order of the slides in the slide master view and don't delete any elements. It's probably also not useful to add elements. Take care with moving and resizing elements; It's probably best to experiment to see what effects you get.

### Slide Template Sequence

The following table shows how each slide type is created.

|Slide Type|Origin|Non-Title Content|
|:--|:---|:-----|
|Processing Summary|Original slide from Template|Metadata: Second Shape|
|Presentation Title|Slide Layout 0|Subtitle: Second Shape|
|Section|Slide Layout 1|Subtitle: Second Shape|
|Graphic With Title|Slide Layout 5|Graphic: New Shape|
|Graphic Without Title|Slide Layout 6|Graphic: New Shape|
|Code|Slide Layout 5|Code: New Shape|
|Content|Slide Layout 2|Bulleted List: Second Shape|
|Table|Slide Layout 5|Table: New Shape|
|Tasks|Slide Layout 2| Bulleted List: Second Shape|

**Notes:**

1. When looking for a title md2pptx looks first for a title shape and, failing that, uses the first shape. It's the Template Designer's responsibility to size and position it sensibly.
2. "New Shape" means md2pptx will create a new shape with, hopefully, sensible position and size.
3. With "Second Shape" it's the Template Designer's responsibility to size and position it sensibly.




