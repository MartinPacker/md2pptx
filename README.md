# md2pptx
Markdown to Powerpoint Converter

**Usage:**

  `md2pptx output.pptx < input.markdown`

User guide to follow.

Before you can use this you need to:

  `pip install python-pptx`

as this code relies on that Python 3 library.

You will probably need to issue the following command from the directory where you install it:

  `chmod +x md2pptx`

I would also suggest you start with a presentation that references Martin Master.pptx in the metadata (before the first blank line). Here is a very simple deck that does exactly that.

```
master: Martin Master.pptx

# This Is A Presentation Title Page

## This Is A Presentation Section Page

### This Is A Bulleted List Page

* One
    * One A
    * One B
* Two
```