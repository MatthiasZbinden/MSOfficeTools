# MS Office Tools
This repo contains small auxiliary tools for the MS Office. I wrote them for
myself, but while searching for a solution, I saw that also others seem to have
the same problem and could profit.

## PPT Globally Change Languages
**ppt_GloballyChangeLang.pptm**

The MS PowerPoint language management for the spell check is lousy. Basically,
you would have to select every text field and change the language to change it.
This gets especially painful if you work in an international company, where
documents are edited in different lingual environments and may contain multiple
languages.

This tool allows to globally apply a language via VBA macro.

### Features
- Chose 1 to many currently open presentations to be processed
- Process regular slides
- Process groups
- Process tables
- Process slide notes
- Process slide masters
- Process charts

### Release Notes
- **2021-08-19 / Version 2 Beta**
   - New features: process masters ([Issue #1](https://github.com/MatthiasZbinden/MSOfficeTools/issues/1)),
     process charts ([Issue #2](https://github.com/MatthiasZbinden/MSOfficeTools/issues/2))
   - Buxfix for Mac OS implemented ([Issue #3](https://github.com/MatthiasZbinden/MSOfficeTools/issues/3))
   - Not yet thoroughly tested. There seem to be crashes depending on Chart Type,
     as some Chart Types do not have all Object Members I want to access, but I
     don't know how to check this in a proper way (without on Error GoTo)
- **2021-01-20 / Version 1**
   - Initial release
   - Stable

### Screenshot
![Screenshot of the Option Dialog](./fig/fig1.png)

## EXPLORER Resize Images at Once
**explorer_MaxLength1200px.reg**

This is a piece of registry which you can use to add some image functionality
you use often. In my case it's "resize the longer side to 1200 px", as I often
have some high-res photos from lab tests, several MB per file, which inflate my
documents and emails. So usually, I select all files, chose this menu entry, and
after, the files are processed.

### Features
- Apply often used function to file types directly.

### Installing
- Using the default function (reduce long image side to 1200 px with Irfan View
  for JPEGs):
  - Ensure IrfanView is installed ([Link](https://www.irfanview.com/))
  - Open the .reg file in a text editor
  - Ensure the path for i_view64.exe is correct
  - Select .reg file, right-click, choose "Merge" (or double-click the .reg
    file)
- Add custom functions
  - Follow the approach inside the .reg file: one registry key is to add the
    context menu
    entry's name, one sub-key with command is for the executable plus
    parameters.
  - IrfanView command line options you'll find in the offline help > Overview >
    Command Line Options.
  - Works of course for any program supporting silent comand line.

### Release Notes
- **2021-08-30 / Version 1**
  - Initial release, stable
  - Resize JPG's longer side to 1200px
  
### Screebshot
![Screenshot of the Option Dialog](./fig/fig2.png)