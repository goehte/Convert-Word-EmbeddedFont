Programm Name: docx_getembfnt
WordML - File Embedded Word Front: re-obfuscate font tool 

This C programm is generated based in the following forum thread: 
https://forum.softmaker.de/viewtopic.php?f=275&t=25741&p=126388#p123933

[ECMA-376-1] 
Documentantion for Office Open XML File Format ECMA-376-1:2016:
https://www.ecma-international.org/publications-and-standards/standards/ecma-376/
Part 1: Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf
Section 17.8.1, Page 670

GUID handling - obfuscation
Example: 001B70DC-AA60-4AD5-90EC-18A0948E1EAE
To:      AE1E8E94-A018-EC90-D54A-60AADC701B00

Build under Linux:
$ gcc main.c -o docx_getembfnt

TOOL USAGE

1. Open *.docx file as ZIP, most simple just rename it to *.zip

2. Unzip obfuscated fonts from the following folder: /word/fonts/ 
   Example file name of a obfuscated TTF font: font1.odttf

3. Open /word/fontTable.xml in a editor or XML viewer

4. Search for related "fontKey" GUID in <w:embedRegular> . e.g. "font1.ottf" == "rId1"  
   Example: <w:embedRegular w:fontKey="{7B19B49C-2336-4F82-AAD2-5D2BAE389560}" r:id="rId1"/>

5. Run: 
   Linux: $ docx_getembfnt <inputfile> <outputfile> <GUID>
   Windows $ docx_getembfnt32.exe <inputfile> <outputfile> <GUID>
   Examples usage (Linux):
    $ ./docx_getembfnt ./examples/font1.odttf ./examples/font1.ttf 7B19B49C-2336-4F82-AAD2-5D2BAE389560
    $ ./docx_getembfnt ./examples/font2.odttf ./examples/font2.ttf 373EDBC8-20C2-49CE-A04A-3A982C578A05
    $ ./docx_getembfnt ./examples/font3.odttf ./examples/font3.ttf 3E0C3015-B393-4155-9571-A242AB103A61
    $ ./docx_getembfnt ./examples/font4.odttf ./examples/font4.ttf 46483AA3-95DD-4E29-834A-6A30BA96940D

6. Output file you should able to insall as new font to your system.

Note: Sure the tool works also "vi­ce ver­sa". So you can make from a font file a obfuscated file.

This source-code is considered to be public-domain (freeware) without any liability and warranty.
-goehte80, Erlangen, 12-Apr-2021