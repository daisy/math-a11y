# Recommended Microsoft Word 365 Settings for Using the Equation Editor with Screen Readers

## Background

People who use screen readers can create mathematical  expressions by using linear input. For example, \sqrt 4 = 2 is the linear expression that can produce the professional appearance of the "square root of four equals two."  Almost all math can be produced in this way. This document identifies the recommended settings in Microsoft Word when using a screen reader.

Most global Word Options apply to all new documents you create.
When you change settings in File > Options, you’re typically adjusting application‑wide preferences. These include things like:
 • Proofing settings (AutoCorrect, spelling/grammar options)
 • Save settings (default file format, AutoRecover intervals)
 • Display settings
 • Advanced editing behavior
 • Accessibility options
These are stored in your user profile (usually in the Windows registry and/or the Normal template, depending on the setting).

We will be  making changes in the options and in the Word equation editor settings. Open a new blank word file.


## Enable Math AutoCorrect Everywhere 
1.  Press Alt + F to open File.
2. Press T for Options.
3.  Arrow down to Proofing.
4.  Tab to AutoCorrect Options and press Enter.
5.  Shift + tab to get to the Math AutoCorrect tab.
6.  There are many options. The ones you want to check are: 
* Check - Use Math AutoCorrect rules outside of math regions.
* Check -Replace text as you type.
This makes typing math much faster and avoids hunting through menus. If this interferes with your screen reader, you can leave it unchecked.

NOTE: You must press OK, which can be found at the end of the settings. You can also move to the main list of options and shift + tab to quickly reach the OK button. You must press the OK button when exiting the options to save them.

##  Structural navigation for reading math in context 

1.  Press Alt + F to open File.
2. Press T for Options.
3.  Arrow down toAdvanced.
4. Tab into advanced and find Show document content:
*Check - Show bookmarks
* Uncheck - Show field codes instead of their values. (Field codes interfere with math reading)
5.	Under Editing options:
* check - Use smart cut and paste
These don’t change math itself, but they prevent Word from collapsing math into unreadable blobs.

## Disable visual-only math features that confuse screen readers

While in the Word options enter the Display:
• uncheck- Show text wrapped within the document window
•uncheck - Show animations in the document
This reduces cursor jumps when a screen reader is tracking math expressions.

## Equation editing options

In the main body of the blank Word document, type alt + "=" (Alt plus the equal symbol.).
This places you inside the Word equation editing field. Type \sqrt 4 = 2.

Now press alt and you will be taken to the equation tab. Press down arrow  and you will be in the options.
Right arrow and you will be in the selection for Unicode or LaTex. we suggest selecting unicode.  Set this as the default.

In the same area, press right arrow to the selection for conversions between linear or professional. Make linear the default.

When you go back to the main document and you are in the equation, hit alt again and check that Unicode and linear are the default options.

## Tips
ctrl + = puts it in professional
ctrl + shift + = puts it in linear

You can press ctrl + a to select all and then press ctrl + = to put all equations into professional for printing or to submit to others who will be expecting professional looking math.



## Sources

* University of Houston guide on typing math in Word