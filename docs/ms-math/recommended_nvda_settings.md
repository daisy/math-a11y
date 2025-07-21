For the Equations in Microsoft Word , please use
the following NVDA and MathCAT settings if possible.

For NVDA to work with digital math content, it needs to have an add-on
installed. The best add-on for speaking math and creating math braille
codes in Nemeth or UEB is the MathCATAdd-on. The DAISY/Math-a11y group’s efforts are
focused on evaluating how well the Microsoft Equations work in Word for
screen reader users. To avoid stumbling across known limitations in
screen readers, please use the following:

- Windows 10 or Windows 11
- NVDA Version 2025.1 or later.
- MathCAT 0.6.10 or later

# NVDA Settings

## Choosing your Voice Synthesizer:

There are usually at least three choices: eSpeak NG, Microsoft Speech
API, and Windows OneCore voices. All synthesizers work.

# Installing MathCAT Add-on:

1.   Open NVDA settings with NVDA key + n and arrow down to the  Tools menu and press enter.
2. Arrow don to the Ad-on store and press enter.
  3. Press shift  tab until you get to the installed Add-ons and arrow over to the Available Add-ons.
  4. Press tab until you get to the Search field and type in MathCAT and press enter.
  6. Press enter and confirm that you want to install.
  7. The MathCAT Add-on will be downloaded and installed.
   8.  NVDA will ask you to restart, select “Yes”
9. NVDA will restart with MathCAT enabled.

# Accessing the MathCAT settings

Note: in case you need it, see the [MathCAT Settings
documentation](https://nsoiffer.github.io/MathCAT/users.html).

To open the MathCAT Preferences dialog in NVDA:

1.  Start NVDA
2.  Use the NVDA Key + n to bring up the NVDA Menu.
    1.  The NVDA Key is usually the Caps lock and the INS key.
    2.  Alternatively, if using the mouse to find the NVDA icon in the
        toolbar. Note that the NVDA icon might be hidden. Use the show
        hidden icons option to bring it up.
3.  Select "Preferences" from the NVDA Menu
4.  Select “MathCAT Settings...” and press enter.

The MathCAT Preferences dialog has three categories of settings: Speech,
Navigation, and Braille.  To move between these three collection of settings, use CTRL + tab to cycle through them. Use tab to move through the options, and arrow keys to make a selection.

Below are the suggested Speech and
Navigation settings:

## Speech Category Settings:

| Setting                       | Set to                 |
|-------------------------------|------------------------|
| Generate Speech for           | Blindness \[Default\]  |
| Language | auto |
| Decimal separator for numbers | Auto|
| Speech Style                  | ClearSpeak \[Default\] |
| Speech Verbosity              | Verbose                |
| Relative speech rate | Set slider lower to slow down speech |
| Pause factor | slider  to pause between items |
| Make a sound when starting/ending math  | check box |
| Subject area to be used when it cannot be determined  | General |
| Speech for chemical equations | Off (H sub 2 O)        |

## Navigation Category Settings:

| Setting | Set to |
|----|----|
| Navigation mode | Enhanced \[Default\] |
| Navigation speech to use when beginning to navigate an equation | Speak \[Default\] |
| Speech amount for navigation | medium |
| Copy math as | MathML |

# Navigating Math with MathCAT

To use MathCAT with NVDA, it is helpful to understand the MathCAT
navigation modes, know the MathCAT navigation commands, and the keys
used to activate them.

## Navigation Modes

- **Enhanced mode**:  navigation is by mathematically meaningful pieces
  (operators, delimiters, and operands)

- **Simple mode**: this moves by words except when you get to a 2D
  notation (fractions, roots, …), then it speaks the entire notation.
  Zooming in lets you explore the 2D notation in the same mode. Zooming
  out or moving out of the 2D notation brings you back to the
  outer/higher level of navigation.

- **Character mode**: The character mode is actually two useful modes –
  word mode and character mode (zoom in to get "real" character mode).
   Moves by words/characters.  This differs for numbers of more than one
  digit and function names such as "sin" that are multiple characters.
  Otherwise, word and character navigation is the same.

## Navigation Commands

For a full list of navigation commands, see the [Navigation Commands
Table](https://nsoiffer.github.io/MathCAT/nav-commands.html#navigation-modes:~:text=Navigation%20Commands%20Table)
in the MathCAT User Guide.

## Activating MathCAT’s Navigation Features

To activate the MathCAT navigation features when reading a web page with
NVDA, navigate to an expression and then press “Enter”. NVDA should
speak “Math”.

## Basic Math Navigation

Typically, you will start at the first term of an expression. The left
and right arrow keys will move you through the terms at your current
navigation level. You might also move up and down levels if needed by
using the up and down arrow keys. alt+ctrl+arrow is used to move around
cells in a table.

Backspace will take you back to where you were, which is not always the
same as moving to the left. For example, if the right arrow moved you
out of a fraction, backspace will take you back to where you were in the
denominator and left arrow will land on the entire fraction.

You will likely find one mode of navigation the most natural for you
most of the time. This can be set in the MathCAT settings. However, at
any time during navigation, you can switch the navigation modes using
shift+up/down arrow. This is useful because each mode of navigation has
its strengths and weaknesses.

# To Copy a Math Expression from a Web Page Using MathCAT

While navigating an expression, “control+c” copies the MathML for the
part of the expression you are currently navigating. If you want to copy
the entire expression, be sure to up-arrow until you hear “Zoomed out
all the way”.

When you press “Control+c”, you should hear NVDA say “Copy as MathML”.

- If instead you hear NVDA say “No selection”, then MathCAT’s navigation
  has not been activated.
  - Make sure you are on the math expression and then press “Enter”.
  - If you hear NVDA say “Math”, then try pressing “control+c” again.
  - If when you pressed “Enter” it didn’t say “Math”, then use NVDA to
    navigate back to the math expression.
