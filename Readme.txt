Keyboard Monitor ActiveX
Version 3.0
Copyright© 2000, Konstantin Tretyakov
http://smartsite.cjb.net



The control allows you to:
- track every key, pressed on the keyboard, whether your program has the focus or not. 
- track some predefined sequences of keypresses (macros). 
- restrict keyboard messages from reaching their destination (this way your app will be the only one, knowing, that a key was pressed, all the others just won't notice it).



Installation:
-To install the control, copy the files
KdbMonitor.exp
KdbMonitor.lib
KdbMonitor.ocx
KtKbdHk3.dll
to your System directory (usually c:\windows\system).

When you deploy your application, that uses the OCX, you may, of course, install these files in the home-directory of your application, but it would be better to put them into the System directory.

POSTCARDWARE (see at the end).


Features:

- Feature #1 (Version 3 & higher) - you can now use up to 300 KeyMon controls simultaneosly. Still there cannot be more than 300 working KeyMons in the system at once (if there are already 300 apps, and each of them uses 1 KeyMon control, then the 301st wouldnt be able to use it).
PLUS there's one bug with this "multi-instancing"!!! Everything is neat with several controls in 1 app (quite a stupid case - 1 app needs only one control). The problems begin, however with several controls in several apps.
For Example:
1)You start 1 app, that uses the control (and calls it KeyMon1). It sets KeyMon1.Active=True
2)You start the 2nd instance of this app, which also enables its KeyMon
3)Everything is fine here, and they both can track the keyboard
4)Now you exit the first app
5)Oops - the control on the 2nd app doesn't track the keyboard anymore
6)If you then make Keymon1.Deactivate (don't mind, that it returns False) and then again KeyMon1.Activate on this app, the control starts functioning again
I tried as hard as I could to remove the bug, but it seems, that it's the problem of Windows. That is - if 1 app loads a DLL, and sets a hook with its HookProc inside the DLL, and then, no matter how many other apps are taking this DLL into usage, as soon as this 1st one is unloaded and the DLL detached from it, the Hook is "closed". Weird.

- Interesting feature - the control works at design time. Try: put a control on the form, then doubleclick it's property "Active" to make it True, then doubleclick its property "Passthru" to make it False. Try typing something. You won't be able to do it until you set Passthru to True or Active to False.

- Weird feature - it works differently when running a program in VB IDE. Try: make a program, that logs every key pressed (that uses the KeyDownEx/KeyUpEx event of the Control). Now set Passthru to False. Make your program activate the key-monitor at startup and deactivate - on unload. Now run your program and try typing something. Your program will not log the keypresses (and it should!). Now compile this app and run it standalone, without VB. Everything will be fine then.

- Stupid feature - it doesn't track the PrintScreen key. I don't know what to do here.

- Dull feature - if you specify a Macro, remember, that it is CASE-SENSITIVE.

- Please report other features at kt@smartsite.cjb.net



Usage:

Properties:

- AcceptRepeats
Specifies, whether every "repeated" key will be reported to you in the control's KeyDown, KeyUp, KeyPress, KeyDownEx, KeyUpEx events.
A "repeated" keypress, is a keypress, that is generated automatically, when a user holds some key down for some time.
It also affects the behaviour of the Macro-searching mechanism.
That is - if you set it to True, then when a user presses and holds some key fopr some time, then the control will report quite a lot of same KeyPress Events.
If you set it to False, then the key will be counted only once when it is pressed and released.
For Example, you specified a macro - "CVC", and set the AcceptRepeats to True.
Then, if some very slow-fingered user enters this combination, but holds every key for too long, the actual entered combination will be like "CCCCVVVVVCCCCC" and won't be recognised as a macro. In this case you should have set AcceptRepeats to False.
On the whole, you will probably rarely need this property to be "True"
However, it may be useful here: If you specify a macro "CCCCCCCCCCCCC" and set AcceptRepeats to True, then the macro will be reported when a user Holds the "C" key for some time.

- DifferNumpad
When a key is pressed, a KeyDownEx and KeyUpEx events tell you his name. If you press key "9" on the numeric keypad (with NumLock On) then
  If DifferNumpad=False, it will be reported as "9"
  If DifferNumpad=True, if will be reported as "{Numpad9}"
That's the difference.

- Passthru
Specifies, whether keyboard messages should reach their destination (the focused window). The control's Key*** events are fired anyway.
That is, if you set it to False, then the only reaction to keypresses in the system will be your KeyMon_Key*** event handlers.
Note: This preperty has effect only when the control's Active property is set to True.
You should set this property to the desired value every time you change the "Active" property, because every time the Keymonitor is activated, this property is reset to "True". (for safety reasons at designtime).
Note2: It really works at designtime, be careful.
Note3: If you have several active Key-Monitors at a time, and at least one of them has his Passthru property set to False, then keyboard messages won't reach their target window, but all the active Keymonitors will still "see" and report these messages.

- Active
Set it to true, to enable the control, and set it to False to disable it.
When Active=False, KeyMon will not notify you of the keypresses.
When you set Activate the KeyMonitor, the Passthru property is automatically reset to "True"

- LogType
You can log Chars and Keys. The difference lies only in alpha-numeric characters (0-9, A-Z).
If you log keys (LogType=kmLogKeys), then, whatever Shift/Caps state is, you will always be reported either an Uppercase letter or a number, when user presses one of these keys.
But when you log Chars (LogType=kmLogChars), then ,say, if you press "A" without Shift and Caps, a lowercase "a" will be reported, or if you press {Shift}+1, not "1", but an "!" will be reported.

-MacroLogType
For Example - you wish some action to be executed after the user presses "C", then "V", then "U". You specify a macro "CVU", and wait for the event, but here may be two ways for the macro to be executed. You may either wait for a KeyDown("C")+KeyDown("V")+KeyDown("U") or for KeyUp("C")+KeyUp("V")+KeyUp("U").
This property will allow you to choose. If you wish a macro to be executed after a number of KeyDowns, than set MacroLogType=kmConsiderKeyDowns, else, set MacroLogType=kmConsiderKeyUps

-Macro
This is a runtime only property. You may specify a number of macros (keystroke-sequences) and when each of them is executed, a Macro event will be fired.
A macro is a string, that consists of a Keystroke identifiers, followed one after another. Remember that a Space is also a key identifier. So, if you specify a macro like: "C C", it will be executed if the user presses "C", then Space, then "C"

Every macro has its name. Like this:

You define a macro:
KeyMon1.Macro("MyMacro1")="{Shift}{Ctrl}A"
Now when exactly this sequence of keystrokes will be tracked, a Macro event will be fired. It will pass you the string ("MyMacro1") as the parameter

You remove a macro like this
KeyMon1.Macro("MyMacro1")=""

NOTE: You specify an exact sequence, eg if you specify a macro "{Shift}{Ctrl}A", then the macro won't be reported if the user presses "{Ctrl}{Shift}A"
The MACRO STRING IS CASE-SENSITIVE!!!
So, if you mistype like this "{SHift}", you macro will never be tracked, because there isn't an identifyer "{SHift}", only "{Shift}"
A list of identifiers below

A Macro cannot be longer than 500 characters



Methods:

- Activate, Deactivate
Same as setting the Active property to True or False respectively
They may return True or False, indicating, whether there were some problems performing the operation (Return value-False) or whether everything went fine (Return value-True).


Events:
- KeyDown, KeyUp, KeyPress
Just normal Keyboard events, the parameters are the same as of any normal VB control, only these events happen no matter which app has the focus.

- KeyUpEx, KeyDownEx
Same as above (also are fired on a keyboard action), but they report a string, representing a key's identifier instead of VB's standard KeyAscii.
A list of identifiers below.







------Key identifiers------
in a KeyDownEx/KeyUpEx event a key pressed is reported to you as a Key identifier.
Also when you specify a macro, you use key identifiers.

The identifiers of characters, numbers, symbols and space are usual.
(the identifier for Space is " ", for a dot is ".", etc...)
The two exceptions are the identifiers of "{" and "}" : they are
 
{LCurlyBracket}
{RCurlyBracket} 

Only remember, that the key identifier returned also depends on the control's LogType property


The identifiers of control keys are the following:
 (some keys are not on every keyboard, so you may not understand their names (at least some of them seemed weird to me))

{Cancel}
{Backspace}
{Tab}
{Clear}
{Enter}
{Esc}
{PgUp}
{PgDn}
{End}
{Home}
{Left}
{Up}
{Right}
{Down}
{Ins}
{Del}
{Help}
{Select}
{Exec}
{Apps}
{Num}
{Scroll}
{Attn}
{CrSel}
{ExSel}
{ErEOF}    
{Play}
{Zoom}
{PA1}
{Pause}
{Print}
{ProcessKey}
{OEMClear}
{Caps}
{Sep}
{Ctrl}
{Shift}
{Alt}


--Anyway, just experiment with the control and you'll understand
only be careful with it, because I can't tell you how may times VB made a GPF while I was making it. VB is DEFINITELY not the language for such stuff.

This stuff is POSTCARDWARE.
That means, that if you find it useful, please, do me a favor: send me a postcard with a view of the place you live, or kinda, signed with a pair of words by you. I'll appreciate that much.

My address:
-----------------
Konstantin Tretyakov
Sopruse 167-86
13413 Tallinn
Estonia
-----------------

Thanx, Yours sincerely, Konstantin Tretyakov

Visit: http://smartsite.cjb.net