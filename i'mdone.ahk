; Setup info

coordmode mouse, relative
DetectHiddenText, Off
SetTitleMatchMode, 2
#Persistent
#NoEnv

RWin::return
LWin::return
#L::return

; Include Typo Correction (Everyone should use this!)

#Include typos.ahk

; Include Jonathan Ploudre's Medical Abbreviations
; Saves about 2 hours of typing a month if you learn them.
; http://unchart.com/2011/how-much-time-can-macros-save/

#Include medicalabbrevs.ahk



; The I'm Done key, closes out with a default action in common windows in e-MDs.
!Space::
#Space::
ifWinActive, Patient Chart
Send !fx
ifwinactive, E&M Coder
ControlClick, TBitBtn5
ifWinActive, Free Text
Send #s
ifWinActive, Memo
gosub, MemoOK
ifWinActive, e-MDs Schedule
Send !x
ifWinActive, e-MDs Chart
Send !fc
ifWinActive, e-MDs DocMan
Send !fx
ifWinActive, Taskman
Send !fs
ifWinActive, New Message
Send !fs
ifWinActive, Reply Message
Send !fs
ifWinActive, Forward Message
Send !fs
ifWinActive, Past Medical History Maintenance
Send !{F4}
ifWinActive, New Reminder
Send !s
ifWinActive, Patient Alert
Send !s
ifWinActive, Image
Send !fi
ifWinActive, Patient Appointments
Send !{F4}
ifWinActive, Growth Charts
Send !{F4}
ifWinActive, Chart Patient Maintenance
Send !s
ifWinActive, Bill Patient Maintenance
Send !s
ifWinActive, Choose Medications
Send !t
IfWinActive, Dictation Box
Send !t
ifWinActive, Edit rule result
click, 314, 449
ifwinActive, Choose Medications for
Send !t
ifWinActive, User Defined Dates
{
send !s
sleep 100
send !s
}
ifwinActive, Select Fax
Send !n
ifWinActive, DDx
click, 720, 650
ifWinActive, Dictation Box
Sendinput !t
IfWinExist, Free Text
{
WinActivate, Free Text
WinWaitActive, Free Text
Send #s
}
return

; Navigates to the Documents
#/::

if WinActive("Patient Chart") or WinActive("e-MDs Chart")
gochart("documents")
return

; Navigates to Lab Results
#.::

if WinActive("Patient Chart") or WinActive("e-MDs Chart")
gochart("labs")
return

; Navigates to the Visit Notes
#,::

if WinActive("Patient Chart") or WinActive("e-MDs Chart")
gochart("visitnotes")
return

; Navigates to flowsheets
#'::

if WinActive("Patient Chart") or WinActive("e-MDs Chart")
Sendinput !l
return

; Navigates to Hs/Visit
#;::
if WinActive("Patient Chart") or WinActive("e-MDs Chart")
Sendinput !v
return

; Used for Reviewing documents from taskman
#Right::

emdschartclick("467", "193")
return

#Left::

emdschartclick("434", "193")
return

; Meaningful Tabs. Tabs through the health summary for quick navigation.
^Tab::
;gosub, collapseHS
chartreviewposition = 1
while (chartreviewposition < 9)
	{
	gosub, collapseHS
	;Ternary var := x>y ? 2 : 3
	ypos := (chartreviewposition = 1) ? 160 : (chartreviewposition = 2) ? 180 : (chartreviewposition = 3) ? 200 : (chartreviewposition = 4) ? 220 : (chartreviewposition = 5) ? 237 : (chartreviewposition = 6) ? 255 : (chartreviewposition = 7) ? 275 : (chartreviewposition = 8) ? 295 : 0
	emdschartclick("77", ypos)
	SoundPlay, %dropboxloc%battlewinfast.wav
	Sleep, 400
	KeyWait, Tab, D
	; Shift tablet goes backwards
	GetKeyState, state, Shift
		{
		if state = D
		chartreviewposition --
		else
		chartreviewposition ++
		}
	}
	; We're done, now open the Medicine Field (My Fav!)
	gosub, collapseHS
	emdschartclick("77", "200")
	SoundPlay, %dropboxloc%battlelossfast.wav
	gosub, emdsrefresh
Return

; Addendum to Log Note.
#a::
ifWinActive, Patient Chart
WinGetText wintext
IfInString, wintext, Log/&Phone/Rx Notes
{
click, 491, 151
}
return

; Delete for Taskman
#d::
ifWinActive, e-MDs TaskMan
Send !ed
return

; Forward in Taskman or Chart
#f::
emdschartclick("977","150")
ifWinActive, e-MDs TaskMan
send !af
return

; 'Kills' or Removes an item from order tracking
#+k::
ifWinActive, Patient Chart
{
WinGetText wintext
IfInString, wintext, La&bs
{
click, 51, 182
Sleep, 100
click, 70, 259
}
}
return

; Clicks the Goldan Man Icon for Account info
#i::
emdschartclick("482","92")
return

; Clicks the Appt Information
#+i::
emdschartclick("744","92")
return

; New Memo or Taskman
#n::
emdschartclick("915","150")
ifWinActive, e-MDs TaskMan
Send !e{left}{enter}
return

; Opens Attactment in Taskman, Opens Unsigned Documents.
#o::
ifWinActive, e-MDs TaskMan
{
	CoordMode, mouse, relative
ControlGetPos, xpos, ypos,,,TPanel9
xpos := xpos + 20
ypos := ypos + 22
click, %xpos%, %ypos%
ypos := ypos + 26
sleep 100
click, %xpos%, %ypos%
}
ifWinActive, e-MDs Chart
Send !fo
ifWinActive, Multi Sign Off
click, 727, 168
IfWinactive, Order Tracking
{
Coordmode, pixel, relative
coordmode, mouse, relative
PixelSearch, ClickX, ClickY, 207, 109, 207, A_ScreenHeight, 0x6b2408
click, ClickX, ClickY, 2
Send {Down}
}
return

; Planned Care Lab. Annotes a lab to Discuss, signs, removes from order tracking
^+p::

IfWinActive, Patient Chart
{
click, 915, 150
WinWaitActive, Memo
Sleep, 100
SendInput, Will discuss at already scheduled visit.
gosub, MemoOK
WinWaitActive, Patient Chart
{
sleep, 100
click, 854, 150
sleep, 100
click, 51, 182
sleep 100
click, 70, 259
sleep, 100
Send !fx
}
}
return

; Reply or Reminder
#r::
ifWinActive, e-MDs TaskMan
Send !ar
ifwinactive, Patient Chart
; Patient Reminder
{
Send !m
sleep 200
click, 39, 125
}
return

; Recall
^r::
ifwinactive, Patient Chart
{
click, 487, 89
sleep 200
Send !M
sleep 100
Send !A
}
return

; Refresh
!r::
gosub, emdsrefresh
return

; Signs Off a Document
#+s::
if WinActive("Patient Chart") or WinActive("e-MDs Chart")
{
emdschartclick("854","150")
SoundPlay, %dropboxloc%clickit.aiff
}
return

; Maps Win-S to work for Alt-S
#s::
Send !s
return


; Update Med Hx for Taskman.
#u::
if WinActive("Patient Chart") or WinActive("e-MDs Chart")
{
emdschartclick("914", "150")
Send Patient history updated.
gosub, MemoOK
sleep 500
emdschartclick("854", "150")
gochart("visiths")
emdschartclick("89", "143")
sleep 100
emdschartclick("83", "219")
sleep 100
emdschartclick("83", "244")
}
return

; The Dragon Macro. Assumes Numpad+ Turns on Microphone.
F12::
NumpadSub::
IfWinActive, Free Text
{
;growl("19-gear.png","HPI Items","Talk to Bill!\n\n* Where\n* Started when\n* Course\n* Associated with\n* Worse with\n* Better with\n* Comorbidities")
WinGetPos, xpos, ypos, width, height, A
run, Notepad
WinWaitActive, Untitled - Notepad
WinMove, A,,%xpos%,%ypos%,%width%,%height%
Send {NumpadAdd}
return
}
IfWinActive, Untitled - Notepad
{
Send {NumpadAdd}
sleep 100
Send {Control down}
Send a
Send x
Send {Control up}
Sleep 100
Send {Alt down}
Send FX
Send {Alt up}
WinWaitNotActive
IfWinExist, Free Text
{
WinActivate, Free Text
WinWaitActive, Free Text
Send {Control down}
Send v
Send {Control up}
send !s
}
return
}
return


; ########################################### Esc/Enter Keys

#ifWinActive, Free Text
Esc::
Send !c
Send !y
return
#a::
Send !a
return

#ifWinActive, E&M Coder
Esc::
click, 488, 436
return
Enter::
click, 562, 436
return

#ifWinActive, Note Conclusion
Esc::
click, 575, 475
return
v::
click, 16, 271
return
Enter::
Click, 511, 481
return

#ifWinActive, Memo
Esc::
click, 518, 428
return

#ifWinActive, Patient Appointments
Esc::
Send !{F4}
return

#ifWinActive, Patient Alert
Esc::
Send !c
return

#ifWinActive, Growth Charts
Esc::
Send !{F4}
return

#ifWinActive, ICD-9 Search
Esc::
Send !{F4}
return


#ifWinActive, Reply Message
Esc::
Send !{F4}
Send !n
return

#ifWinActive, Sending Prescriptions
Enter::
click, 33, 44
return
Esc::
click, 132, 42
return

#ifWinActive, Order Tracking
Enter::
WinGetPos,,,,trackingheight, Order Tracking
PixelSearch, ClickX, ClickY, 182, 121, 182, trackingheight, 0x6b2408, 15
ClickY := ClickY +3
click, %ClickX%, %ClickY%, 2
return

#ifWinActive, Multi Sign Off
Enter::
Coordmode, pixel, relative
coordmode, mouse, relative
PixelSearch, , ClickY, 198, 162, 198, A_ScreenHeight, 0x800000
click, 728, ClickY
return

#ifWinActive, Choose Medications for
Esc::
Send !{F4}
return

#ifWinActive, Printing Explanation
Enter::
SendInput {down 6}
SendInput !o
return

#ifWinActive, e-MDs Chart Option
Esc::
Send !fx
return

#ifWinActive, Template Quick Picks
Esc::
click, 458, 11
return


#IfWinActive, Pick a sick visit
Enter::
Send !s
Return
Esc::
Send !c
Return

#ifWinActive, Splash Screen
~Enter::
sleep, 2000
IfWinActive, Error
{
SoundPlay, %dropboxloc%accessdenied.wav
return
}
SoundPlay, %dropboxloc%login.wav
Return
LButton::
MouseGetPos,xpos,ypos,thecontrol
if (thecontrol = "TButton2")
{
SoundPlay, %dropboxloc%login.wav
msgbox, it worked
}
click,%xpos%, %ypos%
return

; Done with Window Specific Hotkeys
#ifWinActive

; ############################################################ Functions

gochart(location)
{
	ifinstring, location, documents
	{
		send !c
		confirmnav(pcsView)
		Send !u
		confirmnav(Doc&uments)
		; expand all documents. 
		if WinActive("Patient Chart") or WinActive("e-MDs Chart")
		{
			mouseclick, right, 85, 169
			sleep 100
			click, 120, 184
		}
	}
	ifinstring, location, labs
	{
		send !c
		confirmnav(pcsView)
		Send !b
		confirmnav(La&bs)
	}
	ifinstring, location, visitnotes
	{
		send !c
		confirmnav(pcsView)
		Send !n
		confirmnav(&Notes)
	}
	ifinstring, location, visiths
	{
		send !v
		confirmnav(pcsProgressNote)
	}

}
return

confirmnav(navtext)
{
WinGetText, wintext
while (ifNotinString, wintext, %navtext%)
	{
	sleep, 100
	while (ifNotinString, %navtext%, wintext)
	WinGetText, wintext
	}
}
return

MemoOK:
{
	Click, 448, 425
}
Return

collapseHS:
{
emdschartclick("95","149")
return
}

; Corrects for the extra vertical space difference between Chart and e-MDs Chart.
emdschartclick(xpos,ypos)
{
ifWinActive, Patient Chart
click, %xpos%, %ypos%
IfWinActive, e-MDs Chart
{
ypos := ypos + 30
click, %xpos%, %ypos%
}
}

emdsrefresh:
{
if WinActive("Patient Chart") or WinActive("e-MDs Chart")
emdschartclick("815","152")
if WinActive("Order Tracking")
click, 871, 80
if WinActive("e-MDs TaskMan")
click, 400, 56
if WinActive("Multi Sign Off")
ControlClick, TBitBtn1
if WinActive("e-MDs Tracking Board")
click, 68, 55
}
Return
