; Note: This script requires BOM encoding (UTF-8) to display characters properly. 

;version 0.1.1, The 'Sending Case Note Update' version
;version 0.1.2, The 'Dude, Where's My GUI' version
;version 0.1.3, the 'Dakota County' retro version
;version 0.1.4, The 'I think I got all the auto-populating date variables correct THIS time' version
;version 0.1.5, The 'Extra Large serving of date variables because some counties are so far behind' version
;version 0.1.6, The ‘It seems to be working so I fixed some of the Maxis case note text and also changed things behind the scenes that nobody else will notice’ version
;version 0.1.7, The 'Topsy Turvy' version
;version 0.1.8, The 'It still seems to be working so I made more variables into Objects' version
;version 0.1.9, The 'There weren't enough options for missing verifications, so I added more' version
;version 0.2.0, The 'Multi-County' version
;version 0.2.1, The 'I pulled my hair out fixing the Missing Verifications code' version
;version 0.2.2, The 'I really enjoy using Objects, so I did it again' version
;version 0.2.3, The 'DHS finally updated Special Letter, and now some of the options are useful, part 1 of several' version
;version 0.2.4, The 'What happens if this loads off screen and there's no way to get it to move anywhere else' version
;version 0.2.5, The 'I think I finally got the split verification lists to work correctly oh no I think I jinxed it' version
;version 0.2.6, The 'Letter Text now works with mec2functions. Also added Help text.' version
;version 0.2.7, The 'Almost ready for Prime Time so now the name is CaseNotes, not Case Note - App or recert' version
;version 0.2.8, The 'Do or do not, there is no try' version
;version 0.2.9, The 'Imported hotkeys from MEC2Hotkeys' version
;version 0.3.0, The 'Look ma, I'm on GitHub' version
;version 0.3.1, The 'I am so tired of typing questions from redeterminations into Other because someone missed Lump Sum again' version (soonTM)
;version 0.3.2, The 'I don't know what your MAXIS screen is called' version
;version 0.3.3, The 'I had to rearrage code so that Homeless prompting would be added to the Special Letter' version
;version 0.3.4, The 'I prettied up Missing Verifications, changed MV to open/hide on load, and gave each GUI a name' version
;version 0.3.5, The 'If the Special Letter line count ain't right now, it ain't ever gonna be' version
;version 0.3.6, The 'Added some parens and fixed some copy/paste errors in the MissingGuiGUIClose subroutine' version
;version 0.3.7, The 'Redid how coordinates were done. Next is EmployeeInfo/CaseNoteCountyInfo INI' version
;version 0.4.0, The 'I rewrote how settings were done. Fewer read/writes to harddrive. Hurray! Also added Waitlist functionality.' version
;version 0.4.2, The 'Holy crap I finally figured out how to fix the Gui Submit issue.' version
Version := "v0.4.3"

;Future todo ideas:
;Add backup to ini for Case Notes window. Check every minute old info vs new info and write changes to .ini.
;Make a restore button.
;Import from clipboard (when copied from MEC2) (likely mostly same code as restore button)

#Requires AutoHotkey v1+
SetWorkingDir %A_ScriptDir%
#Persistent
#SingleInstance force
#NoTrayIcon
SetTitleMatchMode, RegEx
Global Ini := { CBTPositions: { XClipboard: 0, YClipboard: 0 }
        , CaseNotePositions: { XCaseNotes: 0, YCaseNotes: 0, XVerification: 0, YVerification: 0 }
        , CaseNoteCountyInfo: { CountyNoteInMaxis: 0, CountyFax: A_Space, CountyDocsEmail: A_Space, CountyProviderWorkerPhone: A_Space, CountyEdBSF: A_Space, Waitlist: 1 }
        , EmployeeInfo: { EmployeeName: A_Space, EmployeeCounty: A_Space, EmployeeEmail: A_Space, EmployeePhone: A_Space, EmployeeUseEmail: 0, EmployeeUseMec2Functions: 0, EmployeeBrowser: A_Space, EmployeeMaxis: MAXIS-WINDOW-TITLE } }

SetCoords(CoordObj) {
    For key, value in CoordObj {
        if (Abs(value) > 9000)
            CoordObj[key] := 50
    }
}
SetFromIni() {
    FileRead, StoredIni, %A_MyDocuments%\AHK.ini
    StoredIni := StrReplace(StoredIni, "INI=", "=",, -1)
    Section := ""
    Loop, Parse, StoredIni, `n, `r
    {
        If InStr(A_LoopField, "]",, StrLen(A_LoopField)-1) {
            RegExMatch(A_LoopField, "iO)\[([a-z0-9]+)\]", Section)
        } Else {
            KeyValue := StrSplit(A_LoopField, "=")
            If (KeyValue[2] == "")
                Continue
            Ini[Section[1]][KeyValue[1]] := KeyValue[2]
        }
    }
    SetCoords(Ini.CBTPositions)
    SetCoords(Ini.CaseNotePositions)
}
SetFromIni()

GroupAdd, AutoMail, Automated Mailing Home Page
GroupAdd, AutoMail, ahk_exe obunity.exe,,, Perform Import
CheckGroupAdd()
Global MissingVerifications := {}, ClarifiedVerifications := {}, EmailText := {}, MissingInput := {}, LineCount := 0, Homeless := 0

Global CaseDetails := { DocType: "_DOC?", Eligibility: "_ELIG?", SaEntered: "_SA?", CaseType: "_PRG?", AppType: "_APP?", isHomeless: "", HaveWaitlist: false }
Global CaseNoteEntered := { MEC2Note: 0, MAXISNote: 0 }
Global MaxisNote :=, IdList := ""
Global ConfirmedClear := 0, VerificationWindowOpenedOnce := 0, VerifCat :=, LetterTextNumber := 1, LetterText := {}, MissingHomelessItems := "", SignDate := 0

Global CountySpecificText := { StLouis: { OverIncomeContactInfo: "", CountyName: "St. Louis County" }
, Dakota: { OverIncomeContactInfo: " contact 651-554-6696 and", CountyName: "Dakota County", CustomHotkeys: "
(
    Custom hotkeys for your county exist for the following windows (A ToolTip reminder appears by pressing F1):
    ● OnBase (Alt+4: 'Verifs Due Back' detail),
    ● OnBase (Ctrl+F6-12: Enters keywords on the Perform Import screen)
    ● OnBase (Ctrl+B: Inserts date and case number for mail)
    ● Automated Mailing (Ctrl+B: Inserts date and case number for mail)
    ● Browser (Types an Approved (Alt+F1) or Denied (Alt+F2) app case note. Ctrl+F12 or Alt+F12: worker signature)
    ● Word (Alt+4: Types in first name, case number, and app received date).
    ● MAXIS (Alt+M: Changes the title of the MAXIS window to ""MAXIS"" to enable screen-scraping.)
)" } }

; Date variables
Global DateObject := { ReceivedMDY: "", ReceivedYMD: "", AutodenyYMD: "", ReinstateDate: "" }
Global AutoDenyObject := { AutoDenyExtensionMECnote: "", AutoDenyExtensionDate: "", AutoDenyExtensionSpecLetter: "", AutoDenyMaxisNote: "" }
DateObject.TodayYMD := A_Now
FormatTime, ShortDate, %A_Now%, M/d/yy ; for sending to envelope
Global Received := DateObject.TodayYMD, OverIncomeObj := { overIncomeHHsize: "your size" }
DateObject.TodayMDY := FormatMDY(DateObject.TodayYMD)

If InStr(DateObject.TodayYMD, 0401)
    Menu, Tray, Icon, compstui.dll, 100
Else
    Menu, Tray, Icon, azroleui.dll, 7


;MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION 
Gui MainGui: Font,, Segoe UI
Gui MainGui: Color, a9a9a9, bebebe

Gui, MainGui: Add, Radio, Group Section h17 x12 w75 y+5 gApplicationRadio, Application
Gui, MainGui: Add, Radio, xp y+2 wp h17 gRedeterminationRadio, Redeterm.
;Gui, MainGui: Add, Checkbox, xp y+2 wp h17 vHomeless gHomeless, Homeless
Gui, MainGui: Add, Checkbox, xp y+2 wp h17 vHomeless, Homeless

Gui, MainGui: Add, Radio, Group x+10 ys h17 w78 gMNBenefits vMNBenefitsRadio, MNBenefits
Gui, MainGui: Add, Radio, xp y+2 h17 wp gApp vAppRadio, 3550 App

Gui, MainGui: Add, Radio, Group x+10 ys h17 w58 gBSF vBSF, BSF
Gui, MainGui: Add, Radio, xp y+2 h17 wp gTY vTY, TY
Gui, MainGui: Add, Radio, xp y+2 h17 wp gCCMF vCCMF, CCMF

Gui, MainGui: Add, Radio, Group x+0 ys h17 w80 gPending vPendingRadio, Pending
Gui, MainGui: Add, Radio, xp y+2 h17 wp gEligible vEligibleRadio, Eligible
Gui, MainGui: Add, Radio, xp y+2 h17 wp gIneligible vIneligibleRadio, Ineligible

Gui, MainGui: Add, Radio, Group Hidden x+5 ys h17 vSaApproved gSaApproved, SA Approved
Gui, MainGui: Add, Checkbox, Group Hidden xp yp h17 vManualWaitlistBox, Waitlist
Gui, MainGui: Add, Radio, Hidden xp y+2 h17 vNoSA gNoSA, No SA
Gui, MainGui: Add, Radio, Hidden xp y+2 h17 vNoProvider gNoProvider, No Provider

Gui, MainGui: Add, Text, xp-18 y+9 w200 vAutoDenyStatus,

Gui, MainGui: Add, Text, x420 w35 h20 ys+2,Case &`#
Gui, MainGui: Add, Text, xp y+2 w35 h20, Rec'd:
Gui, MainGui: Add, Text, xp y+2 w35 h20 vSignText Hidden, Signed:

Gui, MainGui: Add, Edit, x+0 ys w75 h17 -Background Limit8 vCaseNumber,
Gui, MainGui: Add, DateTime, xp y+5 w75 h17 vReceived, M/d/yy
Gui, MainGui: Add, DateTime, xp y+5 w75 h17 vSignDate Hidden, M/d/yy

Gui, MainGui: Add, Button, Section x540 ys+0 h17 w65 -TabStop vMEC2NoteButton gMEC2NoteButton, MEC2 Note
Gui, MainGui: Add, Button, xs y+5 h17 w65 -TabStop Hidden vMaxisNoteButton gMaxisNoteButton, Maxis Note
Gui, MainGui: Add, Button, xs y+5 h17 w65 -TabStop vNotepadBackup gNotepadBackup, To Desktop

Gui, MainGui: Add, Button, Section x615 ys+0 h17 w50 -TabStop vClearFormButton gClearFormButton, Clear

SetTextAndResize(newText, fontOptions := "", fontName := "") {
    Gui 9:Font, % fontOptions, % fontName
    Gui 9:Add, Text, Limit200, % newText
    GuiControlGet T, 9:Pos, Static1
    Gui 9:Destroy
    Return TW
}
; 77 characters returns 12h 565w at 3840 x 2160, 150%; returns 12h 539w at 1920 x 1080, 100% ;123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 1234567
OneHundredChars := SetTextAndResize("1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890", "s9", "Lucida Console")
MonoChar := OneHundredChars/100
; 100 characters is 733w at 3840 x 2160
LabelSettings := "xm+5 y+1 w200 "
LabelExampleSettings := "x220 yp+4 h12 w" MonoChar*60 " "
; 995 is about the right size (* .666 = 663)
TextboxSettings := "xm y+1 w" (MonoChar*87)+27 " " ; At w650, WinSpy shows 945 for box with scrollbar, 971 without, and 975 total. (spy * .666 = AHK #s). Which gets 20 for the (~17.4) scrollbar and (~2.6) border
OneRow := "h17 Limit87 "
TwoRows := "h33 "
ThreeRows := "h43 "
FourRows := "h55 "

Gui MainGui: Font, s9, Segoe UI
Gui, MainGui: Margin, 12 12
Gui, MainGui: Add, Text, xm y+45 h0 w0 ; Blank space
Gui, MainGui: Add, Text, % LabelSettings "vHouseholdCompLabel", Household Comp
Gui, MainGui: Add, Text, % LabelExampleSettings "vHouseholdCompLabelExample Hidden", Parent (ID), ChildOne (4, BC), ChildName (age, verif)
Gui, MainGui: Add, Edit, % TextboxSettings TwoRows "vHouseholdComp", 

Gui, MainGui: Add, Text, %LabelSettings% vAddressVerificationLabel, Address Verification
Gui, MainGui: Add, Text, %LabelExampleSettings% vAddressVerificationLabelExample Hidden, 1234 W Minnesota St APT 21, St Paul: ID 5/4/20 (scan date)
Gui, MainGui: Add, Edit, %TextboxSettings% %ThreeRows% vAddressVerification,

Gui, MainGui: Add, Text, %LabelSettings% vSharedCustodyLabel, Shared Custody
Gui, MainGui: Add, Text, %LabelExampleSettings% vSharedCustodyLabelExample Hidden, Absent Parent / Child: Thursday 6pm - Monday 7am
Gui, MainGui: Add, Edit, %TextboxSettings% %FourRows% vSharedCustody, 

Gui, MainGui: Add, Text, %LabelSettings% vSchoolInformationLabel, School information
Gui, MainGui: Add, Text, %LabelExampleSettings% vSchoolInformationLabelExample Hidden, ChildOne, ChildTwo: Wildcat Elementary, M-F 730am - 2pm
Gui, MainGui: Add, Edit, %TextboxSettings% %ThreeRows% vSchoolInformation, 

Gui, MainGui: Add, Text, %LabelSettings% vIncomeLabel, Income 
Gui, MainGui: Add, Text, %LabelExampleSettings% vIncomeLabelExample Hidden Border, Parent - Job: BW avg $1234.56, 43.2hr/wk; annual @ 32098.56
Gui, MainGui: Add, Edit, %TextboxSettings% %FourRows% vIncome, 

Gui, MainGui: Add, Text, %LabelSettings% vChildSupportIncomeLabel, Child Support Income
Gui, MainGui: Add, Text, %LabelExampleSettings% vChildSupportIncomeLabelExample Hidden, 6 month total $2345.67; annual @ 4691.34
Gui, MainGui: Add, Edit, %TextboxSettings% %TwoRows% vChildSupportIncome, 

Gui, MainGui: Add, Text, %LabelSettings% vChildSupportCooperationLabel Border gChildSupportCooperation, Child Support Cooperation
Gui, MainGui: Add, Text, %LabelExampleSettings% vChildSupportCooperationLabelExample Hidden, Absent Parent / Child: Open, cooperating
Gui, MainGui: Add, Edit, %TextboxSettings% %FourRows% vChildSupportCooperation, 

Gui, MainGui: Add, Text, %LabelSettings% vExpensesLabel, Expenses
Gui, MainGui: Add, Text, %LabelExampleSettings% vExpensesLabelExample Hidden, BW Medical $121.23, BW Dental $12.23, BW Vision $2.23
Gui, MainGui: Add, Edit, %TextboxSettings% %TwoRows% vExpenses, 

Gui, MainGui: Add, Text, %LabelSettings% vAssetsLabel, Assets
Gui, MainGui: Add, Text, %LabelExampleSettings% vAssetsLabelExample Hidden, < $1m   or   (blank)
Gui, MainGui: Add, Edit, %TextboxSettings% %OneRow% Limit87 vAssets, 

Gui, MainGui: Add, Text, %LabelSettings% vProviderLabel, Provider
Gui, MainGui: Add, Text, %LabelExampleSettings% vProviderLabelExample Hidden, Kid Kare (PID#, HQ): ChildOne, ChildTwo - Start date 5/4/20
Gui, MainGui: Add, Edit, %TextboxSettings% %TwoRows% vProvider, 

Gui, MainGui: Add, Text, %LabelSettings% vActivityandScheduleLabel, Activity and Schedule
Gui, MainGui: Add, Text, %LabelExampleSettings% vActivityandScheduleLabelExample Hidden, ParentOne - Employment: M-F 9a - 5p (8h x 5d)
Gui, MainGui: Add, Edit, %TextboxSettings% %FourRows% vActivityandSchedule, 

Gui, MainGui: Add, Text, %LabelSettings% vServiceAuthorizationLabel, Service Authorization
Gui, MainGui: Add, Text, %LabelExampleSettings% vServiceAuthorizationLabelExample Hidden, 8h work + 1h travel = 9h/day, 90h/period
Gui, MainGui: Add, Edit, %TextboxSettings% %ThreeRows% vServiceAuthorization, 

Gui, MainGui: Add, Text, %LabelSettings% vNotesLabel, Notes
Gui, MainGui: Add, Edit, %TextboxSettings% %FourRows% vNotes, 

Gui, MainGui: Add, Text, xm+5 y+1 gMissingButton vMissingButtonLabel Border, Missing
Gui, MainGui: Add, Text, %LabelExampleSettings% vMissingButtonLabelExample Hidden, (Click "Missing" to bring up the missing verification list)
Gui, MainGui: Add, Edit, %TextboxSettings% h115 vMissing,

Gui, MainGui: Add, Text, x15 y+4, %Version%
Gui, MainGui: Add, Button, x+20 yp w65 h19 -TabStop gSettingsButton, Settings
Gui, MainGui: Add, Button, x+40 yp wp h19 -TabStop gExamplesButton vExamplesButton, Examples
Gui, MainGui: Add, Button, x+40 yp wp h19 -TabStop gHelpButton, Help
Gui, MainGui: Add, Button, x600 yp wp h19 gMissingButton, Missing

Gui, MainGui: Show, % "x" Ini.CaseNotePositions.xCaseNotes " y" Ini.CaseNotePositions.yCaseNotes, CaseNotes
Gui, MainGui: Show, AutoSize
GuiControl, Focus, HouseholdComp

EditControls := ["HouseholdComp", "SharedCustody", "AddressVerification", "SchoolInformation", "Income", "ChildSupportIncome", "ChildSupportCooperation", "Expenses", "Assets", "Provider", "ActivityandSchedule", "ServiceAuthorization", "Notes", "Missing"]
ExampleLabels := [ "HouseholdCompLabelExample", "AddressVerificationLabelExample", "SharedCustodyLabelExample", "SchoolInformationLabelExample", "IncomeLabelExample", "ChildSupportIncomeLabelExample", "ChildSupportCooperationLabelExample", "ExpensesLabelExample", "AssetsLabelExample", "ProviderLabelExample", "ActivityandScheduleLabelExample", "ServiceAuthorizationLabelExample", "MissingButtonLabelExample" ]
For Index, EditField in EditControls {
    Gui MainGui: Font, s9, Lucida Console ; monospace font
    GuiControl, MainGui: Font, % EditField
}
For Index, Label in ExampleLabels {
    Gui MainGui: Font, s9, Lucida Console
    GuiControl, MainGui: Font, % Label
}
GoSub MissingGui

If (StrLen(Ini.EmployeeInfo.EmployeeName) < 1)
    Gosub SettingsButton
Return

HelpButton:
    Paragraph1 := "
    (
    This app is a template tool for generating Case Notes for CCAP applications and redeterminations 
    and request letters for Special Letters and emails. This tool is not endorsed or sponsored by MN DHS.
    )"
    Paragraph2 := "
    (
    Features:
    ● Auto-formats text to fit within MEC2's Case Notes and Special Letter/Memo fields.
    ● Case Notes are formatted with categories and spacing for consistant alignment. 
    ● User-entered dates calculate the extended auto-deny date (for a minimum of 15ish days from the date processed).
    ● Incorporates document type, approval status, and dates in the notes and verification requests.
    ● Line width is designed to be compatible with the Income Calculator spreadsheet.
    ● Compatible with mec2functions from github.com/MECH2-at-github.
    ● Case Notes can be saved to a text document in the event a case is locked or otherwise inaccessible.
    ● Special Letter requests are broken down into 'clarifications' of checkbox items, and additional items.
    )"
    Paragraph3 := "
    (
    Main window :
    ● [MEC2 Note] - Formats the entire case note and sends the data to MEC2.
        If you are not using mec2functions, it will simulate keypresses to navigate the page. 
        In MEC2: Click 'New' in MEC2 on the CaseNotes webpage. In CaseNotes, click [MEC2 Note].
    ● [Maxis Note] - Visible only if ""Case Note in MAXIS"" is checked in Settings.
        Formats the app date, case status, and verifications list and sends it to Maxis.
        It will activate BlueZone and paste the case note in.
        In BlueZone (Maxis): PF9 to start a new note. In CaseNotes, click [Maxis Note].
    ● [Desktop Backup] - Saves case notes for MEC2, Maxis, the Special Letters, and Email to your desktop.
        In CaseNotes, click [To Desktop]. A text file will be saved using the case number for the file name.
    ● [Clear] - Resets the app. If the case note has not been sent to MEC2/Maxis or saved to file, it will give a
        warning. Otherwise, it will change to [Confirm]. Clicking again will reset the app.
    ● Child Support Cooperation - Click the [Child Support Cooperation] label to copy from Custody.
        The Child Support Cooperation text field must be blank when clicking the button.
    ● Missing Verifications - Click either the [Missing] label or [Missing] button to open the verification window.
    )"
    Paragraph4 := "
    (
      Missing Verifications window:
    ● Checkboxes which cover almost all documents needed, with 3 ""Other"" selections for free-form requests.
    ● Verification types are clustered by categories.
    ● If a checkbox label has (Input) it will open a popup requesting clarifying information.
    ● ""Over-income"" will modify the beginning of the Special Letter text, adding income limit information.
    
    ● [Done] - After selecting the verifications that are missing, [Done] will generate a list and letter/email buttons.
          If the Special Letter text exceeds 30 lines, a second (/third/fourth) letter will be generated.
    ● [Letter 1/2/3/4] - Clicking these will place text on the clipboard to be pasted into the Worker Comments field.
        If ""Use mec2functions"" is checked, [Letter 1] will auto-check and auto-fill fields in MEC2's Special Letter page.
    ● [Email] - Text is send to the clipboard, which can then be pasted into an email. It will include the document type
        that was selected (application/redetermination). For homeless cases, alternate text is generated based on if
        the case is Eligible or Pending.
    )"
    Paragraph5 := "
    (
    Hotkeys:
    ● (Alt+3): Copies the case number to the clipboard. Usable from anywhere when CaseNotes is open.
    ● (Win+m): Inserts Missing Verification text when your browser or an email is the active window.
      Similar to the [Email] or [Letter 1] buttons, but without needing to switch windows.
    ● (Win+left arrow) or (Win+right arrow): Resets CaseNotes location, in the event it is off-screen.
    ● (Ctrl+F12) or (Alt+12): In your browser, types in the worker's name with a separation line (case note signature)
    ● (Ctrl+Alt+a): Shows clipboard text in a popup window. Select and copy text first (such as from a case note).
    )"
    Paragraph6 := "
    (
    Special notes:
    ● CaseNotes will open in the same location it was closed, even if that monitor is no longer connected. See Hotkeys for reset instructions.
    ● All settings for CaseNotes are saved in the My Documents folder, under AHK.ini. Deleting this file will reset all
        saved settings.
    ● Please send any bug reports or feature requests to MECH2.at.github@gmail.com
    )"

    Gui, HelpGui: New, ToolWindow, CaseNotes Help
    Gui, HelpGui: Margin, 12 12
    Gui, Font, s10, Segoe UI
    Gui, HelpGui: Add, Tab3,, Features | CaseNotes | Missing Verifications | Hotkeys and Notes 
    Gui, Tab, 1
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph1
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph2
    Gui, Tab, 2
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph3
    Gui, Tab, 3
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph4
    Gui, Tab, 4
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph5
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph6
    Gui, HelpGui: Add, Text, xm y+15, % CountySpecificText[Ini.EmployeeInfo.EmployeeCounty].CustomHotkeys
    Gui, Tab
    Gui, HelpGui: Add, Button, gHelpGuiClose w70 h25 x375, Close
    Gui, HelpGui:+OwnerMainGui
    Gui, HelpGui: Show,, CaseNotes Help
Return

HelpGuiClose:
    Gui, HelpGui: Destroy
Return

ExamplesButton:
    Gui, MainGui: Submit, NoHide
    GuiControlGet, ExamplesButtonText,, ExamplesButton
    If (ExamplesButtonText = "Examples") {
        For Index, ExampleLabel in ExampleLabels {
            GuiControl, MainGui:Show, % ExampleLabel
        }
        GuiControl, MainGui:Text, ExamplesButton, Restore
    } Else If (ExamplesButtonText = "Restore") {
        For Index, ExampleLabel in ExampleLabels {
            GuiControl, MainGui:Hide, % ExampleLabel
        }
        GuiControl, MainGui:Text, ExamplesButton, Examples
    }
Return

MEC2NoteButton:
	GoSub, MakeCaseNote
return

Return

NotepadBackup:
    MaxisNote := ""
    GoSub, SetEmailText
    GoSub, MakeCaseNote
return

MaxisNoteButton:
    MaxisNote := ""
    GoSub, MakeCaseNote
return

JSONstring(inputString) {
    inputString := StrReplace(inputString, "`n", "\n",,-1)
    inputString := StrReplace(inputString, """", "`\""",,-1)
    return inputString
}

MakeCaseNote:
    Gui, MainGui: Submit, NoHide
    Gui, MissingGui: Submit, NoHide
    GoSub CalcDates
    EditControlsTest := ["HouseholdComp", "SharedCustody", "AddressVerification", "SchoolInformation", "Income", "ChildSupportIncome", "ChildSupportCooperation", "Expenses", "Provider", "ActivityandSchedule", "ServiceAuthorization", "Notes", "Missing"]
    For each, EditField in EditControlsTest {
        %EditField% := st_wordWrap(%EditField%, 87, "")
        %EditField% := StrReplace(%EditField%, "`n", "`n             ")
    }
	If (CaseDetails.Eligibility = "pends" && CaseDetails.DocType = "Redet") {
		CaseDetails.Eligibility := "incomplete"
        CaseDetails.SaEntered := ""
        FormattedSignDate := FormatMDY(SignDate)
        CaseDetails.RedetDue := CaseDetails.Eligibility = "incomplete" ? " (due " FormatMDY(SignDate) ")" : ""
	}
	If (CaseDetails.Eligibility = "pends" || CaseDetails.Eligibility = "ineligible") {
		CaseDetails.SaEntered := ""
	}
    If (OverIncomeMissing && CaseDetails.Eligibility = "ineligible") {
        CaseDetails.Eligibility := "over-income"
    }
    If (InStr(CaseDetails.CaseType, "BSF") && CaseDetails.Eligibility == "ineligible" && Ini.CaseNoteCountyInfo.Waitlist > 1 || WaitlistMissing == 1) {
        CaseDetails.Eligibility .= " - BSF Waitlist"
    }
	MEC2CaseNote := AutoDenyObject.AutoDenyExtensionMECnote " HH COMP:    " HouseholdComp "`n CUSTODY:    " SharedCustody "`n ADDRESS:    " AddressVerification "`n  SCHOOL:    " SchoolInformation  "`n  INCOME:    " Income "`n      CS:    " ChildSupportIncome  "`n CS COOP:    " ChildSupportCooperation  "`nEXPENSES:    " Expenses  "`n  ASSETS:    " Assets "`nPROVIDER:    " Provider "`nACTIVITY:    " ActivityandSchedule "`n      SA:    " ServiceAuthorization "`n   NOTES:    " Notes "`n MISSING:    " Missing "`n=====`n" Ini.EmployeeInfo.EmployeeName
	If (Homeless = 1) {
		CaseDetails.IsHomeless := "*HL "
	} else if (Homeless = 0) {
		CaseDetails.IsHomeless := ""
	}
	If (CaseDetails.DocType = "Application") {
		MEC2NoteTitle := CaseDetails.IsHomeless CaseDetails.CaseType CaseDetails.AppType " rec'd " DateObject.ReceivedMDY ", " CaseDetails.Eligibility CaseDetails.SaEntered
        If (CaseDetails.Eligibility = "pends") {
            MEC2NoteTitle .= " until " AutoDenyObject.AutoDenyExtensionDate
            MaxisNote := "CCAP app rec'd " DateObject.ReceivedMDY ", pend date " AutoDenyObject.AutoDenyExtensionDate ".`n"
        }
	} else if (CaseDetails.DocType = "Redet") {
		MEC2NoteTitle := CaseDetails.CaseType CaseDetails.DocType " rec'd " DateObject.ReceivedMDY ", " CaseDetails.Eligibility CaseDetails.SaEntered CaseDetails.RedetDue
	}
    if (CaseDetails.Eligibility = "elig") {
        IsExpedited := (Homeless = 1) ? " Expedited." : ""
        MaxisNote := "CCAP app rec'd " DateObject.ReceivedMDY ", approved eligible." IsExpedited "`n"
    }
    if (CaseDetails.Eligibility = "ineligible" || CaseDetails.Eligibility = "over-income") {
        MaxisNote := "CCAP app rec'd " DateObject.ReceivedMDY ", denied " DateObject.TodayMDY ".`n"
        if (OverIncomeMissing) {
            MaxisNote .= " Over-income"
        }
        MaxisNote .= "`n"
    }
    if (StrLen(Missing) > 0) {
        MissingMax := st_wordWrap(Missing, 72, "_")
        MissingMax := StrReplace(MissingMax, "`n", "`n* ")
        MissingMax := StrReplace(MissingMax, "* _", "  ")
        MaxisNote .= "Special Letter mailed " DateObject.TodayMDY " requesting:`n* " MissingMax "`n"
    }
    MaxisNote := StrReplace(MaxisNote, "              ", " ")
    MaxisNote .= Ini.EmployeeInfo.EmployeeName

    If (StrLen(MEC2NoteTitle) = 0 || InStr(MEC2NoteTitle, "?")) {
        MsgBox, ,Case Note Error,Select options in the top left before case noting`n     (Document type, Program, Eligibility)
        Return
    }

    If (A_GuiControl = "MEC2NoteButton") {
        StrReplace(MEC2CaseNote, "`n", "`n", MEC2CaseNoteLines) ; Counting new lines
        If (MEC2CaseNoteLines +1 = 31) { ;31 lines, signature lines combined
            MEC2CaseNote := StrReplace(MEC2CaseNote, "`n=====`n", "`n===== ")
        } else If (MEC2CaseNoteLines +1 > 31) {
            MsgBox,,MEC2 Case Note over 30 lines, Notice - Your case note is over 30 lines and will fail to save if not shortened.
        }
		WinActivate % Ini.EmployeeInfo.EmployeeBrowser
		Sleep 500
        MEC2DocType := CaseDetails.DocType = "Redet" ? "Redetermination" : CaseDetails.DocType
        If (Ini.EmployeeInfo.EmployeeUseMec2Functions = 1) {
            jsonCaseNote := "CaseNoteFromAHKJSON{""noteDocType"":""" MEC2DocType """,""noteTitle"":""" JSONstring(MEC2NoteTitle) """,""noteText"":""" JSONstring(MEC2CaseNote) """,""noteElig"":""" CaseDetails.Eligibility """ }"
            Clipboard := jsonCaseNote
            Send, ^v
        } Else If (Ini.EmployeeInfo.EmployeeUseMec2Functions = 0) {
            catNum := { Application: { letter: "A ", pends: 5, elig: 4, denied: 4 }, Redet: { letter: "R ", incomplete: 1, elig: 2, denied: 2 } }
            catLetter := catNum[CaseDetails.DocType].letter
            catNumber := catNum[CaseDetails.DocType][CaseDetails.Eligibility]
            WinActivate % Ini.EmployeeInfo.EmployeeBrowser
            Sleep 1000
            Send {Tab 7}
            Sleep 750
            SendInput, { %catLetter% %catNumber% }
            Sleep 500
            Send {Tab}
            Sleep 500
            SendInput, % MEC2NoteTitle
            Sleep 500
            Send {Tab}
            Sleep 500
            SendInput, % MEC2CaseNote
            Sleep 500
            Send {Tab}
        }
        CaseNoteEntered.MEC2Note := 1
        GuiControl, MainGui:Text, MEC2NoteButton, MEC2 ✔ ; Chr(2714)
        Sleep 500
        Clipboard := CaseNumber
    }
    If (A_GuiControl = "MaxisNoteButton") {
        ;StrReplace(MaxisNote, "`n", "`n", MaxisNoteCaseNoteLines) ; Counting new lines
        Ini.EmployeeInfo.EmployeeMaxis := Ini.EmployeeInfo.EmployeeMaxis = "MAXIS-WINDOW-TITLE" ? "MAXIS" : Ini.EmployeeInfo.EmployeeMaxis
        MaxisWindow := WinExist(Ini.EmployeeInfo.EmployeeMaxis " ahk_exe bzmd.exe")
        If (MaxisWindow = "0x0")
            MaxisWindow := WinExist("BlueZone Mainframe ahk_exe bzmd.exe")

        If (WinExist(ahk_id %MaxisWindow%)) {
            WinActivate, ahk_id %MaxisWindow% ahk_exe bzmd.exe
            Clipboard := MaxisNote
            Sleep 500
            Send, ^v
        }
        ; Test area start
        ;If (MaxisNoteCaseNoteLines > 13) {
        
            ;MaxisNoteArray := StrSplit(MaxisNote, "`n")

            ;MaxisNoteArraySplitter(MaxisNoteArray) {
                ;let tempArray := ""
                ;While (A_Index < 15) {
                    ;tempArray .= MaxisNoteArray[1]
                    ;MaxisNoteArray.RemoveAt(1)
                ;}
                ;return tempArray
            ;}

            ;While (MaxisNoteArray.Length() > 0) {
                ;i := 1
                ;MaxisNotePage%i% := MaxisNoteArraySplitter(MaxisNoteArray)
                ;
            ;}
            
        ;}
        ;MaxisNote := RegExReplace(MaxisNote, "i)(?<=.*`n.*){3}*.`n", "`n4thline")
        ; Test area end

        CaseNoteEntered.MAXISNote := 1
        GuiControl, MainGui:Text, MaxisNoteButton, Maxis ✔ ; Chr(2714)
        sleep 500
        Clipboard := CaseNumber
    }

    If (A_GuiControl = "NotepadBackup") {
		NameTest := RegExMatch(HouseholdComp, "\w+\b", Notepadfilename)
        Notepadfilename := CaseNumber !== "" ? CaseNumber : Notepadfilename
        LetterLabel2 := StrLen(LetterText2) > 0 ? "`n== Letter Page 2 ==`n`n" : ""
        LetterLabel3 := StrLen(LetterText3) > 0 ? "`n`n== Letter Page 3 ==`n`n" : ""
        LetterLabel4 := StrLen(LetterText4) > 0 ? "`n`n== Letter Page 4 ==`n`n" : ""
        MaxisLabel := "`n`n== MAXIS Note ==`n"
        If (Ini.CaseNoteCountyInfo.CountyNoteInMaxis != 1) {
            MaxisLabel := ""
            MaxisNote := ""
        }
        FileAppend, % "====== Case Note Summary ======`n" MEC2NoteTitle "`n`n====== MEC2 Case Note ======`n" MEC2CaseNote "`n`n====== Email ======`n" EmailText.Output "`n`n====== Special Letter 1 ======`n`n" LetterText1 "`n" LetterLabel2 LetterText2 LetterLabel3 LetterText3 LetterLabel4 LetterText4 MaxisLabel MaxisNote "`n`n-------------------------------------------`n`n`n", % A_Desktop "\" Notepadfilename ".txt"
        GuiControl, MainGui:Text, NotepadBackup, Desktop ✔
        CaseNoteEntered.MEC2Note := 1
		CaseNoteEntered.MAXISNote := 1
	}
Return

AllInOneNote:

Return

addFifteenishDays(oldDate) {
	FormatTime, dayNumber, %oldDate%, WDay
	If (dayNumber = 7) {
		oldDate += 17, Days
	}
    Else If (dayNumber > 4) {
        oldDate += 18, Days
    }
	Else {
		oldDate += 16, Days
	}
    Return oldDate
}
AddDays(OldDate, AddedDays) {
    OldDate += AddedDays, Days
    Return OldDate
}
SubtractDates(FutureDate, PastDate) {
    EnvSub, FutureDate, %PastDate%, days
    Return FutureDate
}
FormatMDY(DateYMD) {
    FormatTime, DateMDY, %DateYMD%, M/d/yy
    Return DateMDY
}

CalcDates:
    Gui, MainGui: Submit, NoHide
    DateObject.ReceivedYMD := Received
    DateObject.ReceivedMDY := FormatMDY(DateObject.ReceivedYMD)
    DateObject.AutoDenyYMD := AddDays(DateObject.ReceivedYMD, 29)
    DateObject.AutoDenyMDY := FormatMDY(DateObject.AutoDenyYMD)
    DateObject.RecdPlusFortyfiveYMD := AddDays(DateObject.ReceivedYMD, 44)
    DateObject.TodayPlusFifteenishYMD := addFifteenishDays(DateObject.TodayYMD)
    
    DateObject.ExpeditedNinetyDaysYMD := AddDays(DateObject.ReceivedYMD, 89)
    DateObject.ExpeditedNinetyDaysMDY := FormatMDY(DateObject.ExpeditedNinetyDaysYMD)
    
    DateObject.RecdPlusFifteenishYMD := addFifteenishDays(DateObject.ReceivedYMD)
    DateObject.RecdPlusFifteenishMDY := FormatMDY(DateObject.RecdPlusFifteenishYMD)
    
    NeedsNoExtension := SubtractDates(DateObject.AutoDenyYMD, DateObject.TodayPlusFifteenishYMD)
    NeedsExtension := SubtractDates(DateObject.RecdPlusFortyfiveYMD, DateObject.TodayPlusFifteenishYMD)
    
    DateObject.TodayPlusFifteenishMDY := FormatMDY(DateObject.TodayPlusFifteenishYMD)
    
    ;Redetermination dates
    DateObject.SignedYMD := SignDate
    DateObject.SignedMDY := FormatMDY(SignDate)
    DateObject.RedetCaseCloseYMD := addFifteenishDays(DateObject.SignedYMD)
    DateObject.RedetCaseCloseMDY := FormatMDY(DateObject.RedetCaseCloseYMD)
    DateObject.RedetDocsLastDayMDY := FormatMDY(AddDays(DateObject.RedetCaseCloseYMD, 29))
    
    NeedsFakeExtension := DateObject.RecdPlusFifteenishMDY
    AutoDenyObject.AutoDenyExtensionSpecLetter :=
    If (CaseDetails.DocType = "Application") {
        If (CaseDetails.Eligibility = "pends") {
            If (NeedsNoExtension > -1) {
                AutoDenyObject.AutoDenyExtensionDate := DateObject.AutoDenyMDY
                AutoDenyObject.AutoDenyExtensionSpecLetter := "*You have through " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications."
                GuiControl, MainGui: Text, AutoDenyStatus, Has 15+ days before auto-deny
            } Else If (NeedsExtension > -1) {
                AutoDenyObject.AutoDenyExtensionDate := DateObject.TodayPlusFifteenishMDY
                AutoDenyObject.AutoDenyExtensionMECnote := "Auto-deny extended to " AutoDenyObject.AutoDenyExtensionDate " due to processing < 15 days before auto-deny.`n-`n"
                AutoDenyObject.AutoDenyExtensionSpecLetter := "*You have through " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications."
                GuiControl, MainGui: Text, AutoDenyStatus, % "Extend auto-deny to " AutoDenyObject.AutoDenyExtensionDate
            } Else {
                AutoDenyObject.AutoDenyExtensionDate := DateObject.TodayPlusFifteenishMDY
                AutoDenyObject.AutoDenyExtensionMECnote := "Reinstate date is " AutoDenyObject.AutoDenyExtensionDate " due to processing < 15 days before auto-deny.`n-`n"
                AutoDenyObject.AutoDenyExtensionSpecLetter := "**Please note that you will be mailed an auto-denial notice.`n  You have through " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications.`n  If you are eligible, your case will be reinstated."
                GuiControl, MainGui: Text, AutoDenyStatus, % "Auto-denies tonight, pends until " AutoDenyObject.AutoDenyExtensionDate
            }
        }
        If (Homeless = 1) {
            AutoDenyObject.AutoDenyExtensionDate := DateObject.ExpeditedNinetyDaysMDY
            AutoDenyObject.AutoDenyExtensionSpecLetter := "**You have until " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications"
        }
    }
    
    If (CaseDetails.DocType = "Redet") {
            AutoDenyObject.AutoDenyExtensionSpecLetter := "** If your redetermination is not completed by " DateObject.SignedMDY ",`n   your case will close on " DateObject.RedetCaseCloseMDY ". If it closes,`n   the latest it can be reinstated is " DateObject.RedetDocsLastDayMDY "."
    }
    If (CaseDetails.Eligibility = "elig") {
        AutoDenyObject := {}
    }
return

ApplicationRadio:
	CaseDetails.DocType := "Application"
	GuiControl, MainGui: Text, PendingRadio, Pending
	GuiControl, MainGui: Text, SignText, Signed:
    if (CaseDetails.AppType != "3550") {
        GuiControl, MainGui: Hide, SignText
        GuiControl, MainGui: Hide, SignDate
    }
    If (Ini.CaseNoteCountyInfo.CountyNoteInMaxis = 1) {
        GuiControl, MainGui: Show, MaxisNoteButton
    }
    GuiControl, MainGui: Show, MNBenefitsRadio
    GuiControl, MainGui: Show, AppRadio
Return

RedeterminationRadio:
	CaseDetails.DocType := "Redet"
    GoSub RevertLabels
	GuiControl, MainGui: Text, PendingRadio, Incomplete
	GuiControl, MainGui: Text, SignText, Due:
	GuiControl, MainGui: Show, SignText
	GuiControl, MainGui: Show, SignDate
	GuiControl, MainGui: Hide, AutoDenyStatus
	GuiControl, MainGui: Hide, MaxisNoteButton
	GuiControl, MainGui: Hide, MNBenefitsRadio
	GuiControl, MainGui: Hide, AppRadio
Return

App:
	CaseDetails.AppType := "3550"
    GoSub RevertLabels
	GuiControl, MainGui: Show, SignText
	GuiControl, MainGui: Show, SignDate
return

;Homeless:
    ;Gui, MainGui: Submit, NoHide
;Return

MNBenefits:
	CaseDetails.AppType := "MNB"
    GuiControl, MainGui: Text, HouseholdCompLabel, Household Comp (pages 1, 3-5)
    GuiControl, MainGui: Text, AddressVerificationLabel, Address Verification (page 3)
    GuiControl, MainGui: Text, SharedCustodyLabel, Absent Parent / Child (page 6)
    GuiControl, MainGui: Text, SchoolInformationLabel, School Information (page 7)
    GuiControl, MainGui: Text, IncomeLabel, Income (pages 2, 8-9)
    GuiControl, MainGui: Text, ChildSupportIncomeLabel, Child Support Income (page 9)
    GuiControl, MainGui: Text, ExpensesLabel, Expenses (page 10)
    GuiControl, MainGui: Text, AssetsLabel, Assets (page 10)
    GuiControl, MainGui: Text, ActivityandScheduleLabel, Activity and Schedule (pages 10-11)
    GuiControl, MainGui: Text, ProviderLabel, Provider (pages 12-15)
	GuiControl, MainGui: Hide, SignText
	GuiControl, MainGui: Hide, SignDate
    Gui, Show
return

RevertLabels:
    GuiControl, MainGui: Text, HouseholdCompLabel, Household Comp
    GuiControl, MainGui: Text, AddressVerificationLabel, Address Verification
    GuiControl, MainGui: Text, SharedCustodyLabel, Shared Custody
    GuiControl, MainGui: Text, SchoolInformationLabel, School Information
    GuiControl, MainGui: Text, IncomeLabel, Income
    GuiControl, MainGui: Text, ChildSupportIncomeLabel, Child Support Income
    GuiControl, MainGui: Text, ExpensesLabel, Expenses
    GuiControl, MainGui: Text, AssetsLabel, Assets
    GuiControl, MainGui: Text, ActivityandScheduleLabel, Activity and Schedule
Return

Eligible:
	CaseDetails.Eligibility := "elig"
    GuiControl,, ManualWaitlistBox, 0
    GuiControl, MainGui: Hide, ManualWaitlistBox
    GuiControl, MainGui: Show, SaApproved
    GuiControl, MainGui: Show, NoSA
    GuiControl, MainGui: Show, NoProvider
return

Pending:
	CaseDetails.Eligibility := "pends"
    GuiControl, MainGui: Hide, SaApproved
    GuiControl, MainGui: Hide, NoSA
    GuiControl, MainGui: Hide, NoProvider
    If (InStr(CaseDetails.CaseType, "BSF") && Ini.CaseNoteCountyInfo.Waitlist > 1) {
        GuiControl, MainGui: Show, ManualWaitlistBox
    }
return

Ineligible:
	CaseDetails.Eligibility := "ineligible"
    GuiControl,, ManualWaitlistBox, 0
    GuiControl, MainGui: Hide, ManualWaitlistBox
    GuiControl, MainGui: Hide, SaApproved
    GuiControl, MainGui: Hide, NoSA
    GuiControl, MainGui: Hide, NoProvider
return

SaApproved:
	CaseDetails.SaEntered := " & SA"
return

NoSA:
	CaseDetails.SaEntered := ", no SA"
return

NoProvider:
	CaseDetails.SaEntered := ", no provider"
return

BSF:
	CaseDetails.CaseType := "BSF "
return

TY:
	CaseDetails.CaseType := "TY "
return

CCMF:
	CaseDetails.CaseType := "CCMF "
return

CaseType:
    If (A_GuiControl = BSF)
        CaseDetails.CaseType := "BSF "
    If (A_GuiControl = TY)
        CaseDetails.CaseType := "TY "
    If (A_GuiControl = CCMF)
        CaseDetails.CaseType := "CCMF "
Return

ChildSupportCooperation:
    Gui, MainGui: Submit, NoHide
    If (StrLen(ChildSupportCooperation) = 0) {
        GuiControl, MainGui: Text, ChildSupportCooperation, %SharedCustody%
    }
Return

;==============================================================================================================================================================================================
;MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION  MAIN GUI SECTION  MAIN GUI SECTION  MAIN GUI SECTION 
;==============================================================================================================================================================================================


;==============================================================================================================================================================================================
;VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION  VERIFICATION SECTION  VERIFICATION SECTION 
;==============================================================================================================================================================================================
MissingButton:
    Gui, MissingGui: Restore
    Gui, MissingGui: Show, AutoSize
    Gui, MainGui: Submit, NoHide
    Gui, MissingGui: Submit, NoHide
Return

MissingGui:
    Column1of1 := "xm w390"

    ; 12 + 158 (170) + 240 (410) + 12 = 422
    Column1of2 := "xm w158"
    Column2of2 := "x170 yp+0 w240"

    ; 12 + 118 (130) + 120 (250) + 160 (400) + 12 = 422 uhh... that's 412.
    Column1of3 := "xm w118"
    Column2of3 := "x130 yp+0 w120"
    Column3of3 := "x262 yp+0 w138"
    ; 12 + 118 (130) + 280 (410) + 12 = 422
    Column2and3Of3 := "x130 yp+0 w280"

    LineColor := "0x5" ; https://gist.github.com/jNizM/019696878590071cf739
    TextLine := "x60 y+4 w250 h1 " LineColor
    ;LineColor := "717171"
    ;ProgressLine := "x50 y+4 w250 h1 Background" LineColor
    ;Gui, MissingGui: Add, Progress, %ProgressLine%

    Gui, MissingGui: New,, Missing Verifications
    Gui, MissingGui: Margin, 12 12
    Gui, MissingGui: Add, Checkbox, %Column1of1% vIDmissing gInputBoxAGUIControl, ID (input)
    Gui, MissingGui: Add, Checkbox, %Column1of1% vBCmissing gInputBoxAGUIControl, BC (input)
    Gui, MissingGui: Add, Checkbox, %Column1of1% vBCNonCitizenMissing gInputBoxAGUIControl, BC [non-citizen] (input)
    Gui, MissingGui: Add, Checkbox, %Column1of2% vAddressMissing, Address
    Gui, MissingGui: Add, Checkbox, %Column1of1% vChildSupportFormsMissing gInputBoxAGUIControl, Child Support forms (input)
    Gui, MissingGui: Add, Checkbox, %Column1of1% vChildSupportNoncooperationMissing gInputBoxAGUIControl, CS Non-cooperation (input)
    Gui, MissingGui: Add, Checkbox, %Column1of2% vCustodyScheduleMissing, Custody ("for each child")
    Gui, MissingGui: Add, Checkbox, %Column2of2% vCustodySchedulePlusNamesMissing gInputBoxAGUIControl, Custody (input)
    Gui, MissingGui: Add, Checkbox, %Column1of2% vChildSchoolMissing, Child school information
    Gui, MissingGui: Add, Checkbox, %Column2of2% vChildFTSchoolMissing, Child full-time student status
    Gui, MissingGui: Add, Checkbox, %Column1of2% vMarriageCertificateMissing, Marriage certificate
    Gui, MissingGui: Add, Checkbox, %Column2of2% vLegalNameChangeMissing gInputBoxAGUIControl, Name change (input)
    Gui, MissingGui: Add, Checkbox, %Column1of1% vDependentAdultStudentMissing gInputBoxAGUIControl, Dependent adult child - FT Student, 50`%+ expenses (input)

    Gui, MissingGui: Font, bold ; EARNED INCOME SECTION ==============================================================
    Gui, MissingGui: Add, Text, xm+10 y+22 w110 h1 %LineColor%
    Gui, MissingGui: Add, Text, x+m yp-7, Earned Income
    Gui, MissingGui: Add, Text, x+m yp+7 w115 h1 %LineColor%
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, %Column1of3% vIncomeMissing, Income
    Gui, MissingGui: Add, Checkbox, %Column2of3% vWorkScheduleMissing, Work Schedule
    Gui, MissingGui: Add, Checkbox, %Column3of3% vContractPeriodMissing, Contract Period
    Gui, MissingGui: Add, Checkbox, %Column1of2% vIncomePlusNameMissing gInputBoxAGUIControl, Income (input)
    Gui, MissingGui: Add, Checkbox, %Column2of2% vWorkSchedulePlusNameMissing gInputBoxAGUIControl, Work Schedule (input)
    Gui, MissingGui: Add, Checkbox, %Column1of1% vNewEmploymentMissing, New job at app / end of job search (Wage, dates, hours)
    Gui, MissingGui: Add, Checkbox, %Column1of1% vWorkLeaveMissing, Leave of absence (Dates, pay status, hours, work schedule)
    Gui, MissingGui: Add, Text, %TextLine% ; -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, %Column1of1% vSeasonalWorkMissing, Seasonal employment season length
    Gui, MissingGui: Add, Checkbox, %Column1of1% vSeasonalOffSeasonMissing gInputBoxAGUIControl, Seasonal employment info - app in off-season (input)
    Gui, MissingGui: Add, Text, %TextLine% ; -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, %Column1of2% vSelfEmploymentMissing, Self-Employment Income
    Gui, MissingGui: Add, Checkbox, %Column2of2% vSelfEmploymentScheduleMissing, Self-Employment Schedule
    Gui, MissingGui: Add, Checkbox, %Column1of1% vSelfEmploymentBusinessGrossMissing, Self-Employment Business Gross (if state min wage; <$500k = small business)
    Gui, MissingGui: Add, Text, %TextLine% ; -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, %Column1of2% vExpensesMissing, Expenses
    Gui, MissingGui: Add, Checkbox, %Column2of2% vOverIncomeMissing gInputBoxAGUIControl, Over-income (input)

    Gui, MissingGui: Font, bold ; UNEARNED INCOME SECTION ============================================================
    Gui, MissingGui: Add, Text, xm+10 y+22 w105 h1 %LineColor%
    Gui, MissingGui: Add, Text, x+m yp-7, Unearned Income
    Gui, MissingGui: Add, Text, x+m yp+7 w120 h1 %LineColor%
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, %Column1of2% vChildSupportIncomeMissing, Child Support Income
    Gui, MissingGui: Add, Checkbox, %Column2of2% vSpousalSupportMissing, Spousal Support Income
    Gui, MissingGui: Add, Checkbox, %Column1of2% vRentalMissing, Rental
    Gui, MissingGui: Add, Checkbox, %Column2of2% vDisabilityMissing, STD / LTD 
    Gui, MissingGui: Add, Checkbox, %Column1of2% vAssetsGT1mMissing, Assets (>$1m)
    Gui, MissingGui: Add, Checkbox, %Column2of2% vUnearnedStatementMissing, Blank Unearned Yes/No (statement)
    Gui, MissingGui: Add, Checkbox, %Column1of2% vAssetsBlankMissing, Assets (Blank)
    Gui, MissingGui: Add, Checkbox, %Column2of2% vUnearnedMailedMissing, Blank Unearned Yes/No (mailed back)
    Gui, MissingGui: Add, Checkbox, %Column1of2% vVABenefitsMissing, VA Benefits
    Gui, MissingGui: Add, Checkbox, %Column2of2% vInsuranceBenefitsMissing, Insurance Benefits

    Gui, MissingGui: Font, bold ; ACTIVITY SECTION ===================================================================
    Gui, MissingGui: Add, Text, xm+10 y+22 w130 h1 %LineColor%
    Gui, MissingGui: Add, Text, x+m yp-7, Activity
    Gui, MissingGui: Add, Text, x+m yp+7 w140 h1 %LineColor%
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, %Column1of2% vEdBSFformMissing, BSF/TY Education Form
    Gui, MissingGui: Add, Checkbox, %Column2of2% vEdBSFOneBachelorDegreeMissing, BSF/TY Bachelor's limit notice
    Gui, MissingGui: Add, Checkbox, %Column1of2% vClassScheduleMissing, Class schedule
    Gui, MissingGui: Add, Checkbox, %Column2of2% vTranscriptMissing, Transcript
    Gui, MissingGui: Add, Checkbox, %Column1of2% vEducationEmploymentPlanMissing, ES Plan (CCMF Education)
    Gui, MissingGui: Add, Checkbox, %Column2of2% vStudentStatusOrIncomeMissing, Adult student w/ income (age < 20)
    Gui, MissingGui: Add, Text, %TextLine% ; -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, %Column1of2% y+4 vJobSearchHoursMissing, BSF Job search hours
    Gui, MissingGui: Add, Checkbox, %Column2of2% vSelfEmploymentIneligibleMissing, Self-Employment not enough hours
    Gui, MissingGui: Add, Checkbox, %Column1of2% vEligibleActivityMissing, No Eligible Activity Listed
    Gui, MissingGui: Add, Checkbox, %Column2of2% vEmploymentIneligibleMissing, Employment not enough hours
    Gui, MissingGui: Add, Checkbox, %Column1of2% vESPlanOnlyJSMissing, ES Plan-only JS notice
    Gui, MissingGui: Add, Checkbox, %Column2of2% vActivityAfterHomelessMissing, Activity Req. After 3-Mo Homeless Period

    Gui, MissingGui: Font, bold ; PROVIDER SECTION ===================================================================
    Gui, MissingGui: Add, Text, xm+10 y+22 w125 h1 %LineColor%
    Gui, MissingGui: Add, Text, x+m yp-7, Provider
    Gui, MissingGui: Add, Text, x+m yp+7 w130 h1 %LineColor%
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, %Column1of2% vNoProviderMissing, No Provider Listed
    Gui, MissingGui: Add, Checkbox, %Column2of2% vUnregisteredProviderMissing,Unregistered Provider
    Gui, MissingGui: Add, Checkbox, %Column1of2% vInHomeCareMissing, In-Home Care form
    Gui, MissingGui: Add, Checkbox, %Column2of2% vLNLProviderMissing, LNL Acknowledgement
    Gui, MissingGui: Add, Checkbox, %Column1of2% vStartDateMissing, Provider start date
    Gui, MissingGui: Add, Checkbox, %Column2of2% vProviderForNonImmigrantMissing, Non-citizen/immigrant Provider Reqs.

    Gui, MissingGui: Add, Checkbox, %Column1of1% h50 vOther1 gOther, Other
    Gui, MissingGui: Add, Checkbox, %Column1of1% h50 vOther2 gOther, Other
    Gui, MissingGui: Add, Checkbox, %Column1of1% h50 vOther3 gOther, Other

    Gui, MissingGui: Add, Button, h17 gMissingButtonDoneButton, Done
    Gui, MissingGui: Add, Button, x+20 w40 h17 hidden gEmail vEmail, Email
    Gui, MissingGui: Add, Button, x+20 w42 h17 hidden gLetter vLetter1, Letter 1
    Gui, MissingGui: Add, Button, x+20 w42 h17 hidden gLetter vLetter2, Letter 2
    Gui, MissingGui: Add, Button, x+20 w42 h17 hidden gLetter vLetter3, Letter 3
    Gui, MissingGui: Add, Button, x+20 w42 h17 hidden gLetter vLetter4, Letter 4
    Gui, MissingGui: Show, % "Hide x" Ini.CaseNotePositions.XVerification " y" Ini.CaseNotePositions.YVerification
Return

MissingButtonDoneButton:
	Gui, MainGui: Submit, NoHide
	Gui, MissingGui: Submit, NoHide
    GoSub CalcDates
	MissingVerifications := {}, ClarifiedVerifications := {}, LineCount := 0, EmailText := {}
    EmailTextString := "", LetterText1 := "", LetterText2 := "", LetterText3 := "", LetterText4 := ""
    MEC2DocType := CaseDetails.DocType = "Redet" ? "Redetermination" : CaseDetails.DocType
	GuiControl, MissingGui:Hide, Letter2
	GuiControl, MissingGui:Hide, Letter3
	GuiControl, MissingGui:Hide, Letter4
	MissingVerifications := new OrderedAssociativeArray()
    ClarifiedVerifications := new OrderedAssociativeArray()
    MecCheckboxIds := {}
	LetterTextVar := "LetterText1", LetterNumber := 1, LineNumber := 1, ListItem := 1, ClarifyListItem := 1, EmailListItem := 1, CaseNoteMissing := "", Email := ""
    VerifCat :=, LetterTextNumber := 1, LetterText := {}
    PendingHomelessPreText := "You may be eligible for the homeless policy, which allows us to approve eligibility even though there are verifications we need but do not have. These verifications are still required, and must be received within 90 days of your application date for continued eligibility.`n`nBefore we can approve expedited eligibility, we need`n information that was not on the application:`n"
    
    If OverIncomeMissing {
        OverIncomeMissingText1 := "Using information you provided your case is ineligible as your income is over the limit for a household of " OverIncomeObj.overIncomeHHsize ". The gross limit is $" OverIncomeObj.overIncomeText ".`n"
        OverIncomeMissingText2 := "If your gross income does not match this calculation, you must" CountySpecificText[Ini.EmployeeInfo.EmployeeCounty].OverIncomeContactInfo " submit updated income and expense documents along with the following verifications:`n"
        EmailTextString := "Your Child Care Assistance " MEC2DocType " has been processed.`n`n" OverIncomeMissingText1 OverIncomeMissingText2 "`n"
        MissingVerifications[OverIncomeMissingText1] := 3
        MissingVerifications[OverIncomeMissingText2] := 3
        CaseNoteMissing .= "Household is calculated to be over-income by $" OverIncomeObj.overIncomeDifference " ($" OverIncomeObj.overIncomeReceived " - $" OverIncomeObj.overIncomeLimit ");`n"
    }
    If (Homeless && CaseDetails.Eligibility = "pends" && CaseDetails.DocType = "Application") {
        InputBox, MissingHomelessItems, Homeless Info Missing, What information is needed from the client to approve expedited eligibility?`n`nUse a double space "  " without quotation marks to start a new line.,,,,,,,, % StrReplace(MissingHomelessItems, "`n", "  ")
        If (ErrorLevel = 0) {
            MissingHomelessItems := StrReplace(MissingHomelessItems, "  ", "`n")
            PendingHomelessMissing := getRowCount("  " MissingHomelessItems, 58, "  ")
            MissingVerifications[st_wordwrap(PendingHomelessPreText, 59, " ") "`n"] := 8
            MissingVerifications[PendingHomelessMissing[1] "`n"] := PendingHomelessMissing[2]
            CaseNoteMissing .= "Missing for expedited approval:`n" StrReplace(MissingHomelessItems, "`n", "`n  ") ";`n"
        }
    }
	If (!OverIncomeMissing) {
        EmailText.StartHL := (CaseDetails.Eligibility = "elig") ? "It was approved under the homeless expedited policy which allows us to approve eligibility even though there are verifications we require that we do not have. These verifications are still required, and must be received within 90 days of your application date for continued eligibility." : PendingHomelessPreText MissingHomelessItems

        EmailText.EndHL := (CaseDetails.Eligibility = "elig") ? "`nThe initial approval of child care assistance is 30 hours per week for each child. This amount can be increased once we receive your activity verifications and we determine more assistance is needed.`nIf the provider you select is a “High Quality” provider, meaning they are Parent Aware 3⭐ or 4⭐ rated, or have an approved accreditation, the hours will automatically increase to 50 per week for preschool age and younger children.`nIf you have a 'copay,' the amount the county pays to the provider will be reduced by the copay amount. Many providers charge more than our maximum rates, and you are responsible for your copay and any amounts the county cannot pay." : ""

        EmailText.AreOrWillBe := (Homeless = 1) ? "will be" : "are"

        EmailText.Reason1 := (CaseDetails.Eligibility = "elig") ? "for authorizing assistance hours" : "to determine eligibility or calculate assistance hours"
        EmailText.Reason2 := (Homeless = 1) ? "to determine on-going eligibility or calculate assistance hours after the 90-day period" : EmailText.Reason1
        EmailText.StartAll := "Your Child Care Assistance " MEC2DocType " has been processed. " EmailText.WaitList

        EmailText.Start := (Homeless = 1) ? EmailText.StartAll EmailText.StartHL : EmailText.StartAll
        EmailText.Middle := "`n`nThe following documents or verifications " EmailText.AreOrWillBe " needed " EmailText.Reason2 ":`n`n"

        EmailText.Combined := EmailText.Start EmailText.Middle
	}
    
	If IDmissing {
        IDmissingText := "ID for " MissingInput.IDmissing ";`n"
		ClarifiedVerifications[ClarifyListItem ". " IDmissingText] := 1
        EmailTextString .= EmailListItem ". " IDmissingText
		CaseNoteMissing .= "ID for " MissingInput.IDmissing ";`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfIdentity := 1
    }
	If BCmissing {
        BCmissingText := "Birth date / relationship / citizenship verification for: " MissingInput.BCmissing
		ClarifiedVerifications[ClarifyListItem ". " BCmissingText ";`n"] := 2
        EmailTextString .= EmailListItem ". " BCmissingText ", such as the official birth certificate;`n"
		CaseNoteMissing .= BCmissingText ";`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfBirth := 1
        MecCheckboxIds.proofOfRelation := 1
        MecCheckboxIds.citizenStatus := 1
    }
	If BCNonCitizenMissing {
        BCNonCitizenMissingText := "Birth date / relationship / immigration verification for: " MissingInput.BCNonCitizenMissing ";`n"
		ClarifiedVerifications[ClarifyListItem ". " BCNonCitizenMissingText] := 2
        EmailTextString .= EmailListItem ". " BCNonCitizenMissingText
		CaseNoteMissing .= "Birth date / relationship / immigration verification for: " MissingInput.BCNonCitizenMissing ";`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfBirth := 1
        MecCheckboxIds.proofOfRelation := 1
        MecCheckboxIds.citizenStatus := 1
    }
	If AddressMissing {
        If (Homeless = 1) {
            AddressMissingText := "Verification of current residence, such as a signed statement of your county of residence;`n"
            ClarifiedVerifications[ClarifyListItem ". " AddressMissingText] := 2
            EmailTextString .= EmailListItem ". " AddressMissingText
            ClarifyListItem++
            EmailListItem++
            MecCheckboxIds.proofOfResidence := 1
        } Else If (Homeless = 0) {
            AddressMissingText := "Verification of current residence;`n"
            EmailTextString .= EmailListItem ". " AddressMissingText
            MecCheckboxIds.proofOfResidence := 1
            EmailListItem++
        }
		CaseNoteMissing .= "Address;`n"
        MecCheckboxIds.proofOfResidence := 1
    }
	If ChildSupportFormsMissing {
        If (MissingInput.ChildSupportFormsMissing ~= "^\d$") {
            MissingInput.ChildSupportFormsMissing .= MissingInput.ChildSupportFormsMissing < 2 ? " set" : " sets"
        }
        ChildSupportFormsMissingText := "Cooperation with Child Support forms (" MissingInput.ChildSupportFormsMissing ", sent separately);`n"
        ;CSFMlines := MissingInput.ChildSupportFormsMissing ~= "^\d" ? 1 : 2
		MissingVerifications[ListItem ". " ChildSupportFormsMissingText] := 2 ; CSFMlines
        EmailTextString .= EmailListItem ". " ChildSupportFormsMissingText
		CaseNoteMissing .= "CS forms (" MissingInput.ChildSupportFormsMissing ");`n"
		ListItem++
        EmailListItem++
    }
	If CustodyScheduleMissing {
        CustodyScheduleMissingText := "A statement, written by you that is signed and dated, for each child that has a parent not in your household:`n  A. Stating that you have full custody, or`n  B. Your current Parenting Time (shared custody) schedule `n     listing the days and times of the custody switches;`n"
		MissingVerifications[ListItem ". " CustodyScheduleMissingText] := 5
        EmailTextString .= EmailListItem ". " CustodyScheduleMissingText
		CaseNoteMissing .= "Shared custody / parenting time;`n"
		ListItem++
        EmailListItem++
    }
	If CustodySchedulePlusNamesMissing {
        CustodyScheduleMissingText := "A statement, written by you that is signed and dated, for " MissingInput.CustodySchedulePlusNamesMissing ":`n  A. Stating that you have full custody, or`n  B. Your current Parenting Time (shared custody) schedule `n     listing the days and times of the custody switches;`n"
		MissingVerifications[ListItem ". " CustodyScheduleMissingText] := 5
        EmailTextString .= EmailListItem ". " CustodyScheduleMissingText
		CaseNoteMissing .= "Shared custody / parenting time for " MissingInput.CustodySchedulePlusNamesMissing ";`n"
		ListItem++
        EmailListItem++
    }
    if DependentAdultStudentMissing {
        DependentAdultStudentMissingText := "Verification of full-time student status for " MissingInput.DependentAdultStudentMissing ", verification of their most recent 30 days income, and a signed statement that you provide at least 50% of their financial support;`n"
        MissingVerifications[ListItem ". " DependentAdultStudentMissingText] := 3
        EmailTextString .= EmailListItem ". " DependentAdultStudentMissingText
		CaseNoteMissing .= "Dependant Adult FT school status, income, statement of 50% support;`n"
		ListItem++
        EmailListItem++
    }
	If ChildSchoolMissing {
        ChildSchoolMissingText := "Child's school information (location, grade, start/end times) - does not need to be verification from the school;`n"
        EmailTextString .= EmailListItem ". " ChildSchoolMissingText
		CaseNoteMissing .= "Child school information;`n"
        MecCheckboxIds.childSchoolSchedule := 2
        EmailListItem++
        ;MEC2 text: Child School Schedule- You can provide the school schedule of each child that needs child care by sending a copy of the days and times of school from the school's website or handbook, writing the information on a piece of paper, or telling your worker.
    }
    If ChildFTSchoolMissing {
        ChildFTSchoolMissingText := "Verification of full-time student status for minor children with employment OR their most recent 30 days income (income is not counted if attending school full-time);`n"
        MissingVerifications[ListItem ". " ChildFTSchoolMissingText] := 3
        EmailTextString .= EmailListItem ". " ChildFTSchoolMissingText
		CaseNoteMissing .= "Minor child FT school status or income;`n"
		ListItem++
        EmailListItem++
    }
	If MarriageCertificateMissing {
        MarriageCertificateMissingText := "Marriage verification (example: marriage certificate);`n"
		MissingVerifications[ListItem ". " MarriageCertificateMissingText] := 1
        EmailTextString .= EmailListItem ". " MarriageCertificateMissingText
		CaseNoteMissing .= "Marriage certificate;`n"
		ListItem++
        EmailListItem++
        MecCheckboxIds.proofOfRelation := 1
    }
	If LegalNameChangeMissing {
        LegalNameChangeMissingText := "Legal name change verification for " MissingInput.LegalNameChangeMissing ";`n"
		MissingVerifications[ListItem ". " LegalNameChangeMissingText] := 1
        EmailTextString .= EmailListItem ". " LegalNameChangeMissingText
		CaseNoteMissing .= "Legal name change for " MissingInput.LegalNameChangeMissing ";`n"
		ListItem++
        EmailListItem++
    }
;======================================================
	If IncomeMissing {
        IncomeText := NeedsExtension > -1 ? " your most recent 30 days income" : CaseDetails.DocType = "Redet" ? " 30 days income prior to " DateObject.SignedMDY : " 30 days income prior to " DateObject.ReceivedMDY
        ; IncomeText := if doesn't need extension : elseif redetermination : elseif app needs extension
        IncomeMissingText := "Verification of" IncomeText ";`n"
        ClarifiedVerifications[ClarifyListItem ". Proof of Financial Information: " IncomeMissingText] := 2
        EmailTextString .= EmailListItem ". " IncomeMissingText
		CaseNoteMissing .= "Earned income;`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfFInfo := 1
        ;MEC2 text: Proof of Financial Information- You can provide proof of financial information and income with the last 30 days of check stubs, income tax records, business ledger, award letter, or a letter from your employer with pay rate, number of hours worked per week and how often you are paid.
    }
	If IncomePlusNameMissing {
        IncomeText := NeedsExtension > -1 ? MissingInput.IncomePlusNameMissing "'s most recent 30 days income" : CaseDetails.DocType = "Redet" ? MissingInput.IncomePlusNameMissing "'s 30 days income prior to " DateObject.SignedMDY : MissingInput.IncomePlusNameMissing "'s 30 days income prior to " DateObject.ReceivedMDY
        IncomeMissingText := "Verification of " IncomeText ";`n"
        ClarifiedVerifications[ClarifyListItem ". Proof of Financial Information: " IncomeMissingText] := 2
        EmailTextString .= EmailListItem ". " IncomeMissingText
		CaseNoteMissing .= "Earned income (" MissingInput.IncomePlusNameMissing ");`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfFInfo := 1
    }
	If WorkScheduleMissing {
        WorkScheduleText := NeedsExtension > -1 ? " your work schedule" : CaseDetails.DocType = "Redet" ? " work schedule from " DateObject.SignedMDY : " work schedule from " DateObject.ReceivedMDY
        WorkScheduleMissingText := "Verification of" WorkScheduleText " showing days of the week and start/end times;`n"
        ClarifiedVerifications[ClarifyListItem ". Proof of Activity Schedule: " WorkScheduleMissingText] := 2
        EmailTextString .= EmailListItem ". " WorkScheduleMissingText
		CaseNoteMissing .= "Work schedule;`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfActivitySchedule := 1
        ;MEC2 text: Proof of Activity Schedule- You can provide proof of adult activity schedules with work schedules, school schedules, time cards, or letter from the employer or school with the days and times working or in school. If you have a flexible work schedule, include a statement with typical or possible times worked.
    }
	If WorkSchedulePlusNameMissing {
        WorkScheduleText := NeedsExtension > -1 ? MissingInput.WorkSchedulePlusNameMissing "'s work schedule" : CaseDetails.DocType = "Redet" ? MissingInput.WorkSchedulePlusNameMissing "'s work schedule from " DateObject.SignedMDY : MissingInput.WorkSchedulePlusNameMissing "'s work schedule from " DateObject.ReceivedMDY
        WorkScheduleMissingText := "Verification of " WorkScheduleText " showing days of the week and start/end times;`n"
        ClarifiedVerifications[ClarifyListItem ". Proof of Activity Schedule: " WorkScheduleMissingText] := 2
        EmailTextString .= EmailListItem ". " WorkScheduleMissingText
		CaseNoteMissing .= "Work schedule (" MissingInput.WorkSchedulePlusNameMissing ");`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfActivitySchedule := 1
    }
	If ContractPeriodMissing {
        ContractPeriodMissingText := "Employment Contract Period verification if not full-year;`n"
		MissingVerifications[ListItem ". " ContractPeriodMissingText] := 1
        EmailTextString .= EmailListItem ". " ContractPeriodMissingText
		CaseNoteMissing .= "Employment Contract Period;`n"
		ListItem++
        EmailListItem++
    }
	If NewEmploymentMissing {
        NewEmploymentMissingText := "Verification of employment start date, wage and expected hours per week, and first pay date;`n"
		MissingVerifications[ListItem ". " NewEmploymentMissingText] := 2
        EmailTextString .= EmailListItem ". " NewEmploymentMissingText
		CaseNoteMissing .= "New employment information;`n"
		ListItem++
        EmailListItem++
    }
    If WorkLeaveMissing {
        WorkLeaveMissingText := "Verification of leave of absence, including: `nPaid/unpaid status, start date, and expected: return date, wage, and hours per week. Upon returning, we need your work schedule showing days of the week and start/end times;`n"
		MissingVerifications[ListItem ". " WorkLeaveMissingText] := 4
        EmailTextString .= EmailListItem ". " WorkLeaveMissingText
		CaseNoteMissing .= "Leave of absence details;`n"
		ListItem++
        EmailListItem++
    }
;----------------------------
    If SeasonalWorkMissing {
        SeasonalWorkMissingText := "Verification of seasonal employment expected season length;`n"
		MissingVerifications[ListItem ". " SeasonalWorkMissingText] := 2
        EmailTextString .= EmailListItem ". " SeasonalWorkMissingText
		CaseNoteMissing .= "Seasonal employment season length;`n"
        EmailListItem++
        ListItem++
    }
    If SeasonalOffSeasonMissing {
        ;SeasonalOffSeasonMissing := StrLen(SeasonalOffSeasonMissing) > 0 ? " at " SeasonalOffSeasonMissing : ""
        SeasonalOffSeasonMissing := MissingInput.SeasonalOffSeasonMissing != "" ? " at " MissingInput.SeasonalOffSeasonMissing : ""
        SeasonalOffSeasonMissingText := "Verification of either seasonal employment " SeasonalOffSeasonMissing ", including expected season length and typical wages, or a signed statement that you are no longer an employee at this job.`n Upon returning to work, verification of work schedule will`n be needed, showing days of the week and start/end times;`n"
		MissingVerifications[ListItem ". " SeasonalOffSeasonMissingText] := 6
        EmailTextString .= EmailListItem ". " SeasonalOffSeasonMissingText
		CaseNoteMissing .= "Seasonal employment (applied during off season);`n"
        ListItem++
        EmailListItem++
    }
;----------------------------
	If SelfEmploymentMissing {
        SelfEmploymentMissingText := "Self-employment income such as your recent complete federal tax return. (For new self-employment, state your start date). If you haven't yet filed taxes or your taxes don't represent expected ongoing income, submit monthly reports or ledgers with the most recent full 3 months of gross income;`n"
        ;MEC2 text: Proof of Financial Information- You can provide proof of financial information and income with the last 30 days of check stubs, income tax records, business ledger, award letter, or a letter from your employer with pay rate, number of hours worked per week and how often you are paid. 
		MissingVerifications[ListItem ". " SelfEmploymentMissingText] := 5
        EmailTextString .= EmailListItem ". " SelfEmploymentMissingText
		CaseNoteMissing .= "Self-Employment income;`n"
		ListItem++
        EmailListItem++
    }
	If SelfEmploymentScheduleMissing {
        SelfEmploymentScheduleMissingText := "Written statement of your self-employment work schedule with days of the week and start/end times;`n"
        ;MEC2 text: Proof of Activity Schedule- You can provide proof of adult activity schedules with work schedules, school schedules, time cards, or letter from the employer or school with the days and times working or in school. If you have a flexible work schedule, include a statement with typical or possible times worked.
		MissingVerifications[ListItem ". " SelfEmploymentScheduleMissingText] := 2
        EmailTextString .= EmailListItem ". " SelfEmploymentScheduleMissingText
		CaseNoteMissing .= "Self-Employment work schedule;`n"
		ListItem++
        EmailListItem++
    }
    If SelfEmploymentBusinessGrossMissing {
        SelfEmploymentBusinessGrossMissingText := "Information regarding your self-employment business' annual gross income, if it is less than $500,000 (optional);`n"
        MissingVerifications[ListItem ". " SelfEmploymentBusinessGrossMissingText] := 2
        EmailTextString .= EmailListItem ". " SelfEmploymentBusinessGrossMissingText
		CaseNoteMissing .= "Self-Employment gross (if subject to small/large min wage: <$500k/yr?) - not required;`n"
		ListItem++
        EmailListItem++
    }
;----------------------------
	If ExpensesMissing {
        ExpensesMissingText := "Proof of Deductions: Healthcare Insurance premiums, child support, and spousal support - if not listed on submitted paystubs;`n"
        EmailTextString .= EmailListItem ". " ExpensesMissingText
		CaseNoteMissing .= "Expenses;`n"
        EmailListItem++
        MecCheckboxIds.proofOfDeductions := 1
        ;MEC2 text: Proof of Deductions- You can provide proof of expenses for health insurance premiums (medical, dental, vision), child support paid for a child not living in your home, and spousal support with check stubs, benefit statements or premium statements. 
    }
; over-income is here in the list but has its own sub-routine.
;======================================================
	If ChildSupportIncomeMissing {
        ChildSupportIncomeMissingText := "Verification of your Child Support income;`n"
		MissingVerifications[ListItem ". " ChildSupportIncomeMissingText] := 1
        EmailTextString .= EmailListItem ". " ChildSupportIncomeMissingText
		CaseNoteMissing .= "Child Support income;`n"
		ListItem++
        EmailListItem++
    }
	If SpousalSupportMissing {
        SpousalSupportMissingText := "Verification of your Spousal Support income;`n"
		MissingVerifications[ListItem ". " SpousalSupportMissingText] := 1
        EmailTextString .= EmailListItem ". " SpousalSupportMissingText
		CaseNoteMissing .= "Spousal Support income;`n"
		ListItem++
        EmailListItem++
    }
	If RentalMissing {
        RentalMissingText := "Verification of your rental income;`n"
		MissingVerifications[ListItem ". " RentalMissingText] := 1
        EmailTextString .= EmailListItem ". " RentalMissingText
		CaseNoteMissing .= "Rental income;`n"
		ListItem++
        EmailListItem++
    }
	If DisabilityMissing {
        DisabilityMissingText := "Verification of your disability income;`n"
		MissingVerifications[ListItem ". " DisabilityMissingText] := 1
        EmailTextString .= EmailListItem ". " DisabilityMissingText
		CaseNoteMissing .= "STD / LTD;`n"
		ListItem++
        EmailListItem++
    }
	If InsuranceBenefitsMissing {
        InsuranceBenefitsMissingText := "Verification of your Insurance Benefits income;`n"
		MissingVerifications[ListItem ". " InsuranceBenefitsMissingText] := 1
        EmailTextString .= EmailListItem ". " InsuranceBenefitsMissingText
		CaseNoteMissing .= "Insurance benefits income;`n"
		ListItem++
        EmailListItem++
    }
    If UnearnedStatementMissing {
        UnearnedStatementMissingText := "A statement written by you that is signed and dated, stating if you have any unearned income. Submit verification if yes.`nThis includes: Child/Spousal support, Rentals, Unemployment, RSDI, Insurance payments, VA benefits, Trust income, Contract for deed, Interest, Dividends, Gambling winnings, Inheritance, Capital gains, etc.;`n"
        MissingVerifications[ListItem ". " UnearnedStatementMissingText] := 6
        EmailTextString .= EmailListItem ". " UnearnedStatementMissingText
        CaseNoteMissing .= "Unearned income yes / no questions (statement);`n"
        ListItem++
        EmailListItem++
    }
	If VABenefitsMissing {
        VABenefitsMissingText := "Verification of your VA income;`n"
		MissingVerifications[ListItem ". " VABenefitsMissingText] := 1
        EmailTextString .= EmailListItem ". " VABenefitsMissingText
		CaseNoteMissing .= "VA income;`n"
		ListItem++
        EmailListItem++
    }
    If UnearnedMailedMissing {
        UnearnedMailedMissingText := "Unearned income questions that were not answered (sent separately);`n"
        MissingVerifications[ListItem ". " UnearnedMailedMissingText] := 2
        EmailTextString .= EmailListItem ". " UnearnedMailedMissingText
        CaseNoteMissing .= "Unearned income yes / no questions (mailed back);`n"
        ListItem++
        EmailListItem++
    }
	If AssetsBlankMissing {
        AssetsBlankMissingText := "Written or verbal statement of your assets being either MORE THAN or LESS THAN $1 million;`n"
		MissingVerifications[ListItem ". " AssetsBlankMissingText] := 2
        EmailTextString .= EmailListItem ". " AssetsBlankMissingText
		CaseNoteMissing .= "Assets amount statement;`n"
		ListItem++
        EmailListItem++
    }
	If AssetsGT1mMissing {
        AssetsGT1mMissingText := "Clarification of your assets, which you listed as MORE THAN $1 million;`n"
		MissingVerifications[ListItem ". " AssetsGT1mMissingText] := 1
        EmailTextString .= EmailListItem ". " AssetsGT1mMissingText
		CaseNoteMissing .= "Assets clarification (>$1m on app);`n"
		ListItem++
        EmailListItem++
    }
;======================================================
    ;If (UnearnedUnansweredMissing || LumpSumUnansweredMissing || EmploymentUnansweredMissing || SelfEmploymentUnansweredMissing || AssetsUnansweredMissing) {
        ;UnansweredText := "You did not answer all questions on the " MEC2DocType ".`n Please submit a statement that is written, dated, and signed by you, answering:`n"
        ;CaseNoteUnansweredMissing := Unanswered ?: 
        ;AnsweredMoreThan :=, AnsweredYes :=, AnsweredBoth :=
        ;If UnearnedUnansweredMissing {
            ;UnansweredText .= "Have you received any unearned income in the past 12 months? (Includes: Child/Spousal support, Unemployment, RSDI, Rentals, Insurance payments, RSDI, VA benefits, Contract for deed, Trust income, Interest, Dividends, Worker's comp, Gambling winnings, Inheritance, Capital gains, etc.)`n"
            ;CaseNoteUnansweredMissing .= "Unearned income, "
            ;AnsweredYes := " yes"
        ;}
        ;If LumpSumUnansweredMissing {
            ;UnansweredText .= "Have you received any lump sums in the past 12 months?`n"
            ;CaseNoteUnansweredMissing .= "Lump sum, "
            ;AnsweredYes := " yes"
        ;}
        ;If EmploymentUnansweredMissing {
            ;UnansweredText .= "Is anyone in your household employed?`n"
            ;CaseNoteUnansweredMissing .= "Employment, "
            ;AnsweredYes := " yes"
        ;}
        ;If SelfEmploymentUnansweredMissing {
            ;UnansweredText .= "Is anyone in your household self-employed?`n"
            ;CaseNoteUnansweredMissing .= "Self-employment, "
            ;AnsweredYes := " yes"
        ;}
        ;If ExpensesUnansweredMissing {
            ;UnansweredText .= "Does anyone in your household pay Healthcare premiums, or child/spousal support?`n"
            ;CaseNoteUnansweredMissing .= "Self-employment, "
            ;AnsweredYes := " yes"
        ;}
        ;If AssetsUnansweredMissing {
            ;UnansweredText .= "Are your assets MORE THAN, or are they LESS THAN $1 million?`n"
            ;CaseNoteUnansweredMissing .= "Assets, "
            ;AnsweredMoreThan := " MORE THAN"
        ;}
        ;If ChildSchoolUnansweredMissing {
            ;UnansweredText .= "Are any of your children in school now or starting school in the next 12 months?`n"
            ;CaseNoteUnansweredMissing .= "Child school, "
            ;AnsweredYes := " yes"
        ;}
        ;If AdultSchoolUnansweredMissing {
            ;UnansweredText .= "Do any adults need child care assistance for going to school?`n"
            ;CaseNoteUnansweredMissing .= "Adult school, "
            ;AnsweredYes := " yes"
        ;}
        ;If JobSearchUnansweredMissing {
            ;UnansweredText .= "Do any adults need child care assistance for looking for work?`n"
            ;CaseNoteUnansweredMissing .= "Job search, "
            ;AnsweredYes := " yes"
        ;}
        ;If ActivityHistoryPRI1UnansweredMissing {
            ;UnansweredText .= "In the past 12 months has there been a period of more than 3 months where you (or the other parent in the household) did not work, go to school, or participate in activities as listed on an MFIP/DWP Employment Plan?"
            ;CaseNoteUnansweredMissing .= "Past 12-month activity history question, "
            ;AnsweredYes := " yes"
        ;}
        ;; need str replace CaseNoteUnansweredMissing , " -> ;" (last comma -> semi-colon)
        ;If (StrLen(AnsweredYes AnsweredMoreThan) = 14) {
            ;AnsweredBoth = " or"
        ;}
        ;UnansweredText .= "* Submit verification if you answered" AnsweredYes AnsweredBoth AnsweredMoreThan ". *`n"
        ;UnansweredText := st_wordwrap(UnansweredText, 59, " ")
			;UnansweredTextCount := 0
			;StrReplace(UnansweredText, "`n", "`n", UnansweredTextCount)
			;UnansweredTextCount++
            ;MissingVerifications[ListItem ". " UnansweredText] := UnansweredTextCount
            ;EmailTextString .= EmailListItem ". " UnansweredText
			;CaseNoteMissing .= CaseNoteUnansweredMissing ";`n"
			;ListItem++
            ;EmailListItem++
    ;}
;======================================================
	If EdBSFformMissing {
        EdBSFformMissingText := Ini.CaseNoteCountyInfo.CountyEdBSF " form (sent separately);`n"
		MissingVerifications[ListItem ". " EdBSFformMissingText] := 1
        EmailTextString .= EmailListItem ". " EdBSFformMissingText
		CaseNoteMissing .= Ini.CaseNoteCountyInfo.CountyEdBSF " form;`n"
		ListItem++
        EmailListItem++
    }
	If ClassScheduleMissing {
        ClassScheduleMissingText := "Class schedule with class start/end times and credits;`n"
		MissingVerifications[ListItem ". " ClassScheduleMissingText] := 1
        EmailTextString .= EmailListItem ". " ClassScheduleMissingText
		CaseNoteMissing .= "Adult class schedule;`n"
		ListItem++
        EmailListItem++
    }
	If TranscriptMissing {
        TranscriptMissingText := "Unofficial transcript/academic record;`n"
		MissingVerifications[ListItem ". " TranscriptMissingText] := 1
        EmailTextString .= EmailListItem ". " TranscriptMissingText
		CaseNoteMissing .= "Transcript;`n"
		ListItem++
        EmailListItem++
    }
	If EducationEmploymentPlanMissing {
        EducationEmploymentPlanMissingText := "Cash Assistance Employment Plan listing your education activity and schedule;`n"
		MissingVerifications[ListItem ". " EducationEmploymentPlanMissingText] := 2
        EmailTextString .= EmailListItem ". " EducationEmploymentPlanMissingText
		CaseNoteMissing .= "ES Plan with education activity and schedule;`n"
		ListItem++
        EmailListItem++
    }
    If StudentStatusOrIncomeMissing {
        StudentStatusOrIncomeMissingText := "Verification of your student status of being at least halftime, OR your most recent 30 days income.`n (if you are 19 or under and attending school at least`n   halftime, your income is not counted);`n"
		MissingVerifications[ListItem ". " StudentStatusOrIncomeMissingText] := 4
        EmailTextString .= EmailListItem ". " StudentStatusOrIncomeMissingText
		CaseNoteMissing .= "Halftime+ student status or income (PRI age 19 or under);`n"
		ListItem++
        EmailListItem++
    }
;-------------------------
	If JobSearchHoursMissing {
        JobSearchHoursMissingText := "Job search hours needed per week: Assistance can be approved for 1 to 20 hours of job search each week, limited to a total of 240 hours per calendar year;`n"
		MissingVerifications[ListItem ". " JobSearchHoursMissingText] := 3
        EmailTextString .= EmailListItem ". " JobSearchHoursMissingText
		CaseNoteMissing .= "Job search hours per week;`n"
		ListItem++
        EmailListItem++
    }
    If ESPlanUpdateMissing {
        ESPlanUpdateMissingText := "Updated Employment Plan ...;`n"
		MissingVerifications[ListItem ". " ESPlanUpdateMissingText] := 4
        EmailTextString .= EmailListItem ". " ESPlanUpdateMissingText
		CaseNoteMissing .= "Updated Employment Plan ...;`n"
		ListItem++
        EmailListItem++
    }
	While (A_Index < 4) { ; Other
		If Other%A_Index% {
			TextToPass := StrReplace(Other%A_Index%Input, "  ", "`n")
			CaseNoteMissing .= TextToPass ";`n"
            CountedRows := getRowCount(TextToPass, 57, " ")
			MissingVerifications[ListItem ". " CountedRows[1] ";`n"] := CountedRows[2]
            EmailTextString .= EmailListItem ". " TextToPass ";`n"
			ListItem++
            EmailListItem++
        }
    }
;======================================================
	If InHomeCareMissing {
        InHomeCareMissingText := "In-Home Care form (sent separately) - In-Home Care requires approval by MN DHS;`n"
		MissingVerifications[ListItem ". " InHomeCareMissingText] := 2
        EmailTextString .= EmailListItem ". " InHomeCareMissingText
		CaseNoteMissing .= "In-Home Care form;`n"
		ListItem++
        EmailListItem++
    }
	If LNLProviderMissing {
        LNLProviderMissingText := "Legal Non-Licensed Acknowledgement (sent separately).`n Your provider may not be eligible to be paid for care`n provided prior to completion of specific trainings;`n"
		MissingVerifications[ListItem ". " LNLProviderMissingText] := 3
        EmailTextString .= EmailListItem ". " LNLProviderMissingText
		CaseNoteMissing .= "LNL Acknowledgement form;`n"
		ListItem++
        EmailListItem++
    }
    If StartDateMissing {
        StartDateMissingText := "Start date at your child care provider;`n"
		MissingVerifications[ListItem ". " StartDateMissingText] := 1
        EmailTextString .= EmailListItem ". " StartDateMissingText
		CaseNoteMissing .= "Provider start date;`n"
		ListItem++
        EmailListItem++
    }
	If ChildSupportNoncooperationMissing {
        ChildSupportNoncooperationMissingText := "* You are currently in a non-cooperation status with Child Support. Contact Child Support at " MissingInput.ChildSupportNoncooperationMissing " for details. Child Support cooperation is a requirement for eligibility.`n"
		MissingVerifications[ChildSupportNoncooperationMissingText] := 3
        EmailTextString .= ChildSupportNoncooperationMissingText
		CaseNoteMissing .= "Cooperation status with Child Support, CS number: " MissingInput.ChildSupportNoncooperationMissing ";`n"
    }
	If EdBSFOneBachelorDegreeMissing {
        EdBSFOneBachelorDegreeMissingText := "* Unless listed on a Cash Assistance Employment Plan, education is an eligible activity only up to your first bachelor's degree, plus CEUs (no additional degrees).`n"
		MissingVerifications[EdBSFOneBachelorDegreeMissingText] := 3
        EmailTextString .= EdBSFOneBachelorDegreeMissingText
		CaseNoteMissing .= "* Client informed only up to first bachelor's degree is BSF/TY eligible;`n"
    }

    EligibleActivityWithJSText := "Eligible activities are:`n  A. Employment of 20+ hours per week (10+ for FT students)`n  B. Education with an approved plan`n  C. Job Search up to 20 hours per week`n  D. Activities on a Cash Assistance Employment Plan."
    EligibleActivityWithoutJSText := "Eligible activities are:`n  A. Employment of 20+ hours per week (10+ for FT students)`n  B. Education with an approved plan`n  C. Activities on a Cash Assistance Employment Plan."

    If SelfEmploymentIneligibleMissing {
        SelfEmploymentIneligibleMissingText := "* Your self-employment does not meet activity requirements. Self-employment hours are calculated using 50% of recent gross income, or gross minus expenses on tax return divided by minimum wage. " EligibleActivityWithJSText "`n"
		MissingVerifications[SelfEmploymentIneligibleMissingText] := 8
        EmailTextString .= SelfEmploymentIneligibleMissingText
		CaseNoteMissing .= "Self-employment hours meeting minimum requirement, or other eligible activity;`n"
    }
    If EligibleActivityMissing {
        EligibleActivityMissingText := "* You did not select an eligible activity on the " MEC2DocType ". " EligibleActivityWithJSText "`n"
		MissingVerifications[EligibleActivityMissingText] := 6
        EmailTextString .= EligibleActivityMissingText
		CaseNoteMissing .= "Eligible activity (none selected on form);`n"
    }
    If EmploymentIneligibleMissing {
        EmploymentIneligibleMissingText := "* Your employment does not meet eligible activity requirements. " EligibleActivityWithJSText "`nYou can submit up to 6 months of recent paystubs to average above 20 hours.`n"
		MissingVerifications[EmploymentIneligibleMissingText] := 8
        EmailTextString .= EmploymentIneligibleMissingText
		CaseNoteMissing .= "Employment hours meeting minimum requirement, or other eligible activity;`n"
    }
    If ESPlanOnlyJSMissing {
        ESPlanOnlyJSMissingText := "* While you have an Employment Plan, assistance hours cannot be approved for job search unless it is listed on the Plan"
		MissingVerifications[ESPlanOnlyJSMissingText ";`n"] := 2
        EmailTextString .= ESPlanOnlyJSMissingText ". Contact your Job Counselor to have an updated Plan written if job search hours are needed;`n"
		CaseNoteMissing .= "Client has ES Plan - informed JS hours are required to be on the Plan;`n"
    }
	If ActivityAfterHomelessMissing {
        ActivityAfterHomelessMissingText := "* At the end of the 90-day homeless exemption period, you must have an eligible activity to keep your Child Care Assistance case open. " EligibleActivityWithoutJSText "`n"
		MissingVerifications[ActivityAfterHomelessMissingText] := 6
        EmailTextString .= ActivityAfterHomelessMissingText
		CaseNoteMissing .= "Eligible activity after the 3-month homeless period;`n"
    }
	If NoProviderMissing {
        NoProviderMissingText := "* Once you have a daycare provider, please notify me with the provider’s name, location, and the start date.`n`n   If you need help locating a daycare provider, contact Parent Aware at 888-291-9811 or www.parentaware.org/search`n"
        EmailTextString .= NoProviderMissingText
		CaseNoteMissing .= "Provider;`n"
        MecCheckboxIds.providerInformation := 1
    }
    ;*   Provider Information- If you have a child care provider, send the provider's name, address and start date (if known). Visit www.parentaware.org for help finding a provider. Care is not approved until you get a Service Authorization.
	If UnregisteredProviderMissing {
        UnregisteredProviderMissingText := "* Your daycare provider is not registered with Child Care Assistance. Please have them call " Ini.CaseNoteCountyInfo.CountyProviderWorkerPhone " to register.`n"
		MissingVerifications[UnregisteredProviderMissingText] := 2
        EmailTextString .= UnregisteredProviderMissingText
		CaseNoteMissing .= "Registered provider;`n"
    }
    If ProviderForNonImmigrantMissing {
        ProviderForNonImmigrantMissingText := "* If your child is not a US citizen, Lawful Permanent Resident, Lawfully residing non-citizen, or fleeing persecution, assistance can only be approved at a daycare that is subject to public educational standards.`n"
        MissingVerifications[ProviderForNonImmigrantMissingText] := 4
        EmailTextString .= ProviderForNonImmigrantMissingText
        CaseNoteMissing .= "Provider subject to Public Educational Standards (4.15), if child not citizen/immigrant;`n"
    }
    CaseDetails.HaveWaitlist := ( InStr(CaseDetails.CaseType, "BSF") && CaseDetails.Eligibility == "ineligible" && Ini.CaseNoteCountyInfo.Waitlist > 1 )
    If (!CaseDetails.HaveWaitlist) {
        FaxAndEmailWrapped := FaxAndEmailText()
        FaxAndEmailWrapped := getRowCount(FaxAndEmailWrapped, 60, " ")
        AutoDeny := getRowCount(AutoDenyObject.AutoDenyExtensionSpecLetter, 60, "")
        ClarifiedVerifications[ "NewLineAutoreplace" FaxAndEmailWrapped[1] "`nNewLineAutoreplace" AutoDeny[1] ] := FaxAndEmailWrapped[2]+AutoDeny[2]
        EmailTextString .= AutoDeny[1] 
    }

    MecCheckboxIds.other := 1
    IdList := ""
    For key, value in MecCheckboxIds {
        If StrLen(IdList) > 1
            IdList .= ","
        IdList .= key
    }
    
    InsertAtOffset := (CaseDetails.Eligibility = "pends" && Homeless) ? 2 : 0
    If ( !OverIncomeMissing && !CaseDetails.HaveWaitlist && !ManualWaitlistBox && MissingVerifications.Length() > (0 + InsertAtOffset) ) {
        If (StrLen(IdList) > 5 || InsertAtOffset = 2) { ; "other" will always add at least 5
            MissingVerifications.InsertAt(1 + InsertAtOffset, "__In addition to the above, please submit following items:__`n", 1)
        }
        Else If (StrLen(IdList) = 5) {
            MissingVerifications.InsertAt(1+ InsertAtOffset, "_____________Please submit the following items:_____________`n", 1)
        }
    }
    WaitlistText := ""
    If (CaseDetails.HaveWaitlist || ManualWaitlistBox) {
        WaitlistNumber := Ini.CaseNoteCountyInfo.Waitlist -1
        WaitlistPriorities := { 1: "• Are attending High School, GED, or ESL classes;`n" }
        WaitlistPriorities.2 := WaitlistPriorities.1 "• Families in which an applicant is a veteran;`n"
        WaitlistPriorities.3 := WaitlistPriorities.2 "• Families which don't qualify for other priorities;`n"
        WaitlistText := "
(
Due to limited funding, new eligibility for CCAP in " CountySpecificText[Ini.EmployeeInfo.EmployeeCounty].CountyName " is currently limited to those who:
• Received Cash Assistance (MFIP/DWP) within the past year;
• Are applying for and are approved for Cash Assistance;
" WaitlistPriorities[WaitlistNumber] "`n
)"
        EmailText.Waitlist := "`n`n" WaitlistText "`n" (ManualWaitlistBox ? SubmitOnlyCommentItemsText "`n" : "")
        SubmitOnlyCommentItemsText := "If you meet one of the above criteria, please submit the following items:`n", SubmitCommentAndCheckboxItemsText := "If you meet one of the above criteria, in addition to items above the Worker Comments, please submit the following:`n"
        If (ManualWaitlistBox) {
            WaitlistText .= StrLen(IdList) == 5 ? SubmitOnlyCommentItemsText : SubmitCommentAndCheckboxItemsText
        }
        WaitlistText := getRowCount(WaitlistText, 60, "")
        MissingVerifications.InsertAt(1, WaitlistText[1] "`n", WaitlistText[2])
        CaseNoteMissing .= "Approved MFIP/DWP or meet current Waitlist criteria;`n"
    }
    If (ClarifiedVerifications.Length() > 1) {
        ClarifiedVerifications.InsertAt(1, "__Clarification of items listed above the Worker Comments:__`n", 1)
    }
    GoSub SetEmailText
    ArrayLines := 0
    VerifCat := "Missing"
    GoSub ListifyMissing
    ArrayLines := CountLines(ClarifiedVerifications)
    VerifCat := "Clarified"
    GoSub ListifyMissing
	CaseNoteMissing := SubStr(CaseNoteMissing, 1, -1)
	While LetterText%A_Index% {
        If (InStr(LetterText%A_Index%, "__Clarification",,2)) {
            StrReplace(st_wordWrap(LetterText%A_Index%, 60, ""), "`n", "`n", LetterLineCount)
            If (LetterLineCount < 27) {
                LetterText%A_Index% := StrReplace(LetterText%A_Index%, "__Clarification", "`n__Clarification")
            }
        }
		TempVar := "Letter" . A_Index
		GuiControl, MissingGui:Show, %TempVar%
    }
	GuiControl, MainGui:Text, Missing, %CaseNoteMissing%
	GuiControl, MissingGui:Show, Email
	WinActivate, CaseNotes
Return

IncrementLetterPage:
    LetterTextNumber++
    LetterText[LetterTextNumber] .= "                   Continued on letter " LetterTextNumber
    LetterText[LetterTextNumber] .= "                  Continued from letter " LetterTextNumber-1 "`n"
Return

FaxAndEmailText() {
    FaxInfo := (StrLen(Ini.CaseNoteCountyInfo.CountyFax) > 1) ? "faxed to " Ini.CaseNoteCountyInfo.CountyFax : ""
    EmailInfo := (StrLen(Ini.CaseNoteCountyInfo.CountyDocsEmail) > 1) ? "emailed to " Ini.CaseNoteCountyInfo.CountyDocsEmail : ""
    FaxAndEmail := (StrLen(FaxInfo) > 1 && StrLen(EmailInfo) > 1) ? " and " : ""
    Return ((StrLen(FaxInfo) > 1 || StrLen(EmailInfo) > 1))
    ? " Documents can also be " FaxInfo . FaxAndEmail . EmailInfo ". Please include your case number." : ""
}

CountLines(VerificationArray) {
    TotalLines := 0
    For key, value in VerificationArray {
        TotalLines := TotalLines + value
    }
    Return TotalLines
}

ListifyMissing:
    VerificationList := (VerifCat = "Clarified") ? ClarifiedVerifications : MissingVerifications
    If ((LineCount + ArrayLines) > 30) { ; puts ClarifiedVerifications on the next letter if it will exceed the current letter's available space
        LetterTextNumber++
        LetterTextPassed[LetterTextNumber] .= "                   Continued on letter " LetterTextNumber
        LetterTextPassed[LetterTextNumber] .= "                  Continued from letter " LetterTextNumber-1 "`n"
        
        LetterNumber++
        %LetterTextVar% .= "                   Continued on letter " LetterNumber
        LetterTextVar := "LetterText" . LetterNumber
        LineCount := 1 ; For the continued from line
        %LetterTextVar% .= "                  Continued from letter " LetterNumber-1 "`n"
    }
    For key, value in VerificationList {
        If (InStr(key, "NewLineAutoreplace")) { ; last key in group
            LineCountPlusFaxed := LineCount + value
            If (LineCountPlusFaxed = 30) {
                key := StrReplace(key, "NewLineAutoreplace", "")
                key := StrReplace(key, "NewLineAutoreplace", "")
            } Else If (LineCountPlusFaxed > 30) {
                key := StrReplace(key, "NewLineAutoreplace", "`n")
                key := StrReplace(key, "NewLineAutoreplace", "`n")
                
                    LetterTextNumber++
                    LetterText[LetterTextNumber] .= "                   Continued on letter " LetterTextNumber
                    LetterText[LetterTextNumber] .= "                  Continued from letter " LetterTextNumber-1 "`n"
                    
                    LetterNumber++
                    %LetterTextVar% .= "                   Continued on letter " LetterNumber
                    LetterTextVar := "LetterText" . LetterNumber
                    %LetterTextVar% .= "                  Continued from letter " LetterNumber-1 "`n"
            } Else {
                While (InStr(key, "NewLineAutoreplace") && LineCountPlusFaxed < 30) {
                    key := StrReplace(key, "NewLineAutoreplace", "`n",,1)
                    LineCountPlusFaxed++
                }
                key := StrReplace(key, "NewLineAutoreplace", "")
            }
            %LetterTextVar% .= key
        } Else { ; does not contain "NewLineAutoreplace"
            If ((LineCount + value) > 29) {
            
                    LetterTextNumber++
                    LetterText[LetterTextNumber] .= "                   Continued on letter " LetterTextNumber
                    LetterText[LetterTextNumber] .= "                  Continued from letter " LetterTextNumber-1 "`n"
                    
                    LetterNumber++
                    %LetterTextVar% .= "                   Continued on letter " LetterNumber
                    LetterTextVar := "LetterText" . LetterNumber
                    %LetterTextVar% .= "                  Continued from letter " LetterNumber-1 "`n"
                    %LetterTextVar% .= key
                    LineCount := (value + 1) ; For the continued from line
            } Else {
                LineCount += value
                %LetterTextVar% .= key
            }
        }
    }
Return

Email:
    Clipboard := EmailText.Output
    WinActivate, Message - 
Return

SetEmailText:
	Gui, MainGui: Submit, NoHide
	Gui, MissingGui: Submit, NoHide
    EmailTextString := StrReplace(EmailTextString, "`n ", " ")
    EmailTextString := StrReplace(EmailTextString, "    ", " ")
    EmailTextString := StrReplace(EmailTextString, "   ", " ")
    EmailTextString := StrReplace(EmailTextString, "`n*", "`n`n*")
    EmailTextString := StrReplace(EmailTextString, "sent separately", "see attached")
	;EmailText.Output := (Homeless = 1) ? EmailText.Combined EmailTextString EmailText.EndHL : EmailText.Combined EmailTextString
	EmailText.Output := EmailText.Combined EmailTextString EmailText.EndHL
Return

Letter:
	Gui, MainGui: Submit, NoHide
	Gui, MissingGui: Submit, NoHide
    If (Homeless = 1 && CaseDetails.Eligibility = "pends" && StrLen(MissingHomelessItems) < 1) {
        GoSub MissingButtonDoneButton
    }
    WinActivate % Ini.EmployeeInfo.EmployeeBrowser
    Sleep 500
	LetterGUINumber := "LetterText" SubStr(A_GuiControl, 0)
    If (Ini.EmployeeInfo.EmployeeUseMec2Functions = 1) {
        CaseStatus := InStr(CaseDetails.DocType, "?") ? "" : (CaseDetails.DocType = "Redet") ? "Redetermination" : (Homeless = 1) ? "Homeless App" : CaseDetails.DocType
        jsonLetterText := "LetterTextFromAHKJSON{""LetterText"":""" JSONstring(%LetterGUINumber%) """,""CaseStatus"":""" CaseStatus """,""IdList"":""" IdList """ }"
        Clipboard := jsonLetterText
        Send, ^v
    } Else {
        Clipboard := %LetterGUINumber%
        Send, ^v
    }
    Sleep 500
    Clipboard := CaseNumber
Return

Other:
    OtherName := A_GuiControl
    Gui, Submit, NoHide
	If (%A_GuiControl% == 0)
		Return
    Gui, OtherGui: New,, Other Verification
    Gui, OtherGui: Margin, 12 12
    Gui, OtherGui: Font, s9, Lucida Console
    Gui, OtherGui: Add, Text, w525, Additional Input Required: State what the client needs to submit.
    Gui, OtherGui: Add, Edit, % "v" OtherName "Input h100 " TextboxSettings, % %OtherName%Input
    Gui, OtherGui: Add, Button, gSaveOther, Save
    Gui, OtherGui: Show
    Gui, OtherGui:+OwnerMissingGui
Return
SaveOther:
    Gui, Submit, NoHide
    GuiControl, MissingGui:, % OtherName, % %OtherName%Input
    Gui, OtherGui: Destroy
Return
OtherGuiGuiClose:
    GuiControl, MissingGui:, % OtherName, 0
    Gui, OtherGui: Destroy
Return

InputBoxAGUIControl:
    Gui, Submit, NoHide
    Gui +OwnDialogs
    PromptText := ""
    If (%A_GuiControl% == 0) ; unchecked
        Return

    ;v2: convert to switch:
    If (A_GuiControl = "IDmissing")
        PromptText := "Who is ID needed for?`n`nExample: 'Susanne, Robert Sr'"
    Else If (A_GuiControl = "BCmissing")
        PromptText := "Who is birth verification needed for?`n`nExample: 'Susie, Bobby Jr'"
    Else If (A_GuiControl = "IncomePlusNameMissing")
        PromptText := "Who is the income verification needed for?"
    Else If (A_GuiControl = "CustodySchedulePlusNamesMissing")
        PromptText := "Who is the schedule needed for? `n'...stating the current parenting time schedule for: ____________'`n`nExample: 'Susie and Bobby Jr' or 'your children'"
    Else If (A_GuiControl = "WorkSchedulePlusNameMissing")
        PromptText := "Who is the work schedule needed for?"
    Else If (A_GuiControl = "DependentAdultStudentMissing")
        PromptText := "Who is the adult dependent student?"
    Else If (A_GuiControl = "ChildSupportFormsMissing")
        PromptText := "Enter the number of sets of Child Support forms needed`nor the names of the absent parent/children.`n`nExample: 'Robert Sr / Susie, Bobby Jr' or '2'"
    Else If (A_GuiControl = "ChildSupportNoncooperationMissing")
        PromptText := "What is the phone number of the Child Support officer?"
    Else If (A_GuiControl = "LegalNameChangeMissing")
        PromptText := "Who is the name change proof needed for?"
    Else If (A_GuiControl = "SeasonalOffSeasonMissing")
        PromptText := "Who is the employer? (optional)"
    Else if (A_GuiControl = "OverIncomeMissing")
        PromptText := "Without dollar signs, enter the calculated income less expenses, income limit, and household size.`nOnly type numbers separated by spaces - no commas or periods.`n`n(Example: 76392 49605 3)"

    InputBox, InputBoxInput, Additional Input Required, % PromptText, , , , , , , ,% MissingInput[A_GuiControl]
    MissingInput[A_GuiControl] := InputBoxInput
	If ErrorLevel {
        GuiControl, MissingGui:, % A_GuiControl, 0
		Return
    }
    If (StrLen(MissingInput[A_GuiControl]) == 0) {
        MissingInput[A_GuiControl] := "(input)"
        GuiControl, MissingGui:, % A_GuiControl, 0 ; uncheck box if input is blank
    }
    GuiControlGet, VerificationName,,%A_GuiControl%, Text ; VerificationName := %A_GuiControl% text value
    ;v2: convert to switch:
    If (InStr(VerificationName, "(")) {
        VerificationName := SubStr(VerificationName, 1, InStr(VerificationName, " (") -1)
    }
    Else If (InStr(VerificationName, " for ")) {
        VerificationName := SubStr(VerificationName, 1, InStr(VerificationName, " for ") -1)
    }
    Else If (InStr(VerificationName, " at ")) {
        VerificationName := SubStr(VerificationName, 1, InStr(VerificationName, " at ") -1)
    }

    If (A_GuiControl = "ChildSupportNoncooperationMissing")
        GuiControl,,%A_GuiControl%, % VerificationName " - CS phone: " MissingInput[A_GuiControl]
    Else If (A_GuiControl = "ChildSupportFormsMissing") {
        SetWording :=
        If (StrLen(MissingInput[A_GuiControl]) == 1) {
            SetWording .= MissingInput[A_GuiControl] < 2 ? " set" : " sets"
        } else {
            SetWording := ""
        }
        GuiControl,, %A_GuiControl%, % "Child Support forms - " MissingInput[A_GuiControl] SetWording
    }
    Else if (A_GuiControl = "OverIncomeMissing")
        OverIncomeSub(MissingInput[A_GuiControl])
    Else 
        GuiControl,, %A_GuiControl%, % VerificationName " for " MissingInput[A_GuiControl]
Return

OverIncomeSub(OverIncomeString) {
    overIncomeEntriesArray := StrSplit(OverIncomeString, A_Space, ",", -1)
    If (StrLen(overIncomeEntriesArray[3]) > 0) {
        OverIncomeObj.overIncomeHHsize := overIncomeEntriesArray[3]
    }
    OverIncomeObj.overIncomeReceived := Round(StrReplace(overIncomeEntriesArray[1], ","))
    OverIncomeObj.overIncomeLimit := StrReplace(overIncomeEntriesArray[2], ",")
    OverIncomeObj.overIncomeText := OverIncomeObj.overIncomeLimit ", your income is calculated as $" OverIncomeObj.overIncomeReceived
    OverIncomeObj.overIncomeDifference := OverIncomeObj.overIncomeReceived - OverIncomeObj.overIncomeLimit
    GuiControl,, %A_GuiControl%, % "Over-income by $" OverIncomeObj.overIncomeDifference
}
;=============================================================================================================================
;VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION 
;=============================================================================================================================



;=========================================================================================================================================================
;ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION
;=========================================================================================================================================================

CBTGuiClose:
    WinGetPos, XClipboardGet, YClipboardGet,,, Clipboard_Text
    If (XClipboardGet == "")
        Return
    If ((XClipboardGet - Ini.CBTPositions.XClipboard + YClipboardGet - Ini.CBTPositions.YClipboard) != 0) {
        CoordObjOut := {}
        For Key, Value in ["XClipboard", "YClipboard"] {
            CoordObjOut[Value] := %Value%Get
            Ini.CBTPositions[Value] := %Value%Get
        }
        coordString := CheckCoordValues(CoordObjOut)
        IniWrite, %coordString%, %A_MyDocuments%\AHK.ini, CBTPositions
    }
    Gui, CBT: Destroy
Return

MissingGuiGuiClose:
    WinGetPos, XVerificationGet, YVerificationGet,,, A
    For Key, Value in ["XVerification", "YVerification"] {
        Ini.CaseNotePositions[Value] := %Value%Get
    }
	Gui, MissingGui: Hide
Return

MainGuiGuiClose:
    GoSub, ClearFormButton
Return

ClearFormButton:
    ClosingPromptText := ""
    ClearingForm := A_GuiControl == "ClearFormButton" ? 1 : 0
    If (ConfirmedClear > 0) {
        SaveCoordsAndRun(1)
    }
    ClosingPromptText .= CaseNoteEntered.MEC2Note == 0 ? " MEC2" : ""
    If (CaseNoteEntered.MAXISNote == 0 && Ini.CaseNoteCountyInfo.CountyNoteInMaxis == 1 && CaseDetails.DocType == "Application") {
        ClosingPromptText .= StrLen(ClosingPromptText) > 0 ? " or MAXIS" : " MAXIS"
    }
    If (ClearingForm) {
        If (StrLen(ClosingPromptText) > 0) {
            MsgBox, 4, Case Note Prompt, % "Case note not entered in" ClosingPromptText ". `nClear form anyway?"
            IfMsgBox Yes
                SaveCoordsAndRun(1)
            Return
        }
        GuiControl, MainGui: Text, ClearFormButton, Confirm
        Gui, Font, s9, Segoe UI
        GuiControl, MainGui: Font, ClearFormButton
        ConfirmedClear++
    }
    If (!ClearingForm) {
        If (StrLen(ClosingPromptText) > 0) {
            MsgBox, 4, Case Note Prompt, % "Case note not entered in" ClosingPromptText ". `nExit anyway?"
            IfMsgBox Yes
                SaveCoordsAndRun()
            Return
        }
        Else If (StrLen(ClosingPromptText) == 0) {
            SaveCoordsAndRun()
        }
    }
Return

SaveCoordsAndRun(ReOpen := 0) {
	WinGetPos, XCaseNotesGet, YCaseNotesGet,,, CaseNotes
	WinGetPos, XVerificationGet, YVerificationGet,,, Missing Verifications
    If (XVerificationGet == "") {
        XVerificationGet := Ini.CaseNotePositions.XVerification, YVerificationGet := Ini.CaseNotePositions.YVerification
    }
    If ((XCaseNotesGet - Ini.CaseNotePositions.XCaseNotes + YCaseNotesGet - Ini.CaseNotePositions.YCaseNotes + XVerificationGet - Ini.CaseNotePositions.XVerification + YVerificationGet - Ini.CaseNotePositions.YVerification) != 0) {
        CoordObjOut := {}
        For Key, Value in ["XVerification", "YVerification", "XCaseNotes", "YCaseNotes"] {
            CoordObjOut[Value] := %Value%Get
        }
        coordString := CheckCoordValues(CoordObjOut)
        IniWrite, %coordString%, %A_MyDocuments%\AHK.ini, CaseNotePositions
    }
    If (ReOpen == 1) {
        Run %A_ScriptName%
    } Else {
        ExitApp
    }
}

CheckCoordValues(CoordObjIn) {
    coordStringReturn :=
    For Key, Value in CoordObjIn {
        coordStringReturn .= Key "=" (Abs(Value) < 9999 && Value != "" ? Value : 0) "`n"
    }
    Return coordStringReturn
}

;=====================================================================================================================================================
;SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION
;=====================================================================================================================================================
SettingsButton:
    CountyContact := {}
    CountyContact.Dakota := { Email: "EEADOCS@co.dakota.mn.us", Fax: "651-306-3187", ProviderWorker: "651-554-5764", EdBSF: "Training Request for Childcare", CountyNoteInMaxis: 1 }
    CountyContact.StLouis := { Email: "ess@stlouiscountymn.gov", Fax: "218-733-2976", ProviderWorker: "218-726-2064", EdBSF: "SLC CCAP Education Plan", CountyNoteInMaxis: 0 }
    Ini.CaseNoteCountyInfo.CountyFax := Ini.CaseNoteCountyInfo.CountyFax != " " ? Ini.CaseNoteCountyInfo.CountyFax : CountyContact[Ini.EmployeeInfo.EmployeeCounty].Fax
    Ini.CaseNoteCountyInfo.CountyDocsEmail := Ini.CaseNoteCountyInfo.CountyDocsEmail != " " ? Ini.CaseNoteCountyInfo.CountyDocsEmail : CountyContact[Ini.EmployeeInfo.EmployeeCounty].Email
    Ini.CaseNoteCountyInfo.CountyProviderWorkerPhone := Ini.CaseNoteCountyInfo.CountyProviderWorkerPhone != " " ? Ini.CaseNoteCountyInfo.CountyProviderWorkerPhone : CountyContact[Ini.EmployeeInfo.EmployeeCounty].ProviderWorker
    Ini.CaseNoteCountyInfo.CountyEdBSF := Ini.CaseNoteCountyInfo.CountyEdBSF != " " ? Ini.CaseNoteCountyInfo.CountyEdBSF : CountyContact[Ini.EmployeeInfo.EmployeeCounty].EdBSF
    
    EditboxOptions := "x200 yp-3 h18 w200 "
    CheckboxOptions := "x200 yp-3 h18 w20 "
    TextLabelOptions := "xm w170 h18 Right"
    Gui, Font,, Lucida Console
    Gui, Color, 989898, a9a9a9
    Gui, CSG: Margin, 12 12
    Gui, CSG: New, AlwaysOnTop ToolWindow,
    Gui, CSG: Add, Text, % TextLabelOptions " y12",Worker Name:
    Gui, CSG: Add, Edit, % EditboxOptions "vEmployeeNameWrite", % Ini.EmployeeInfo.EmployeeName
    Gui, CSG: Add, Text, % TextLabelOptions, Worker Phone:
    Gui, CSG: Add, Edit, % EditboxOptions "vEmployeePhoneWrite", % Ini.EmployeeInfo.EmployeePhone
    Gui, CSG: Add, Text, % TextLabelOptions, Worker Email:
    Gui, CSG: Add, Edit, % EditboxOptions "vEmployeeEmailWrite", % Ini.EmployeeInfo.EmployeeEmail
    Gui, CSG: Add, Text, % TextLabelOptions, Use Worker Email in Letters:
    Gui, CSG: Add, CheckBox, % "vEmployeeUseEmailWrite " CheckboxOptions "Checked" Ini.EmployeeInfo.EmployeeUseEmail
    Gui, CSG: Add, Text, % TextLabelOptions,  Using mec2functions:
    Gui, CSG: Add, CheckBox, % "vEmployeeUseMec2FunctionsWrite gWorkerUsingMec2Functions " CheckboxOptions "Checked" Ini.EmployeeInfo.EmployeeUseMec2Functions
    Gui, CSG: Add, ComboBox, x+10 yp vEmployeeBrowserWrite Choose1 R4 Hidden, % Ini.EmployeeInfo.EmployeeBrowser "|Google Chrome|Mozilla Firefox|Microsoft Edge"
    If (Ini.EmployeeInfo.EmployeeUseMec2Functions = 1) {
        GuiControl, CSG: Show, EmployeeBrowserWrite
    }
    Gui, CSG: Add, Text, h0 w0 y+10
    Gui, CSG: Add, Text, % TextLabelOptions,  Select a county to auto-populate
    Gui, CSG: Add, ComboBox, % EditboxOptions "vEmployeeCountyWrite gCountySelection Choose1 R4", % Ini.EmployeeInfo.EmployeeCounty "|Dakota|StLouis|Not Listed"
    Gui, CSG: Add, Text, % TextLabelOptions, Case Note in MAXIS:
    Gui, CSG: Add, CheckBox, % "vCountyNoteInMaxisWrite gCountyNoteInMaxis " CheckboxOptions "Checked" Ini.CaseNoteCountyInfo.CountyNoteInMaxis
    Gui, CSG: Add, Edit, x+10 yp h18 w170 vEmployeeMaxisWrite Hidden, % Ini.EmployeeInfo.EmployeeMaxis
    If (Ini.CaseNoteCountyInfo.CountyNoteInMaxis = 1) {
        GuiControl, CSG: Show, EmployeeMaxisWrite
    }
    Gui, CSG: Add, Text, % TextLabelOptions, Fax Number:
    Gui, CSG: Add, Edit, % EditboxOptions "vCountyFaxWrite", % Ini.CaseNoteCountyInfo.CountyFax
    Gui, CSG: Add, Text, % TextLabelOptions, County Documents Email:
    Gui, CSG: Add, Edit, % EditboxOptions "vCountyDocsEmailWrite", % Ini.CaseNoteCountyInfo.CountyDocsEmail
    Gui, CSG: Add, Text, % TextLabelOptions, Provider Worker Phone:
    Gui, CSG: Add, Edit, % EditboxOptions "vCountyProviderWorkerPhoneWrite", % Ini.CaseNoteCountyInfo.CountyProviderWorkerPhone
    Gui, CSG: Add, Text, % TextLabelOptions, BSF Education Form Name:
    Gui, CSG: Add, Edit, % EditboxOptions "vCountyEdBSFWrite", % Ini.CaseNoteCountyInfo.CountyEdBSF
    Gui, CSG: Add, Text, % TextLabelOptions, Waiting List Priority
    Gui, CSG: Add, DropDownList, % EditboxOptions "vWaitlistWrite R4 AltSubmit Choose" Ini.CaseNoteCountyInfo.Waitlist, None|HS / GED / ESL|A PRI is a veteran|All others
    Gui, CSG: Add, Button, w80 gUpdateIniFile, Save
    Gui, CSG:+OwnerMainGui
    Gui, CSG: Show,w450,Update CaseNotes Settings
Return

WorkerUsingMec2Functions:
    Gui, CSG: Submit, NoHide
    GuiControlGet, WorkerBrowser,,EmployeeUseMec2FunctionsWrite
    If (WorkerBrowser = 0) {
        GuiControl, CSG: Hide, EmployeeBrowserWrite
        Return
    }
    GuiControl, CSG: Show, EmployeeBrowserWrite
Return

CountyNoteInMaxis:
    Gui, CSG: Submit, NoHide
    GuiControlGet, MaxisChecked,,CountyNoteInMaxisWrite
    If (MaxisChecked = 0) {
        GuiControl, CSG: Hide, EmployeeMaxisWrite
        Return
    }
    GuiControl, CSG: Show, EmployeeMaxisWrite
Return

UpdateIniFile:
    Gui, CSG: Submit, NoHide
    ;If (CountyNoteInMaxisWrite && EmployeeMaxisWrite == "MAXIS-WINDOW-TITLE") {
        ;change border of EmployeeMaxisWrite, blink, dance, return?
    ;}
    SettingsArrays := { EmployeeInfo: [ "EmployeeName", "EmployeePhone", "EmployeeEmail", "EmployeeUseEmail", "EmployeeUseMec2Functions", "EmployeeBrowser", "EmployeeCounty", "EmployeeMaxis" ]
    , CaseNoteCountyInfo: [ "CountyFax", "CountyDocsEmail", "CountyProviderWorkerPhone", "CountyEdBSF", "CountyNoteInMaxis", "Waitlist" ] }
    For Key, Value in SettingsArrays {
        UpdateIniFileText(Key, Value)
    }
    CheckGroupAdd()
    Gui, Destroy
Return

UpdateIniFileText(Section, IniArray) {
    IniSettingsValues := ""
    For Key, Value in IniArray {
        IniSettingsValues .= Value "=" %Value%Write "`n"
        Ini[Section][Value] := %Value%Write
    }
    IniWrite, % IniSettingsValues, % A_MyDocuments "\AHK.ini", % Section
}
CheckGroupAdd() {
    If (Ini.EmployeeInfo.EmployeeBrowser != "")
        GroupAdd, Browser, % Ini.EmployeeInfo.EmployeeBrowser
    If (Ini.EmployeeInfo.EmployeeMaxis != "MAXIS-WINDOW-TITLE")
        GroupAdd, Maxis, % Ini.EmployeeInfo.EmployeeMaxis
}

CountySelection:
    Gui, CSG: Submit, NoHide
    Ini.EmployeeInfo.EmployeeCounty := EmployeeCountyWrite
    GuiControl, CSG: Text, CountyFaxWrite, % CountyContact[Ini.EmployeeInfo.EmployeeCounty].Fax
    GuiControl, CSG: Text, CountyProviderWorkerPhoneWrite, % CountyContact[Ini.EmployeeInfo.EmployeeCounty].ProviderWorker
    GuiControl, CSG: Text, CountyDocsEmailWrite, % CountyContact[Ini.EmployeeInfo.EmployeeCounty].Email
    GuiControl, CSG: Text, CountyEdBSFWrite, % CountyContact[Ini.EmployeeInfo.EmployeeCounty].EdBSF
    GuiControl,, CountyNoteInMaxisWrite, % CountyContact[Ini.EmployeeInfo.EmployeeCounty].CountyNoteInMaxis
    GoSub, CountyNoteInMaxis
Return

getRowCount(Text, columns, indentString) {
    indentString := StrLen(indentString) > 0 ? indentString : ""
    Text := st_wordwrap(Text, columns, indentString)
    StrReplace(Text, "`n", "`n", xCount)
    Return [Text, xCount +1]
}
;================================================================================================================================================================
;BORROWED FUNCTIONS SECTION BORROWED FUNCTIONS SECTION BORROWED FUNCTIONS SECTION BORROWED FUNCTIONS SECTION 
st_wordWrap(string, column, indentChar) { ; String Things - Common String & Array Functions
    If (StrLen(string) < 1)
        Return
    indentLength := StrLen(indentChar)
    Loop, Parse, string, `n, `r
    {
        If (StrLen(A_LoopField) > column) {
            pos := 0
            Loop, Parse, A_LoopField, %A_Space% ; A_LoopField is the individual word
            {
            CombinedLen := pos + (loopLength := StrLen(A_LoopField))
                If (CombinedLen <= column) {
                    out .= A_LoopField (CombinedLen < column ? " " : "")
                    , pos += (loopLength + 1) ; += word + space
                } Else {
                    pos := (indentLength + loopLength + 1) ; := indent + word + space
                    , out .= "`n" indentChar A_LoopField (NewLen < column ? " " : "")
                }
            }
            out .= "`n"
        } Else {
            If (StrLen(A_LoopField) > 0) {
                out .= indentChar A_LoopField "`n"
            }
        }
    }
    Return SubStr(RegExReplace(out, " ", "", , 1, -1), 1, -1)
}

Class OrderedAssociativeArray { ; Capt Odin https://www.autohotkey.com/boards/viewtopic.php?t=37083
	__New() {
		ObjRawSet(this, "__Data", {})
		ObjRawSet(this, "__Order", [])
	}
	__Get(args*) {
		return this.__Data[args*]
	}
	__Set(args*) {
		key := args[1]
		val := args.Pop()
		if(args.Length() < 2 && this.__Data.HasKey(key)) {
			this.Delete(key)
		}
		if(!this.__Data.HasKey(key)) {
			this.__Order.Push(key)
		}
		this.__Data[args*] := val
		return val
	}
	InsertAt(pos, key, val) {
		this.__Order.InsertAt(pos, key)
		this.__Data[key] := val
	}
	RemoveAt(pos) {
		val := this.__Data[this.__Order[pos]]
		this.__Data.Delete(this.__Order[pos])
		this.__Order.RemoveAt(pos)
		return val
	}
	Delete(key) {
		for i, v in this.__Order {
			if(key == v) {
				return this.RemoveAt(i)
			}
		}
	}
	Length() {
		return this.__Order.Length()
	}
	HasKey(key) {
		return this.__Data.HasKey(key)
	}
	_NewEnum() {
		return new OrderedAssociativeArray.Enum(this.__Data, this.__Order)
	}
	Class Enum {
		__New(Data, Order) {
			this.Data := Data
			this.oEnum := Order._NewEnum()
		}
		Next(ByRef key, ByRef val := "") {
			res := this.oEnum.next(, key)
			val := this.Data[key]
			return res
		}
	}
}
;BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION
;===========================================================================================================================================================================
;=================================================================================================================================================================================
;HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION - HOTKEYS SECTION
^+t::
Return

!3::
	Gui, MainGui: Submit, NoHide
    Clipboard := CaseNumber
Return

#m::
    ;Global EmailText, Ini, CaseDetails, CaseNumber
    If WinActive("Message" ahk_exe Outlook.exe) {
        Clipboard := EmailText.Output
        Send, ^v
    } Else If WinActive(Ini.EmployeeInfo.EmployeeBrowser) {
        Gui, MainGui: Submit, NoHide
        Gui, MissingGui: Submit, NoHide
        Sleep 500
        If (Ini.EmployeeInfo.EmployeeUseMec2Functions = 1) {
            CaseStatus := InStr(CaseDetails.DocType, "?") ? "" : (Homeless = 1) ? "Homeless App" : (CaseDetails.DocType = "Redet") ? "Redetermination" : CaseDetails.DocType
            jsonLetterText := JSONstring("LetterTextFromAHKJSON{""LetterText"":""" LetterText1 """,""CaseStatus"":""" CaseStatus """,""IdList"":""" IdList """ }")
            Clipboard := jsonLetterText
            Send, ^v
        } Else {
            Clipboard := LetterText1
            Send, ^v
        }
    }
    Sleep 500
    Clipboard := CaseNumber
Return

RemoveToolTip:
    ToolTip
return

;Shows Clipboard text in a AHK GUI
!^a::
    If WinExist("Clipboard_Text") {
        Gui, CBT: Destroy
    }
    Gui, CBT: New
    Gui, Color, Silver, C0C0C0
    Gui, Font, s11, Lucida Console
    Gui, CBT: Add, Edit, ReadOnly -VScroll vClipboardContents, %clipboard%
    GuiControl, CBT: font, ClipboardContents
    Gui, CBT: Show, % "x" Ini.CBTPositions.XClipboard " y" Ini.CBTPositions.YClipboard, Clipboard_Text
    ControlSend,,{End}, Clipboard_Text
Return

#IfWinActive CaseNotes
    PgDn::
        ControlFocus,,\d ahk_exe obunity.exe
        ControlSend,,^{PgDn}, \d ahk_exe obunity.exe
        Sleep 150
        WinActivate, CaseNotes
    Return
    PgUp::
        ControlFocus,,\d ahk_exe obunity.exe
        ControlSend,,^{PgUp}, \d ahk_exe obunity.exe
        Sleep 150
        WinActivate, CaseNotes
    Return

    #Left::
    #Right::
        MsgBox,4, Reset Position?, Do you want to reset CaseNotes' position?, 10
        IfMsgBox, Yes
            ResetPositions()
        IfMsgBox, Timeout
            ResetPositions()
        Return
    Return
#If
ResetPositions() {
    WinMove, CaseNotes,, 0, 0
    WinMove, Missing Verifications,,0,0
    XCaseNotes := 0
    YCaseNotes := 0
    XVerification := 0
    YVerification := 0
}

#IfWinActive ahk_group Browser
    ^F12:: ;CtrlF12/AltF12 Add worker signature
    !F12::
        SendInput `n=====`n
        Send, % Ini.EmployeeInfo.EmployeeName
    Return
#If
If (Ini.EmployeeInfo.EmployeeCounty = "Dakota") {
    #IfWinActive ahk_exe WINWORD.EXE
        F1::
            ToolTip,
            (
    Alt+4: Starting from the name field, moves to and enters date,
             case number, and client's first name.
            ), 0, 0
            SetTimer, RemoveToolTip, -5000
        Return
        !4::
            Gui, MainGui: Submit, NoHide
            ReceivedDate := FormatMDY(Received)
            RegExMatch(HouseholdComp, "^\w+\b", NameMatch)
            SendInput, {Down 2}
            Sleep 400
            SendInput, % ReceivedDate
            Sleep 400
            SendInput, {Up}
            Sleep 400
            SendInput, % CaseNumber
            Sleep 400
            SendInput, {Up}
            Sleep 400
            SendInput, % NameMatch " "
        Return
    #If

    OnBaseImportKeys(CaseNum, DocType, DetailText, DetailTabs=1, ToolTipHelp="") {
        SendInput, {Tab 2}
        Sleep 250
        SendInput, % DocType
        Sleep 1000
        SendInput, {Tab 4}
        Sleep 1500
        SendInput, NO
        Sleep 250
        SendInput, {Tab}
        Sleep 500
        SendInput, % CaseNum
        Sleep 500
        SendInput, {Tab %DetailTabs%}
        Sleep 750
        SendInput, % DetailText
        Sleep 200
        If (StrLen(ToolTipHelp) > 0) {
            CaretY := A_CaretY + 40
            ToolTip, % "`n  " ToolTipHelp "  `n ", % A_CaretX, % CaretY
            SetTimer, RemoveToolTip, -5000   
        }
    }

    ;Ctrl+: OnBase docs (OnBaseImportKeys("Text to get doc type", "Details Text", "Tab presses from Case # to details field")
    #IfWinActive Perform Import
        F1:: 
            ToolTip, CTRL+ `n F6: RSDI `n F7: SMI ID `n F8: PRISM GCSC `n F9: CS $ Calc `nF10: Income Calc `nF11: The Work # `nF12: CCAPP Letter, 0, 0
            SetTimer, RemoveToolTip, -8000
        Return
        ^F6::
            Gui, MainGui: Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 ssi", "{Text}RSDI ", 3, "Member#, Member Name")
        Return
        ^F7::
            Gui, MainGui: Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3001 other id", "{Text}SMI ", 3, "Member#, Member Name")
        Return
        ^F8::
            Gui, MainGui: Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 child support", "{Text}GCSC ", 1, "Y/N, Child(ren) Member#")
        Return
        ^F9::
            Gui, MainGui: Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 wo", "{Text}CCAP CS INCOME CALC")
        Return
        ^F10::
            Gui, MainGui: Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 wo", "{Text}CCAP INCOME CALC")
        Return
        ^F11::
            Gui, MainGui: Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 other - in", "{Text}W# ", 3, "Member#, Employer")
        Return
        ^F12::
            Gui, MainGui: Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 edak 3813", "{Text}OUTBOUND")
        Return
    #If

    #IfWinActive ahk_group AutoMail ; OnBase, excluding "Perform Import"
        F1::
            ToolTipText := WinActive("Automated Mailing Home Page")
            ? "Ctrl+B: Types in the current date and case number." : "
            (
    Ctrl+B: (Mail) Types in the current date and case number, clicks Yes. Works best from Custom Query.
              Step 1: Select documents in the query.
              Step 2: Right Click -> Send To -> Envelope.
              Step 3: Click 'Create Envelope'

    Alt+4: (Keywords) Enters 'VERIFS DUE BACK' + verif due date. 'Details' keyword field must be active.
            )"
            ToolTip, % ToolTipText,0,0
            SetTimer, RemoveToolTip, -8000
        Return
        ^b::
            Gui, MainGui: Submit, NoHide
            SendInput, % ShortDate " " CaseNumber
            Sleep 500
            If ( WinActive("ahk_exe obunity.exe") ) {
                Sleep 250
                Send {Tab}
                Send {Enter}
                MsgBox, 4100, Case Open Mail, Reminder: First add documents to envelope.`n`nOpen / switch to Automated Mailing?
                    IfMsgBox Yes
                        If (WinExist("Automated Mailing Home Page") ) {
                            WinActivate, "Automated Mailing Home Page"
                                Return
                        } Else {
                            run http://webapp4/AutomatedMailingPRD/#step-1
                        }
            }
            
        Return
        !4::
            SendInput, % "VERIFS DUE BACK " AutoDenyObject.AutoDenyExtensionDate
        Return
    #If

    #IfWinActive ahk_group Maxis
        ^m::
            WinSetTitle, ahk_exe bzmd.exe,,S1 - MAXIS
            Send ^{m}
        Return
    #If


    #IfWinActive ahk_group Browser
        F1::
            ToolTip,
            (
    Alt+F1: Reviewed/Approved application (Start New case note first)
    Alt+F2: Reviewed/Denied application (Start New case note first)

    Ctrl/Alt+F12: Add worker signature to case note
            )
            , 0, 0
            SetTimer, RemoveToolTip, -5000
        Return

        !F1::
            InputBox, ApprovedDate, Enter Approved Date, Approved eligible results effective _____.
            InputBox, SaApprovalInfo, Enter Service Authorization Details, Service Authorization _______. `n`nExamples: `n  approved effective 1/1/25 `n  not approved
            If (Ini.EmployeeInfo.EmployeeUseMec2Functions = 0)
                Send {Tab 7}
            Else
                Send {Tab}
            Sleep 750
            Send {A 4}
            Sleep 500
            Send {Tab}
            Sleep 500
            SendInput, Reviewed application requirements - approved elig
            Sleep 500,
            Send, {Tab}
            Sleep 500,
            SendInput, % "Reviewed case for verifications that are required at application. Verifications were received.`n-`nApproved eligible results effective " ApprovedDate ".`n-`nService Authorization " SaApprovalInfo ".`n=====`n" Ini.EmployeeInfo.EmployeeName
        Return

        !F2::
            If (Ini.EmployeeInfo.EmployeeUseMec2Functions = 0)
                Send {Tab 7}
            Else
                Send {Tab}
            Sleep 750
            Send {A 4}
            Sleep 500
            Send {Tab}
            Sleep 500
            SendInput, Reviewed application requirements - app denied
            Sleep 500,
            Send, {Tab}
            Sleep 500,
            SendInput, % "Reviewed case for documents that are required at application. Documents were not received.`n-`nApplication was denied by MEC2 and remains denied.`n=====`n" Ini.EmployeeInfo.EmployeeName
            Sleep 1000,
            Send, !{s}
        Return
    #If
}
DisplayResult(Result) {
    ToolTip, %Result%
    SetTimer, RemoveToolTip, -2000
}