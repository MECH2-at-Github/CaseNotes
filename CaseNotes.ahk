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

Version := "v0.3.6"

;Future todo ideas:
;Add backup to ini for Case Notes window. Check every minute old info vs new info and write changes to .ini.
;Make a restore button.
;Import from clipboard (when copied from MEC2) (likely mostly same code as restore button)

SetWorkingDir %A_ScriptDir%
#Persistent
#SingleInstance force
#NoTrayIcon
SetTitleMatchMode, RegEx

IniRead, WorkerNameRead, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeName, %A_Space%
IniRead, WorkerCounty, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeCounty,%A_Space%
IniRead, WorkerEmailRead, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeEmail, %A_Space%
IniRead, WorkerPhoneRead, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeePhone, %A_Space%
IniRead, UseWorkerEmailRead, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeUseEmail, 0
IniRead, UseMec2FunctionsRead, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeUseMec2Functions, 0
IniRead, WorkerBrowserRead, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeBrowser, Google Chrome
IniRead, WorkerMaxisRead, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeMaxis, MAXIS-WINDOW-TITLE
;; County Specific Items:
IniRead, CountyNoteInMaxisRead, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyNoteInMaxis, 0
IniRead, CountyFaxRead, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyFax, %A_Space%
IniRead, CountyDocsEmailRead, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyDocsEmail, %A_Space%
IniRead, ProviderWorkerPhoneRead, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyProviderWorkerPhone, %A_Space%
IniRead, CountyEdBSFformRead, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyEdBSF, %A_Space%

IniRead, XClipboard, %A_MyDocuments%\AHK.ini, ClipboardContents, XClipboardINI, 0
IniRead, YClipboard, %A_MyDocuments%\AHK.ini, ClipboardContents, YClipboardINI, 0

IniRead, XCaseNotes, %A_MyDocuments%\AHK.ini, CaseNotePositions, XCaseNotesINI, 0
IniRead, YCaseNotes, %A_MyDocuments%\AHK.ini, CaseNotePositions, YCaseNotesINI, 0
If (XCaseNotes < -3000 || YCaseNotes < -3000) {
    XCaseNotes := 0
    YCaseNotes := 0
}
;Declaring Global Variables
MissingVerifications := {}, ClarifiedVerifications := {}, EmailText := {}, LineCount := 0, Homeless := 0

CaseDetails := { DocType: "_DOC?", Eligibility: "_ELIG?", SaEntered: "_SA?", CaseType: "_PRG?", AppType: "_APP?", isHomeless: "" }
SignDate := 0
CaseNoteEntered := { MEC2Note: 0, MAXISNote: 0 }
MaxisNote :=, IdList := ""
ConfirmedClear := 0
VerificationWindowOpenedOnce := 0
VerifCat :=, LetterTextNumber := 1, LetterText := {}, MissingHomelessItems := ""

CountySpecificText := {}
CountySpecificText.Dakota := { OverIncomeContactInfo: " contact 651-554-6696 and", CustomHotkeys: "Custom hotkeys for your county exist for the following windows (A ToolTip reminder appears by pressing F1):`nOnBase (Alt+4, 'Verifs Due Back' detail),`nOnBase (Perform Import) - For printing docs to OnBase (Ctrl+F6-12)`nOnBase (Ctrl+B, Inserts date and case number for mail)`nAutomated Mailing (Ctrl+B, Inserts date and case number for mail)`nBrowser (Types in case note for app denied/approved, and worker signature)`nWord (Types in first name, case number, and app received date)." }
CountySpecificText.StLouis := { OverIncomeContactInfo: "" }

; Date variables
DateObject := { ReceivedMDY: "", ReceivedYMD: "", AutodenyYMD: "", ReinstateDate: "" }
AutoDenyObject := { AutoDenyExtensionMECnote: "", AutoDenyExtensionDate: "", AutoDenyExtensionSpecLetter: "", AutoDenyExtraLines: "", AutoDenyMaxisNote: "" }
DateObject.TodayYMD := A_Now
FormatTime, ShortDate, %A_Now%, M/d/yy ; for sending to envelope
Received := DateObject.TodayYMD
DateObject.TodayMDY := FormatMDY(DateObject.TodayYMD)
OverIncomeObj := { overIncomeHHsize: "your size" } ;:= "your size"

If InStr(DateObject.TodayYMD, 0401)
    Menu, Tray, Icon, compstui.dll, 100
Else
    Menu, Tray, Icon, azroleui.dll, 7


;MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION MAIN GUI SECTION 
Gui MainGui: Font,, Segoe UI
Gui MainGui: Color, a9a9a9, bebebe

Gui, MainGui: Add, Radio, Group Section h17 x12 w75 y+5 gApplicationRadio, Application
Gui, MainGui: Add, Radio, xp y+2 wp h17 gRedeterminationRadio, Redeterm.
Gui, MainGui: Add, Checkbox, xp y+2 wp h17 vHomeless gHomeless, Homeless

Gui, MainGui: Add, Radio, Group x+10 ys h17 w78 gMNBenefits vMNBenefitsRadio, MNBenefits
Gui, MainGui: Add, Radio, xp y+2 h17 wp gApp vAppRadio, 3550 App

Gui, MainGui: Add, Radio, Group x+10 ys h17 w58 gBSF vBSF, BSF
Gui, MainGui: Add, Radio, xp y+2 h17 wp gTY vTY, TY
Gui, MainGui: Add, Radio, xp y+2 h17 wp gCCMF vCCMF, CCMF

Gui, MainGui: Add, Radio, Group x+0 ys h17 w80 gPending vPendingRadio, Pending
Gui, MainGui: Add, Radio, xp y+2 h17 wp gEligible vEligibleRadio, Eligible
Gui, MainGui: Add, Radio, xp y+2 h17 wp gIneligible vIneligibleRadio, Ineligible

Gui, MainGui: Add, Radio, Group Hidden x+5 ys h17 vSaApproved gSaApproved, SA Approved
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
    Gui 9:Font, %fontOptions%, %fontName%
    Gui 9:Add, Text, Limit200, %newText%
    GuiControlGet T, 9:Pos, Static1
    Gui 9:Destroy
    Return TW
}
; 77 characters returns 12h 565w at 3840 x 2160, 150%; returns 12h 539w at 1920 x 1080, 100% ;123456789 123456789 123456789 123456789 123456789 123456789 123456789 123456789 1234567
OneHundredChars := SetTextAndResize("1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890", "s9", "Lucida Console")
MonoChar := OneHundredChars/100
; 100 characters is 733w at 3840 x 2160
LabelSettings := "xm+5 y+1 w200"
LabelExampleSettings := "x220 yp+4 h12 w" MonoChar*60
; 995 is about the right size (* .666 = 663)
TextboxSettings := "xm y+1 w" (MonoChar*87)+27 ; At w650, WinSpy shows 945 for box with scrollbar, 971 without, and 975 total. (spy * .666 = AHK #s). Which gets 20 for the (~17.4) scrollbar and (~2.6) border
OneRow := "h17 Limit87"
TwoRows := "h33"
ThreeRows := "h43"
FourRows := "h55"

Gui MainGui: Font, s9, Segoe UI
Gui, MainGui: Margin, 12 12
Gui, MainGui: Add, Text, xm y+45 h0 w0 ; Blank space
Gui, MainGui: Add, Text, %LabelSettings% vHouseholdCompLabel, Household Comp
Gui, MainGui: Add, Text, %LabelExampleSettings% vHouseholdCompLabelExample Hidden, Parent (ID), ChildOne (4, BC), ChildName (age, verif)
Gui, MainGui: Add, Edit, %TextboxSettings% %TwoRows% vHouseholdComp, 

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

Gui, MainGui: Show, x%XCaseNotes% y%YCaseNotes%, CaseNotes
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

If (StrLen(WorkerNameRead) < 1)
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
    Main window functions:
    ● [MEC2 Note] - Formats the entire case note and sends the data to MEC2.
        If you are not using mec2functions, it will simulate keypresses to navigate the page. 
        In MEC2: Click 'New' in MEC2 on the CaseNotes webpage. In CaseNotes, click [MEC2 Note].
    ● [Maxis Note] - Formats the app date, case status, and verifications list and sends it to Maxis.
        It will activate BlueZone and paste the case note in.
        In BlueZone (Maxis): PF9 to start a new note. In CaseNotes, click [Maxis Note].
    ● [Desktop Backup] - Saves case notes for MEC2, Maxis, the Special Letters, and Email to your desktop.
        In CaseNotes, click [To Desktop]. A text file will be saved using the case number for the file name.
    ● [Clear] - Resets the app. If the case note has not been sent to MEC2/Maxis or saved to file, it will give a
        warning. Otherwise, it will change to [Confirm]. Clicking again will reset the app.
    ● Child Support Cooperation - Click the [Child Support Cooperation] label to copy from Custody.
        The Child Support Cooperation text field must be blank when clicking the button.
    ● Missing Verifications - Click either the [Missing] label or [Missing] button to open the verification window.
      
      Missing Verifications:
    ● Checkboxes cover almost all documents needed, with 3 [Other] selections for free-form requests.
    ● Verification types are clustered by categories.
    ● If a checkbox label has (Input) it will open a popup requesting clarifying information.
    ● [Over-income] will modify the beginning of the Special Letter text, adding income limit information.
    
    ● [Done] - After selecting the verifications that are missing, [Done] will generate a list and letter/email buttons.
          If the Special Letter text exceeds 30 lines, a second (/third/fourth) letter will be generated.
    ● [Letter 1/2/3/4] - Clicking these will place text on the clipboard to be pasted into the Worker Comments field.
        If mec2functions is enabled, [Letter 1] will auto-check and auto-fill fields in MEC2's Special Letter page.
    ● [Email] - Text is send to the clipboard, which can then be pasted into an email. It will include the document type
        that was selected (application/redetermination). For homeless cases, alternate text is generated based on if
        the case is Eligible or Pending.
    )"
    Paragraph4 := "
    (
    Hotkeys:
    ● [Alt+3]: Copies the case number to the clipboard. Usable from anywhere when CaseNotes is open.
    ● [Win+left arrow] or [Win+right arrow]: Resets CaseNotes location, in the event it is off-screen.
    ● [Ctrl+F12] or [Alt+12]: In MEC2, types in the worker's name with a separation line (case note signature)
    ● [Ctrl+Alt+a]: When in MEC2, select and copy text (such as from a case note). Hotkey popups copied text in a separate window.
    )"
    Paragraph5 := "
    (
    Special notes:
    ● CaseNotes will open in the same location it was closed, even if that monitor is no longer connected. See Hotkeys.
    ● All settings for CaseNotes are saved in the My Documents folder, under AHK.ini. Deleting this file will reset all
        saved settings.
    ● Please send any bug reports or feature requests to MECH2.at.github@gmail.com
    )"

    Gui, HelpGui: New, ToolWindow, CaseNotes Help
    Gui, HelpGui: Margin, 12 12
    Gui, Font, s10, Segoe UI
    Gui, HelpGui: Add, Tab3,, Features | Main| Hotkeys and Notes 
    Gui, Tab, 1
    Gui, HelpGui: Add, Text, xm y+5, % Paragraph1
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph2
    Gui, Tab, 2
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph3
    Gui, Tab, 3
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph4
    Gui, HelpGui: Add, Text, xm y+15, % Paragraph5
    Gui, HelpGui: Add, Text, xm y+15, % CountySpecificText[WorkerCounty].CustomHotkeys
    Gui, Tab
    Gui, HelpGui: Add, Button, gCloseHelp w70 h25 x375, Close
    Gui, HelpGui: Show,, CaseNotes Help
Return

CloseHelp:
    Gui, HelpGui: Destroy
Return

ExamplesButton:
    Gui, Submit, NoHide
    
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

CBTGuiClose:
    WinGetPos, XClipboardContentsGet, YClipboardContentsGet,,, A
    If (XClipboardContentsGet - XClipboard <> 0)
        IniWrite, %XClipboardContentsGet%, %A_MyDocuments%\AHK.ini, ClipboardContents, XClipboardINI
    If (YClipboardContentsGet - YClipboard <> 0)
        IniWrite, %YClipboardContentsGet%, %A_MyDocuments%\AHK.ini, ClipboardContents, YClipboardINI
    Gui, CBT: Destroy
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

MakeCaseNote:
	Gui, Submit, NoHide
    GoSub CalcDates
    EditControlsTest := ["HouseholdComp", "SharedCustody", "AddressVerification", "SchoolInformation", "Income", "ChildSupportIncome", "ChildSupportCooperation", "Expenses", "Provider", "ActivityandSchedule", "ServiceAuthorization", "Notes", "Missing"]
    For each, EditField in EditControlsTest {
        %EditField% := st_wordWrap(%EditField%, 88, "")
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
	MEC2CaseNote := AutoDenyObject.AutoDenyExtensionMECnote " HH COMP:    " HouseholdComp "`n CUSTODY:    " SharedCustody "`n ADDRESS:    " AddressVerification "`n  SCHOOL:    " SchoolInformation  "`n  INCOME:    " Income "`n      CS:    " ChildSupportIncome  "`n CS COOP:    " ChildSupportCooperation  "`nEXPENSES:    " Expenses  "`n  ASSETS:    " Assets "`nPROVIDER:    " Provider "`nACTIVITY:    " ActivityandSchedule "`n      SA:    " ServiceAuthorization "`n   NOTES:    " Notes "`n MISSING:    " Missing "`n=====`n" WorkerNameRead
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
        MissingMax := st_wordWrap(Missing, 73, "_")
        MissingMax := StrReplace(MissingMax, "`n", "`n* ")
        MissingMax := StrReplace(MissingMax, "* _", "  ")
        MaxisNote .= "Special Letter mailed " DateObject.TodayMDY " requesting:`n* " MissingMax "`n"
    }
    MaxisNote := StrReplace(MaxisNote, "              ", " ")
    MaxisNote .= WorkerNameRead

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
		WinActivate % WorkerBrowserRead
		Sleep 500
        MEC2DocType := CaseDetails.DocType = "Redet" ? "Redetermination" : CaseDetails.DocType
        If (UseMec2FunctionsRead = 1) {
            concatCaseNote := "CaseNoteFromAHKSPLIT" MEC2DocType "SPLIT" MEC2NoteTitle "SPLIT" MEC2CaseNote
            Clipboard := concatCaseNote
            Send, ^v
        } Else If (UseMec2FunctionsRead = 0) {
            WinActivate % WorkerBrowserRead
            Sleep 1000
            Send {Tab 7}
            Sleep 750
            If (CaseDetails.DocType = "Application") {
                Send {A 4}
            } else if (CaseDetails.DocType = "Redet") {
                Send R
            }
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
        WorkerMaxisRead := WorkerMaxisRead = "MAXIS-WINDOW-TITLE" ? "MAXIS" : WorkerMaxisRead
        MaxisWindow := WinExist(WorkerMaxisRead)
        If (MaxisWindow = "0x0")
            MaxisWindow := WinExist("BlueZone Mainframe")

        If (WinExist(ahk_id %MaxisWindow%)) {
            WinActivate ahk_id %MaxisWindow%
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
        If (CountyNoteInMaxisRead != 1) {
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
    Gui, Submit, NoHide
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
    AutoDenyObject.AutoDenyExtraLines := 0, AutoDenyObject.AutoDenyExtensionSpecLetter :=
    If (CaseDetails.DocType = "Application") {
        If (CaseDetails.Eligibility = "pends") {
            If (NeedsNoExtension > -1) {
                AutoDenyObject.AutoDenyExtensionDate := DateObject.AutoDenyMDY
                AutoDenyObject.AutoDenyExtensionSpecLetter := "NewLineAutoreplaceTwo`n**You have through " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications."
                AutoDenyObject.AutoDenyExtraLines := 1
                GuiControl, MainGui: Text, AutoDenyStatus, Has 15+ days before auto-deny
            } Else If (NeedsExtension > -1) {
                AutoDenyObject.AutoDenyExtensionDate := DateObject.TodayPlusFifteenishMDY
                AutoDenyObject.AutoDenyExtensionMECnote := "Auto-deny extended to " AutoDenyObject.AutoDenyExtensionDate " due to processing < 15 days before auto-deny.`n-`n"
                AutoDenyObject.AutoDenyExtensionSpecLetter := "NewLineAutoreplaceTwo`n**You have through " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications."
                AutoDenyObject.AutoDenyExtraLines := 1
                GuiControl, MainGui: Text, AutoDenyStatus, % "Extend auto-deny to " AutoDenyObject.AutoDenyExtensionDate
            } Else {
                AutoDenyObject.AutoDenyExtensionDate := DateObject.TodayPlusFifteenishMDY
                AutoDenyObject.AutoDenyExtensionMECnote := "Reinstate date is " AutoDenyObject.AutoDenyExtensionDate " due to processing < 15 days before auto-deny.`n-`n"
                AutoDenyObject.AutoDenyExtensionSpecLetter := "NewLineAutoreplaceTwo`n**Please note that you will be mailed an auto-denial notice.`n  You have through " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications.`n  If you are eligible, your case will be reinstated."
                AutoDenyObject.AutoDenyExtraLines := 3
                GuiControl, MainGui: Text, AutoDenyStatus, % "Auto-denies tonight, pends until " AutoDenyObject.AutoDenyExtensionDate
            }
        }
        If (Homeless = 1) {
            AutoDenyObject.AutoDenyExtensionDate := DateObject.ExpeditedNinetyDaysMDY
            AutoDenyObject.AutoDenyExtensionSpecLetter := "NewLineAutoreplaceTwo`n**You have until " AutoDenyObject.AutoDenyExtensionDate " to submit required verifications."
            AutoDenyObject.AutoDenyExtraLines := 3
        }
    }
    
    If (CaseDetails.DocType = "Redet") {
            AutoDenyObject.AutoDenyExtensionSpecLetter := "NewLineAutoreplaceTwo`n*  If your redetermination is not completed by " DateObject.SignedMDY ",`n   your case will close on " DateObject.RedetCaseCloseMDY ". If it closes,`n   the latest it can be reinstated is " DateObject.RedetDocsLastDayMDY "."
            AutoDenyObject.AutoDenyExtraLines := 3
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
    If (CountyNoteInMaxisRead = 1) {
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

Homeless:
    Gui, Submit, NoHide
Return

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
    GuiControl, MainGui: Show, SaApproved
    GuiControl, MainGui: Show, NoSA
    GuiControl, MainGui: Show, NoProvider
return

Pending:
	CaseDetails.Eligibility := "pends"
    GuiControl, MainGui: Hide, SaApproved
    GuiControl, MainGui: Hide, NoSA
    GuiControl, MainGui: Hide, NoProvider
return

Ineligible:
	CaseDetails.Eligibility := "ineligible"
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
    Gui, Submit, NoHide
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
    Gui, Submit, NoHide
Return

MissingGui:
    Column1of1 := "xm w390"

    ; 12 + 158 (170) + 240 (410) + 12 = 422
    Column1of2 := "xm w158"
    Column2of2 := "x170 yp+0 w240"

    ; 12 + 118 (130) + 120 (250) + 160 (400) + 12 = 422
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

    IniRead, XVerification, %A_MyDocuments%\AHK.ini, CaseNotePositions, XVerificationINI, 0
    IniRead, YVerification, %A_MyDocuments%\AHK.ini, CaseNotePositions, YVerificationINI, 0
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
    Gui, MissingGui: Add, Checkbox, %Column1of1% vSeasonalOffSeasonMissing, Seasonal employment info - app in off-season (input)
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
    Gui, MissingGui: Add, Checkbox, %Column1of2% vInsuranceBenefitsMissing, Insurance Benefits 
    Gui, MissingGui: Add, Checkbox, %Column2of2% vUnearnedStatementMissing, Blank Unearned Yes/No (statement)
    Gui, MissingGui: Add, Checkbox, %Column1of2% vAssetsBlankMissing, Assets (Blank)
    Gui, MissingGui: Add, Checkbox, %Column2of2% vUnearnedMailedMissing, Blank Unearned Yes/No (mailed back)
    Gui, MissingGui: Add, Checkbox, %Column1of2% vVABenefitsMissing, VA Benefits
    Gui, MissingGui: Add, Checkbox, %Column2of2% vAssetsGT1mMissing, Assets (>$1m)

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

    Gui, MissingGui: Show, Hide x%XVerification% y%YVerification%
Return

MissingButtonDoneButton:
	Gui, Submit, NoHide
    GoSub CalcDates
	MissingVerifications := {}, ClarifiedVerifications := {}, LineCount := 0
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
    PendingHomelessText1 := "You may be eligible for the homeless policy, which allows us to approve eligibility even though there are verifications we need but do not have. These verifications are still required, and must be received within 90 days of your application date for continued eligibility.`n`nBefore we can approve eligibility, we need information that you did not put on the application:`n"

    If OverIncomeMissing {
        OverIncomeMissingText1 := "Using information you provided your case is ineligible as your income is over the limit for a household of " OverIncomeObj.overIncomeHHsize ". The gross limit is $" OverIncomeObj.overIncomeText ".`n"
        OverIncomeMissingText2 := "If your gross income does not match this calculation, you must" CountySpecificText[WorkerCounty].OverIncomeContactInfo " submit updated income and expense documents along with the following verifications:`n"
        EmailTextString := "Your Child Care Assistance " MEC2DocType " has been processed.`n`n" OverIncomeMissingText1 OverIncomeMissingText2 "`n"
        MissingVerifications[OverIncomeMissingText1] := 3
        MissingVerifications[OverIncomeMissingText2] := 3
        CaseNoteMissing .= "Household is calculated to be over-income by $" OverIncomeObj.overIncomeDifference " ($" OverIncomeObj.overIncomeReceived " - $" OverIncomeObj.overIncomeLimit ");`n"
    }
    If (Homeless = 1 && CaseDetails.Eligibility = "pends") {
        InputBox, MissingHomelessItems, Homeless Info Missing, What information is needed from the client to approve expedited eligibility?`n`nUse a double space "  " without quotation marks to start a new line.,,,,,,,, % StrReplace(MissingHomelessItems, "`n", "  ")
        If (ErrorLevel = 0) {
            MissingHomelessItems := StrReplace(MissingHomelessItems, "  ", "`n")
            PendingHomelessMissing := getRowCount(MissingHomelessItems, 57, "   ")
            MissingVerifications[st_wordwrap(PendingHomelessText1, 59, " ") "`n"] := 8
            MissingVerifications[PendingHomelessMissing[1] "`n"] := PendingHomelessMissing[2]
            CaseNoteMissing .= "Missing for expedited approval:`n" PendingHomelessMissing[1] "`n"
        }
    }

	EmailText := {}
	If (!InStr(Missing, "over-income")) {
        EmailText.StartHL := (CaseDetails.Eligibility = "elig") ? "It was approved under the homeless expedited policy which allows us to approve eligibility even though there are verifications we require that we do not have. These verifications are still required, and must be received within 90 days of your application date for continued eligibility." : PendingHomelessText1 MissingHomelessItems

        EmailText.EndHL := (CaseDetails.Eligibility = "elig") ? "`nThe initial approval of child care assistance is 30 hours per week for each child. This amount can be increased once we receive your activity verifications and we determine more assistance is needed.`nIf the provider you select is a “High Quality” provider, meaning they are Parent Aware 3⭐ or 4⭐ rated, or have an approved accreditation, the hours will automatically increase to 50 per week for preschool age and younger children.`nIf you have a 'copay,' the amount the county pays to the provider will be reduced by the copay amount. Many providers charge more than our maximum rates, and you are responsible for your copay and any amounts the county cannot pay." : ""

        EmailText.AreOrWillBe := (Homeless = 1) ? "will be" : "are"

        EmailText.Reason1 := (CaseDetails.Eligibility = "elig") ? "for authorizing assistance hours" : "to determine eligibility or calculate assistance hours"
        EmailText.Reason2 := (Homeless = 1) ? "to determine on-going eligibility or calculate assistance hours after the 90-day period" : EmailText.Reason1
        EmailText.StartAll := "Your Child Care Assistance " MEC2DocType " has been processed. "

        EmailText.Start := (Homeless = 1) ? EmailText.StartAll EmailText.StartHL : EmailText.StartAll
        EmailText.Middle := "`n`nThe following documents or verifications " EmailText.AreOrWillBe " needed " EmailText.Reason2 ":`n`n"

        EmailText.Combined := EmailText.Start EmailText.Middle
	}
    
	If IDmissing {
        IDmissingText := "ID for " IDmissingInput ";`n"
		ClarifiedVerifications[ClarifyListItem ". " IDmissingText] := 1
        EmailTextString .= EmailListItem ". " IDmissingText
		CaseNoteMissing .= "ID for " IDmissingInput ";`n"
        ClarifyListItem++
        EmailListItem++
        MecCheckboxIds.proofOfIdentity := 1
    }
	If BCmissing {
        BCmissingText := "Birth date / relationship / citizenship verification for: " BCmissingInput
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
        BCNonCitizenMissingText := "Birth date / relationship / immigration verification for: " BCNonCitizenMissingInput ";`n"
		ClarifiedVerifications[ClarifyListItem ". " BCNonCitizenMissingText] := 2
        EmailTextString .= EmailListItem ". " BCNonCitizenMissingText
		CaseNoteMissing .= "Birth date / relationship / immigration verification for: " BCNonCitizenMissingInput ";`n"
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
        If (ChildSupportFormsMissingInput ~= "^\d$") {
            ChildSupportFormsMissingInput .= ChildSupportFormsMissingInput < 2 ? " set" : " sets"
        }
        ChildSupportFormsMissingText := "Cooperation with Child Support forms (" ChildSupportFormsMissingInput ", sent separately);`n"
        ;CSFMlines := ChildSupportFormsMissingInput ~= "^\d" ? 1 : 2
		MissingVerifications[ListItem ". " ChildSupportFormsMissingText] := 2 ; CSFMlines
        EmailTextString .= EmailListItem ". " ChildSupportFormsMissingText
		CaseNoteMissing .= "CS forms (" ChildSupportFormsMissingInput ");`n"
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
        CustodyScheduleMissingText := "A statement, written by you that is signed and dated, for " CustodySchedulePlusNamesMissingInput ":`n  A. Stating that you have full custody, or`n  B. Your current Parenting Time (shared custody) schedule `n     listing the days and times of the custody switches;`n"
		MissingVerifications[ListItem ". " CustodyScheduleMissingText] := 5
        EmailTextString .= EmailListItem ". " CustodyScheduleMissingText
		CaseNoteMissing .= "Shared custody / parenting time for " CustodySchedulePlusNamesMissingInput ";`n"
		ListItem++
        EmailListItem++
    }
    if DependentAdultStudentMissing {
        DependentAdultStudentMissingText := "Verification of full-time student status for " DependentAdultStudentMissingInput ", verification of their most recent 30 days income, and a signed statement that you provide at least 50% of their financial support;`n"
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
        LegalNameChangeMissingText := "Legal name change verification for " LegalNameChangeMissingInput ";`n"
		MissingVerifications[ListItem ". " LegalNameChangeMissingText] := 1
        EmailTextString .= EmailListItem ". " LegalNameChangeMissingText
		CaseNoteMissing .= "Legal name change for " LegalNameChangeMissingInput ";`n"
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
        IncomeText := NeedsExtension > -1 ? IncomePlusNameMissingInput "'s most recent 30 days income" : CaseDetails.DocType = "Redet" ? IncomePlusNameMissingInput "'s 30 days income prior to " DateObject.SignedMDY : IncomePlusNameMissingInput "'s 30 days income prior to " DateObject.ReceivedMDY
        IncomeMissingText := "Verification of " IncomeText ";`n"
        ClarifiedVerifications[ClarifyListItem ". Proof of Financial Information: " IncomeMissingText] := 2
        EmailTextString .= EmailListItem ". " IncomeMissingText
		CaseNoteMissing .= "Earned income (" IncomePlusNameMissingInput ");`n"
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
        WorkScheduleText := NeedsExtension > -1 ? WorkSchedulePlusNameMissingInput "'s work schedule" : CaseDetails.DocType = "Redet" ? WorkSchedulePlusNameMissingInput "'s work schedule from " DateObject.SignedMDY : WorkSchedulePlusNameMissingInput "'s work schedule from " DateObject.ReceivedMDY
        WorkScheduleMissingText := "Verification of " WorkScheduleText " showing days of the week and start/end times;`n"
        ClarifiedVerifications[ClarifyListItem ". Proof of Activity Schedule: " WorkScheduleMissingText] := 2
        EmailTextString .= EmailListItem ". " WorkScheduleMissingText
		CaseNoteMissing .= "Work schedule (" WorkSchedulePlusNameMissingInput ");`n"
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
        SeasonalOffSeasonMissing := StrLen(SeasonalOffSeasonMissing) > 0 ? " at " SeasonalOffSeasonMissing : ""
        SeasonalOffSeasonMissingText := "Verification of either seasonal employment" SeasonalOffSeasonMissing ", including expected season length and typical wages, or a signed statement that you are no longer an employee at this job.`n Upon returning to work, verification of work schedule will be needed, showing days of the week and start/end times;`n"
		MissingVerifications[ListItem ". " SeasonalOffSeasonMissingText] := 6
        EmailTextString .= EmailListItem ". " SeasonalOffSeasonMissingText
		CaseNoteMissing .= "Seasonal employment (applied during off season);`n"
        ListItem++
        EmailListItem++
    }
;----------------------------
	If SelfEmploymentMissing {
        SelfEmploymentMissingText := "Self-employment income and expenses, such as federal tax return/schedules, or if less than a full tax year of self-employment or if last year's taxes don't represent expected ongoing income, a report or ledger with the most recent 3 months of gross income;`n"
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
        ExpensesMissingText := "Proof of Deductions: Healthcare premiums, child support, and spousal support, if not listed on submitted paystubs;`n"
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
        ;UnansweredText := st_wordwrap(UnansweredText, 57, " ")
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
        EdBSFformMissingText := CountyEdBSFformRead " form (sent separately);`n"
		MissingVerifications[ListItem ". " EdBSFformMissingText] := 1
        EmailTextString .= EmailListItem ". " EdBSFformMissingText
		CaseNoteMissing .= CountyEdBSFformRead " form;`n"
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
            EmailTextString .= EmailListItem ". " CountedRows[1] ";`n"
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
        ChildSupportNoncooperationMissingText := "* You are currently in a non-cooperation status with Child Support. Contact Child Support at " ChildSupportNoncooperationMissingInput " for details. Child Support cooperation is a requirement for eligibility.`n"
		MissingVerifications[ChildSupportNoncooperationMissingText] := 3
        EmailTextString .= ChildSupportNoncooperationMissingText
		CaseNoteMissing .= "Cooperation status with Child Support, CS number: " ChildSupportNoncooperationMissingInput ";`n"
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
        EligibleActivityMissingText := "* You did not select an eligible activity on the application. " EligibleActivityWithJSText "`n"
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
        ESPlanOnlyJSMissingText := "* While you have an Employment Plan, assistance hours cannot be approved for job search unless it is listed on the Plan;`n"
		MissingVerifications[ESPlanOnlyJSMissingText] := 2
        EmailTextString .= ESPlanOnlyJSMissingText
		CaseNoteMissing .= "Client has ES Plan - informed JS hours are required to be on the Plan;`n"
    }
	If ActivityAfterHomelessMissing {
        ActivityAfterHomelessMissingText := "* At the end of the 90-day homeless exemption period, you must have an eligible activity to keep your Child Care Assistance case open. " EligibleActivityWithoutJSText "`n"
		MissingVerifications[ActivityAfterHomelessMissingText] := 6
        EmailTextString .= ActivityAfterHomelessMissingText
		CaseNoteMissing .= "Eligible activity after the 3-month homeless period;`n"
    }
	If NoProviderMissing {
        NoProviderMissingText := "* Once you have a daycare provider, please notify me with the provider’s name and location, and the start date.`n`n   If you need help locating a daycare provider, contact Parent Aware at 888-291-9811 or www.parentaware.org/search`n"
        EmailTextString .= NoProviderMissingText
		CaseNoteMissing .= "Provider;`n"
        MecCheckboxIds.providerInformation := 1
    }
    ;*   Provider Information- If you have a child care provider, send the provider's name, address and start date (if known). Visit www.parentaware.org for help finding a provider. Care is not approved until you get a Service Authorization.
	If UnregisteredProviderMissing {
        UnregisteredProviderMissingText := "* Your daycare provider is not registered with Child Care Assistance. Please have them call " ProviderWorkerPhoneRead " to register.`n"
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
    
    ClarifiedVerifications["NewLineAutoreplaceOne Documents can also be faxed to " CountyFaxRead " or emailed to`n " CountyDocsEmailRead ". Please include your case number." AutoDenyObject.AutoDenyExtensionSpecLetter] := 2+AutoDenyObject.AutoDenyExtraLines

    MecCheckboxIds.other := 1
    IdList := ""
    For key, value in MecCheckboxIds {
        If StrLen(IdList) > 1
            IdList .= ","
        IdList .= key
    }
    If !OverIncomeMissing {
        InsertFollowing := (Homeless = 1 && CaseDetails.Eligibility = "pends") ? 3 : 1
        MissingMod := (Homeless = 1 && CaseDetails.Eligibility = "pends") ? 2 : 0
        If (MissingVerifications.Length() > (0 + MissingMod)) {
            If (StrLen(IdList) > 5 || InsertFollowing = 3) { ; "other" will always add at least 5
                MissingVerifications.InsertAt(InsertFollowing, "__In addition to the above, please submit following items:__`n", 1)
            }
            Else If (StrLen(IdList) = 5) {
                MissingVerifications.InsertAt(InsertFollowing, "_____________Please submit the following items:_____________`n", 1)
            }
        }
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
        If (InStr(key, "faxed")) { ; faxed = last key in group
            LineCountPlusFaxed := LineCount + value
            If (LineCountPlusFaxed = 30) {
                key := StrReplace(key, "NewLineAutoreplaceTwo", "")
                key := StrReplace(key, "NewLineAutoreplaceOne", "")
            } Else If (LineCountPlusFaxed > 30) {
                key := StrReplace(key, "NewLineAutoreplaceTwo", "`n")
                key := StrReplace(key, "NewLineAutoreplaceOne", "`n")
                
                    LetterTextNumber++
                    LetterText[LetterTextNumber] .= "                   Continued on letter " LetterTextNumber
                    LetterText[LetterTextNumber] .= "                  Continued from letter " LetterTextNumber-1 "`n"
                    
                    LetterNumber++
                    %LetterTextVar% .= "                   Continued on letter " LetterNumber
                    LetterTextVar := "LetterText" . LetterNumber
                    %LetterTextVar% .= "                  Continued from letter " LetterNumber-1 "`n"
            } Else If (LineCountPlusFaxed < 30) {
                key := StrReplace(key, "NewLineAutoreplaceTwo", "`n")
                If (LineCountPlusFaxed < 29) {
                    key := StrReplace(key, "NewLineAutoreplaceOne", "`n")
                } Else {
                    key := StrReplace(key, "NewLineAutoreplaceOne", "")
                }
            }
            %LetterTextVar% .= key
        } Else { ; does not contain faxed
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
Return

SetEmailText:
    Gui, Submit, NoHide
    EmailTextString := StrReplace(EmailTextString, "`n ", " ")
    EmailTextString := StrReplace(EmailTextString, "    ", " ")
    EmailTextString := StrReplace(EmailTextString, "   ", " ")
    EmailTextString := StrReplace(EmailTextString, "`n*", "`n`n*")
    EmailTextString := StrReplace(EmailTextString, "sent separately", "see attached")
	EmailText.Output := (Homeless = 1) ? EmailText.Combined EmailTextString EmailText.EndHL : EmailText.Combined EmailTextString
Return

Letter:
    Gui, Submit, NoHide
    If (Homeless = 1 && CaseDetails.Eligibility = "pends" && StrLen(MissingHomelessItems) < 1) {
        GoSub MissingButtonDoneButton
    }
    WinActivate % WorkerBrowserRead
    Sleep 500
	LetterGUINumber := "LetterText" . SubStr(A_GuiControl, 0)
    If (UseMEC2FunctionsRead = 1) {
        CaseStatus := InStr(CaseDetails.DocType, "?") ? "" : (Homeless = 1) ? "Homeless App" : (CaseDetails.DocType = "Redet") ? "Redetermination" : CaseDetails.DocType
        concatLetterText := "LetterTextFromAHKSPLIT" %LetterGUINumber% "SPLIT" CaseStatus "SPLIT" IdList
        Clipboard := concatLetterText
        Send, ^v
    } Else {
        Clipboard := %LetterGUINumber%
        Send, ^v
    }
    Sleep 500
    Clipboard := CaseNumber
Return

Other:
	Gui, Submit, NoHide
	If (%A_GuiControl% = 0)
		Return
	InputBoxAGUIControlVariable := %A_GuiControl%Input
	InputBox, %A_GuiControl%Input, Additional Input Required, Enter the requested item below. Use double space "  " to make a new line.,, 500,,,,,,%InputBoxAGUIControlVariable%
	If ErrorLevel
		Return
	InputBoxAGUIControlVariable := %A_GuiControl%Input
	GuiControl,,%A_GuiControl%, %InputBoxAGUIControlVariable%
	If (StrLen(%A_GuiControl%Input) > 150) {
		Gui, Font, s7, Segoe UI
		GuiControl, Font, %A_GuiControl%
	}
Return

InputBoxAGUIControl:
    Gui, Submit, NoHide
    If (%A_GuiControl% = 0)
        Return
    InputBoxEntry := SubStr(A_GuiControl, 1, InStr(A_GuiControl, "Missing")+7)
    InputBoxAGUIControlVariable := %InputBoxEntry%Input

    If (InputBoxEntry = "IDmissing")
        PromptText := "Who is ID needed for?`n`nExample: 'Susanne, Robert Sr'"
    Else If (InputBoxEntry = "BCmissing")
        PromptText := "Who is birth verification needed for?`n`nExample: 'Susie, Bobby Jr'"
    Else If (InputBoxEntry = "IncomePlusNameMissing")
        PromptText := "Who is the income verification needed for?"
    Else If (InputBoxEntry = "CustodySchedulePlusNamesMissing")
        PromptText := "Who is the schedule needed for? `n'...stating the current parenting time schedule for: ____________'`n`nExample: 'Susie and Bobby Jr' or 'your children'"
    Else If (InputBoxEntry = "WorkSchedulePlusNameMissing")
        PromptText := "Who is the work schedule needed for?"
    Else If (InputBoxEntry = "DependentAdultStudentMissing")
        PromptText := "Who is the adult dependent student?"
    Else If (InputBoxEntry = "ChildSupportFormsMissing")
        PromptText := "Enter the number of sets of Child Support forms needed`nor the names of the absent parent/children.`n`nExample: 'Robert Sr / Susie, Bobby Jr' or '2'"
    Else If (InputBoxEntry = "ChildSupportNoncooperationMissing")
        PromptText := "What is the phone number of the Child Support officer?"
    Else If (InputBoxEntry = "LegalNameChangeMissing")
        PromptText := "Who is the name change proof needed for?"
    Else If (InputBoxEntry = "SeasonalOffSeasonMissing")
        PromptText := "Who is the employer? (optional)"
    Else if (InputBoxEntry = "OverIncomeMissing")
        PromptText := "Without dollar signs, enter the calculated income less expenses, income limit, and household size.`nOnly type numbers separated by spaces - no commas or periods.`n`n(Example: 76392 49605 3)"

    InputBox, %InputBoxEntry%Input, Additional Input Required, %PromptText%, , , , , , , ,%InputBoxAGUIControlVariable%
    If ErrorLevel
        Return

    InputBoxAGUIControlVariable := %InputBoxEntry%Input
    GuiControlGet, VerificationName,,%InputBoxEntry%, Text
    If (InStr(VerificationName, "(")) {
        RemoveAmount := InStr(VerificationName, " (")
        VerificationName := SubStr(VerificationName, 1, RemoveAmount-1)
    }
    If (InStr(VerificationName, " for ")) {
        RemoveAmount := InStr(VerificationName, " for ")
        VerificationName := SubStr(VerificationName, 1, RemoveAmount-1)
    }
    If (InStr(VerificationName, " at ")) {
        RemoveAmount := InStr(VerificationName, " at ")
        VerificationName := SubStr(VerificationName, 1, RemoveAmount-1)
    }

    If (InputBoxEntry = "ChildSupportNoncooperationMissing")
        GuiControl,,%InputBoxEntry%, %VerificationName% - CS phone: %InputBoxAGUIControlVariable%
    Else If (InputBoxEntry = "ChildSupportFormsMissing") {
        SetWording :=
        If (InputBoxAGUIControlVariable ~= "\d") {
            setWording .= InputBoxAGUIControlVariable < 2 ? " set" : " sets"
        } else {
            SetWording := ""
        }
        GuiControl,,%InputBoxEntry%, Child Support forms - %InputBoxAGUIControlVariable% %SetWording%
    }
    Else if (InputBoxEntry = "OverIncomeMissing")
        GoSub OverIncomeSub
    Else 
        GuiControl,,%InputBoxEntry%, %VerificationName% for %InputBoxAGUIControlVariable%
Return

OverIncomeSub:
    overIncomeEntriesArray := StrSplit(InputBoxAGUIControlVariable, A_Space, ",", -1)
    If (StrLen(overIncomeEntriesArray[3]) > 0) {
        OverIncomeObj.overIncomeHHsize := overIncomeEntriesArray[3]
    }
    OverIncomeObj.overIncomeReceived := Round(StrReplace(overIncomeEntriesArray[1], ","))
    OverIncomeObj.overIncomeLimit := StrReplace(overIncomeEntriesArray[2], ",")
    OverIncomeObj.overIncomeText := OverIncomeObj.overIncomeLimit ", your income is calculated as $" OverIncomeObj.overIncomeReceived
    OverIncomeObj.overIncomeDifference := OverIncomeObj.overIncomeReceived - OverIncomeObj.overIncomeLimit
    GuiControl,, %InputBoxEntry%, % "Over-income by $" OverIncomeObj.overIncomeDifference
Return
getRowCount(Text, columns, indentString) {
    indentString := StrLen(indentString) > 0 ? indentString : ""
    Text := st_wordwrap(Text, columns, indentString)
    StrReplace(Text, "`n", "`n", xCount)
    Return [Text, xCount +1]
}
;=============================================================================================================================
;VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION VERIFICATION SECTION 
;=============================================================================================================================



;=========================================================================================================================================================
;ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION
;=========================================================================================================================================================

MissingGuiGuiClose:
    Gosub, SaveMissingGuiGUICoords
	Gui, MissingGui: Hide
Return

MainGuiGuiClose:
    GoSub, SaveMainGuiCoords
    Gosub, SaveMissingGuiGUICoords
    GoSub, ClearFormButton
Return

SaveMainGuiCoords:
	WinGetPos, XCaseNotesGet, YCaseNotesGet,,, CaseNotes
	If ((XCaseNotesGet - XCaseNotes) <> 0 && Abs(XCaseNotesGet) < 9999)
		IniWrite, %XCaseNotesGet%, %A_MyDocuments%\AHK.ini, CaseNotePositions, XCaseNotesINI
	If ((YCaseNotesGet - YCaseNotes) <> 0 && Abs(YCaseNotesGet) < 9999)
		IniWrite, %YCaseNotesGet%, %A_MyDocuments%\AHK.ini, CaseNotePositions, YCaseNotesINI
Return

SaveMissingGuiGUICoords:
	WinGetPos, XVerificationsGet, YVerificationsGet,,, Missing Verifications
	If ((XVerificationsGet - XVerification) <> 0 && Abs(XVerificationsGet) < 9999)
		IniWrite, %XVerificationsGet%, %A_MyDocuments%\AHK.ini, CaseNotePositions, XVerificationINI
	If ((YVerificationsGet - YVerification) <> 0 && Abs(YVerificationsGet) < 9999)
		IniWrite, %YVerificationsGet%, %A_MyDocuments%\AHK.ini, CaseNotePositions, YVerificationINI
Return

ClearFormButton:
    GoSub, SaveMainGuiCoords
    Gosub, SaveMissingGuiGUICoords
    ClearingForm := A_GuiControl = "ClearFormButton" ? 1 : 0
    If (ConfirmedClear > 0) {
        run %A_ScriptName%
    }
    PromptText :=
    If (CaseNoteEntered.MEC2Note = 0) {
        PromptText .= " MEC2"
    }
    If (CaseNoteEntered.MAXISNote = 0 && CountyNoteInMaxisRead = 1 && CaseDetails.DocType = "Application") {
        If (StrLen(PromptText) > 0) {
            PromptText .= " or"
        }
        PromptText .= " MAXIS"
    }
    If (ClearingForm) {
        If (StrLen(PromptText) > 0) {
            MsgBox, 4, Case Note Prompt, Case note not entered in%PromptText%. `nClear form anyway?
            IfMsgBox Yes
                run %A_ScriptName%
            Return
        }
        GuiControl, MainGui: Text, ClearFormButton, Confirm
        Gui, Font, s9, Segoe UI
        GuiControl, MainGui: Font, ClearFormButton
        ConfirmedClear++
    }
    If (!ClearingForm) {
        If (StrLen(PromptText) = 0)
            ExitApp
        If (StrLen(PromptText) > 0) {
            MsgBox, 4, Case Note Prompt, Case note not entered in%PromptText%. `nExit anyway?
            IfMsgBox Yes
                ExitApp
            Return
        }
    }
Return

;=====================================================================================================================================================
;SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION
;=====================================================================================================================================================
SettingsButton:
    CountyContact := {}
    CountyContact.Dakota := { Email: "EEADOCS@co.dakota.mn.us", Fax: "651-306-3187", ProviderWorker: "651-554-5764", EdBSF: "Training Request for Childcare", CountyNoteInMaxis: 1 }
    CountyContact.StLouis := { Email: "ess@stlouiscountymn.gov", Fax: "218-733-2976", ProviderWorker: "218-726-2064", EdBSF: "SLC CCAP Education Plan", CountyNoteInMaxis: 0 }
    CountyFaxRead := (StrLen(CountyFaxRead) > 0) ? CountyFaxRead : CountyContact[WorkerCounty].Fax
    CountyDocsEmailRead := (StrLen(CountyDocsEmailRead) > 0) ? CountyDocsEmailRead : CountyContact[WorkerCounty].Email
    ProviderWorkerPhoneRead := (StrLen(ProviderWorkerPhoneRead) > 0) ? ProviderWorkerPhoneRead : CountyContact[WorkerCounty].ProviderWorker
    CountyEdBSFformRead := (StrLen(CountyEdBSFformRead) > 0) ? CountyEdBSFformRead : CountyContact[WorkerCounty].EdBSF
    UseWorkerEmailReadini := (InStr(UseWorkerEmailRead, "0")) ? "Checked0" : "Checked"
    UseMec2FunctionsReadini := (InStr(UseMec2FunctionsRead, "0")) ? "Checked0" : "Checked"
    CountyNoteInMaxisReadini := (InStr(CountyNoteInMaxisRead, "0")) ? "Checked0" : "Checked"
    
    EditboxOptions := "x200 yp-3 h18 w200"
    CheckboxOptions := "x200 yp-3 h18 w20"
    TextLabelOptions := "xm w170 h18 Right"
    Gui, Font,, Lucida Console
    Gui, Color, 989898, a9a9a9
    Gui, CSG: Margin, 12 12
    Gui, CSG: New, AlwaysOnTop ToolWindow,
    Gui, CSG: Add, Text, %TextLabelOptions% y12,Worker Name:
    Gui, CSG: Add, Edit, %EditboxOptions% vWorkerNameWrite,%WorkerNameRead%
    Gui, CSG: Add, Text, %TextLabelOptions%,Worker Phone:
    Gui, CSG: Add, Edit, %EditboxOptions% vWorkerPhoneWrite,%WorkerPhoneRead%
    Gui, CSG: Add, Text, %TextLabelOptions%,Worker Email:
    Gui, CSG: Add, Edit, %EditboxOptions% vWorkerEmailWrite,%WorkerEmailRead%
    Gui, CSG: Add, Text, %TextLabelOptions%,Use Worker Email in Letters:
    Gui, CSG: Add, CheckBox, %CheckboxOptions% vWorkerEmailYesWrite %UseWorkerEmailReadini%
    Gui, CSG: Add, Text, %TextLabelOptions%,Using mec2functions:
    Gui, CSG: Add, CheckBox, %CheckboxOptions% vWorkerUsingMec2FunctionsWrite gWorkerUsingMec2Functions %UseMec2FunctionsReadini%
    Gui, CSG: Add, ComboBox, x+10 yp vWorkerBrowserWrite Choose1 R4 Hidden, %WorkerBrowserRead%|Google Chrome|Mozilla Firefox|Microsoft Edge
    If (UseMec2FunctionsReadini = "Checked") {
        GuiControl, CSG: Show, WorkerBrowserWrite
    }
    Gui, CSG: Add, Text, h0 w0 y+10
    Gui, CSG: Add, Text, %TextLabelOptions%,Select a county to auto-populate
    Gui, CSG: Add, ComboBox, %EditboxOptions% vWorkerCountyWrite gCountySelection Choose1 R4, %WorkerCounty%|Dakota|StLouis|Other
    Gui, CSG: Add, Text, %TextLabelOptions%,Case Note in MAXIS:
    Gui, CSG: Add, CheckBox, %CheckboxOptions% vCountyNoteInMaxisWrite gCountyNoteInMaxis %CountyNoteInMaxisReadini%
    Gui, CSG: Add, Edit, x+10 yp h18 w170 vWorkerMaxisWrite Hidden, %WorkerMaxisRead%
    If (CountyNoteInMaxisReadini = "Checked") {
        GuiControl, CSG: Show, WorkerMaxisWrite
    }
    Gui, CSG: Add, Text, %TextLabelOptions%,Fax Number:
    Gui, CSG: Add, Edit, %EditboxOptions% vCountyFaxWrite, %CountyFaxRead%
    Gui, CSG: Add, Text, %TextLabelOptions%,County Documents Email:
    Gui, CSG: Add, Edit, %EditboxOptions% vCountyDocsEmailWrite, %CountyDocsEmailRead%
    Gui, CSG: Add, Text, %TextLabelOptions%,Provider Worker Phone:
    Gui, CSG: Add, Edit, %EditboxOptions% vCountyProviderPhoneWrite, %ProviderWorkerPhoneRead%
    Gui, CSG: Add, Text, %TextLabelOptions%,BSF Education Form Name:
    Gui, CSG: Add, Edit, %EditboxOptions% vCountyEdBSFformWrite, %CountyEdBSFformRead%
    Gui, CSG: Add, Button, w80 gUpdateIniFile, Save
    Gui, CSG: Show,w450,Change Settings
Return

WorkerUsingMec2Functions:
    Gui, Submit, NoHide
    GuiControlGet, WorkerBrowser,,WorkerUsingMec2FunctionsWrite
    If (WorkerBrowser = 0) {
        GuiControl, CSG: Hide, WorkerBrowserWrite
        Return
    }
    GuiControl, CSG: Show, WorkerBrowserWrite
Return
CountyNoteInMaxis:
    Gui, Submit, NoHide
    GuiControlGet, MaxisChecked,,CountyNoteInMaxisWrite
    If (MaxisChecked = 0) {
        GuiControl, CSG: Hide, WorkerMaxisWrite
        Return
    }
    GuiControl, CSG: Show, WorkerMaxisWrite
Return

UpdateIniFile:
    Gui, Submit, NoHide
    ;If (CountyNoteInMaxisWrite && WorkerMaxisWrite = "MAXIS-WINDOW-TITLE") {
        ;change border of WorkerMaxisWrite, blink, dance, return?
    ;}
    IniWrite, %WorkerNameWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeName
    WorkerNameRead := WorkerNameWrite
    
    IniWrite, %WorkerPhoneWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeePhone
    WorkerPhoneRead := WorkerPhoneWrite
    
    IniWrite, %WorkerEmailWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeEmail
    WorkerEmailRead := WorkerEmailWrite
    
    IniWrite, %WorkerEmailYesWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeUseEmail
    UseWorkerEmailRead := WorkerEmailYesWrite
    
    IniWrite, %WorkerUsingMec2FunctionsWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeUseMec2Functions
    UseMec2FunctionsRead := WorkerUsingMec2FunctionsWrite
    
    IniWrite, %WorkerBrowserWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeBrowser
    WorkerBrowserRead := WorkerBrowserWrite
    
    IniWrite, %WorkerCountyWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeCounty
    WorkerCounty := WorkerCountyWrite

    IniWrite, %CountyNoteInMaxisWrite%, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyNoteInMaxis
    CountyNoteInMaxisRead := CountyNoteInMaxisWrite
    
    IniWrite, %WorkerMaxisWrite%, %A_MyDocuments%\AHK.ini, EmployeeInfo, EmployeeMaxis
    WorkerMaxisRead := WorkerMaxisWrite
    
    IniWrite, %CountyFaxWrite%, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyFax
    CountyFaxRead := CountyFaxWrite
    
    IniWrite, %CountyDocsEmailWrite%, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyDocsEmail
    CountyDocsEmailRead := CountyDocsEmailWrite
    
    IniWrite, %CountyProviderPhoneWrite%, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyProviderWorkerPhone
    ProviderWorkerPhoneRead := CountyProviderPhoneWrite
    
    IniWrite, %CountyEdBSFformWrite%, %A_MyDocuments%\AHK.ini, CaseNoteCountyInfo, CountyEdBSF
    CountyEdBSFformRead := CountyEdBSFformWrite

    Gui, Destroy
Return

CountySelection:
    Gui, Submit, NoHide
    WorkerCounty := WorkerCountyWrite
    GuiControl, CSG: Text, CountyFaxWrite, % CountyContact[WorkerCounty].Fax
    GuiControl, CSG: Text, CountyProviderPhoneWrite, % CountyContact[WorkerCounty].ProviderWorker
    GuiControl, CSG: Text, CountyDocsEmailWrite, % CountyContact[WorkerCounty].Email
    GuiControl, CSG: Text, CountyEdBSFformWrite, % CountyContact[WorkerCounty].EdBSF
    GuiControl,, CountyNoteInMaxisWrite, % CountyContact[WorkerCounty].CountyNoteInMaxis
    GoSub, CountyNoteInMaxis
Return

;================================================================================================================================================================
;BORROWED FUNCTIONS SECTION BORROWED FUNCTIONS SECTION BORROWED FUNCTIONS SECTION BORROWED FUNCTIONS SECTION 
st_wordWrap(string, column, indentChar) { ; String Things - Common String & Array Functions
    indentLength := StrLen(indentChar)
    Loop, Parse, string, `n, `r
    {
        If (StrLen(A_LoopField) > column) {
            pos := 1
            Loop, Parse, A_LoopField, %A_Space% ; A_LoopField is the individual word
            {
                If (pos + (loopLength := StrLen(A_LoopField)) <= column) {
                    out .= (A_Index = 1 ? "" : " ") A_LoopField
                    , pos += loopLength + 1
                } Else {
                    pos := loopLength + 1 + indentLength
                    , out .= "`n" indentChar A_LoopField
                }
            }
            out .= "`n"
        } Else {
            If (StrLen(A_LoopField) > 0) {
                out .= indentChar A_LoopField "`n"
            }
        }
    }
    Return SubStr(out, 1, -1)
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
!3::
    Gui, Submit, NoHide
    Clipboard := CaseNumber
Return

#v::
    If WinActive(ahk_exe Outlook.exe) {
        SendInput, % EmailText.Output
    }
    Else If WinActive(WorkerBrowserRead) {
        Gui, Submit, NoHide
        Sleep 500
        If (UseMEC2FunctionsRead = 1) {
            CaseStatus := InStr(CaseDetails.DocType, "?") ? "" : (Homeless = 1) ? "Homeless App" : (CaseDetails.DocType = "Redet") ? "Redetermination" : CaseDetails.DocType
            concatLetterText := "LetterTextFromAHKSPLIT" %Letter1% "SPLIT" CaseStatus "SPLIT" IdList
            Clipboard := concatLetterText
            Send, ^v
        } Else {
            Clipboard := %Letter1%
            Send, ^v
        }
        Sleep 500
        Clipboard := CaseNumber
    }
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
    Gui, CBT: Show, x%XClipboard% y%YClipboard%, Clipboard_Text
    ControlSend,,{End}, Clipboard_Text
Return

#IfWinActive, ahk_class AutoHotkeyGUI
    ^PgDn::
        ControlFocus,,\d ahk_exe obunity.exe
        ControlSend,,^{PgDn}, \d ahk_exe obunity.exe
        If (!WinActive(ahk_class AutoHotkeyGUI)) {
            WinActivate ahk_class AutoHotkeyGUI
        }
    Return
    ^PgUp::
        ControlFocus,,\d ahk_exe obunity.exe
        ControlSend,,^{PgUp}, \d ahk_exe obunity.exe
        If (!WinActive(ahk_class AutoHotkeyGUI)) {
            WinActivate ahk_class AutoHotkeyGUI
        }
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

#If WinActive(WorkerBrowserRead)
    ^F12:: ;CtrlF12/AltF12 Add worker signature
    !F12::
        SendInput `n=====`n
        Send, %WorkerNameRead%
    Return
#If

If (WorkerCounty = "Dakota") {
    #If WinActive("ahk_exe WINWORD.EXE")
        F1::
            ToolTip,
            (
    Alt+4: Starting from the name field, moves to and enters date,
             case number, and client's first name.
            ), 0, 0
            SetTimer, RemoveToolTip, -5000
        Return
        !4::
            Gui, Submit, NoHide
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
    #If WinActive("Perform Import")
        F1:: 
            ToolTip, CTRL+ `n F6: RSDI `n F7: SMI ID `n F8: PRISM GCSC `n F9: CS $ Calc `nF10: Income Calc `nF11: The Work # `nF12: CCAPP Letter, 0, 0
            SetTimer, RemoveToolTip, -8000
        Return
        ^F6::
            Gui, Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 ssi", "{Text}RSDI ", 3, "Member#, Member Name")
        Return
        ^F7::
            Gui, Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3001 other id", "{Text}SMI ", 3, "Member#, Member Name")
        Return
        ^F8::
            Gui, Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 child support", "{Text}GCSC ", 1, "Y/N, Child(ren) Member#")
        Return
        ^F9::
            Gui, Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 wo", "{Text}CCAP CS INCOME CALC")
        Return
        ^F10::
            Gui, Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 wo", "{Text}CCAP INCOME CALC")
        Return
        ^F11::
            Gui, Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 other - in", "{Text}W# ", 3, "Member#, Employer")
        Return
        ^F12::
            Gui, Submit, NoHide
            OnBaseImportKeys(CaseNumber, "3003 edak 3813", "{Text}OUTBOUND")
        Return
    #If

    #If WinActive("Automated Mailing Home Page") || WinActive("ahk_exe obunity.exe",,"Perform Import") ; OnBase, excluding "Perform Import"
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
            Gui, Submit, NoHide
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

    #If WinActive(WorkerMaxisRead)
        ^m::
            WinSetTitle, ahk_exe bzmd.exe,,S1 - MAXIS
            Send ^{m}
        Return
    #If


    #If WinActive(WorkerBrowserRead)
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
            If (UseMec2FunctionsRead = 0)
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
            SendInput, Reviewed case for verifications that are required at application. Verifications were received.`n-`nApproved eligible results effective DATEAPPROVEDGOESHERE.`n-`nService Authorization APPROVEDorNOTandEFFECTIVEDATEIFAPPROVED.`n=====`n%WorkerNameRead%
        Return

        !F2::
            If (UseMec2FunctionsRead = 0)
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
            SendInput, Reviewed case for documents that are required at application. Documents were not received.`n-`nApplication was denied by MEC2 and remains denied.`n=====`n%WorkerNameRead%
            Sleep 1000,
            Send, !{s}
        Return
    #If
}
DisplayResult(Result) {
    ToolTip, %Result%
    SetTimer, RemoveToolTip, -2000
}