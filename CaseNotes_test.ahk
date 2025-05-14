; Note: This script requires BOM encoding (UTF-8) to display characters properly. 

;version 0.1.1, The 'Sending Case Note Update' version
;version 0.1.2, The 'Dude, Where's My GUI' version
;version 0.1.3, the 'Dakota County' retro version
;version 0.1.4, The 'I think I got all the auto-populating date variables correct THIS time' version
;version 0.1.5, The 'Extra Large serving of date variables because some counties are so far behind' version
;version 0.1.6, The ‘It seems to be working so I fixed some of the MAXIS case note text and also changed things behind the scenes that nobody else will notice’ version
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
;version 0.2.9, The 'Imported hotkeys from mec2Hotkeys' version
;version 0.3.0, The 'Look ma, I'm on GitHub' version
;version 0.3.1, The 'I am so tired of typing questions from redeterminations into Other because someone missed Lump Sum again' version (soonTM)
;version 0.3.2, The 'I don't know what your MAXIS screen is called' version
;version 0.3.3, The 'I had to rearrage code so that Homeless prompting would be added to the Special Letter' version
;version 0.3.4, The 'I prettied up Missing Verifications, changed MV to open/hide on load, and gave each GUI a name' version
;version 0.3.5, The 'If the Special Letter line count ain't right now, it ain't ever gonna be' version
;version 0.3.6, The 'Added some parens and fixed some copy/paste errors in the MissingGuiGUIClose subroutine' version
;version 0.3.7, The 'Redid how coordinates were done. Next is employeeInfo/caseNoteCountyInfo ini' version
;version 0.4.0, The 'I rewrote how settings were done. Fewer read/writes to harddrive. Hurray! Also added Waitlist functionality.' version
;version 0.4.2, The 'Holy crap I finally figured out how to fix the Gui Submit issue.' version
;version 0.4.3, The 'I changed most AHK built-in function commands to % variable "string"' version
;version 0.5.0, The 'Every subroutine was rewritten as a function and it still works' version
Version := "v0.5.83"

;Future todo ideas:
;Add backup to ini for Case Notes window. Check every minute old info vs new info and write changes to .ini.
;Make a restore button.
;Import from clipboard (when copied from MEC2) (likely mostly same code as restore button)
;Restructure MissingVerifications
;'Other': Add a checkbox to check to make it an asterisk instead of numbered item

#Requires AutoHotkey v1+
SetWorkingDir % A_ScriptDir
#Persistent
#SingleInstance force
#NoTrayIcon
SetTitleMatchMode, RegEx

; Rule for AHKv1 GUI functions and variables: If you are doing a "Gui, Submit" the function needs to be declared Global.

Global verbose := A_ScriptName == "CaseNotes_dev.ahk" ? 1 : 0, sq := "²", cm := "✔", cs := ", "
Global signRecDateVisible, signRecDateVisible1, signRecDateVisible2, signRecDateVisible3

;Settings
Global ini := { cbtPositions: { xClipboard: 0, yClipboard: 0 }
    , caseNotePositions: { xCaseNotes: 0, yCaseNotes: 0, xVerification: 0, yVerification: 0 }
    , caseNoteCountyInfo: { countyNoteInMaxis: 0, countyFax: A_Space, countyDocsEmail: A_Space, countyProviderWorkerPhone: A_Space, countyEdBSF: A_Space, Waitlist: 1 }
    , employeeInfo: { employeeName: A_Space, employeeCounty: A_Space, employeeEmail: A_Space, employeePhone: A_Space, employeeUseEmail: 0, employeeUseMec2Functions: 0, employeeBrowser: A_Space, employeeMaxis: MAXIS-WINDOW-TITLE } }
setFromIni()
GroupAdd, autoMailGroup, % "Automated Mailing Home Page"
GroupAdd, autoMailGroup, % "ahk_exe obunity.exe",,, % "Perform Import"
checkGroupAdd()

Global CH100x30 := guiEditAreaSize("1234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890`n2`n3`n4`n5`n6`n7`n8`n9`n0`n1`n2`n3`n4`n5`n6`n7`n8`n9`n0`n1`n2`n3`n4`n5`n6`n7`n8`n9`n0", "s9", "Lucida Console"),
    wCH := CH100x30[1]/100, hCH := CH100x30[2]/30
Global CH87 := guiEditAreaSize("123456789012345678901234567890123456789012345678901234567890123456789012345678901234567`n2", "s9", "Lucida Console"), wCH87 := CH87[1]
Global zoomPPI := A_ScreenDPI/96
Global pad := 2, scrollbar := 24, margin := 12, oneRow := "h" hCH+pad*2 " Limit87", twoRows := "h" (hCH*2.7)+pad, threeRows := "h" (hCH*3)+pad*3, fourRows := "h" (hCH*4)+pad*3

Global countySpecificText := { StLouis: { OverIncomeContactInfo: "", CountyName: "St. Louis County" }
    , Dakota: { OverIncomeContactInfo: " contact 651-554-6696 and", CountyName: "Dakota County", customHotkeys: "
(
    Custom hotkeys for your county exist for the following windows (A ToolTip reminder appears by pressing F1):
    ● OnBase (Alt+4: 'Verifs Due Back' detail)
    ● OnBase (Ctrl+F6-12: Enters keywords on the Perform Import screen)
    ● OnBase (Ctrl+B: Inserts date and case number for mail)
    ● Automated Mailing (Ctrl+B: Inserts date and case number for mail)
    ● Browser (Types an Approved (Alt+F1) or Denied (Alt+F2) app case note. Ctrl+F12 or Alt+F12: worker signature)
    ● Word (Alt+4: Types in first name, case number, and app received date)
    ● MAXIS (Alt+M: Changes the title of the MAXIS window to ""MAXIS"" to enable screen-scraping)
)" } }

;Base globals
Global caseDetails := { docType: "_DOC?", eligibility: "_ELIG?", saEntered: "_SA?", caseType: "_PRG?", appType: "_APP?", haveWaitlist: false, newChanges: true }
Global caseNoteEntered := { mec2NoteEntered: 0, maxisNoteEntered: 0, confirmedClear: 0 }
Global dateObject := { todayYMD: A_Now, todayMDY: formatMDY(A_Now), receivedMDY: "", receivedYMD: "", autoDenyYMD: "" }
Global autoDenyObject := { autoDenyExtensionMECnote: "", autoDenyExtensionDate: "", autoDenyExtensionSpecLetter: "" }
Global editControls := ["01HouseholdCompEdit", "02SharedCustodyEdit", "03AddressVerificationEdit", "04SchoolInformationEdit", "05IncomeEdit", "06ChildSupportIncomeEdit"
    , "07ChildSupportCooperationEdit", "08ExpensesEdit", "09AssetsEdit", "10ProviderEdit", "11ActivityAndScheduleEdit", "12ServiceAuthorizationEdit", "13NotesEdit", "14MissingEdit"]
Global exampleLabels := [ "01HouseholdCompEditLabelExample", "02SharedCustodyEditLabelExample", "03AddressVerificationEditLabelExample", "04SchoolInformationEditLabelExample", "05IncomeEditLabelExample", "06ChildSupportIncomeEditLabelExample"
    , "07ChildSupportCooperationEditLabelExample", "08ExpensesEditLabelExample", "09AssetsEditLabelExample", "10ProviderEditLabelExample", "11ActivityAndScheduleEditLabelExample", "12ServiceAuthorizationEditLabelExample", "14MissingEditLabelExample" ]
Global emailText := { stillRequiredText: "These verifications are still required, and must be received within 90 days of your application date for continued eligibility.", pendingHomelessPreText: "You may be eligible for the homeless policy, which allows us to approve eligibility even though there are verifications we need but do not have. " emailText.stillRequiredText "`n`nBefore we can approve expedited eligibility, we need information that was not on the application:", initialApproval: "`nThe initial approval of child care assistance is 30 hours per week for each child. This amount can be increased once we receive your activity verifications and we determine more assistance is needed.`nIf the provider you select is a “High Quality” provider, meaning they are Parent Aware 3⭐ or 4⭐ rated, or have an approved accreditation, the hours will automatically increase to 50 per week for preschool age and younger children.`nIf you have a 'copay,' the amount the county pays to the provider will be reduced by the copay amount. Many providers charge more than our maximum rates, and you are responsible for your copay and any amounts the county cannot pay.", approvedWithMissing: "It was approved under the homeless expedited policy which allows us to approve eligibility even though there are verifications we require that we do not have. " }
setIcon()
;Missing Verification globals
Global emailTextObject := {}, missingInput := {}, otherMissing := {}, letterText := {}, letterTextNumber := 1, missingHomelessItems := "", idList := "", lineCount := 0
Global overIncomeObj := { overIncomeHHsize: "your size" }
Global missingInputObject := { IDmissing: { baseText: "ID", inputAdject: " for ", promptText: "Who is ID needed for?`n`nExample: 'Susanne, Robert Sr'", strRem: 3 }
    , BCmissing: { baseText: "BC", inputAdject: " for ", promptText: "Who is birth verification needed for?`n`nExample: 'Susie, Bobby Jr'" }
    , BCNonCitizenMissing: { baseText: "BC [non-citizen]", inputAdject: " for ", promptText: "Who is birth verification needed for?`n`nExample: 'Susie, Bobby Jr'" }
    , PaternityMissing: { baseText: "Paternity", inputAdject: " for ", promptText: "Who is paternity verification needed for?`n`nExample: 'Susie, Bobby Jr' or 'Robert / Bobby Jr'" }
    , IncomePlusNameMissing: { baseText: "Income", inputAdject: " for ", promptText: "Who is the income verification needed for?" }
    , CustodySchedulePlusNamesMissing: { baseText: "Custody", inputAdject: " for ", promptText: "Who is the schedule needed for? `n'...stating the current parenting time schedule for: ____________'`n`nExample: 'Susie and Bobby Jr' or 'your children'" }
    , WorkSchedulePlusNameMissing: { baseText: "Work Schedule", inputAdject: " for ", promptText: "Who is the work schedule needed for?" }
    , DependentAdultStudentMissing: { baseText: "Dependent adult child - FT Student, 50{U+0025}+ expenses", inputAdject: " for ", promptText: "Who is the adult dependent student?" }
    , ChildSupportFormsMissing: { baseText: "Child Support forms", inputAdject: " - ", promptText: "Enter the number of sets of Child Support forms needed`nor the names of the absent parent/children.`n`nExample: 'Robert / Susie, Bobby Jr' or '2'" }
    , ChildSupportNoncooperationMissing: { baseText: "CS Non-cooperation", inputAdject: " - CSO phone: ", promptText: "What is the phone number of the Child Support officer?" }
    , LegalNameChangeMissing: { baseText: "Name change", inputAdject: " for ", promptText: "Who is the name change proof needed for?" }
    , SeasonalOffSeasonMissing: { baseText: "Seasonal employment info - app in off-season", inputAdject: " for ", promptText: "Who is the employer? (optional)" }
    , overIncomeMissing: { baseText: "Over-income", inputAdject: " by $", promptText: "Without dollar signs, enter the calculated income less expenses, income limit, and household size.`nOnly type numbers separated by spaces - no commas or periods.`n`n(Example: 76392 49605 3)" } }


buildMainGui()
buildMissingGui()
buildSettingsGui(1)
Global caseNotesMonCenter := getMonCenter("CaseNotes")

Return

;==============================================================================================================================================================================================
;MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION
buildMainGui() {
    Global
    Gui +OwnDialogs
    local labelSettings := "xm+5 y+1 w200"
    local labelExampleSettings := "x220 yp+4 h12 w" wCH*60 " "
    local textboxSettings := "xm y+1 w" (wCH87)+scrollbar+pad

    Gui, MainGui: Font,, % "Segoe UI"
    Gui, MainGui: Color, % "a9a9a9", % "bebebe"

    Gui, MainGui: Add, Radio, % "Group Section h17 x12 w75 y+5 gsetDocType vApplicationRadio", % "Application"
    Gui, MainGui: Add, Radio, % "xp y+2 wp h17 gsetDocType vRedeterminationRadio", % "Redeterm."
    Gui, MainGui: Add, Checkbox, % "xp y+2 wp h17 vHomeless gnewChangesTrue", % "Homeless"

    Gui, MainGui: Add, Radio, % "Group x+10 ys h17 w78 gsetAppType vMNBenefitsRadio", % "MNBenefits"
    Gui, MainGui: Add, Radio, % "xp y+2 h17 wp gsetAppType vPaperAppRadio", % "3550 App"

    Gui, MainGui: Add, Radio, % "Group x+10 ys h17 w58 gsetCaseType vBSF", % "BSF"
    Gui, MainGui: Add, Radio, % "xp y+2 h17 wp gsetCaseType vTY", % "TY"
    Gui, MainGui: Add, Radio, % "xp y+2 h17 wp gsetCaseType vCCMF", % "CCMF"

    Gui, MainGui: Add, Radio, % "Group x+0 ys h17 w80 gsetEligibility vPendingRadio", % "Pending"
    Gui, MainGui: Add, Radio, % "xp y+2 h17 wp gsetEligibility vEligibleRadio", % "Eligible"
    Gui, MainGui: Add, Radio, % "xp y+2 h17 wp gsetEligibility vIneligibleRadio", % "Ineligible"

    Gui, MainGui: Add, Radio, % "Group Hidden x+5 ys h17 vSaApproved gsetSA", % "SA Approved"
    Gui, MainGui: Add, Checkbox, % "Group Hidden xp yp h17 vmanualWaitlistBox", % "Waitlist"
    Gui, MainGui: Add, Radio, % "Hidden xp y+2 h17 vNoSA gsetSA", % "No SA"
    Gui, MainGui: Add, Radio, % "Hidden xp y+2 h17 vNoProvider gsetSA", % "No Provider"

    Gui, MainGui: Add, Text, % "xp-18 y+9 w200 vautoDenyStatus",

    Gui, MainGui: Add, Text, % "x420 w35 h20 ys+2", % "Case #"
    Gui, MainGui: Add, Text, % "xp y+2 w35 h20", % "Rec'd:"
    Gui, MainGui: Add, Text, % "xp y+2 w35 h20 vSignOrDueLabel Hidden", % "Signed:"

    Gui, MainGui: Add, Edit, % "x+0 ys w70 h17 -Background Limit8 vcaseNumber",
    Gui, MainGui: Add, DateTime, % "xp y+5 w70 h17 vReceivedDate gnewChangesTrue", % "M/d/yy"
    Gui, MainGui: Add, DateTime, % "xp y+5 w70 h17 vSignOrDueDate gnewChangesTrue Hidden", % "M/d/yy"

    Gui, MainGui: Add, Button, % "Section x535 ys+0 h17 w70 -TabStop vmec2NoteButton goutputCaseNote", % "MEC" sq " Note"
    Gui, MainGui: Add, Button, % "xs y+5 h17 w70 -TabStop Hidden vmaxisNoteButton goutputCaseNote", % "MAXIS Note"
    Gui, MainGui: Add, Button, % "xs y+5 h17 w70 -TabStop vnotepadNoteButton goutputCaseNote", % "To Desktop"
    Gui, MainGui: Add, Button, % "Section x615 ys+0 h17 w50 -TabStop vClearFormButton gMainGuiGuiClose", % "Clear"

    Gui, MainGui: Font, s9, % "Segoe UI"
    Gui, MainGui: Margin, % marginW
    Gui, MainGui: Add, Text, % "xm y+45 h0 w0" ; Blank space
    Gui, MainGui: Add, Text, % labelSettings " v01HouseholdCompEditLabel", % "Household Comp"
    Gui, MainGui: Add, Text, % labelExampleSettings " v01HouseholdCompEditLabelExample Hidden", % "Parent (ID), ChildOne (4, BC), ChildName (age, verif)"
    Gui, MainGui: Add, Edit, % textboxSettings " " twoRows " v01HouseholdCompEdit", ;v2 add LoseFocus gui_event to force spacing.

    Gui, MainGui: Add, Text, % labelSettings " v03AddressVerificationEditLabel", % "Address Verification"
    Gui, MainGui: Add, Text, % labelExampleSettings " v03AddressVerificationEditLabelExample Hidden", % "1234 W Minnesota St APT 21, St Paul: ID 5/4/20 (scan date)"
    Gui, MainGui: Add, Edit, % textboxSettings " " threeRows " v03AddressVerificationEdit",

    Gui, MainGui: Add, Text, % labelSettings " v02SharedCustodyEditLabel", % "Shared Custody"
    Gui, MainGui: Add, Text, % labelExampleSettings " v02SharedCustodyEditLabelExample Hidden", % "Absent Parent / Child: Thursday 6pm - Monday 7am"
    Gui, MainGui: Add, Edit, % textboxSettings " " fourRows " v02SharedCustodyEdit",

    Gui, MainGui: Add, Text, % labelSettings " v04SchoolInformationEditLabel", % "School information"
    Gui, MainGui: Add, Text, % labelExampleSettings " v04SchoolInformationEditLabelExample Hidden", % "ChildOne, ChildTwo: Wildcat Elementary, M-F 730am - 2pm"
    Gui, MainGui: Add, Edit, % textboxSettings " " threeRows " v04SchoolInformationEdit",

    Gui, MainGui: Add, Text, % labelSettings " v05IncomeEditLabel", % "Income"
    Gui, MainGui: Add, Text, % labelExampleSettings " v05IncomeEditLabelExample Hidden Border", % "Parent - Job: BW avg $1234.56, 43.2hr/wk; annual @ 32098.56"
    Gui, MainGui: Add, Edit, % textboxSettings " " fourRows " v05IncomeEdit",

    Gui, MainGui: Add, Text, % labelSettings " v06ChildSupportIncomeEditLabel", % "Child Support Income"
    Gui, MainGui: Add, Text, % labelExampleSettings " v06ChildSupportIncomeEditLabelExample Hidden", % "6 month total $2345.67; annual @ 4691.34"
    Gui, MainGui: Add, Edit, % textboxSettings " " twoRows " v06ChildSupportIncomeEdit",

    Gui, MainGui: Add, Text, % labelSettings " v07ChildSupportCooperationEditLabel Border gcopySharedCustodyEditToCSCoopEdit", % "Child Support Cooperation"
    Gui, MainGui: Add, Text, % labelExampleSettings " v07ChildSupportCooperationEditLabelExample Hidden", % "Absent Parent / Child: Open, cooperating"
    Gui, MainGui: Add, Edit, % textboxSettings " " fourRows " v07ChildSupportCooperationEdit",

    Gui, MainGui: Add, Text, % labelSettings " v08ExpensesEditLabel", % "Expenses"
    Gui, MainGui: Add, Text, % labelExampleSettings " v08ExpensesEditLabelExample Hidden", % "BW Medical $121.23, BW Dental $12.23, BW Vision $2.23"
    Gui, MainGui: Add, Edit, % textboxSettings " " twoRows " v08ExpensesEdit",

    Gui, MainGui: Add, Text, % labelSettings " v09AssetsEditLabel", % "Assets"
    Gui, MainGui: Add, Text, % labelExampleSettings " v09AssetsEditLabelExample Hidden", % "< $1m   or   (blank)"
    Gui, MainGui: Add, Edit, % textboxSettings " " oneRow " Limit87 v09AssetsEdit",

    Gui, MainGui: Add, Text, % labelSettings " v10ProviderEditLabel", % "Provider"
    Gui, MainGui: Add, Text, % labelExampleSettings " v10ProviderEditLabelExample Hidden", % "Kid Kare (PID#, HQ): ChildOne, ChildTwo - Start date 5/4/20"
    Gui, MainGui: Add, Edit, % textboxSettings " " twoRows " v10ProviderEdit",

    Gui, MainGui: Add, Text, % labelSettings " v11ActivityAndScheduleEditLabel", % "Activity and Schedule"
    Gui, MainGui: Add, Text, % labelExampleSettings " v11ActivityAndScheduleEditLabelExample Hidden", % "ParentOne - Employment: M-F 9a - 5p (8h x 5d)"
    Gui, MainGui: Add, Edit, % textboxSettings " " fourRows " v11ActivityAndScheduleEdit", 

    Gui, MainGui: Add, Text, % labelSettings " v12ServiceAuthorizationEditLabel", % "Service Authorization"
    Gui, MainGui: Add, Text, % labelExampleSettings " v12ServiceAuthorizationEditLabelExample Hidden", % "8h work + 1h travel = 9h/day, 90h/period"
    Gui, MainGui: Add, Edit, % textboxSettings " " threeRows " v12ServiceAuthorizationEdit", 

    Gui, MainGui: Add, Text, % labelSettings " v13NotesEditLabel", % "Notes"
    Gui, MainGui: Add, Edit, % textboxSettings " " fourRows " v13NotesEdit",

    Gui, MainGui: Add, Text, % "xm+5 y+1 gshowMissingVerifs v14MissingEditLabel Border", % "Missing"
    Gui, MainGui: Add, Text, % labelExampleSettings " v14MissingEditLabelExample Hidden", % "(Click ""Missing"" to bring up the missing verification list)"
    Gui, MainGui: Add, Edit, % textboxSettings " h" hCH*10+pad " v14MissingEdit",

    Gui, MainGui: Add, Text, % "x15 y+4", % Version
    Gui, MainGui: Add, Button, % "x+20 yp w65 h19 -TabStop gbuildSettingsGui", % "Settings"
    Gui, MainGui: Add, Button, % "x+40 yp wp h19 -TabStop gexamplesButton vexamplesButtonText", % "Examples"
    Gui, MainGui: Add, Button, % "x+40 yp wp h19 -TabStop gbuildHelpGui", % "Help"
    Gui, MainGui: Add, Button, % "x600 yp wp h19 gshowMissingVerifs", % "Missing"

    Gui, MainGui: Show, % "x" ini.caseNotePositions.xCaseNotes " y" ini.caseNotePositions.yCaseNotes, CaseNotes
    Gui, MainGui: Show, AutoSize

    For i, editField in editControls {
        Gui, MainGui: Font, s9, % "Lucida Console" ; monospace font
        GuiControl, MainGui: Font, % editField
    }    
    For i, catLabel in exampleLabels {
        Gui, MainGui: Font, s9, % "Lucida Console"
        GuiControl, MainGui: Font, % catLabel
    }
}
; v2 convert to switch
setDocType() {
    Gui, Submit, NoHide
    If (A_GuiControl == "ApplicationRadio") {
        caseDetails.docType := "Application"
        GuiControl, MainGui: Text, PendingRadio, % "Pending"
        GuiControl, MainGui: Text, SignOrDueLabel, % "Signed:"
        If (caseDetails.appType != "3550") {
            GuiControl, MainGui: Hide, SignOrDueLabel
            GuiControl, MainGui: Hide, SignOrDueDate
        }
        If (ini.caseNoteCountyInfo.countyNoteInMaxis == 1) {
            GuiControl, MainGui: Show, maxisNoteButton
        }
        GuiControl, MainGui: Show, MNBenefitsRadio
        GuiControl, MainGui: Show, PaperAppRadio
    } Else If (A_GuiControl == "RedeterminationRadio") {
        caseDetails.docType := "Redet"
        revertLabels()
        GuiControl, MainGui: Text, PendingRadio, % "Incomplete"
        GuiControl, MainGui: Text, SignOrDueLabel, % "Due:"
        GuiControl, MainGui: Show, SignOrDueLabel
        GuiControl, MainGui: Show, SignOrDueDate
        GuiControl, MainGui: Hide, autoDenyStatus
        GuiControl, MainGui: Hide, maxisNoteButton
        GuiControl, MainGui: Hide, MNBenefitsRadio
        GuiControl, MainGui: Hide, PaperAppRadio
    }
    checkWaitlist()
    newChangesTrue()
}
setAppType() {
    Gui, Submit, NoHide
    If (A_GuiControl == "PaperAppRadio") {
        caseDetails.appType := "3550"
        revertLabels()
        GuiControl, MainGui: Show, SignOrDueLabel
        GuiControl, MainGui: Show, SignOrDueDate
    } Else If (A_GuiControl == "MNBenefitsRadio") {
        caseDetails.appType := "MNB"
        GuiControl, MainGui: Text, 01HouseholdCompEditLabel, % "Household Comp (pages 1, 3-5)"
        GuiControl, MainGui: Text, 03AddressVerificationEditLabel, % "Address Verification (page 3)"
        GuiControl, MainGui: Text, 02SharedCustodyEditLabel, % "Absent Parent / Child (page 6)"
        GuiControl, MainGui: Text, 04SchoolInformationEditLabel, % "School Information (page 7)"
        GuiControl, MainGui: Text, 05IncomeEditLabel, % "Income (pages 2, 8-9)"
        GuiControl, MainGui: Text, 06ChildSupportIncomeEditLabel, % "Child Support Income (page 9)"
        GuiControl, MainGui: Text, 08ExpensesEditLabel, % "Expenses (page 10)"
        GuiControl, MainGui: Text, 09AssetsEditLabel, % "Assets (page 10)"
        GuiControl, MainGui: Text, 11ActivityAndScheduleEditLabel, % "Activity and Schedule (pages 10-11)"
        GuiControl, MainGui: Text, 10ProviderEditLabel, % "Provider (pages 12-15)"
        GuiControl, MainGui: Hide, SignOrDueLabel
        GuiControl, MainGui: Hide, SignOrDueDate
        Gui, Show
    }
    checkWaitlist()
}
; v2 convert to switch
setEligibility() {
    Gui, Submit, NoHide
    If (A_GuiControl == "EligibleRadio") {
        caseDetails.eligibility := "elig"
        checkWaitlist()
        GuiControl, MainGui: Show, SaApproved
        GuiControl, MainGui: Show, NoSA
        GuiControl, MainGui: Show, NoProvider
    } Else If (A_GuiControl == "PendingRadio") {
        caseDetails.eligibility := "pends"
        GuiControl, MainGui: Hide, SaApproved
        GuiControl, MainGui: Hide, NoSA
        GuiControl, MainGui: Hide, NoProvider
        checkWaitlist()
    } Else If (A_GuiControl == "IneligibleRadio") {
        caseDetails.eligibility := "ineligible"
        GuiControl, MainGui: Hide, SaApproved
        GuiControl, MainGui: Hide, NoSA
        GuiControl, MainGui: Hide, NoProvider
        checkWaitlist()
    }
    newChangesTrue()
}
setCaseType() {
    caseDetails.caseType := A_GuiControl
    checkWaitlist()
}
setSA() {
    Gui, Submit, NoHide
    caseDetails.saEntered := A_GuiControl == "SaApproved" ? " & SA" : A_GuiControl == "NoSA" ? ", no SA" : A_GuiControl == "NoProvider" ? ", no provider" : ""
}
checkWaitlist() {
    If (ini.caseNoteCountyInfo.Waitlist > 1 && caseDetails.caseType == "BSF" && caseDetails.docType == "Application" && caseDetails.eligibility == "pends") {
        GuiControl, MainGui: Show, manualWaitlistBox
    } Else {
        GuiControl, MainGui: Hide, manualWaitlistBox
        GuiControl,, manualWaitlistBox, 0
    }
}
newChangesTrue() {
    caseDetails.newChanges := true
}
revertLabels() {
    GuiControl, MainGui: Text, 01HouseholdCompEditLabel, % "Household Comp"
    GuiControl, MainGui: Text, 03AddressVerificationEditLabel, % "Address Verification"
    GuiControl, MainGui: Text, 02SharedCustodyEditLabel, % "Shared Custody"
    GuiControl, MainGui: Text, 04SchoolInformationEditLabel, % "School Information"
    GuiControl, MainGui: Text, 05IncomeEditLabel, % "Income"
    GuiControl, MainGui: Text, 06ChildSupportIncomeEditLabel, % "Child Support Income"
    GuiControl, MainGui: Text, 08ExpensesEditLabel, % "Expenses"
    GuiControl, MainGui: Text, 09AssetsEditLabel, % "Assets"
    GuiControl, MainGui: Text, 11ActivityAndScheduleEditLabel, % "Activity and Schedule"
}
copySharedCustodyEditToCSCoopEdit() {
    Global
    Gui, MainGui: Submit, NoHide
    If (StrLen(02SharedCustodyEdit) > 0 && StrLen(07ChildSupportCooperationEdit) == 0) {
        Loop, Parse, 02SharedCustodyEdit, `n, `r
        {
            colonLoc := InStr(A_LoopField, ":",, 0, 1)
            outputText .= SubStr(A_LoopField, 1, colonLoc) " `n"
        }
        outputText := Trim(outputText, "`n")
        GuiControl, MainGui: Text, 07ChildSupportCooperationEdit, % outputText
        GuiControl, MainGui: Focus, 07ChildSupportCooperationEdit
        Send, {End}
    }
}
;MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION - MAIN GUI SECTION
;==============================================================================================================================================================================================

;=====================================================================================================================================================================================================
;BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION
makeCaseNote() {
    Global
    Gui, MainGui: Submit, NoHide
    Gui, MissingGui: Submit, NoHide
    If (caseDetails.newChanges) {
        missingVerifsDoneButton()
    }
    local finishedCaseNote := {}, originalMissingEdit := 14MissingEdit
    local caseDetailsModified := { caseType: caseDetails.caseType, appType: caseDetails.appType, docType: caseDetails.docType, eligibility: caseDetails.eligibility, saEntered: caseDetails.saEntered }
    ;v2: continuation section 
    editFields := { 01HouseholdCompEdit: " HH COMP:    ", 02SharedCustodyEdit: " CUSTODY:    ", 03AddressVerificationEdit: " ADDRESS:    ", 04SchoolInformationEdit: "  SCHOOL:    ", 05IncomeEdit: "  INCOME:    ", 06ChildSupportIncomeEdit: "      CS:    ", 07ChildSupportCooperationEdit: " CS COOP:    ", 08ExpensesEdit: "EXPENSES:    ", 09AssetsEdit: "  ASSETS:    ", 10ProviderEdit: "PROVIDER:    ", 11ActivityAndScheduleEdit: "ACTIVITY:    ", 12ServiceAuthorizationEdit: "      SA:    ", 13NotesEdit: "   NOTES:    ", 14MissingEdit: " MISSING:    " }
    ;For i, pattern in [ "i)([a-z])([0-9])", "i)([a-z0-9])(\()", "i)(\))([a-z0-9])" ] {
    For i, pattern in [ "i)([a-z])([0-9])", "i)(\))([a-z0-9])" ] {
        01HouseholdCompEdit := RegExReplace(01HouseholdCompEdit, pattern, "$1 $2")
    }
    finishedCaseNote.mec2CaseNote := autoDenyObject.autoDenyExtensionMECnote 
    For editField, label in editFields {
        finishedCaseNote.mec2CaseNote .= label stWordWrap(%editField%, 100, "             ", 1, 1) "`n"
    }
    finishedCaseNote.mec2CaseNote .= "=====`n" ini.employeeInfo.employeeName
    finishedCaseNote.eligibility := caseDetails.eligibility
	If (caseDetails.eligibility == "pends" && caseDetailsModified.docType == "Redet") {
		caseDetailsModified.eligibility := "incomplete (due " formatMDY(SignOrDueDate) ")"
        finishedCaseNote.eligibility := "incomplete"
	}
	If (caseDetails.eligibility != "elig") {
		caseDetailsModified.saEntered := ""
	}
    If (overIncomeMissing && caseDetailsModified.eligibility == "ineligible") {
        caseDetailsModified.eligibility := "over-income"
        finishedCaseNote.eligibility := "ineligible"
    }
    If ( (caseDetailsModified.docType == "Application" && caseDetailsModified.caseType == "BSF" && caseDetailsModified.eligibility == "ineligible" && ini.caseNoteCountyInfo.Waitlist > 1) || WaitlistMissing == 1) {
        caseDetailsModified.eligibility := caseDetails.eligibility " - BSF Waitlist"
        finishedCaseNote.eligibility := WaitListMissing == 1 ? "pends" : "ineligible"
    }
    If (Homeless == 1 && caseDetailsModified.docType == "Application") {
		caseDetailsModified.caseType := "*HL " caseDetails.caseType
	}
	If (caseDetailsModified.docType == "Application") {
		finishedCaseNote.mec2NoteTitle := caseDetailsModified.caseType " " caseDetailsModified.appType " rec'd " dateObject.receivedMDY ", " caseDetailsModified.eligibility caseDetailsModified.saEntered
        If (caseDetails.eligibility == "pends") {
            finishedCaseNote.mec2NoteTitle .= " until " autoDenyObject.autoDenyExtensionDate
            finishedCaseNote.maxisNote := "CCAP app rec'd " dateObject.receivedMDY ", pend date " autoDenyObject.autoDenyExtensionDate ".`n" ; MAXIS
        }
    ; MAXIS ----------------------------------------
        If (caseDetails.eligibility == "elig") {
            finishedCaseNote.maxisNote := "CCAP app rec'd " dateObject.receivedMDY ", approved eligible." (Homeless == 1 ? " Expedited." : "") "`n"
        }
        If (caseDetailsModified.eligibility == "ineligible" || caseDetailsModified.eligibility == "over-income") {
            finishedCaseNote.maxisNote := "CCAP app rec'd " dateObject.receivedMDY ", denied " dateObject.todayMDY ".`n"
            If (overIncomeMissing) {
                finishedCaseNote.maxisNote .= " Over-income"
            }
            finishedCaseNote.maxisNote .= "`n"
        }
        If (StrLen(originalMissingEdit) > 0) {
            missingMax := stWordWrap(originalMissingEdit, 74, "* ", 4)
            finishedCaseNote.maxisNote .= "Special Letter mailed " dateObject.todayMDY " requesting:`n" missingMax "`n"
        }
        finishedCaseNote.maxisNote .= ini.employeeInfo.employeeName
    ; MAXIS ----------------------------------------
	} Else If (caseDetailsModified.docType == "Redet") {
		finishedCaseNote.mec2NoteTitle := caseDetailsModified.caseType " " caseDetailsModified.docType " rec'd " dateObject.receivedMDY ", " caseDetailsModified.eligibility caseDetailsModified.saEntered
	}
    If ( InStr(finishedCaseNote.mec2NoteTitle, "?") ) {
        MsgBox,, % "Case Note Error", % "Select options at the top before case noting.`n  (Document type, Program, Eligibility, etc.)"
        Return false
    }


    Return finishedCaseNote
}
outputCaseNote() {
    Global
    madeCaseNote := makeCaseNote()
    If (!madeCaseNote) {
        Return
    }
    ; v2 Convert to switch
    If (A_GuiControl == "mec2NoteButton") {
        outputCaseNoteMec2(madeCaseNote)
    } Else If (A_GuiControl == "maxisNoteButton") {
        outputCaseNoteMaxis(madeCaseNote)
    } Else If (A_GuiControl == "notepadNoteButton") {
        outputCaseNoteNotepad(madeCaseNote)
	}
    sleep 500
    Clipboard := caseNumber
}
outputCaseNoteMec2(sendingCaseNote) {
    Global
    StrReplace(sendingCaseNote.mec2CaseNote, "`n", "`n", mec2CaseNoteLines) ; Counting lines
    If (mec2CaseNoteLines > 29) { ; off by 1 - last line won't have a newline
        sendingCaseNote.mec2CaseNote := buildOversizedNoteGui(sendingCaseNote.mec2CaseNote)
        If (StrLen(sendingCaseNote.mec2CaseNote) < 2) {
            Sleep 200
            WinActivate, % "CaseNotes"
            Return
        }
    }
    WinActivate % ini.employeeInfo.employeeBrowser
    WinWaitActive, % ini.employeeInfo.employeeBrowser,, 5
    mec2docType := caseDetails.docType == "Redet" ? "Redetermination" : caseDetails.docType
    If (ini.employeeInfo.employeeUseMec2Functions == 1) {
        jsonCaseNote := "CaseNoteFromAHKJSON{""noteDocType"":""" mec2docType """,""noteTitle"":""" JSONstring(sendingCaseNote.mec2NoteTitle) """,""noteText"":""" JSONstring(sendingCaseNote.mec2CaseNote) """,""noteElig"":""" sendingCaseNote.eligibility """ }"
        Clipboard := jsonCaseNote
        Sleep 500
        Send, ^v
    } Else If (ini.employeeInfo.employeeUseMec2Functions == 0) {
        catNum := { Application: { letter: "A", pends: 5, elig: 4, denied: 4 }, Redet: { letter: "R", incomplete: 1, elig: 2, denied: 2 } }
        catLetter := catNum[caseDetails.docType].letter
        catNumber := catNum[caseDetails.docType][caseDetails.eligibility]
        Sleep 500
        Send {Tab 7}
        Sleep 750
        SendInput, { %catLetter% %catNumber% }
        Sleep 500
        Send {Tab}
        Sleep 500
        SendInput, % sendingCaseNote.mec2NoteTitle
        Sleep 500
        Send {Tab}
        Sleep 500
        SendInput, % sendingCaseNote.mec2CaseNote
        Sleep 500
        Send {Tab}
    }
    caseNoteEntered.mec2NoteEntered := 1
    GuiControl, MainGui:Text, % "mec2NoteButton", % "MEC2 " cm
}
outputCaseNoteMaxis(sendingCaseNote) {
    Global
    Clipboard := sendingCaseNote.maxisNote
    maxisWindow := WinExist("ahk_group maxisGroup") ; returns an ID
    If (maxisWindow == "0x0") {
        MsgBox, 1, % "MAXIS window not found", % "MAXIS window not found. Open or click on the MAXIS window, then click the OK button."
        IfMsgBox, Cancel
            Return
        maxWin := WinExist("ahk_exe bzmd.exe")
        WinGetTitle, maxWinName, % "ahk_id " maxWin
        GroupAdd, maxisGroup, % maxWinName
        ini.employeeInfo.employeeMaxis := maxWinName
        IniWrite, % maxWinName, % A_MyDocuments "\AHK.ini", EmployeeInfo, EmployeeMaxis
    }
; Current code (works)
    ;If (maxisWindow) {
        ;WinActivate, % "ahk_id " maxisWindow
        ;Clipboard := sendingCaseNote.maxisNote
        ;WinWaitActive, % "ahk_id " maxisWindow,, 5
        ;Sleep 500
        ;Send, ^v
    ;}
; End current code
; Test area start

    maxisNote := sendingCaseNote.maxisNote
    maxisNoteLength := maxisNote.Length() -1, maxisNoteOutput := "", j := 1
    For i, textLine in maxisNote {
        maxisNoteOutput .= textLine "`n"
        j++
        If (j == 14 && i < maxisNoteLength) {
            maxisNoteOutput .= "...continued on next page"
            doPasteInMaxis(maxisNoteOutput, maxisWindow)
            j := 1
            maxisNoteOutput := ""
            Send, {F9}
            Sleep 1000
        }
    }
    doPasteInMaxis(maxisNoteOutput, maxisWindow)

; Test area end

    caseNoteEntered.maxisNoteEntered := 1
    GuiControl, MainGui:Text, maxisNoteButton, % "MAXIS " cm
}
outputCaseNoteNotepad(sendingCaseNote) {
    Global
    local notepadFileName := caseNumber !== "" ? caseNumber : ""
    If (notepadFileName == "") {
        RegExMatch(HouseholdComp, "\w+\W", notepadFileName)
    }
    local letterNotepad := ""
    For i, letterTextValue in letterText {
        letterNotepad .= StrLen(LetterText[i]) > 0 ? "`n====== Special Letter " i " ======`n" letterTextValue "`n" : ""
    }
    FileAppend, % "====== Case Note Summary ======`n" sendingCaseNote.mec2NoteTitle "`n`n====== MEC2 Case Note ===== `n" sendingCaseNote.mec2CaseNote "`n`n===== Email ===== `n" emailTextObject.output "`n" letterNotepad "`n" (ini.caseNoteCountyInfo.countyNoteInMaxis == 1 ? "`n===== MAXIS Note =====`n" sendingCaseNote.maxisNote "`n" : "") "`n-------------------------------------------`n`n`n", % A_Desktop "\" notepadFileName ".txt"
    If (verbose) {
        buildOversizedNoteGui("====== Case Note Summary ======`n" sendingCaseNote.mec2NoteTitle "`n`n====== MEC2 Case Note ===== `n" sendingCaseNote.mec2CaseNote "`n`n===== Email ===== `n" emailTextObject.output "`n" letterNotepad "`n" (ini.caseNoteCountyInfo.countyNoteInMaxis == 1 ? "`n===== MAXIS Note =====`n" sendingCaseNote.maxisNote "`n" : "") "`n-------------------------------------------`n")
    }
    GuiControl, MainGui:Text, notepadNoteButton, % "Desktop " cm
    caseNoteEntered.mec2NoteEntered := 1
    caseNoteEntered.maxisNoteEntered := 1
}
doPasteInMaxis(textString, ByRef maxisWindow) {
    textString := Trim(textString, "`n")
    clipboard := textString
    WinActivate, % "ahk_id " maxisWindow
    WinWaitActive, % "ahk_id " maxisWindow
    Send, ^v
    Return 1
}

buildOversizedNoteGui(oversizedCaseNote) {
    Global
    Gui, OversizedNoteGui: New,, % "Oversized Note"
    Gui, OversizedNoteGui: Color, % "a9a9a9", % "bebebe"
    Gui, OversizedNoteGui: Margin, % marginW
    Gui, OversizedNoteGui: Font, s10, % "Segoe UI"
    Gui, OversizedNoteGui: Add, Text, x90, % "Your case note exceeds the maximum line count of 30. Edit your note here or return to CaseNotes."
    Gui, OversizedNoteGui: Font, s9, % "Lucida Console"
    Gui, OversizedNoteGui: Add, Edit, % "xm voversizedEdit goversizedEditChange w" CH100x30[1]+scrollbar+pad " h" CH100x30[2]+pad*2, % oversizedCaseNote
    Gui, OversizedNoteGui: Font, s10, % "Segoe UI"
    Gui, OversizedNoteGui: Add, Text, , % "Line Count: "
    Gui, OversizedNoteGui: Add, Text, % "voversizedLineCount x+p", "--"
    Gui, OversizedNoteGui: Add, Button, % "vsaveOversizedButton gOversizedNoteGuiGuiClose x+185", % "Send to MEC" sq
    Gui, OversizedNoteGui: Add, Button, % "vreturnOversizedButton gOversizedNoteGuiGuiClose x+30", % "Return to CaseNotes"
    Gui, OversizedNoteGui: Show, % "Hide x" ini.caseNotePositions.xCaseNotes " y" ini.caseNotePositions.yCaseNotes
    GuiControlGet oversizedEditSize, OversizedNoteGui:POS, oversizedEdit
    oversizedNoteGuiX := caseNotesMonCenter[1] - ((oversizedEditSizeW*zoomPPI + scrollbar)/2), oversizedNoteGuiY := caseNotesMonCenter[2] - (oversizedEditSizeH*zoomPPI/2)
    oversizedEditChange()
    GuiControl, OversizedNoteGui:Text, oversizedEdit, % oversizedCaseNote
    Gui, OversizedNoteGui: Show, % "x" oversizedNoteGuiX " y" oversizedNoteGuiY
    Gui, OversizedNoteGui: +OwnerMainGui
    Gui, MissingGui: +Disabled
    Gui, MainGui: +Disabled
    Send {end}
    oversizedGuiCloseControl :=
    WinWaitClose, % "Oversized Note"
    Gui, OversizedNoteGui: Submit
    Return oversizedGuiCloseControl == "saveOversizedButton" ? oversizedEdit : ""
}
OversizedNoteGuiGuiClose() {
    Global oversizedGuiCloseControl := A_GuiControl
    Gui, OversizedNoteGui: Submit, NoHide
    Gui, MissingGui: -Disabled
    Gui, MainGui: -Disabled
    Gui, OversizedNoteGui: Destroy
}
oversizedEditChange() {
    Global
    Gui, OversizedNoteGui: Submit, NoHide
    StrReplace(oversizedEdit, "`n", "`n", oversizedNoteLines)
    GuiControl, OversizedNoteGui: Text, oversizedLineCount, % oversizedNoteLines+1
}

JSONstring(inputString) {
    inputString := StrReplace(inputString, "\", "\\",, -1)
    inputString := StrReplace(inputString, "`n", "\n",, -1)
    inputString := StrReplace(inputString, """", "\""",, -1)
    Return inputString
}
;BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION - BUILD AND SEND SECTION
;=====================================================================================================================================================================================================

;===========================================================================================================================================================================================
;DATES SECTION DATES SECTION DATES SECTION DATES SECTION DATES SECTION DATES SECTION DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  
addFifteenishDays(oldDate) {
	FormatTime, dayNumber, % oldDate, WDay
    Return dayNumber == 7 ? addDays(oldDate, 17) : dayNumber > 4 ? addDays(oldDate, 18) : addDays(oldDate, 16)
}
addDays(origDate, addedDays) {
    origDate += addedDays, Days
    Return origDate
}
subtractDates(futureDate, pastDate) {
    EnvSub, futureDate, % pastDate, days
    Return futureDate
}
formatMDY(inputDate) {
    FormatTime, dateMDY, % inputDate, % "M/d/yy"
    Return dateMDY ;v2 can this be combined?
}
calcDates() {
    Global
    Gui, MainGui: Submit, NoHide
    autoDenyObject.autoDenyExtensionSpecLetter :=

    dateObject.receivedYMD := ReceivedDate
    dateObject.receivedMDY := formatMDY(dateObject.receivedYMD)
    dateObject.autoDenyYMD := addDays(dateObject.receivedYMD, 29)
    dateObject.recdPlusFortyfiveYMD := addDays(dateObject.receivedYMD, 44)
    dateObject.todayPlusFifteenishYMD := addFifteenishDays(dateObject.todayYMD)
    dateObject.recdPlusFifteenishYMD := addFifteenishDays(dateObject.receivedYMD)
    dateObject.needsNoExtension := subtractDates(dateObject.autoDenyYMD, dateObject.todayPlusFifteenishYMD)
    dateObject.needsExtension := subtractDates(dateObject.recdPlusFortyfiveYMD, dateObject.todayPlusFifteenishYMD)

    If (caseDetails.docType == "Application") {
        If (caseDetails.eligibility == "pends") {
            If (dateObject.needsNoExtension > -1) {
                autoDenyObject.autoDenyExtensionDate := formatMDY(dateObject.autoDenyYMD)
                autoDenyObject.autoDenyExtensionSpecLetter := "*You have through " autoDenyObject.autoDenyExtensionDate " to submit required verifications."
                GuiControl, MainGui: Text, autoDenyStatus, % "Has 15+ days before auto-deny"
            } Else If (dateObject.needsExtension > -1) {
                autoDenyObject.autoDenyExtensionDate := formatMDY(dateObject.todayPlusFifteenishYMD)
                autoDenyObject.autoDenyExtensionMECnote := "Auto-deny extended to " autoDenyObject.autoDenyExtensionDate " due to processing < 15 days before auto-deny.`n-`n"
                autoDenyObject.autoDenyExtensionSpecLetter := "*You have through " autoDenyObject.autoDenyExtensionDate " to submit required verifications."
                GuiControl, MainGui: Text, autoDenyStatus, % "Extend auto-deny to " autoDenyObject.autoDenyExtensionDate
            } Else {
                autoDenyObject.autoDenyExtensionDate := formatMDY(dateObject.todayPlusFifteenishYMD)
                autoDenyObject.autoDenyExtensionMECnote := "Reinstate date is " autoDenyObject.autoDenyExtensionDate " due to processing < 15 days before auto-deny.`n-`n"
                autoDenyObject.autoDenyExtensionSpecLetter := "*Please note that you will be mailed an auto-denial notice.`n  You have through " autoDenyObject.autoDenyExtensionDate " to submit required verifications.`n  If you are eligible, your case will be reinstated."
                GuiControl, MainGui: Text, autoDenyStatus, % "Auto-denies tonight, pends until " autoDenyObject.autoDenyExtensionDate
            }
        } Else If (Homeless == 1) {
            dateObject.ExpeditedNinetyDaysYMD := addDays(dateObject.receivedYMD, 89)
            autoDenyObject.autoDenyExtensionDate := formatMDY(dateObject.ExpeditedNinetyDaysYMD)
            autoDenyObject.autoDenyExtensionSpecLetter := "*You have through " autoDenyObject.autoDenyExtensionDate " to submit required verifications."
        } Else {
            GuiControl, MainGui: Text, autoDenyStatus, % ""
            autoDenyObject.autoDenyExtensionSpecLetter := ""
        }
    }
    If (caseDetails.docType == "Redet") {
        dateObject.RedetDueYMD := SignOrDueDate

        If (caseDetails.eligibility != "elig") {
        
            dateObject.RedetDueMDY := formatMDY(SignOrDueDate)
            dateObject.RedetCaseCloseYMD := addFifteenishDays(dateObject.RedetDueYMD)
            dateObject.RedetCaseCloseMDY := formatMDY(dateObject.RedetCaseCloseYMD)
            dateObject.RedetDocsLastDayYMD := addDays(dateObject.RedetCaseCloseYMD, 29)
            dateObject.RedetDocsLastDayMDY := formatMDY(dateObject.RedetDocsLastDayYMD)
        
            If (dateObject.todayYMD > dateObject.RedetDocsLastDayYMD) {
                autoDenyObject.autoDenyExtensionSpecLetter := "** Your case has closed due to failure to complete the redetermination process. Because your redetermination was not completed within 30 days of closure, your case cannot be reinstated. To be eligible for CCAP, you must reapply by completing an application. The date you submit an application is the earliest date of your eligibility. If you have received Cash Assistance (MFIP, DWP) within the last 12 months, you may be eligible for limited backdating."
            } Else {
                autoDenyObject.autoDenyExtensionSpecLetter := dateObject.todayYMD < dateObject.RedetDueYMD
                ? "** If your redetermination is not completed by " dateObject.RedetDueMDY ", "
                : "** Your redetermination was not completed by " dateObject.RedetDueMDY " and "
                autoDenyObject.autoDenyExtensionSpecLetter .= dateObject.todayYMD < dateObject.RedetCaseCloseYMD
                ? "your case will close on " dateObject.RedetCaseCloseMDY ". If it closes, the latest it can be reinstated is " dateObject.RedetDocsLastDayMDY "."
                : "your case has closed. To reinstate your case, you must complete the redetermination process by " dateObject.RedetDocsLastDayMDY "."
            }
        } Else If (caseDetails.eligibility == "elig" && NoSA) {
            scheduleMissing := (WorkSchedulePlusNameMissing + WorkScheduleMissing + CustodyScheduleMissing + CustodySchedulePlusNamesMissing + SelfEmploymentScheduleMissing + ClassScheduleMissing)
            providerIssue := (UnregisteredProviderMissing + InHomeCareMissing + LNLProviderMissing + StartDateMissing)
            autoDenyObject.autoDenyExtensionSpecLetter := "** Your redetermination is approved and your case remains eligible. "
            If (scheduleMissing) {
                autoDenyObject.autoDenyExtensionSpecLetter .= "Assistance hours at your provider"
                autoDenyObject.autoDenyExtensionSpecLetter .= dateObject.todayYMD < dateObject.RedetDueYMD ? " have ended" : " will end"
                autoDenyObject.autoDenyExtensionSpecLetter .= " until we receive the above items.`n"
            }
            If (providerIssue) {
                autoDenyObject.autoDenyExtensionSpecLetter .= "The above items must be resolved before assistance hours at your provider can be approved.`n"
            }
        }
    }
}
;DATES SECTION DATES SECTION DATES SECTION DATES SECTION DATES SECTION DATES SECTION DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  DATES SECTION  
;===========================================================================================================================================================================================

;=====================================================================================================================================================================================
;VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION
showMissingVerifs() {
    Gui, MissingGui: Restore
    Gui, MissingGui: Show, AutoSize
    Gui, MainGui: Submit, NoHide
    Gui, MissingGui: Submit, NoHide
}
buildMissingGui() {
    Global
    local column1of1 := "xm w390"
    local column1of2 := "xm w158", column2of2 := "x170 yp+0 w240", 
    local column1of3 := "xm w118", column2of3 := "x130 yp+0 w120", column3of3 := "x262 yp+0 w138", column2and3Of3 := "x130 yp+0 w280"

    local lineColor := "0x5" ; https://gist.github.com/jNizM/019696878590071cf739
    local textLine := "x60 y+4 w250 h1 " lineColor
    ;-- Alternate method for lines:
    ;lineColor := "717171"
    ;ProgressLine := "x50 y+4 w250 h1 Background" lineColor
    ;Gui, MissingGui: Add, Progress, % ProgressLine

    Gui, MissingGui: New,, % "Missing Verifications"
    Gui, MissingGui: Margin, % marginW
    Gui, MissingGui: Add, Checkbox, % column1of1 " vIDmissing ginputBoxAGUIControl", % "ID (input)"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vBCmissing ginputBoxAGUIControl", % "BC (input)"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vBCNonCitizenMissing ginputBoxAGUIControl", % "BC [non-citizen] (input)"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vPaternityMissing ginputBoxAGUIControl", % "Paternity (input)"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vAddressMissing", % "Address"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vChildSupportFormsMissing ginputBoxAGUIControl", % "Child Support forms (input)"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vChildSupportNoncooperationMissing ginputBoxAGUIControl", % "CS Non-cooperation (input)"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vCustodyScheduleMissing", % "Custody (""for each child"")"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vCustodySchedulePlusNamesMissing ginputBoxAGUIControl", % "Custody (input)"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vChildSchoolMissing", % "Child school information"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vChildFTSchoolMissing", % "Child full-time student status"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vMarriageCertificateMissing", % "Marriage certificate"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vLegalNameChangeMissing ginputBoxAGUIControl", % "Name change (input)"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vDependentAdultStudentMissing ginputBoxAGUIControl", % "Dependent adult child - FT Student, 50`%+ expenses (input)"

    Gui, MissingGui: Font, bold ;-- EARNED INCOME SECTION ==============================================================
    Gui, MissingGui: Add, Text, % "xm+10 y+22 w110 h1 " lineColor
    Gui, MissingGui: Add, Text, % "x+m yp-7", % "Earned Income"
    Gui, MissingGui: Add, Text, % "x+m yp+7 w115 h1 " lineColor
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, % column1of3 " vIncomeMissing", % "Income"
    Gui, MissingGui: Add, Checkbox, % column2of3 " vWorkScheduleMissing", % "Work Schedule"
    Gui, MissingGui: Add, Checkbox, % column3of3 " vContractPeriodMissing", % "Contract Period"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vIncomePlusNameMissing ginputBoxAGUIControl", % "Income (input)"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vWorkSchedulePlusNameMissing ginputBoxAGUIControl", % "Work Schedule (input)"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vNewEmploymentMissing", % "New job at app / end of job search (Wage, dates, hours)"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vWorkLeaveMissing", % "Leave of absence (Dates, pay status, hours, work schedule)"
    Gui, MissingGui: Add, Text, % textLine ;-- -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, % column1of1 " vSeasonalWorkMissing", % "Seasonal employment season length"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vSeasonalOffSeasonMissing ginputBoxAGUIControl", % "Seasonal employment info - app in off-season (input)"
    Gui, MissingGui: Add, Text, % textLine ;-- -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, % column1of2 " vSelfEmploymentMissing", % "Self-Employment Income"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vSelfEmploymentScheduleMissing", % "Self-Employment Schedule"
    Gui, MissingGui: Add, Checkbox, % column1of1 " vSelfEmploymentBusinessGrossMissing", % "Self-Employment Business Gross (if state min wage; <$500k = small business)"
    Gui, MissingGui: Add, Text, % textLine ;-- -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, % column1of2 " vExpensesMissing", % "Expenses"
    Gui, MissingGui: Add, Checkbox, % column2of2 " voverIncomeMissing ginputBoxAGUIControl", % "Over-income (input)"

    Gui, MissingGui: Font, bold ;-- UNEARNED INCOME SECTION ============================================================
    Gui, MissingGui: Add, Text, % "xm+10 y+22 w105 h1 " lineColor
    Gui, MissingGui: Add, Text, % "x+m yp-7", % "Unearned Income"
    Gui, MissingGui: Add, Text, % "x+m yp+7 w120 h1 " lineColor
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, % column1of2 " vChildSupportIncomeMissing", % "Child Support Income"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vSpousalSupportMissing", % "Spousal Support Income"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vRentalMissing", % "Rental"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vDisabilityMissing", % "STD / LTD "
    Gui, MissingGui: Add, Checkbox, % column1of2 " vAssetsGT1mMissing", % "Assets (>$1m)"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vUnearnedStatementMissing", % "Blank Unearned Yes/No (statement)"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vAssetsBlankMissing", % "Assets (Blank)"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vUnearnedMailedMissing", % "Blank Unearned Yes/No (mailed back)"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vVABenefitsMissing", % "VA Benefits"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vInsuranceBenefitsMissing", % "Insurance Benefits"

    Gui, MissingGui: Font, bold ;-- ACTIVITY SECTION ===================================================================
    Gui, MissingGui: Add, Text, % "xm+10 y+22 w130 h1 " lineColor
    Gui, MissingGui: Add, Text, % "x+m yp-7", % "Activity"
    Gui, MissingGui: Add, Text, % "x+m yp+7 w140 h1 " lineColor
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, % column1of2 " vEdBSFformMissing", % "BSF/TY Education Form"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vEdBSFOneBachelorDegreeMissing", % "BSF/TY Bachelor's limit notice"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vClassScheduleMissing", % "Class schedule"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vTranscriptMissing", % "Transcript"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vEducationEmploymentPlanMissing", % "ES Plan (CCMF Education)"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vStudentStatusOrIncomeMissing", % "Adult student w/ income (age < 20)"
    Gui, MissingGui: Add, Text, % textLine ;-- -------------------------------------------------------------------------
    Gui, MissingGui: Add, Checkbox, % column1of2 " vJobSearchHoursMissing", % "BSF Job search hours"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vSelfEmploymentIneligibleMissing", % "Self-Employment not enough hours"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vEligibleActivityMissing", % "No Eligible Activity Listed"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vEmploymentIneligibleMissing", % "Employment not enough hours"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vESPlanOnlyJSMissing", % "ES Plan-only JS notice"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vActivityAfterHomelessMissing", % "Activity Req. After 3-Mo Homeless Period"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vUnableToProvideCareMissing", % "Unable to Provide Care form"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vMightBeUnableToProvideCareMissing", % "One parent might be Unable to Provide Care"

    Gui, MissingGui: Font, bold ;-- PROVIDER SECTION ===================================================================
    Gui, MissingGui: Add, Text, % "xm+10 y+22 w125 h1 " lineColor
    Gui, MissingGui: Add, Text, % "x+m yp-7", % "Provider"
    Gui, MissingGui: Add, Text, % "x+m yp+7 w130 h1 " lineColor
    Gui, MissingGui: Font

    Gui, MissingGui: Add, Checkbox, % column1of2 " vNoProviderMissing", % "No Provider Listed"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vUnregisteredProviderMissing", % "Unregistered Provider"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vInHomeCareMissing", % "In-Home Care form"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vLNLProviderMissing", % "LNL Acknowledgement"
    Gui, MissingGui: Add, Checkbox, % column1of2 " vStartDateMissing", % "Provider start date"
    Gui, MissingGui: Add, Checkbox, % column2of2 " vProviderForNonImmigrantMissing", % "Non-citizen/immigrant Provider Reqs."

    Gui, MissingGui: Add, Checkbox, % column1of1 " h50 votherInput1 gotherGUI", % "Other"
    Gui, MissingGui: Add, Checkbox, % column1of1 " h50 votherInput2 gotherGUI", % "Other"
    Gui, MissingGui: Add, Checkbox, % column1of1 " h50 votherInput3 gotherGUI", % "Other"

    Gui, MissingGui: Add, Button, % "h17 gmissingVerifsDoneButton", % "Done"
    Gui, MissingGui: Add, Button, % "x+20 w40 h17 hidden gemailButtonClick vemailButton", % "Email"
    Gui, MissingGui: Add, Button, % "x+20 w42 h17 hidden gletterButtonClick vletter1", % "Letter 1"
    Gui, MissingGui: Add, Button, % "x+20 w42 h17 hidden gletterButtonClick vletter2", % "Letter 2"
    Gui, MissingGui: Add, Button, % "x+20 w42 h17 hidden gletterButtonClick vletter3", % "Letter 3"
    Gui, MissingGui: Add, Button, % "x+20 w42 h17 hidden gletterButtonClick vletter4", % "Letter 4"
    Gui, MissingGui: Show, % "Hide x" ini.caseNotePositions.xVerification " y" ini.caseNotePositions.yVerification
}

missingVerifsDoneButton() {
    Global
	Gui, MainGui: Submit, NoHide
	Gui, MissingGui: Submit, NoHide
    calcDates()
    
    emailTextString := "", emailTextObject := {}
    mecCheckboxIds := {}
    letterNumber := 1, letterTextNumber := 1, letterText := {}
    lineNumber := 1, missingListEnum := 1, clarifiedListEnum := 1, emailListEnum := 1, lineCount := 0
    caseNoteMissingText := ""
	local missingVerifications := {}, clarifiedVerifications := {}
	missingVerifications := new orderedAssociativeArray()
    clarifiedVerifications := new orderedAssociativeArray()
    mec2docType := caseDetails.docType == "Redet" ? "Redetermination" : caseDetails.docType
	GuiControl, MissingGui: Hide, % "letter2"
	GuiControl, MissingGui: Hide, % "letter3"
	GuiControl, MissingGui: Hide, % "letter4"
    
    emailTextObject.StartAll := "Your Child Care Assistance " mec2docType " has been " (caseDetails.eligibility == "elig" ? "approved. " : "processed. ")
    If overIncomeMissing {
        overIncomeMissingText1 := "Using information you provided, your case is ineligible as your income is over the limit for a household of " overIncomeObj.overIncomeHHsize ". The gross limit is $" overIncomeObj.overIncomeText ".`n"
        overIncomeMissingText2 := "If your gross income does not match this calculation, you must" countySpecificText[ini.employeeInfo.employeeCounty].OverIncomeContactInfo " submit income and eligible expense documents along with the following verifications:`n"
        emailTextObject.StartAll .= "`n`n" overIncomeMissingText1 overIncomeMissingText2 "`n"
        missingVerifications[overIncomeMissingText1] := 3
        missingVerifications[overIncomeMissingText2] := 3
        caseNoteMissingText .= "Household is calculated to be over-income by $" overIncomeObj.overIncomeDifference " ($" overIncomeObj.overIncomeReceived " - $" overIncomeObj.overIncomeLimit ");`n"
    } Else If (Homeless && caseDetails.docType == "Application") {
        If (caseDetails.eligibility == "pends") {
            InputBox, missingHomelessItems, % "Homeless App - Info Missing", % "Eligibility is marked as Pending. What information is needed from the client to approve expedited eligibility?`n`nUse a double space ""  "" without quotation marks to start a new line.",,,,,,,, % StrReplace(missingHomelessItems, "`n", "  ")
            If (ErrorLevel == 0) {
                missingHomelessItems := StrReplace(missingHomelessItems, "  ", "`n")
                pendingHomelessMissing := getRowCount("  " missingHomelessItems, 60, "  ")
                missingVerifications[stWordWrap(emailText.pendingHomelessPreText, 60, " ") "`n"] := 8
                missingVerifications[pendingHomelessMissing[1] "`n"] := pendingHomelessMissing[2]
                caseNoteMissingText .= "Missing for expedited approval:`n" StrReplace(missingHomelessItems, "`n", "`n  ") ";`n"
            }
        } Else If (caseDetails.eligibility == "elig") {
            emailTextObject.StartHL := (caseDetails.eligibility == "elig") ? emailText.approvedWithMissing "`n" emailText.stillRequiredText : PendingHomelessPreText missingHomelessItems
            emailTextObject.EndHL := (caseDetails.eligibility == "elig") ? emailText.initialApproval : ""
        }

        emailTextObject.AreOrWillBe := (Homeless == 1) ? "will be" : "are"
        emailTextObject.Reason1 := (caseDetails.eligibility == "elig") ? "for authorizing assistance hours" : "to determine eligibility or calculate assistance hours"
        emailTextObject.Reason2 := (Homeless == 1) ? "to determine on-going eligibility or calculate assistance hours after the 90-day period" : emailTextObject.Reason1
    }
    emailTextObject.Start := (Homeless == 1) ? emailTextObject.StartAll emailTextObject.StartHL : emailTextObject.StartAll
    emailTextObject.Middle := !overIncomeMissing ? "`n`nThe following documents or verifications " emailTextObject.AreOrWillBe " needed " emailTextObject.Reason2 ":`n`n" : ""
    emailTextObject.Combined := emailTextObject.Start emailTextObject.Middle

    parseMissingVerifications(missingVerifications, missingListEnum, clarifiedVerifications, clarifiedListEnum, emailTextString, emailListEnum, caseNoteMissingText)



    caseDetails.haveWaitlist := (caseDetails.caseType == "BSF" && caseDetails.eligibility == "ineligible" && ini.caseNoteCountyInfo.Waitlist > 1)
    If (!caseDetails.haveWaitlist) {
        faxAndEmail := faxAndEmailText()
        faxAndEmailWrapped := getRowCount(faxAndEmail, 60, "")
        autoDenyTextAndLines := getRowCount(autoDenyObject.autoDenyExtensionSpecLetter, 60, "")
        clarifiedVerifications[ "NEWLINE" faxAndEmail "`nNEWLINE" autoDenyObject.autoDenyExtensionSpecLetter ] := faxAndEmailWrapped[2]+autoDenyTextAndLines[2]+1
        emailTextString .= autoDenyObject.autoDenyExtensionSpecLetter
    }

    idList := "other"
    For checkboxId in mecCheckboxIds {
        idList .= "," checkboxId
    }
    Trim(idList, ",")
    insertAtOffset := (caseDetails.eligibility == "pends" && Homeless) ? 2 : 0
    If ( !overIncomeMissing && !caseDetails.haveWaitlist && !manualWaitlistBox && missingVerifications.Length() > (0 + insertAtOffset) ) {
        If (StrLen(idList) > 5 || insertAtOffset == 2) { ; "other" will always add at least 5
            missingVerifications.InsertAt(1 + insertAtOffset, "In addition to the above, please submit the following items:`n", 1)
        }
        Else If (StrLen(idList) == 5) {
            missingVerifications.InsertAt(1+ insertAtOffset, "_____________Please submit the following items:_____________`n", 1)
        }
    }
    waitlistText := ""
    If (caseDetails.haveWaitlist || manualWaitlistBox) {
        waitlistNumber := ini.caseNoteCountyInfo.Waitlist -1
        waitlistPriorities := { 1: "• Are attending High School, GED, or ESL classes;`n" }
        waitlistPriorities.2 := waitlistPriorities.1 "• Families in which an applicant is a veteran;`n"
        waitlistPriorities.3 := waitlistPriorities.2 "• Families which don't qualify for other priorities;`n"
        waitlistText := "
(
Due to limited funding, new eligibility for CCAP in " countySpecificText[ini.employeeInfo.employeeCounty].CountyName " is currently limited to those who:
• Received Cash Assistance (MFIP/DWP) within the past year;
• Are applying for and are approved for Cash Assistance;
" waitlistPriorities[waitlistNumber] "`n
)"
        emailTextObject.Waitlist := "`n`n" waitlistText "`n" (manualWaitlistBox ? submitOnlyCommentItemsText "`n" : "")
        submitOnlyCommentItemsText := "If you meet one of the above criteria, please submit the following items:`n"
        submitCommentAndCheckboxItemsText := "If you meet one of the above criteria, in addition to items above the Worker Comments, please submit the following:`n"
        If (manualWaitlistBox) {
            waitlistText .= StrLen(idList) == 5 ? submitOnlyCommentItemsText : submitCommentAndCheckboxItemsText
        }
        waitlistText := getRowCount(waitlistText, 60, "")
        missingVerifications.InsertAt(1, waitlistText[1] "`n", waitlistText[2])
        caseNoteMissingText .= "Approved MFIP/DWP or meet current Waitlist criteria;`n"
    }
    If (clarifiedVerifications.Length() > 1) {
        clarifiedVerifications.InsertAt(1, (missingVerifications.Length() > 1 ? "`n" : "") "__Clarification of items listed above the Worker Comments:__`n", 1)
        ;clarifiedVerifications.InsertAt(1, "__Clarification of items listed above the Worker Comments:__`n", 1)
    }
    emailTextObject.output := setEmailText(emailTextString)
    listifyMissingObject := { 1missing: { arrayLines: 0, verificationList: missingVerifications }, 2clarified: { arrayLines: countLines(clarifiedVerifications), verificationList: clarifiedVerifications } }
    For objName, objContents in listifyMissingObject {
        listifyMissing( objContents )
    }
	caseNoteMissingText := Trim(caseNoteMissingText, "`n") ; removes the last new line
    For i, letterTextContents in letterText {
        If (InStr(letterTextContents, "__Clarification",,2)) {
            StrReplace(stWordWrap(letterTextContents, 60, ""), "`n", "`n", letterLineCount)
            If (letterLineCount < 27) {
                letterText[i] := StrReplace(letterTextContents, "__Clarification", "`n__Clarification")
            }
        }
		GuiControl, MissingGui:Show, % "letter" . i
    }
	GuiControl, MainGui: Text, 14MissingEdit, % caseNoteMissingText
	GuiControl, MissingGui: Show, % "emailButton"
	WinActivate, CaseNotes
    caseDetails.newChanges := false
}

parseMissingVerifications(ByRef missingVerifications, ByRef missingListEnum, ByRef clarifiedVerifications, ByRef clarifiedListEnum, ByRef emailTextString, ByRef emailListEnum, ByRef caseNoteMissingText) {
	Global
	If IDmissing {
        missingText := "ID for " missingInput.IDmissing ";`n"
		clarifiedVerifications[clarifiedListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= missingText
        clarifiedListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfIdentity := 1
    }
	If BCmissing {
        missingText := "Birth date / relationship / citizenship verification for: " missingInput.BCmissing
		clarifiedVerifications[clarifiedListEnum ". " missingText ";`n"] := 2
        emailTextString .= emailListEnum ". " missingText ". (Example: official Birth Certificate);`n"
		caseNoteMissingText .= missingText ";`n"
        clarifiedListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfBirth := 1
        mecCheckboxIds.proofOfRelation := 1
        mecCheckboxIds.citizenStatus := 1
    }
	If BCNonCitizenMissing {
        missingText := "Birth date / relationship / immigration verification for: " missingInput.BCNonCitizenMissing ";`n"
		clarifiedVerifications[clarifiedListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= missingText
        clarifiedListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfBirth := 1
        mecCheckboxIds.proofOfRelation := 1
        mecCheckboxIds.citizenStatus := 1
    }
    If PaternityMissing {
        missingText := "Paternity verification for: " missingInput.PaternityMissing ". (Examples: official Birth Certificate, Recognition of Parentage form. If the father is not listed, contact me for alternatives.)"
        tempText := getRowCount(missingListEnum ". " missingText ";", 60, "")
        missingVerifications[tempText[1] "`n"] := tempText[2]
        emailTextString .= emailListEnum ". " missingText ";`n"
		caseNoteMissingText .= "Paternity for  " missingInput.PaternityMissing ";`n"
        missingListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfRelation := 1
    }
	If AddressMissing {
        If (Homeless) {
            missingText := "Verification of current residence, such as a signed statement of your county of residence;`n"
            clarifiedVerifications[clarifiedListEnum ". " missingText] := 2
            emailTextString .= emailListEnum ". " missingText
            caseNoteMissingText .= "Address (homeless);`n"
            clarifiedListEnum++
            emailListEnum++
        } Else {
            missingText := "Verification of current residence;`n"
            emailTextString .= emailListEnum ". " missingText
            caseNoteMissingText .= "Address;`n"
            emailListEnum++
        }
        mecCheckboxIds.proofOfResidence := 1
    }
	If ChildSupportFormsMissing {
        If (missingInput.ChildSupportFormsMissing ~= "^\d$") {
            missingInput.ChildSupportFormsMissing .= missingInput.ChildSupportFormsMissing < 2 ? " set" : " sets"
        }
        missingText := "Cooperation with Child Support forms (" missingInput.ChildSupportFormsMissing ", sent separately);`n"
		missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "CS forms (" missingInput.ChildSupportFormsMissing ");`n"
		missingListEnum++
        emailListEnum++
    }
    sharedCustodyOptionsText := "  A. Stating that you have full custody, or`n  B. Your current Parenting Time (shared custody) schedule `n     listing the days and times of the custody switches;"
	If CustodyScheduleMissing {
        missingText := "A statement, written by you that is signed and dated, for each child needing CCAP that has a parent not in your household:`n" sharedCustodyOptionsText "`n"
		missingVerifications[missingListEnum ". " missingText] := 5
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Shared custody / parenting time;`n"
		missingListEnum++
        emailListEnum++
    }
	If CustodySchedulePlusNamesMissing {
        missingText := "A statement, written by you that is signed and dated, for " missingInput.CustodySchedulePlusNamesMissing ":`n" sharedCustodyOptionsText "`n"
		missingVerifications[missingListEnum ". " missingText] := 5
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Shared custody / parenting time for " missingInput.CustodySchedulePlusNamesMissing ";`n"
		missingListEnum++
        emailListEnum++
    }
    if DependentAdultStudentMissing {
        missingText := "Verification of full-time student status for " missingInput.DependentAdultStudentMissing ", verification of their most recent 30 days income, and a signed statement that you provide at least 50% of their financial support;`n"
        missingVerifications[missingListEnum ". " missingText] := 3
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Dependant Adult FT school status, income, statement of 50% support;`n"
		missingListEnum++
        emailListEnum++
    }
	If ChildSchoolMissing {
        missingText := "Child's school information (location, grade, start/end times) - does not need to be verified by the school;`n"
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Child school information;`n"
        mecCheckboxIds.childSchoolSchedule := 2
        emailListEnum++
        ;MEC2 text: Child School Schedule- You can provide the school schedule of each child that needs child care by sending a copy of the days and times of school from the school's website or handbook, writing the information on a piece of paper, or telling your worker.
    }
    If ChildFTSchoolMissing {
        missingText := "Verification of full-time student status for minor children with employment OR their most recent 30 days income (income is not counted if attending school full-time);`n"
        missingVerifications[missingListEnum ". " missingText] := 3
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Minor child FT school status or income;`n"
		missingListEnum++
        emailListEnum++
    }
	If MarriageCertificateMissing {
        missingText := "Marriage verification (example: marriage certificate);`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Marriage certificate;`n"
		missingListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfRelation := 1
    }
	If LegalNameChangeMissing {
        missingText := "Legal name change verification for " missingInput.LegalNameChangeMissing ";`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Legal name change for " missingInput.LegalNameChangeMissing ";`n"
		missingListEnum++
        emailListEnum++
    }
;======================================================
	If IncomeMissing {
        tempText := dateObject.needsExtension > -1 ? " your most recent 30 days income" : caseDetails.docType == "Redet" ? " 30 days income prior to " dateObject.RedetDueMDY : " 30 days income prior to " dateObject.receivedMDY
        ; IncomeText := if doesn't need extension : elseif redetermination : elseif app needs extension
        missingText := "Verification of" tempText ";`n"
        clarifiedVerifications[clarifiedListEnum ". Proof of Financial Information: " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Earned income;`n"
        clarifiedListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfFInfo := 1
        ;MEC2 text: Proof of Financial Information- You can provide proof of financial information and income with the last 30 days of check stubs, income tax records, business ledger, award letter, or a letter from your employer with pay rate, number of hours worked per week and how often you are paid.
    }
	If IncomePlusNameMissing {
        tempText := dateObject.needsExtension > -1 ? missingInput.IncomePlusNameMissing "'s most recent 30 days income" : caseDetails.docType == "Redet" ? missingInput.IncomePlusNameMissing "'s 30 days income prior to " dateObject.RedetDueMDY : missingInput.IncomePlusNameMissing "'s 30 days income prior to " dateObject.receivedMDY
        missingText := "Verification of " tempText ";`n"
        clarifiedVerifications[clarifiedListEnum ". Proof of Financial Information: " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Earned income (" missingInput.IncomePlusNameMissing ");`n"
        clarifiedListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfFInfo := 1
    }
	If WorkScheduleMissing {
        tempText := dateObject.needsExtension > -1 ? " your work schedule" : caseDetails.docType == "Redet" ? " work schedule from " dateObject.RedetDueMDY : " work schedule from " dateObject.receivedMDY
        missingText := "Verification of" tempText " showing days of the week and start/end times;`n"
        clarifiedVerifications[clarifiedListEnum ". Proof of Activity Schedule: " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Work schedule;`n"
        clarifiedListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfActivitySchedule := 1
        ;MEC2 text: Proof of Activity Schedule- You can provide proof of adult activity schedules with work schedules, school schedules, time cards, or letter from the employer or school with the days and times working or in school. If you have a flexible work schedule, include a statement with typical or possible times worked.
    }
	If WorkSchedulePlusNameMissing {
        tempText := dateObject.needsExtension > -1 ? missingInput.WorkSchedulePlusNameMissing "'s work schedule" : caseDetails.docType == "Redet" ? missingInput.WorkSchedulePlusNameMissing "'s work schedule from " dateObject.RedetDueMDY : missingInput.WorkSchedulePlusNameMissing "'s work schedule from " dateObject.receivedMDY
        missingText := "Verification of " tempText " showing days of the week and start/end times;`n"
        clarifiedVerifications[clarifiedListEnum ". Proof of Activity Schedule: " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Work schedule (" missingInput.WorkSchedulePlusNameMissing ");`n"
        clarifiedListEnum++
        emailListEnum++
        mecCheckboxIds.proofOfActivitySchedule := 1
    }
	If ContractPeriodMissing {
        missingText := "Employment Contract Period verification if not full-year;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Employment Contract Period;`n"
		missingListEnum++
        emailListEnum++
    }
	If NewEmploymentMissing {
        missingText := "Verification of employment start date, wage and expected hours per week, and first pay date;`n"
		missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "New employment information;`n"
		missingListEnum++
        emailListEnum++
    }
    If WorkLeaveMissing {
        missingText := "Verification of leave of absence, including: `nPaid/unpaid status, start date, and expected: return date, wage, and hours per week. Upon returning, we need your work schedule showing days of the week and start/end times;`n"
		missingVerifications[missingListEnum ". " missingText] := 4
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Leave of absence details;`n"
		missingListEnum++
        emailListEnum++
    }
;-----------------------------------------------------------------
    If SeasonalWorkMissing {
        missingText := "Verification of seasonal employment expected season length;`n"
		missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Seasonal employment season length;`n"
        emailListEnum++
        missingListEnum++
    }
    If SeasonalOffSeasonMissing {
        ;tempText := StrLen(SeasonalOffSeasonMissing) > 0 ? " at " SeasonalOffSeasonMissing : ""
        tempText := missingInput.SeasonalOffSeasonMissing != "" ? " at " missingInput.SeasonalOffSeasonMissing : ""
        missingText := "Verification of either seasonal employment " tempText ", including expected season length and typical wages, or a signed statement that you are no longer an employee at this job.`n Upon returning to work, verification of work schedule will`n be needed, showing days of the week and start/end times;`n"
		missingVerifications[missingListEnum ". " missingText] := 6
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Seasonal employment (applied during off season);`n"
        missingListEnum++
        emailListEnum++
    }
;-----------------------------------------------------------------
	If SelfEmploymentMissing {
        missingText := "Self-employment income such as your recent complete federal tax return. (For new self-employment, state your start date). If you haven't yet filed taxes or your taxes don't represent expected ongoing income, submit monthly reports or ledgers with the most recent full 3 months of gross income;`n"
        ;MEC2 text: Proof of Financial Information- You can provide proof of financial information and income with the last 30 days of check stubs, income tax records, business ledger, award letter, or a letter from your employer with pay rate, number of hours worked per week and how often you are paid. 
		missingVerifications[missingListEnum ". " missingText] := 5
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Self-Employment income;`n"
		missingListEnum++
        emailListEnum++
    }
	If SelfEmploymentScheduleMissing {
        missingText := "Written statement of your self-employment work schedule with days of the week and start/end times;`n"
        ;MEC2 text: Proof of Activity Schedule- You can provide proof of adult activity schedules with work schedules, school schedules, time cards, or letter from the employer or school with the days and times working or in school. If you have a flexible work schedule, include a statement with typical or possible times worked.
		missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Self-Employment work schedule;`n"
		missingListEnum++
        emailListEnum++
    }
    If SelfEmploymentBusinessGrossMissing {
        missingText := "Information regarding your self-employment business' annual gross income, if it is less than $500,000 (optional);`n"
        missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Self-Employment gross (if subject to small/large min wage: <$500k/yr?) - not required;`n"
		missingListEnum++
        emailListEnum++
    }
;-----------------------------------------------------------------
	If ExpensesMissing {
        missingText := "Proof of Expenses: Healthcare Insurance premiums, child support, and spousal support - if not listed on submitted paystubs;`n"
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Expenses;`n"
        emailListEnum++
        mecCheckboxIds.proofOfDeductions := 1
        ;MEC2 text: Proof of Deductions- You can provide proof of expenses for health insurance premiums (medical, dental, vision), child support paid for a child not living in your home, and spousal support with check stubs, benefit statements or premium statements. 
    }
;==================================================================================
	If ChildSupportIncomeMissing {
        missingText := "Verification of your Child Support income;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Child Support income;`n"
		missingListEnum++
        emailListEnum++
    }
	If SpousalSupportMissing {
        missingText := "Verification of your Spousal Support income;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Spousal Support income;`n"
		missingListEnum++
        emailListEnum++
    }
	If RentalMissing {
        missingText := "Verification of your rental income;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Rental income;`n"
		missingListEnum++
        emailListEnum++
    }
	If DisabilityMissing {
        missingText := "Verification of your disability income;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "STD / LTD;`n"
		missingListEnum++
        emailListEnum++
    }
	If InsuranceBenefitsMissing {
        missingText := "Verification of your Insurance Benefits income;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Insurance benefits income;`n"
		missingListEnum++
        emailListEnum++
    }
    If UnearnedStatementMissing {
        missingText := "A statement written by you that is signed and dated, stating if you have any unearned income. Submit verification if yes.`nThis includes: Child/Spousal support, Rentals, Unemployment, RSDI, Insurance payments, VA benefits, Trust income, Contract for deed, Interest, Dividends, Gambling winnings, Inheritance, Capital gains, etc.;`n"
        missingVerifications[missingListEnum ". " missingText] := 6
        emailTextString .= emailListEnum ". " missingText
        caseNoteMissingText .= "Unearned income yes / no questions (statement);`n"
        missingListEnum++
        emailListEnum++
    }
	If VABenefitsMissing {
        missingText := "Verification of your VA income;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "VA income;`n"
		missingListEnum++
        emailListEnum++
    }
    If UnearnedMailedMissing {
        missingText := "Unearned income questions that were not answered (sent separately);`n"
        missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
        caseNoteMissingText .= "Unearned income yes / no questions (mailed back);`n"
        missingListEnum++
        emailListEnum++
    }
	If AssetsBlankMissing {
        missingText := "Written or verbal statement of your assets being either MORE THAN or LESS THAN $1 million;`n"
		missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Assets amount statement;`n"
		missingListEnum++
        emailListEnum++
    }
	If AssetsGT1mMissing {
        missingText := "Clarification of your assets, which you listed as MORE THAN $1 million;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Assets clarification (>$1m on app);`n"
		missingListEnum++
        emailListEnum++
    }
;======================================================
	If EdBSFformMissing {
        missingText := ini.caseNoteCountyInfo.countyEdBSF " form (sent separately);`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= ini.caseNoteCountyInfo.countyEdBSF " form;`n"
		missingListEnum++
        emailListEnum++
    }
	If ClassScheduleMissing {
        missingText := "Class schedule with class start/end times and credits;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Adult class schedule;`n"
		missingListEnum++
        emailListEnum++
    }
	If TranscriptMissing {
        missingText := "Unofficial post-secondary transcript/academic record;`n"
		missingVerifications[missingListEnum ". " missingText] := 1
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Post-secondary transcript;`n"
		missingListEnum++
        emailListEnum++
    }
	If EducationEmploymentPlanMissing {
        missingText := "Cash Assistance Employment Plan listing your education activity and schedule;`n"
		missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "ES Plan with education activity and schedule;`n"
		missingListEnum++
        emailListEnum++
    }
    If StudentStatusOrIncomeMissing {
        missingText := "Verification of your student status of being at least halftime, OR your most recent 30 days income.`n (if you are 19 or under and attending school at least`n   halftime, your income is not counted);`n"
		missingVerifications[missingListEnum ". " missingText] := 4
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Halftime+ student status or income (PRI age 19 or under);`n"
		missingListEnum++
        emailListEnum++
    }
;-------------------------
	If JobSearchHoursMissing {
        missingText := "Job search hours needed per week: Assistance can be approved for 1 to 20 hours of job search each week, limited to a total of 240 hours per calendar year;`n"
		missingVerifications[missingListEnum ". " missingText] := 3
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Job search hours per week;`n"
		missingListEnum++
        emailListEnum++
    }
    If ESPlanUpdateMissing {
        missingText := "Updated Employment Plan ...;`n"
		missingVerifications[missingListEnum ". " missingText] := 4
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "Updated Employment Plan ...;`n"
		missingListEnum++
        emailListEnum++
    }
    If UnableToProvideCareMissing {
        missingText := "Parent Medical Condition Form (DHS-6305) or documentation from a licensed practitioner with the parent's condition and limitations, which children, and the time period.`n"
		missingVerifications[missingText] := 3
        emailTextString .= missingText
		caseNoteMissingText .= "`n"
    }
    For i, missingText in otherMissing {
        If (!otherInput%i%) { ; not checked
            continue
        }
        missingText := RTrim(missingText, ".")
        caseNoteMissingText .= missingText ";`n"
        tempText := getRowCount(missingListEnum ". " missingText ";", 60, "")
        missingVerifications[tempText[1] "`n"] := tempText[2]
        emailTextString .= emailListEnum ". " missingText ";`n"
        missingListEnum++
        emailListEnum++
    }
;======================================================
	If InHomeCareMissing {
        missingText := "In-Home Care form (sent separately) - In-Home Care requires approval by MN DHS;`n"
		missingVerifications[missingListEnum ". " missingText] := 2
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "In-Home Care form;`n"
		missingListEnum++
        emailListEnum++
    }
	If LNLProviderMissing {
        missingText := "Legal Non-Licensed Acknowledgement (sent separately). Your provider may not be eligible to be paid for care provided prior to them completing age specific trainings. The ‘Health and Safety Resources’ documents do not need to be completed or returned;"
		missingVerifications[missingListEnum ". " missingText] := 5
        emailTextString .= emailListEnum ". " missingText
		caseNoteMissingText .= "LNL Acknowledgement form;`n"
		missingListEnum++
        emailListEnum++
    }
    If StartDateMissing {
        StartDateMissingText := "Start date at your child care provider;`n"
		missingVerifications[missingListEnum ". " StartDateMissingText] := 1
        emailTextString .= emailListEnum ". " StartDateMissingText
		caseNoteMissingText .= "Provider start date;`n"
		missingListEnum++
        emailListEnum++
    }
	If ChildSupportNoncooperationMissing {
        missingText := "* You are currently in a non-cooperation status with Child Support. Contact Child Support at " missingInput.ChildSupportNoncooperationMissing " for details. Child Support cooperation is a requirement for eligibility.`n"
		missingVerifications[missingText] := 3
        emailTextString .= missingText
		caseNoteMissingText .= "Cooperation status with Child Support, CS number: " missingInput.ChildSupportNoncooperationMissing ";`n"
    }
	If EdBSFOneBachelorDegreeMissing {
        missingText := "* Unless listed on a Cash Assistance Employment Plan, education is an eligible activity only up to your first bachelor's degree, plus CEUs (no additional degrees).`n"
		missingVerifications[missingText] := 3
        emailTextString .= missingText
		caseNoteMissingText .= "* Client informed only up to first bachelor's degree is BSF/TY eligible;`n"
    }

    EligibleActivityWithJSText := "Eligible activities are:`n  A. Employment of 20+ hours per week (10+ for FT students)`n  B. Education with an approved plan`n  C. Job Search up to 20 hours per week`n  D. Activities on a Cash Assistance Employment Plan"
    EligibleActivityWithoutJSText := "Eligible activities are:`n  A. Employment of 20+ hours per week (10+ for FT students)`n  B. Education with an approved plan`n  C. Activities on a Cash Assistance Employment Plan"

    If SelfEmploymentIneligibleMissing {
        missingText := "* Your self-employment does not meet activity requirements. Self-employment hours are calculated using 50% of recent gross income, or gross minus expenses on tax return divided by minimum wage. " EligibleActivityWithJSText "`n"
		missingVerifications[missingText] := 8
        emailTextString .= missingText
		caseNoteMissingText .= "Self-employment hours meeting minimum requirement, or other eligible activity;`n"
    }
    If EligibleActivityMissing {
        missingText := "* You did not select an eligible activity on the " mec2docType ". " EligibleActivityWithJSText "`n"
		missingVerifications[missingText] := 6
        emailTextString .= missingText
		caseNoteMissingText .= "Eligible activity (none selected on form);`n"
    }
    If EmploymentIneligibleMissing {
        missingText := "* Your employment does not meet eligible activity requirements. " EligibleActivityWithJSText "`nYou can submit up to 6 months of recent paystubs to meet the requirement.`n"
		missingVerifications[missingText] := 8
        emailTextString .= missingText
		caseNoteMissingText .= "Employment hours meeting minimum requirement, or other eligible activity;`n"
    }
    If ESPlanOnlyJSMissing {
        missingText := "* While you have an Employment Plan, assistance hours cannot be approved for job search unless it is listed on the Plan"
		missingVerifications[missingText ";`n"] := 2
        emailTextString .= missingText ". Contact your Job Counselor to have an updated Plan written if job search hours are needed;`n"
		caseNoteMissingText .= "Client has ES Plan - informed JS hours are required to be on the Plan;`n"
    }
	If ActivityAfterHomelessMissing {
        missingText := "* At the end of the 90-day homeless exemption period, you must have an eligible activity to keep your Child Care Assistance case open. " EligibleActivityWithoutJSText "`n"
		missingVerifications[missingText] := 6
        emailTextString .= missingText
		caseNoteMissingText .= "Eligible activity after the 3-month homeless period;`n"
    }
    If MightBeUnableToProvideCareMissing {
        missingText := "* If a parent is unable to provide care for a child, they can be exempted from the activity requirement. A licensed practitioner must either complete DHS-6305, or document the parent's condition and limitations, which children, and the time period as per form DHS-6305. Contact me for details.`n"
		missingVerifications[missingText] := 5
        emailTextString .= missingText
		caseNoteMissingText .= "`n"
    }
	If NoProviderMissing {
        missingText := "`n* Once you have a daycare provider, please notify me with the provider’s name, location, and the start date.`n   If you need help locating a daycare provider, contact Parent Aware at 888-291-9811 or www.parentaware.org/search`n"
        emailTextString .= missingText
		caseNoteMissingText .= "Provider;`n"
        mecCheckboxIds.providerInformation := 1
    }
    ;*   Provider Information- If you have a child care provider, send the provider's name, address and start date (if known). Visit www.parentaware.org for help finding a provider. Care is not approved until you get a Service Authorization.
	If UnregisteredProviderMissing {
        missingText := "* Your daycare provider is not registered with Child Care Assistance. Please have them call " ini.caseNoteCountyInfo.countyProviderWorkerPhone " to register.`n"
		missingVerifications[missingText] := 2
        emailTextString .= missingText
		caseNoteMissingText .= "Registered provider;`n"
    }
    If ProviderForNonImmigrantMissing {
        missingText := "* If your child is not a US citizen, Lawful Permanent Resident, Lawfully residing non-citizen, or fleeing persecution, assistance can only be approved at a daycare that is subject to public educational standards (Head Start, pre-K, school age program).`n"
        missingVerifications[missingText] := 4
        emailTextString .= missingText
        caseNoteMissingText .= "Provider subject to Public Educational Standards (4.15), if child not citizen/immigrant;`n"
    }
        ;If (UnearnedUnansweredMissing || LumpSumUnansweredMissing || EmploymentUnansweredMissing || SelfEmploymentUnansweredMissing || AssetsUnansweredMissing) {
        ;UnansweredText := "You did not answer all questions on the " mec2docType ".`n Please submit a statement that is written, dated, and signed by you, answering:`n"
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
        ;If (StrLen(AnsweredYes AnsweredMoreThan) == 14) {
            ;AnsweredBoth == " or"
        ;}
        ;UnansweredText .= "* Submit verification if you answered" AnsweredYes AnsweredBoth AnsweredMoreThan ". *`n"
        ;UnansweredText := stWordWrap(UnansweredText, 60, " ")
			;UnansweredTextCount := 0
			;StrReplace(UnansweredText, "`n", "`n", UnansweredTextCount)
			;UnansweredTextCount++
            ;missingVerifications[missingListEnum ". " UnansweredText] := UnansweredTextCount
            ;emailTextString .= emailListEnum ". " UnansweredText
			;caseNoteMissingText .= CaseNoteUnansweredMissing ";`n"
			;missingListEnum++
            ;emailListEnum++
    ;}
}

incrementLetterPage(ByRef letterTextNumber) {
    letterText[letterTextNumber] .= "                   Continued on letter " letterTextNumber+1
    letterText[letterTextNumber+1] .= "                  Continued from letter " letterTextNumber "`n"
    letterTextNumber++
}
faxAndEmailText() {
    contactMethodsList := { countyFax: "faxed to ", countyDocsEmail: "emailed to " }
    contactMethods := []
    For method, methodText in contactMethodsList {
        contactInfo := ini.caseNoteCountyInfo[method]
        If (StrLen(contactInfo) > 1) {
            contactMethods.push(methodText contactInfo)
        }
    }
    contactMethodsLength := contactMethods.Length()
    If (contactMethodsLength == 0) {
        Return
    }
    specLetterText := "Documents can mailed"
    For i, contactText in contactMethods {
        specLetterText .= (i == contactMethodsLength) ? (contactMethodsLength > 1 ? "," : "") " and " : ", "
        specLetterText .= contactText
    }
    specLetterText .= ". Please include your case number."
    Return specLetterText
}
countLines(VerificationArray) {
    totalLines := 0
    For i, lineCountAmt in VerificationArray {
        totalLines += lineCountAmt
    }
    Return totalLines
}
listifyMissing(verifObj) {
    Global
    If ((lineCount + verifObj.arrayLines) > 30) { ; puts clarifiedVerifications on the next letter if it will exceed the current letter's available space
        incrementLetterPage(letterTextNumber)
        lineCount := 1 ; For the continued from line
    }
    For missingText, missingLineCount in verifObj.verificationList {
        If (InStr(missingText, "NEWLINE")) { ; last missingText in group
            lineCountPlusFaxed := lineCount + missingLineCount
            If (lineCountPlusFaxed == 30) {
                missingText := StrReplace(missingText, "NEWLINE", "")
                missingText := StrReplace(missingText, "NEWLINE", "")
            } Else If (lineCountPlusFaxed > 30) {
                missingText := StrReplace(missingText, "NEWLINE", "`n")
                missingText := StrReplace(missingText, "NEWLINE", "`n")
                incrementLetterPage(letterTextNumber)
            } Else {
                While (InStr(missingText, "NEWLINE") && lineCountPlusFaxed < 30) {
                    missingText := StrReplace(missingText, "NEWLINE", "`n",,1)
                    lineCountPlusFaxed++
                }
                missingText := StrReplace(missingText, "NEWLINE", "")
            }
            letterText[letterTextNumber] .= missingText
        } Else { ; does not contain "NEWLINE"
            If ((lineCount + missingLineCount) > 29) {
            
                incrementLetterPage(letterTextNumber)
                
                letterText[letterTextNumber] .= missingText
                lineCount := (missingLineCount + 1) ; For the continued from line
            } Else {
                lineCount += missingLineCount
                letterText[letterTextNumber] .= missingText
            }
        }
    }
}
emailButtonClick() {
    Clipboard := getFirstName() emailTextObject.output
    WinActivate, % "Message - "
    Send, ^v
}
setEmailText(emailTextStringIn) {
	Gui, MainGui: Submit, NoHide
	Gui, MissingGui: Submit, NoHide
    emailTextStringOut := StrReplace(emailTextStringIn, "`n ", " ")
    emailTextStringOut := StrReplace(emailTextStringOut, "    ", " ")
    emailTextStringOut := StrReplace(emailTextStringOut, "   ", " ")
    emailTextStringOut := StrReplace(emailTextStringOut, "  ", "`n  ")
    emailTextStringOut := StrReplace(emailTextStringOut, "`n*", "`n`n*")
    emailTextStringOut := StrReplace(emailTextStringOut, "sent separately", "see attached")
	return emailTextObject.Combined emailTextStringOut emailTextObject.EndHL
}
letterButtonClick(letterGUINumber := 1) {
    Global
	Gui, MainGui: Submit, NoHide
	Gui, MissingGui: Submit, NoHide
	letterGUINumber := LTrim(A_GuiControl, "letter")
    If (Homeless == 1 && caseDetails.eligibility == "pends" && StrLen(missingHomelessItems) < 1) {
        missingVerifsDoneButton()
    }
    thisLetterText := Trim(letterText[letterGUINumber], "`n")
    If (ini.employeeInfo.employeeUseMec2Functions == 1) {
        ;caseStatus := InStr(caseDetails.docType, "?") ? "" : (caseDetails.docType == "Redet") ? "Redetermination" : (Homeless == 1) ? "Homeless App" : caseDetails.docType
        caseStatus := InStr(caseDetails.docType, "?") ? "" : (caseDetails.docType == "Redet") ? "Redetermination" : (Homeless == 1 && caseDetails.eligibility == "elig") ? "Homeless App" : caseDetails.docType
        jsonLetterText := "LetterTextFromAHKJSON{""LetterText"":""" JSONstring(thisLetterText) """,""CaseStatus"":""" caseStatus """,""IdList"":""" idList """ }"
        Clipboard := jsonLetterText
    } Else {
        Clipboard := thisLetterText
    }
    WinActivate % ini.employeeInfo.employeeBrowser
    Sleep 500
    Send, ^v
    Sleep 500
    Clipboard := caseNumber
}
otherGUI() {
    Global
    Gui, MissingGui: Submit, NoHide
    If (%A_GuiControl% == 0) { ; unchecked
        Return
    }
    local otherEditWidth := (wCH*60)+scrollbar+pad
    local otherNumber := Trim(A_GuiControl, "otherInput")
    Gui, OtherGui: New,, % "Other Verification"
    Gui, OtherGui: Margin, % marginW
    Gui, OtherGui: Font, s9, % "Lucida Console"
    Gui, OtherGui: Add, Text,, % "Additional Input Required:`nState what the client needs to submit."
    Gui, OtherGui: Add, Edit, % "h100 w" otherEditWidth " votherEdit", % otherMissing[otherNumber]
    Gui, OtherGui: Add, Button, % "gOtherGuiGuiClose x175", % "Save"
    Gui, OtherGui: Add, Button, % "gOtherGuiGuiClose x+20 yp", % "Close"
    Gui, OtherGui: Add, Edit, % "Hidden votherInputID", % otherNumber
    Gui, OtherGui: Show, % "Hide x" ini.caseNotePositions.xCaseNotes " y" ini.caseNotePositions.yCaseNotes
    GuiControlGet editSize, OtherGui:POS, otherEdit
    guiX := caseNotesMonCenter[1] - ((editSizeW*zoomPPI + scrollbar)/2), guiY := caseNotesMonCenter[2] - (editSizeH*zoomPPI/2)
    Gui, OtherGui: Show, % "x" guiX " y" guiY
    Gui, OtherGui: Show, AutoSize
    Gui, OtherGui: +OwnerMissingGui
    Gui, MissingGui: +Disabled
}
OtherGuiGuiClose() {
    Global
    Gui, OtherGui: Submit, NoHide
    If (A_GuiControl == "Save" && otherEdit != "") {
        GuiControl, MissingGui: Text, % "otherInput" . otherInputID, % otherEdit
        otherMissing[otherInputID] := Trim(otherEdit, "`n")
    } Else {
        If (otherEdit == "") {
            GuiControl, MissingGui: Text, % "otherInput" . otherInputID, Other
        }
        GuiControl, MissingGui:, % "otherInput" . otherInputID, 0
    }
    Gui, OtherGui: Destroy
    Gui, MissingGui: -Disabled
    WinActivate, % "Missing Verifications"
}
inputBoxAGUIControl() {
    Global
    Gui, Submit, NoHide
    Gui +OwnDialogs
    If (%A_GuiControl% == 0) { ; unchecked
        Return
    }
    inputBoxDefaultText := A_GuiControl == "ChildSupportFormsMissing" ? Trim(missingInput[A_GuiControl], "sets") : Trim(missingInput[A_GuiControl], " (input)")
    InputBox, inputBoxInput, % "Additional Input Required", % missingInputObject[A_GuiControl].promptText,,,,,,,, % inputBoxDefaultText
	If (ErrorLevel) {
        GuiControl, MissingGui:, % A_GuiControl, 0
		Return
    }
    If (StrLen(inputBoxInput) == 0) {
        GuiControl, MissingGui: Text, % A_GuiControl, % missingInputObject[A_GuiControl].baseText " (input)"
        GuiControl, MissingGui:, % A_GuiControl, 0 ; uncheck if blank
        Return
    }
    If (A_GuiControl == "ChildSupportFormsMissing") {
        inputBoxInput .= StrLen(inputBoxInput) == 1 ? ( (inputBoxInput < 2 ? " set" : " sets") ) : ""
    } Else If (A_GuiControl == "overIncomeMissing") {
        overIncomeSub(inputBoxInput)
    }
    GuiControl, MissingGui: Text, % A_GuiControl, % missingInputObject[A_GuiControl].baseText missingInputObject[A_GuiControl].inputAdject inputBoxInput
    missingInput[A_GuiControl] := inputBoxInput ; set to global object
}
overIncomeSub(overIncomeString) {
    overIncomeEntriesArray := StrSplit(overIncomeString, A_Space, ",", -1)
    If (StrLen(overIncomeEntriesArray[3]) > 0) {
        overIncomeObj.overIncomeHHsize := overIncomeEntriesArray[3]
    }
    overIncomeObj.overIncomeReceived := Round(StrReplace(overIncomeEntriesArray[1], ","))
    overIncomeObj.overIncomeLimit := StrReplace(overIncomeEntriesArray[2], ",")
    overIncomeObj.overIncomeText := overIncomeObj.overIncomeLimit "; your income is calculated as $" overIncomeObj.overIncomeReceived
    overIncomeObj.overIncomeDifference := overIncomeObj.overIncomeReceived - overIncomeObj.overIncomeLimit
    GuiControl, MissingGui: Text, % A_GuiControl, % "Over-income by $" overIncomeObj.overIncomeDifference
    missingInput[A_GuiControl] := inputBoxInput
}
getFirstName() {
    Global
    Gui, MainGui: Submit, NoHide
    RegExMatch(01HouseholdCompEdit, "[A-Za-z\-]+", firstName)
    addedName := StrLen(firstName) ? firstName ", `n`n" : ""
    Return addedName
}
;VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION - VERIFICATION SECTION
;=====================================================================================================================================================================================

;===================================================================================================================================================================================
;ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION
MainGuiGuiClose() {
    Global
    closingPromptText := ""
    If (caseNoteEntered.confirmedClear > 0) {
        coordSaveAndExitApp(1)
    }
    closingPromptText .= caseNoteEntered.mec2NoteEntered == 0 ? " MEC2" : ""
    If (ini.caseNoteCountyInfo.countyNoteInMaxis == 1 && caseDetails.docType == "Application" && caseNoteEntered.maxisNoteEntered == 0) {
        closingPromptText .= StrLen(closingPromptText) > 0 ? " or MAXIS" : " MAXIS"
    }
    If (A_GuiControl == "ClearFormButton") {
        If (StrLen(closingPromptText) > 0) {
            MsgBox, 4, % "Case Note Prompt", % "Case note not entered in" closingPromptText ". `nClear form anyway?"
            IfMsgBox No
                Return 1
            coordSaveAndExitApp(1)
        }
        GuiControl, MainGui: Text, ClearFormButton, % "Confirm"
        Gui, Font, s9, Segoe UI
        GuiControl, MainGui: Font, ClearFormButton
        caseNoteEntered.confirmedClear++
    } Else {
        If (StrLen(closingPromptText) == 0) {
            coordSaveAndExitApp()
        } Else {
            MsgBox, 4, % "Case Note Prompt", % "Case note not entered in" closingPromptText ". `nExit anyway?"
            IfMsgBox No
                Return 1
            coordSaveAndExitApp()
        }
    }
}
MissingGuiGuiClose() {
    WinGetPos, XVerificationGet, YVerificationGet,,, A
    For i, coordName in ["xVerification", "yVerification"] {
        ini.caseNotePositions[Value] := %coordName%Get
    }
	Gui, MissingGui: Hide
}
CBTGuiClose() {
    WinGetPos, xClipboardGet, yClipboardGet,,, % "Clipboard Text"
    If (xClipboardGet == "")
        Return
    If ((xClipboardGet - ini.cbtPositions.xClipboard + yClipboardGet - ini.cbtPositions.yClipboard) != 0) {
        coordObjOut := {}
        For i, coordName in ["xClipboard", "yClipboard"] {
            coordObjOut[coordName] := %coordName%Get
            ini.cbtPositions[coordName] := %coordName%Get
        }
        coordString := coordStringify(coordObjOut)
        IniWrite, %coordString%, %A_MyDocuments%\AHK.ini, cbtPositions
    }
    Gui, CBT: Destroy
}
HelpGuiGuiClose() {
    Gui, HelpGui: Destroy
}
;ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION - ON WINDOW CLOSE SECTION
;===================================================================================================================================================================================

;========================================================================================================================================================================
;SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION
buildSettingsGui(checkOnOpen := 0) {
    Global
    If (checkOnOpen == 1 && StrLen(ini.employeeInfo.employeeName) > 0) { 
        Return
    }
    countyContact := { Default: { Email: "" } } ;v2 reset to continuation version
    countyContact.Dakota := { Email: "EEADOCS@co.dakota.mn.us", Fax: "651-306-3187", ProviderWorker: "651-554-5764", EdBSF: "Training Request for Childcare", countyNoteInMaxis: 1 }
    countyContact.StLouis := { Email: "ess@stlouiscountymn.gov", Fax: "218-733-2976", ProviderWorker: "218-726-2064", EdBSF: "SLC CCAP Education Plan", countyNoteInMaxis: 0 }
    ini.caseNoteCountyInfo.countyFax := ini.caseNoteCountyInfo.countyFax != " " ? ini.caseNoteCountyInfo.countyFax : countyContact[ini.employeeInfo.employeeCounty].Fax
    ini.caseNoteCountyInfo.countyDocsEmail := ini.caseNoteCountyInfo.countyDocsEmail != " " ? ini.caseNoteCountyInfo.countyDocsEmail : countyContact[ini.employeeInfo.employeeCounty].Email
    ini.caseNoteCountyInfo.countyProviderWorkerPhone := ini.caseNoteCountyInfo.countyProviderWorkerPhone != " " ? ini.caseNoteCountyInfo.countyProviderWorkerPhone : countyContact[ini.employeeInfo.employeeCounty].ProviderWorker
    ini.caseNoteCountyInfo.countyEdBSF := ini.caseNoteCountyInfo.countyEdBSF != " " ? ini.caseNoteCountyInfo.countyEdBSF : countyContact[ini.employeeInfo.employeeCounty].EdBSF
    
    local editboxOptions := "x200 yp-3 h18 w200"
    local checkboxOptions := "x200 yp-3 h18 w20"
    local textLabelOptions := "xm w170 h18 Right"
    Gui, Font,, % "Lucida Console"
    Gui, Color, % "989898", % "a9a9a9"
    Gui, SettingsGui: Margin, % marginW
    Gui, SettingsGui: New, AlwaysOnTop ToolWindow,
    Gui, SettingsGui: Add, Text, % textLabelOptions " y12", % "Worker Name:"
    Gui, SettingsGui: Add, Edit, % editboxOptions " vEmployeeNameWrite", % ini.employeeInfo.employeeName
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Worker Phone:"
    Gui, SettingsGui: Add, Edit, % editboxOptions " vEmployeePhoneWrite", % ini.employeeInfo.employeePhone
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Worker Email:"
    Gui, SettingsGui: Add, Edit, % editboxOptions " vEmployeeEmailWrite", % ini.employeeInfo.employeeEmail
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Use Worker Email in Letters:"
    Gui, SettingsGui: Add, CheckBox, % "vEmployeeUseEmailWrite " checkboxOptions " Checked" ini.employeeInfo.employeeUseEmail
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Using mec2functions:"
    Gui, SettingsGui: Add, CheckBox, % "vEmployeeUseMec2FunctionsWrite gworkerUsingMec2Functions " checkboxOptions " Checked" ini.employeeInfo.employeeUseMec2Functions
    Gui, SettingsGui: Add, ComboBox, % "x+10 yp vEmployeeBrowserWrite Choose1 R4 Hidden", % ini.employeeInfo.employeeBrowser "|Google Chrome|Mozilla Firefox|Microsoft Edge"
    If (ini.employeeInfo.employeeUseMec2Functions == 1) {
        GuiControl, SettingsGui: Show, EmployeeBrowserWrite
    }
    Gui, SettingsGui: Add, Text, h0 w0 y+10
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Select a county to auto-populate"
    Gui, SettingsGui: Add, ComboBox, % editboxOptions " vEmployeeCountyWrite gcountySelection Choose1 R4", % ini.employeeInfo.employeeCounty "|Dakota|StLouis|Not Listed"
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Case Note in MAXIS:"
    Gui, SettingsGui: Add, CheckBox, % "vcountyNoteInMaxisWrite gcountyNoteInMaxis " checkboxOptions " Checked" ini.caseNoteCountyInfo.countyNoteInMaxis
    Gui, SettingsGui: Add, Edit, % "x+10 yp h18 w170 vEmployeeMaxisWrite Hidden", % ini.employeeInfo.employeeMaxis
    If (ini.caseNoteCountyInfo.countyNoteInMaxis == 1) {
        GuiControl, SettingsGui: Show, EmployeeMaxisWrite
    }
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Fax Number:"
    Gui, SettingsGui: Add, Edit, % editboxOptions " vCountyFaxWrite", % ini.caseNoteCountyInfo.countyFax
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "County Documents Email:"
    Gui, SettingsGui: Add, Edit, % editboxOptions " vCountyDocsEmailWrite", % ini.caseNoteCountyInfo.countyDocsEmail
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Provider Worker Phone:"
    Gui, SettingsGui: Add, Edit, % editboxOptions " vCountyProviderWorkerPhoneWrite", % ini.caseNoteCountyInfo.countyProviderWorkerPhone
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "BSF Education Form Name:"
    Gui, SettingsGui: Add, Edit, % editboxOptions " vCountyEdBSFWrite", % ini.caseNoteCountyInfo.countyEdBSF
    Gui, SettingsGui: Add, Text, % textLabelOptions, % "Waiting List Priority"
    Gui, SettingsGui: Add, DropDownList, % editboxOptions " vWaitlistWrite R4 AltSubmit Choose" ini.caseNoteCountyInfo.Waitlist, % "None|HS / GED / ESL|A PRI is a veteran|All others"
    Gui, SettingsGui: Add, Button, % "w80 gupdateIniFile", % "Save"
    Gui, SettingsGui:+OwnerMainGui
    Gui, SettingsGui: Show,w450, % "Update CaseNotes Settings"
}
workerUsingMec2Functions() {
    GuiControlGet, EmployeeUseMec2FunctionsWrite
    If (EmployeeUseMec2FunctionsWrite == 0) {
        GuiControl, SettingsGui: Hide, EmployeeBrowserWrite
        Return
    }
    GuiControl, SettingsGui: Show, EmployeeBrowserWrite
}
countySelection() {
    Global
    GuiControlGet, EmployeeCountyWrite
    ini.employeeInfo.employeeCounty := EmployeeCountyWrite
    GuiControl, SettingsGui: Text, CountyFaxWrite, % countyContact[ini.employeeInfo.employeeCounty].Fax
    GuiControl, SettingsGui: Text, CountyProviderWorkerPhoneWrite, % countyContact[ini.employeeInfo.employeeCounty].ProviderWorker
    GuiControl, SettingsGui: Text, CountyDocsEmailWrite, % countyContact[ini.employeeInfo.employeeCounty].Email
    GuiControl, SettingsGui: Text, CountyEdBSFWrite, % countyContact[ini.employeeInfo.employeeCounty].EdBSF
    GuiControl,, countyNoteInMaxisWrite, % countyContact[ini.employeeInfo.employeeCounty].countyNoteInMaxis
    countyNoteInMaxis()
}
countyNoteInMaxis() {
    GuiControlGet, countyNoteInMaxisWrite
    If (countyNoteInMaxisWrite == 0) {
        GuiControl, SettingsGui: Hide, EmployeeMaxisWrite
        Return
    }
    GuiControl, SettingsGui: Show, EmployeeMaxisWrite
}
updateIniFile() {
    Gui, SettingsGui: Submit, NoHide
    ;If (countyNoteInMaxisWrite && EmployeeMaxisWrite == "MAXIS-WINDOW-TITLE") { change border of EmployeeMaxisWrite, blink, dance, return? }
     ;v2 reset to continuation version
    settingsArrays := { employeeInfo: [ "employeeName", "employeePhone", "employeeEmail", "employeeUseEmail", "employeeUseMec2Functions", "employeeBrowser", "employeeCounty", "employeeMaxis" ], caseNoteCountyInfo: [ "countyFax", "countyDocsEmail", "countyProviderWorkerPhone", "countyEdBSF", "countyNoteInMaxis", "Waitlist" ] }
    For section, settingArray in settingsArrays {
        updateIniFileText(section, settingArray)
    }
    checkGroupAdd()
    Gui, SettingsGui: Destroy
}
updateIniFileText(section, settingArray) {
    iniSettingsValues := ""
    For i, settingName in settingArray {
        iniSettingsValues .= settingName "=" %settingName%Write "`n"
        ini[section][settingName] := %settingName%Write
    }
    IniWrite, % iniSettingsValues, % A_MyDocuments "\AHK.ini", % section
}
;SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION - SETTINGS SECTION
;========================================================================================================================================================================

;=======================================================================================================================================================================================================
;EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION  EXAMPLES/HELP SECTION  EXAMPLES/HELP SECTION 
buildHelpGui() {
    local widestText
    Paragraph1 := "
    (
    This app is a template tool for generating Case Notes for CCAP applications, redeterminations, and request letters
    for Special Letters and emails. This tool is not endorsed or sponsored by MN DCYF.
    )"
    Paragraph2 := "
    (
    Features:
    ● Auto-formats text to fit within MEC" sq "'s Case Notes and Special Letter/Memo fields.
       and is designed to be compatible with the Income Calculator spreadsheet.
    ● Case Notes are formatted with categories and spacing for consistant alignment. 
    ● User-entered dates calculate the extended auto-deny date (for a minimum of 15ish days from the date processed).
    ● Incorporates document type, approval status, and dates in the notes and verification requests.
    ● Compatible with mec2functions from github.com/MECH2-at-github.
    ● Case Notes can be saved to a text document in the event a case is locked or otherwise inaccessible.
    ● Special Letter requests are broken down into 'clarifications' of checkbox items, and additional items.
    )"
    Paragraph3 := "
    (
    Main window :
    ● [MEC" sq " Note] - Formats the entire case note and sends the data to MEC" sq ".
        If you are not using mec2functions, it will simulate keypresses to navigate the page. 
        In MEC" sq ": Click 'New' in MEC" sq " on the CaseNotes webpage. In CaseNotes, click [MEC" sq " Note].
    ● [MAXIS Note] - Visible only if ""Case Note in MAXIS"" is checked in Settings.
        Formats the app date, case status, and verifications list and sends it to MAXIS.
        It will activate BlueZone and paste the case note in.
        In BlueZone (MAXIS): PF9 to start a new note. In CaseNotes, click [MAXIS Note].
    ● [Desktop Backup] - Saves case notes for MEC" sq ", MAXIS, the Special Letters, and Email to your desktop.
        In CaseNotes, click [To Desktop]. A text file will be saved using the case number for the file name.
    ● [Clear] - Resets the app. If the case note has not been sent to MEC" sq "/MAXIS or saved to file, it will give a
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
        If ""Use mec2functions"" is checked, [Letter 1] will auto-check and auto-fill fields in MEC" sq "'s Special Letter page.
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
    ● CaseNotes will open in the same location it was closed, even if that monitor is no longer connected.
       See Hotkeys for reset instructions.
    ● All settings for CaseNotes are saved in the My Documents folder, under AHK.ini.
       Deleting this file will reset all saved settings.
    ● Please send any bug reports or feature requests to MECH2.at.github@gmail.com
    )"
    helpXY := "xm y+15"
    Gui, HelpGui: New, ToolWindow, % "CaseNotes Help"
    Gui, HelpGui: Margin, % marginW
    Gui, Font, s10, % "Segoe UI"
    Gui, HelpGui: Add, Tab3,, % "Features | CaseNotes | Missing Verifications | Hotkeys and Notes"
    Gui, Tab, 1
    Gui, HelpGui: Add, Text, % helpXY, % Paragraph1
    Gui, HelpGui: Add, Text, % helpXY, % Paragraph2
    Gui, Tab, 2
    Gui, HelpGui: Add, Text, % helpXY, % Paragraph3
    Gui, Tab, 3
    Gui, HelpGui: Add, Text, % helpXY, % Paragraph4
    Gui, Tab, 4
    Gui, HelpGui: Add, Text, % helpXY, % Paragraph5
    Gui, HelpGui: Add, Text, % helpXY, % Paragraph6
    Gui, HelpGui: Add, Text, % helpXY, % countySpecificText[ini.employeeInfo.employeeCounty].customHotkeys
    Gui, Tab
    Gui, HelpGui: Add, Button, % "gHelpGuiGuiClose w70 h25", % "Close"
    Gui, HelpGui:+OwnerMainGui
    Gui, HelpGui: Show, % "Hide x" ini.caseNotePositions.xCaseNotes " y" ini.caseNotePositions.yCaseNotes
    GuiControlGet editSize, HelpGui:POS, Static5
    guiX := caseNotesMonCenter[1] - ((editSizeW*zoomPPI)/2), guiY := caseNotesMonCenter[2] - (editSizeH*zoomPPI/2)
    GuiControl, HelpGui: Move, % "Close", % "x" (editSizeW/2)
    Gui, HelpGui: Show, % "x" guiX " y" guiY
}
examplesButton() {
    GuiControlGet, examplesButtonText
    If (examplesButtonText == "Examples") {
        For i, exampleLabel in exampleLabels {
            GuiControl, MainGui:Show, % exampleLabel
        }
        GuiControl, MainGui:Text, examplesButtonText, % "Restore"
    } Else If (examplesButtonText == "Restore") {
        For i, exampleLabel in exampleLabels {
            GuiControl, MainGui:Hide, % exampleLabel
        }
        GuiControl, MainGui:Text, examplesButtonText, % "Examples"
    }
}
;EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION EXAMPLES/HELP SECTION  EXAMPLES/HELP SECTION  EXAMPLES/HELP SECTION 
;=======================================================================================================================================================================================================

;=======================================================================================================================================================================
;MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS
getMonCenter(windowName) {
    winHandle := WinExist(windowName)
    VarSetCapacity(monitorInfo, 40), NumPut(40, monitorInfo)
    monitorHandle := DllCall("MonitorFromWindow", "Ptr", winHandle, "UInt", 0x2)
    DllCall("GetMonitorInfo", "Ptr", monitorHandle, "Ptr", &monitorInfo)
    monLeft      := NumGet(monitorInfo, 20, "Int") ; Left
    monTop       := NumGet(monitorInfo, 24, "Int") ; Top
    monRight     := NumGet(monitorInfo, 28, "Int") ; Right
    monBottom    := NumGet(monitorInfo, 32, "Int") ; Bottom
    return [halfMath(monLeft, monRight), halfMath(monTop, monBottom)]
}
halfMath(tl, br) {
    return tl + (br - tl)/2
}
setIcon() {
    If InStr(dateObject.todayYMD, 0401) { ; icon
        Menu, Tray, Icon, compstui.dll, 100
    } Else {
        Menu, Tray, Icon, azroleui.dll, 7
    }
}
setFromIni() {
    FileRead, storedIni, % A_MyDocuments "\AHK.ini"
    storedIni := StrReplace(storedIni, "ini=", "=",, -1)
    section := ""
    Loop, Parse, storedIni, `n, `r
    {
        If InStr(A_LoopField, "]",, StrLen(A_LoopField)-1) { ; ends with "]"
            RegExMatch(A_LoopField, "iO)\[([a-z0-9]+)\]", section)
        } Else {
            keyValue := StrSplit(A_LoopField, "=")
            If (keyValue[2] == "")
                Continue
            ini[Section[1]][keyValue[1]] := keyValue[2]
        }
    }
    coordResetIfOver9000(ini.cbtPositions)
    coordResetIfOver9000(ini.caseNotePositions)
}
guiEditAreaSize(newText, fontOptions := "", fontName := "") {
    Gui, 9:Font, % fontOptions, % fontName
    Gui, 9:Add, Text, Limit200, % newText
    Gui, 9: Show, % "Hide x" ini.caseNotePositions.xCaseNotes " y" ini.caseNotePositions.yCaseNotes
    GuiControlGet T, 9:Pos, Static1
    Gui, 9:Destroy
    Return [TW, TH]
}
checkGroupAdd() {
    If (ini.employeeInfo.employeeBrowser != "") {
        GroupAdd, browserGroup, % ini.employeeInfo.employeeBrowser
    }
    If (ini.employeeInfo.employeeMaxis != "MAXIS-WINDOW-TITLE" && StrLen(ini.employeeInfo.employeeMaxis) > 1) {
        GroupAdd, maxisGroup, % ini.employeeInfo.employeeMaxis
    }
}
removeToolTip() {
    ToolTip
}
timedToolTip(string, duration) {
    ToolTip, % string, 0, 0
    SetTimer, removeToolTip, % "-" duration
}
resetPositions() {
    WinMove, % "CaseNotes",, 0, 0
    WinMove, % "Missing Verifications",,0,0
    xCaseNotes := 0
    yCaseNotes := 0
    xVerification := 0
    yVerification := 0
}
coordSaveAndExitApp(reOpen := 0) {
	WinGetPos, xCaseNotesGet, yCaseNotesGet,,, % "CaseNotes"
	WinGetPos, xVerificationGet, yVerificationGet,,, % "Missing Verifications"
    If (xVerificationGet == "") {
        xVerificationGet := ini.caseNotePositions.xVerification, yVerificationGet := ini.caseNotePositions.yVerification
    }
    If ((xCaseNotesGet - ini.caseNotePositions.xCaseNotes + yCaseNotesGet - ini.caseNotePositions.yCaseNotes + XVerificationGet - ini.caseNotePositions.xVerification + YVerificationGet - ini.caseNotePositions.yVerification) != 0) {
        coordObjOut := {}
        For i, coordName in ["xVerification", "yVerification", "xCaseNotes", "yCaseNotes"] {
            coordObjOut[coordName] := %coordName%Get
        }
        coordString := coordStringify(coordObjOut)
        IniWrite, % coordString, % A_MyDocuments "\AHK.ini", % "caseNotePositions"
    }
    If (reOpen == 1) {
        Reload
    } Else {
        ExitApp
    }
}
coordResetIfOver9000(ByRef coordObj) {
    For i, coordName in coordObj {
        If (Abs(coordName) > 9000)
            coordObj[i] := 50
    }
}
coordStringify(coordObjIn) {
    coordResetIfOver9000(coordObjIn)
    For coordName, coordValue in coordObjIn {
        coordString.= coordName "=" coordValue "`n"
    }
    Return coordString
}
;getRowCount({originalString, maxColumns, indentString:="", indentRow:=4, reduceRowButDoNotIndent:=0}) {
getRowCount(originalString, maxColumns, indentString:="", indentRow:=4) {
    ;indentString := StrLen(indentString) > 0 ? indentString : ""
    textString := stWordWrap(originalString, maxColumns, indentString, indentRow)
    ;textString := stWordWrap({ originalString: originalString, maxColumns: maxColumns, indentString: indentString, indentRow: indentRow })
    StrReplace(textString, "`n", "`n", xCount)
    Return [textString, xCount +1]
}
;function has been rewritten to allow for increased indenting options
;stWordWrap({originalString, maxColumns, indentString:="", indentRow:=4, reduceRowButDoNotIndent:=0}) { ; indentRow: 4 = all lines, 3 = first line of each paragraph, 2 = only very first, 1 =  all but first, 0 = none
stWordWrap(originalString, maxColumns, indentString:="", indentRow:=4, reduceRowButDoNotIndent:=0) { ; indentRow: 4 = all lines, 3 = first line of each paragraph, 2 = only very first, 1 =  all but first, 0 = none
    If (!StrLen(originalString)) {
        Return
    }
    indentLength := StrLen(indentString), indentRow = indentLength > 0 ? indentRow : 0
    completeString := RTrim(originalString, "`n"), firstLine := 1
    If (
        !InStr(completeString, "`n")
        && (reduceRowButDoNotIndent ? maxColumns - indentLength : maxColumns)
        >= ( StrLen(completeString) + (indentRow > 1 ? indentLength : 0) ) ) {
        Return (indentRow > 1 ? indentString : "") completeString
    }
    Loop, Parse, completeString, `n, `r ; A_LoopField == sentence;
    {
        RegExMatch(A_LoopField, "^\s{1,}", manualIndent)
        trimmedSentence := Trim(A_LoopField)
        maxLoopColumns := maxColumns, column := 0
        firstLineOfPara := 1
        Loop, Parse, trimmedSentence, %A_Space% ; A_LoopField == word;
        {
            If (column == 0) {
                If (firstLineOfPara) {
                    out .= manualIndent
                    maxLoopColumns -= StrLen(manualIndent)
                }
                maxLoopColumns -= (firstLine && reduceRowButDoNotIndent) ? indentLength : 0
            }
            wordLength := StrLen(A_LoopField)
            lineLength := column + wordLength
            If (lineLength > maxLoopColumns ) {
                out .= "`n", column := 0, firstLine := 0, firstLineOfPara := 0
            }
            isFirstLineOfPara := (indentRow > 3 && firstLineOfPara)
            test .= A_LoopField ": " isFirstLineOfPara "`n"
            indentThisLine := ((column == 0 && indentRow > 0) && (indentRow == 4 || (!firstLine && indentRow == 1) || (indentRow > 1 && firstLine) || (indentRow > 2 && firstLineOfPara) ))
            out .= ( (indentThisLine ? indentString : "") . A_LoopField " " )
            column += ( (indentThisLine ? indentLength : 0) + wordLength + 1 )
        }
        firstLine := 0
        out := RTrim(RTrim(out), "`n") "`n" ; ensures it ends with a new line
    }
    ;MsgBox % out
    Return RTrim(out, "`n")
}
;MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS - MISC FUNCTIONS
;=======================================================================================================================================================================

;===========================================================================================================================================================================
;BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION - BORROWED FUNCTIONS SECTION 
Class orderedAssociativeArray { ; Capt Odin https://www.autohotkey.com/boards/viewtopic.php?t=37083
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
		return new orderedAssociativeArray.Enum(this.__Data, this.__Order)
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
	Gui, MainGui: Submit, NoHide
    Clipboard := caseNumber
Return

#m::
    If WinActive("Message" ahk_exe Outlook.exe) {
        Clipboard := getFirstName() emailTextObject.output
        Send, ^v
    } Else If WinActive(ini.employeeInfo.employeeBrowser) {
        Gui, MainGui: Submit, NoHide
        Gui, MissingGui: Submit, NoHide
        Sleep 200
        If (ini.employeeInfo.employeeUseMec2Functions == 1) {
            caseStatus := InStr(caseDetails.docType, "?") ? "" : (caseDetails.docType == "Redet") ? "Redetermination" : (Homeless == 1) ? "Homeless App" : caseDetails.docType
            jsonLetterText := "LetterTextFromAHKJSON{""LetterText"":""" JSONstring(letterText[1]) """,""CaseStatus"":""" caseStatus """,""IdList"":""" idList """ }"
            Clipboard := jsonLetterText
            Sleep 200
            Send, ^v
        } Else {
            Clipboard := LetterText[1]
            Sleep 200
            Send, ^v
        }
    }
    Sleep 500
    Clipboard := caseNumber
Return

;Shows Clipboard text in an AHK GUI
!^a::
    If WinExist("Clipboard Text") {
        Gui, CBT: Destroy
    }
    Gui, CBT: New
    Gui, Color, Silver, C0C0C0
    Gui, Font, s11, Lucida Console
    Gui, CBT: Add, Edit, % "ReadOnly -VScroll vClipboardContents", % clipboard
    GuiControl, CBT: font, % "ClipboardContents"
    Gui, CBT: Show, % "x" ini.cbtPositions.xClipboard " y" ini.cbtPositions.yClipboard, Clipboard Text
    ControlSend,,{End}, % "Clipboard Text"
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
        MsgBox,4, % "Reset Position?", % "Do you want to reset CaseNotes' position?", 10
        IfMsgBox, Yes
            resetPositions()
        IfMsgBox, Timeout
            resetPositions()
        Return
    Return
#If

#IfWinActive Other Verification
    Esc:: OtherGuiGuiClose()
#If

#IfWinActive ahk_group browserGroup
    ^F12:: ;CtrlF12/AltF12 Add worker signature
    !F12::
        SendInput % "`n=====`n"
        Send, % ini.employeeInfo.employeeName
    Return
#If
If (ini.employeeInfo.employeeCounty == "Dakota") {
    #IfWinActive ahk_exe WINWORD.EXE ; Word file not in use anymore?
        F1::
            toolTipText := "Alt+4: Starting from the name field, moves to and enters date,         case number, and client's first name."
            timedToolTip(toolTipText, 8000)
        Return
        !4::
            Gui, MainGui: Submit, NoHide
            RegExMatch(01HouseholdCompEdit, "^\w+\b", NameMatch)
            SendInput, {Down 2}
            Sleep 400
            SendInput, % ReceivedDate
            Sleep 400
            SendInput, {Up}
            Sleep 400
            SendInput, % caseNumber
            Sleep 400
            SendInput, {Up}
            Sleep 400
            SendInput, % NameMatch " "
        Return
    #If

    onBaseImportKeys(CaseNum, docType, DetailText, DetailTabs=1, ToolTipHelp="") {
        SendInput, {Tab 2}
        Sleep 250
        SendInput, % docType
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
            SetTimer, removeToolTip, -5000   
        }
    }

    ;Ctrl+: OnBase docs (onBaseImportKeys("Text to get doc type", "Details Text", "Tab presses from Case # to details field")
    #IfWinActive Perform Import
        F1:: 
            toolTipText := "CTRL+ `n F6: RSDI `n F7: SMI ID `n F8: PRISM GCSC `n F9: CS $ Calc `nF10: Income Calc `nF11: The Work # `nF12: CCAPP Letter"
            timedToolTip(toolTipText, 8000)
        Return
        ^F6::
            Gui, MainGui: Submit, NoHide
            onBaseImportKeys(caseNumber, "3003 ssi", "{Text}RSDI ", 3, "Member#, Member Name")
        Return
        ^F7::
            Gui, MainGui: Submit, NoHide
            onBaseImportKeys(caseNumber, "3001 other id", "{Text}SMI ", 3, "Member#, Member Name")
        Return
        ^F8::
            Gui, MainGui: Submit, NoHide
            onBaseImportKeys(caseNumber, "3003 child support", "{Text}GCSC ", 1, "Y/N, Child(ren) Member#")
        Return
        ^F9::
            Gui, MainGui: Submit, NoHide
            onBaseImportKeys(caseNumber, "3003 wo", "{Text}CCAP CS INCOME CALC")
        Return
        ^F10::
            Gui, MainGui: Submit, NoHide
            onBaseImportKeys(caseNumber, "3003 wo", "{Text}CCAP INCOME CALC")
        Return
        ^F11::
            Gui, MainGui: Submit, NoHide
            onBaseImportKeys(caseNumber, "3003 other - in", "{Text}W# ", 3, "Member#, Employer")
        Return
        ^F12::
            Gui, MainGui: Submit, NoHide
            onBaseImportKeys(caseNumber, "3003 edak 3813", "{Text}OUTBOUND")
        Return
    #If

    #IfWinActive ahk_group autoMailGroup ; OnBase, excluding "Perform Import"
        F1::
            toolTipText := WinActive("Automated Mailing Home Page")
            ? "Ctrl+B: Types in the current date and case number." : "
            (
Ctrl+B: (Mail) Types in the current date and case number, clicks Yes. Works best from Custom Query.
          Step 1: Select documents in the query.
          Step 2: Right Click -> Send To -> Envelope.
          Step 3: Click 'Create Envelope'

Alt+4: (Keywords) Enters 'VERIFS DUE BACK' + verif due date. 'Details' keyword field must be active.
            )"
            timedToolTip(toolTipText, 8000)
        Return
        ^b::
            Gui, MainGui: Submit, NoHide
            FormatTime, shortDate, % dateObject.todayYMD, % "M/d/yy"
            SendInput, % shortDate " " caseNumber
            Clipboard := caseNumber
            Sleep 500
            If ( WinActive("ahk_exe obunity.exe") ) {
                Sleep 250
                Send {Tab}
                Send {Enter}
                MsgBox, 4100, % "Case Open Mail", % "Reminder: First add documents to envelope.`n`nOpen / switch to Automated Mailing?"
                    IfMsgBox Yes
                        If (WinExist("Automated Mailing Home Page") ) {
                            WinActivate, % "Automated Mailing Home Page"
                                Return
                        } Else {
                            run % "http://webapp4/AutomatedMailingPRD/#step-1"
                        }
            }
            
        Return
        !4::
            SendInput, % "VERIFS DUE BACK " autoDenyObject.autoDenyExtensionDate
        Return
    #If

    #IfWinActive ahk_group maxisGroup
        ^m::
            WinSetTitle, ahk_exe bzmd.exe,, % "S1 - MAXIS"
            Send ^{m}
        Return
    #If


    #IfWinActive ahk_group browserGroup
        F1::
            timedToolTip("
            (
    Alt+F1: Reviewed/Approved application (Start New case note first)
    Alt+F2: Reviewed/Denied application (Start New case note first)

    Ctrl/Alt+F12: Add worker signature to case note"
            ), 8000)
        Return

        !F1::
            InputBox, approvedDate, % "Enter Approved Date", % "Approved eligible results effective _____."
            InputBox, saApprovalInfo, % "Enter Service Authorization Details", % "Service Authorization _______. `n`nExamples: `n  approved effective 1/1/25 `n  not approved"
            reviewString := "Reviewed case for verifications that are required at application. Verifications were received.`n-`nApproved eligible results effective " approvedDate ".`n-`nService Authorization " saApprovalInfo ".`n=====`n" ini.employeeInfo.employeeName
            noteTitle := "Reviewed application requirements - approved elig"
            If (ini.employeeInfo.employeeUseMec2Functions == 1) {
                jsonCaseNote := "CaseNoteFromAHKJSON{""notedocType"":""Application Approved"",""noteTitle"":""" noteTitle """,""noteText"":""" JSONstring(reviewString) """,""noteElig"":""elig"" }"
                Clipboard := jsonCaseNote
                Send, ^v
            } Else {
                Send {Tab 7}
                Sleep 750
                Send {A 4}
                Sleep 500
                Send {Tab}
                Sleep 500
                SendInput, % noteTitle
                Sleep 500,
                Send, {Tab}
                Sleep 500,
                SendInput, % reviewString
                Sleep 1000,
            }
            Send, !{s}
        Return

        !F2::
            reviewString := "Reviewed case for documents that are required at application. Documents were not received.`n-`nApplication was denied by MEC2 and remains denied.`n=====`n" ini.employeeInfo.employeeName
            noteTitle := "Reviewed application requirements - app denied"
            If (ini.employeeInfo.employeeUseMec2Functions == 1) {
                jsonCaseNote := "CaseNoteFromAHKJSON{""notedocType"":""Application Approved"",""noteTitle"":""" noteTitle """,""noteText"":""" JSONstring(reviewString) """,""noteElig"":""ineligible"" }"
                Clipboard := jsonCaseNote
                Send, ^v
            } Else {
                Send {Tab 7}
                Sleep 750
                Send {A 4}
                Sleep 500
                Send {Tab}
                Sleep 500
                SendInput, % noteTitle
                Sleep 500,
                Send, {Tab}
                Sleep 500,
                SendInput, % reviewString
                Sleep 1000,
            }
            Send, !{s}
        Return
    #If
}