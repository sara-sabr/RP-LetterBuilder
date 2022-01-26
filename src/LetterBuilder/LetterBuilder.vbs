' User modifiable Constants
' =======================================
const FN_CAD = "CAD.xlsb"
const FN_CONFIG = "configuration.properties"
const MOD_CONFIG = "info.properties"
const MOD_SUFFIX_EN = "-EN.docx"
const MOD_SUFFIX_FR = "-FR.docx"
const MOD_ATTACHMENTS_EN = "AttachmentsEN"
const MOD_ATTACHMENTS_FR = "AttachmentsFR"

' Configs
' =======================================
const CONFIG_GD_EPOST = "epost.email"

' System Constants
' =======================================
const EXCEL_FILTER_MATCH = 0
const EXCEL_FILTER_CONTAINS = 1
const EXCEL_FILTER_STARTS = 2
const EXCEL_FILTER_ENDS = 3
const CAD_SHEET_DED = "DED"
const CAD_SHEET_ERN = "ERN"
const CAD_SHEET_PER = "PER"
const CAD_SHEET_CAD = "CAD"
const CAD_SHEET_JOB = "JOB"
const CAD_SHEET_BEN = "BEN"
const LANG_ENGLISH = 0
const LANG_FRENCH = 1

' Translations
' Key is DOM ID
' Value is an array where (English, French)
' ==========================================
Dim I18N
Set I18N = CreateObject("Scripting.Dictionary")
I18N.Add "applicationTitle", 		Array("Letter Builder", "Générateur de lettres")
I18N.Add "pageTitle", 				Array("Letter Builder", "Générateur de lettres")
I18N.Add "page1Subtitle",   		Array("Are the files installed correctly?", "Les fichiers sont-ils installés correctement ?")
I18N.Add "page1Header",     		Array("What letter would you like to work on?", "Sur quelle lettre aimeriez-vous travailler?")
I18N.Add "page1LetterType",     	Array("Letter Type", "Type de lettre")
I18N.Add "page1LanguageSelect",     Array("Language of Letter", "Langue de la lettre")
I18N.Add "page1OptionEnglish",      Array("English", "Anglais")
I18N.Add "page1OptionFrench",       Array("French", "Français")
I18N.Add "page1DefaultLetter",    	Array("Select from List", "Sélectionner dans la liste")
I18N.Add "page1DefaultLanguage",    Array("Select from List", "Sélectionner dans la liste")
I18N.Add "page1EffectiveDate",     	Array("Effective Date of Pay Action (dd/mm/yyyy)", "Date d’effet de l'action de paye (jj/mm/aaaa)")
I18N.Add "page1ButtonStart",     	Array("Lets get started", "Débuter")
I18N.Add "page1CadFileLabel",    	Array("CAD", "TBAC")
I18N.Add "page1CadFileErrorMsg",   	Array("Missing <span id=""fileCadName""></span> in folder.", "<span id=""fileCadName""></span> est manquant dans le dossier.")
I18N.Add "page1EnLtrLabel",    		Array("English Letter Template", "Gabarit Lettre - Anglais")
I18N.Add "page1EnLtrErrorMsg",    	Array("Missing <span id=""fileEnLetterName""></span> in folder.", "<span id=""fileEnLetterName""></span> est manquant dans le dossier.")
I18N.Add "page1FrLtrLabel",    		Array("French Letter Template", "Gabarit Lettre - Français")
I18N.Add "page1FrLtrErrorMsg",    	Array("Missing  <span id=""fileFrLetterName""></span> in folder.", "<span id=""fileFrLetterName""></span> est manquant dans le dossier.")
' I18N.Add "page1EnPSHCPOptLabel",    Array("English PSHCP Options Form", "Formulaire d’options RSSFP - Anglais")
' I18N.Add "page1EnPSHCPOptErrorMsg", Array("Missing  <span id=""fileEnOptionsName""></span> in folder.", "<span id=""fileEnOptionsName""></span> est manquant dans le dossier.")
' I18N.Add "page1FrPSHCPOptLabel",    Array("French PSHCP Options Form", "Formulaire d’options RSSFP - Français")
' I18N.Add "page1FrPSHCPOptErrorMsg", Array("Missing  <span id=""fileFrOptionsName""></span> in folder.", "<span id=""fileFrOptionsName""></span> est manquant dans le dossier.")
I18N.Add "page1RetryButton",		Array("Retry", "Réessayer")
I18N.Add "page2Subtitle",			Array("Enter PRI in CAD", "Inscrire le CIDP dans le TBAC")
I18N.Add "page2ContentP1",			Array("The CAD Excel spreadsheet is now open.", "Le fichier excel TBAC est maintenant ouvert.")
I18N.Add "page2ContentL1",			Array("Acknowledge the CAD disclaimer (<b>IF SHOWN</b>)", "Cliquez sur Accepter et continuer, <b>si la page de Mise en garde – TBAC apparaît.</b>")
I18N.Add "page2ContentL2",			Array("Navigate to the “CAD” tab of the file and enter in a single PRI into Cell B2", "Dans le fichier excel TBAC, inscrivez un CIDP à la cellule B2")
I18N.Add "page2ContentL3",			Array("Ensure the following codes on the CAD are checked DED, PER, JOB, ERN, and BEN", "Assurez-vous que les onglets suivants sont cochés : DED, PER, JOB, ERN, BEN.")
I18N.Add "page2ContentL4",			Array("Press “Auto Load” on CAD to populate employee data", "Cliquez sur Chargement automatique pour télécharger les données de l’employé")
I18N.Add "page2ContentL5",			Array("Once complete, click next on this screen to continue", "Une fois complété, cliquez sur Suivant au bas de cet écran pour continuer")
I18N.Add "page2ButtonStart",     	Array("Lets get started", "Débuter")
I18N.Add "page3ButtonNext",     	Array("Next", "Suivant")
I18N.Add "page3Subtitle",			Array("Verify Information", "Vérifier les informations")
I18N.Add "PayListLabel",			Array("Pay List", "Liste de paye")
I18N.Add "PRILabel",				Array("PRI", "CIDP")
I18N.Add "EENameLabel",				Array("Employee Name", "Nom de l'employé(e)")
'I18N.Add "EEStreetLabel",			Array("Employee Address", "Adresse de l'employé(e)")
'I18N.Add "EECityLabel",				Array("Employee City", "Ville de l'employé(e)")
'I18N.Add "EEProvinceLabel",			Array("Employee Province", "Province de l'employé(e)")
'I18N.Add "EEPostalCodeLabel",		Array("Employee Postal Code", "Code postale de l'employé(e)")
I18N.Add "ReasonLabel",				Array("Reason", "Raison")
I18N.Add "inputEffectiveDate",		Array("Effective Date of Pay Action (dd/mm/yyyy)", "Date d’effet de l'action de paye (jj/mm/aaaa)")
I18N.Add "PersonalEmailLabel",		Array("Personal Email", "Courriel personnel")
I18N.Add "CaseNumberLabel",			Array("Case Number", "Numéro de cas")
I18N.Add "Page3InstructionsLabel",	Array("Letter Customizations", "Personnalisation de la lettre")
I18N.Add "Page3InstructionsText",	Array("Content with checkmarks are loaded by default based on the information collected from the CAD. Please review and select or unselect as needed. Selecting field will add to the generated letter, unselected fields will not be displayed on the letter.", "Le contenu coché est chargé par défaut selon les informations recueillies dans le CAD. Veuillez les vérifier et les sélectionner ou les désélectionner si nécessaire. Les champs sélectionnés seront ajoutés à la lettre générée. Les champs non sélectionnés ne seront pas affichés sur la lettre.")
I18N.Add "PensionandDSBLabel",		Array("Pension", "Pension")
I18N.Add "DisabilityInsuranceLabel",Array("Disability Insurance", "Assurance invalidité")
I18N.Add "PSMIPLabel",				Array("PSMIP", "RACGFP")
I18N.Add "IANLabel",				Array("IAN", "NIO")
I18N.Add "PSHCPLabel",				Array("PSHCP", "RSSFP")
I18N.Add "PSHCPNoLabel",			Array("PSHCP Number", "Numéro RSSFP")
I18N.Add "PSHCPLevelLabel",			Array("PSHCP Level", "Niveau RSSFP")
I18N.Add "DCPLabel",			    Array("DCP Status", "Statut RSD")
I18N.Add "DCPPlanNoLabel",			Array("DCP Plan Number", "Numéro de régime RSD")
I18N.Add "DCP-Option-Default",		Array("Select", "Sélectionnez")
I18N.Add "CertificateNumberLabel",	Array("Certificate Number", "Numéro de certificat")
I18N.Add "UnionInsuranceLabel",		Array("Union Insurance", "Assurance syndicale")
I18N.Add "BilingualBonusLabel",		Array("Bilingual Bonus","Prime au bilinguisme")
I18N.Add "AnnualandSickLeaveLabel",	Array("Annual and Sick Leave", "Congés annuels et de maladie")
I18N.Add "AWWLabel",				Array("AWW", "SDT (Semaine désignée de travail)")
I18N.Add "CompensatoryLeaveLabel",	Array("Compensatory Leave", "Congés compensatoires")
I18N.Add "PayRevisionLabel",		Array("Pay Revision", "Révisions salariales")
I18N.Add "ContinuousServiceLabel",	Array("Continuous Service", "Service continu")
I18N.Add "TermEmploymentLabel",		Array("Term Employment", "Emploi déterminé")
I18N.Add "UnionDuesLabel",			Array("Union Dues", "Cotisations syndicales")
I18N.Add "RALabel",					Array("RA", "Association récréative")
I18N.Add "CreditUnionLabel",		Array("Credit Union", "Coopérative de crédit")
I18N.Add "GCWCCLabel",				Array("GOC Workplace Charitable Campaign", "Campagne de charité en milieu de travail (CCMTGC)")
I18N.Add "GarnishmentsLabel",		Array("Garnishments", "Saisie-arrêt")
I18N.Add "StudentLoansLabel",		Array("Student Loans", "Prêt étudiant")
I18N.Add "LIALabel",				Array("LIA", "CER")
I18N.Add "SelffundedLabel",			Array("Self funded", "Retenues pour congé autofinancé")
I18N.Add "ParentalLabel",			Array("Parental Leave", "Fin du congé de maternité/parental")
I18N.Add "GradualLabel",			Array("Gradual Return to Work", "Retour progressif au travail")
I18N.Add "ParkingLabel",			Array("Parking", "Retenues pour les frais de stationnement")
I18N.Add "InvCheck-CasualorStudentLabel",	Array("Casual/Student", "Occasionnel/étudiant")
I18N.Add "BackToBackLabel",	        Array("Back to Back LWOP", "CNP dos à dos")
I18N.Add "PartTimeEmployeeLabel",	Array("Part-Time", "À temps partiel")
I18N.Add "PSHCPLevel1",				Array("Level 1", "Niveau 1")
I18N.Add "PSHCPLevel2",				Array("Level 2", "Niveau 2")
I18N.Add "PSHCPLevel3", 			Array("Level 3", "Niveau 3")
'	I18N.Add "page3InvalidDate",		Array("Missing or Invalid Date", "Date invalide ou manquant")
'	I18N.Add "page3InvalidDCPPlan",		Array("Missing or Invalid DCP Plan Number", "Numéro de régime RSD invalide ou manquant")
I18N.Add "page3-reloadbutton-text",	Array("Reset Information", "Informations sur le repos")
I18N.Add "page3-reloadcad-text",	Array("Reset Information", "Informations sur le repos")
I18N.Add "page3-reloadbutton",		Array("Reset Employee Information", "Réinitialiser l'employé")
I18N.Add "page3-reloadcad",			Array("Reset Letter Customizations", "Lettre de réinitialisation")
I18N.Add "page3ButtonGenerate", 	Array("Generate Letter", "Générer lettre")
I18N.Add "page3Instruction",	 	Array("Please verify employee information:", "Svp vérifier les informations de l’employé:")
I18N.Add "page4Subtitle",			Array("Review and Export", "Réviser et exporter")
I18N.Add "page4ButtonNext",			Array("Export to PDF", "Exporter en PDF")
I18N.Add "page4Instructions",	 	Array("The Letter generation is now complete. Please review content, before exporting to PDF", "Le processus est completé. Veuillez réviser le document avant d’exporter en PDF")
I18N.Add "page5Instructions",	 	Array("Successfully exported letter to PDF.", "Lettre exportée en PDF avec succès.")
I18N.Add "page5Subtitle",			Array("Ready to be sent to EPOST", "Prêt à être envoyé à Postel")
I18N.Add "page5Subnote1",			Array("An email template had been created and ready to be sent to EPOST GD Box.", "Un gabarit de courriel a été créé et est prêt à être envoyé par Postel.")
I18N.Add "page5Subnote2",			Array("Review the content of the email and send when ready.", "Révisez le contenu du courriel et procédez à l’expédition.")
I18N.Add "page5Subnote3",			Array("Press “Generate next letter” to proceed with processing a new letter.", "Cliquez sur ‘’Générer une autre lettre’’ pour débuter la création d’une nouvelle lettre. ")
I18N.Add "page5ButtonNext",			Array("Generate next letter", "Générer une autre lettre")

' Reference to Workbook
Dim excelWorkbook
' Reference to Excel Application
Dim excelApplication
Set excelApplication = Nothing
' Reference to Word Application
Dim wordApplication
' Reference to Word document generated.
Dim wordDocument
' Reference to PDF absolute file path generated.
Dim pdfFilePath
' File name of PDF
Dim pdfFileName
'Employee record
Dim employeeRecord
' Document Language
Dim documentLanguage

' Configuration settings
Dim configurationSettings
Set configurationSettings = CreateObject("Scripting.Dictionary")
' List of installed modules
Dim modulesAvailable
Set modulesAvailable = CreateObject("Scripting.Dictionary")

' Shell object
Set objShell = CreateObject("WScript.Shell")
currentDirectory = objShell.CurrentDirectory
' List of Attachments
Dim attachmentsEN, attachmentsFR
Set attachmentsEN = CreateObject("System.Collections.ArrayList")
Set attachmentsFR = CreateObject("System.Collections.ArrayList")

' Properties
' =======================================
Dim cadFile, letterEnFile, letterFrFile, workingLetterFile, configFile
cadFile = currentDirectory & "\" & FN_CAD
configFile = currentDirectory & "\" & FN_CONFIG

' Debug Flags
' =======================================
Dim debugExcelRowFilter
debugExcelRowFilter = False

' Classes
' =======================================
' Search Expression
Class SearchExpression
    ' Column index to apply the filter (starts at 1)
    Public ColumnIndex
    ' Simple expressions, generally =, <>, <, >. Checks FilterExpression before
    ' BetweenExpression. So, if you defined both, the FilterExpression will
    ' take precedence.
    Public FilterExpression
    ' Complex expressions where value at index is between A and B
    Public BetweenExpressionA
    ' Complex expressions where value at index is between A and B
    Public BetweenExpressionB
End Class
' Functions
' =======================================
' Open the CAD excel file.
Sub OpenCAD()
    ' Open CAD if not open.
    If excelApplication Is Nothing Then
        Set excelApplication = CreateObject("Excel.Application")
        With excelApplication
            .Left = (screen.Width/2) * 0.75
            .Top = 0
            .Width = (screen.Width/2) * 0.75
            .Height = (screen.Height) * 0.75 - 30
            .Visible = True
            .DisplayAlerts = False
        End With
        set excelWorkbook = excelApplication.Workbooks.Open(cadFile)
        excelWorkbook.Worksheets("CAD").Activate
    End If
End Sub
' Reset all objects at the start
Sub ResetData()
    If Not(IsEmpty(cadDict)) Then
        cadDict.RemoveAll
    End If
    If Not(IsEmpty(employeeRecord)) Then
        employeeRecord.RemoveAll
    End If
    configurationSettings.RemoveAll
    attachmentsEN.Clear
    attachmentsFR.Clear
    
    document.getElementById("selectLetter").value = "--"
    document.getElementById("selectLanguage").value = "--"
    document.getElementById("inputEffectiveDate").value = ""
    document.getElementById("cad-form").reset()
    document.getElementById("pdfFileName").innerHTML = "Placeholder"

    placeholderPolyfill("inputEffectiveDate")
End Sub

' Close CAD
Sub CloseCAD()
    excelWorkbook.Close(0)
    Set excelWorkbook = Nothing
    excelApplication.Quit
    Set excelApplication = Nothing
End Sub
'Check for available modules and add them to a list
Sub CheckAvailableModules()
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim altSuffix, altModule, fileSuffix, languageCode, moduleName, propertyFile, selectElement, optionElement
    languageCode = document.getElementById("html").lang
    Set selectElement = document.getElementById("selectLetter")

    Dim langArr(2)
    If languageCode = "fr" Then
        fileSuffix = MOD_SUFFIX_FR
        altSuffix = MOD_SUFFIX_EN
    Else
        fileSuffix = MOD_SUFFIX_EN
        altSuffix = MOD_SUFFIX_FR
    End If

    If selectElement.length = 1 Then
        For Each objFolder in fso.GetFolder(".\Templates\").SubFolders
            langArr(0) = ""
            langArr(1) = ""

            For Each file in objFolder.Files
                ' Grab all files that end with -EN.docx and -FR.docx in all the template folders
                
                If instr(file.name, fileSuffix) <> 0 Then
                    moduleName = Replace(file.name, fileSuffix, "")
                    langArr(0) = moduleName
                ElseIf instr(file.name, altSuffix) <> 0 Then
                    altModule = Replace(file.name, altSuffix, "")
                    langArr(1) = altModule
                ElseIf instr(file.name, MOD_CONFIG) <> 0 Then
                    propertyFile = file.Path
                End If

                ' Dynamically add module label to I18N
                If langArr(0) <> "" AND langArr(1) <> "" Then
                    I18N.Add "letter-" & langArr(0) , langArr
                End If
            Next
            If Not(IsNull(moduleName)) And Not(IsNull(propertyFile)) Then
                modulesAvailable.Add moduleName, propertyFile
                Set optionElement = document.createElement("option")
                optionElement.id=   "letter-" & moduleName
                optionElement.value = moduleName
                optionElement.text = moduleName
                selectElement.Add optionElement
            End If
        Next
    End If
    Set fso = Nothing
End Sub
' Verify installation is good.
Sub CheckInstall()
    LoadSelectedModule(document.getElementById("selectLetter").value)
    Call UpdateValidationStatusEntry("page1CadFileLabel", cadFile)
    Call UpdateValidationStatusEntry("page1EnLtrLabel", letterEnFile)
    Call UpdateValidationStatusEntry("page1FrLtrLabel", letterFrFile)
    If FileExists(cadFile) AND _
        FileExists(letterEnFile) AND _
        FileExists(letterFrFile) Then
        document.getElementById("page2ButtonStart").disabled = false
        document.getElementById("recheck-install").className = "page1-recheckButtonbar hidden"
    Else
        document.getElementById("page2ButtonStart").disabled = true
        document.getElementById("recheck-install").className = "page1-recheckButtonbar"
    End If
End Sub

' Translate function. The language code must be en or fr
Sub TranslateScreen(languageCode, e)
    ' Flag the page lang code.
    document.getElementById("html").setAttribute "lang", languageCode

    arrayIdx = LANG_ENGLISH

    if languageCode = "fr" Then
        arrayIdx = LANG_FRENCH
    End If

    'Flip all IDs
    For Each messageKey in I18N.keys
        localizedArray = I18N(messageKey)
        If document.getElementById(messageKey).getAttribute("idTextAsTitle") Then
            document.getElementById(messageKey).title = localizedArray(arrayIdx)
        Else		
            document.getElementById(messageKey).innerHTML = localizedArray(arrayIdx)		
        End If
    Next
    If languageCode = "fr" Then
        document.getElementById("languageEnglish").focus()
        document.getElementById("fip").src = "assets/img/fip-fr.png"
    Else
        document.getElementById("languageFrench").focus()
        document.getElementById("fip").src = "assets/img/fip-en.png"
    End If
    document.getElementById("fileCadName").innerText = FN_CAD
    document.getElementById("fileEnLetterName").innerText = FN_ENGLISH_LETTER
    document.getElementById("fileFrLetterName").innerText = FN_FRENCH_LETTER
    If Not e is Nothing Then
        e.preventDefault()
    End If
End Sub
' Populate Fields based on CAD
' FLAGS [
' 0: Populate only checkbox inputs
' 1: Populate entire Form
']
Sub PopulateScreenCADFields(flag)
    ' Show loading spinner
    ShowProgressOverlay(true)

     Select Case flag
        Case 0
            window.setTimeOut "RunPopulateScreenCADFields(0)", 2000, "VBScript"
        Case 1
            window.setTimeOut "RunPopulateScreenCADFields(1)", 2000, "VBScript"
        Case Else
    End Select
    
    
End Sub

Sub RunPopulateScreenCADFields(flag)
    Set employeeRecord = GetCADData(0)
    Set employeeTextInputs = document.getElementById("cad-text-inputs")
    Set employeeCheckedInputs = document.getElementById("cad-checked-inputs")

    If flag = 0 Then
        UpdateFieldsInElement employeeTextInputs, employeeRecord
    ElseIf flag = 1 Then
        UpdateFieldsInElement employeeCheckedInputs, employeeRecord
    End If

    If document.getElementById("selectLanguage").value = "French" Then
        document.getElementById("Reason").value = configurationSettings("reason.fr")
    Else
        document.getElementById("Reason").value = configurationSettings("reason.en")
    End If

    employeeTextInputs.style.display = "inline"
    employeeCheckedInputs.style.display = "inline"

    ' Hide loading spinner
    ShowProgressOverlay(false)
End Sub

Sub UpdateFieldsInElement(element, employeeRecord)
    For Each inputField In element.getElementsByClassName("cadField")
        fieldName = inputField.name
        inputField.value = ""
        If employeeRecord.Exists(fieldName) Then
            tagName = inputField.tagName
            Select Case  tagName
                Case "INPUT"
                    If inputField.type = "text" Then
                        inputField.value = employeeRecord(fieldName)
                    ElseIf inputField.type = "checkbox" Then
                        If employeeRecord(fieldName) <> "" Then
                        inputField.checked = true
                        Else
                        inputField.checked = false
                        End If
                        If inputField.hasAttribute("childField") Then
                            ShowCadField inputField.getAttribute("childField"), inputField.checked
                        End If
                    End If
                Case "SELECT"
                    inputField.value = employeeRecord(fieldName)
            End Select
        End If

        If(inputField.hasAttribute("placeholder")) Then
            If(Len(inputField.value) > 0) Then
                document.getElementById(inputField.id + "-placeholder").textContent = ""
            Else
                placeholderText = inputField.getAttribute("placeholder")
                document.getElementById(inputField.id + "-placeholder").textContent = placeholderText
            End If
        End If
    Next

    validateForm()
End Sub

Sub ShowCadField(element, visible)
    Set field = document.getElementById(element)
    If visible Then
        field.style.display = "block"
    Else	
        field.style.display = "none"
    End If
End Sub

' Set document language
Sub SetDocumentLanguage()
    Dim docLang
    docLang = document.getElementById("selectLanguage").value

    If docLang = "French" OR docLang = "Français" Then
        documentLanguage = Array("Français", "French")
    Else 
        documentLanguage = Array("Anglais", "English")
    End If

End Sub
' Generate English Letter
Sub GenerateEnglishLetter()
    GenerateLetter(letterEnFile)
End Sub
' Generate French Letter
Sub GenerateFrenchLetter()
    GenerateLetter(letterFrFile)
End Sub

' Adds in the slashes while typing in the effective Date
Sub EffectiveDateFormatting(newEvent)
    Set currentDate = document.getElementById(newEvent.target.name)
    key = Mid(newEvent.target.value, Len(currentDate.value), Len(currentDate.value)) 
    
    If(IsNumeric(key) AND Len(currentDate.value) <= 10) Then
        If Len(currentDate.value) = 2 OR Len(currentDate.value) = 5 Then
        tempDate = currentDate.value & "/"
        currentDate.value = tempDate
        End If
    Else
        currentDate.value = Mid(currentDate.value, 1, Len(currentDate.value) -1) 
    End If

    ValidatePage1()
End Sub
' Generate Letter Helper
Sub GenerateLetter(letterFile)
    ToggleWordGenerateButtons(false)
    Set employeeRecord = PopulateFieldsManualInput()

    CreateTempLetter(letterFile)

    ' Open word
    Set wordApplication = CreateObject("Word.Application")
    With wordApplication
        .Left = (screen.Width/2) * 0.75
        .Top = 0
        .Width = (screen.Width/2) * 0.75
        .Height = (screen.Height) * 0.75 - 30
        .Visible = True
        .DisplayAlerts = False
    End With
    Set wordDocument = wordApplication.Documents.Open(workingLetterFile)
    employeeRecord.Add "Date" ,   Day(Date) & "/" & Month(Date) & "/" & Year(Date)
    For Each key in employeeRecord.keys
        SearchAndRep key, employeeRecord(key)
        SearchAndRemoveIfs key, employeeRecord(key)
    Next
    ToggleWordGenerateButtons(true)
End Sub
' Create the PDF file.
Sub CreatePDF()
    Set employeeRecord = PopulateFieldsManualInput()
    pdfFileName = GetFileNameToUse() & ".pdf"

    pdfFilePath = currentDirectory & "\GeneratedLetters\" & pdfFileName
    wordDocument.SaveAs pdfFilePath, 17 'Index for PDF
    wordDocument.Close(0) '0 is don't save changes https://docs.microsoft.com/en-us/office/vba/api/word.wdsaveoptions
    ' Clean up
    Set wordDocument = Nothing
    window.setTimeOut "CleanUpWord()", 2000
End Sub
' Open PDF File.
Sub OpenPDFFile()
    objShell.Run("""" & pdfFilePath & """")
End Sub
' Clean up Word
Sub CleanUpWord()
    ' Clean up
    wordApplication.Quit
    Set wordApplication = Nothing
    DeleteTempLetter
End Sub
' Find and Replace all key value pairs
Sub SearchAndRep(searchTerm, replaceTerm)
    Set myRange = wordApplication.ActiveDocument.Content
    With myRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "<<" & searchTerm & ">>"
        .Replacement.Text = replaceTerm
        .Execute ,,,,,,,,,,2 ' ReplaceAllFlag
    End With
    Set myRange = Nothing
End Sub
' Bulk remove entire section blocks of <<IF VAR>> .... <<END IF VAR>>
Sub SearchAndRemoveIfs(searchTerm, termData)
    Set myRange = wordApplication.ActiveDocument.Content
    If IsEmpty(termData) Or _
        termData = false Or _
        termData = "" Or _
        IsNull(termData) Then
        ' Remove Entire section
        With myRange.Find
            .ClearFormatting
            .MatchWildcards = True
            .IgnoreSpace = True
            .Replacement.ClearFormatting
            .Text = "?\<\< IF " & searchTerm & " \>\>*\<\< END IF " & searchTerm & " \>\>"
            .Replacement.Text = ""
            .Execute ,,,,,,,,,,2 ' ReplaceAllFlag
        End With
    Else
        ' Remove start <<IF term>>
        With myRange.Find
            .ClearFormatting
            .MatchWildcards = True
            .IgnoreSpace = True
            .Replacement.ClearFormatting
            .Text = "?\<\< IF " & searchTerm & " \>\>"
            .Replacement.Text = "^11"
            .Execute ,,,,,,,,,,2 ' ReplaceAllFlag
        End With
        ' Remove end start <<END IF term>>
        With myRange.Find
            .ClearFormatting
            .IgnoreSpace = True
            .MatchWildcards = True
            .Replacement.ClearFormatting
            .Text = "?\<\< END IF " & searchTerm & " \>\>"
            .Replacement.Text = ""
            .Execute ,,,,,,,,,,2 ' ReplaceAllFlag
        End With
    End If
    Set myRange = Nothing
End Sub

' Validate attachments
Function ValidateAttachments()
    If CInt(configurationSettings("attachments.en")) = attachmentsEN.Count AND _
        CInt(configurationSettings("attachments.fr")) = attachmentsFR.Count Then
        ValidateAttachments = True
    Else
        ValidateAttachments = False
    End If
End Function
Function GetEffectiveDate()
    GetEffectiveDate = document.getElementById("inputEffectiveDate").value
End Function
Function GetEffectiveDateUserInput()
    GetEffectiveDateUserInput = document.getElementById("EffectiveDate").value
End Function
' Load module information
Function LoadSelectedModule(moduleName)
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' Reset the configuration settings on Load
    Set configurationSettings = Nothing
    Set configurationSettings = CreateObject("Scripting.Dictionary")
    Dim moduleLocation, moduleProperties

    moduleProperties = modulesAvailable(moduleName)

   

    LoadConfig(configFile)
    LoadConfig(moduleProperties)
    moduleLocation = Replace(moduleProperties, "\" & MOD_CONFIG, "")
    Set objFolder = fso.GetFolder(moduleLocation)
    For Each file in objFolder.Files
        If instr(file.name, MOD_SUFFIX_FR) Then
            letterFrFile = file.Path
        ElseIf instr(file.name, MOD_SUFFIX_EN) Then
            letterEnFile = file.Path
        End If
    Next
    For Each subfolder in objFolder.SubFolders 
        For Each file in subfolder.Files
            If instr(subfolder, MOD_ATTACHMENTS_EN) Then
                attachmentsEN.Add file.Path
            ElseIf instr(subfolder, MOD_ATTACHMENTS_FR) Then
                attachmentsFR.Add file.Path
            End If
        Next
    Next

    Set fso = Nothing
End Function
' Read configuration files
Function LoadConfig(filePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set fileHandler = fso.OpenTextFile(filePath)
    Dim line, keyValueDelimIdx, key, value
    Do Until fileHandler.AtEndOfStream
        line = fileHandler.ReadLine
        line = Trim(line)
        ' See if comment line
        keyValueDelimIdx = InStr(line, "#")
        if keyValueDelimIdx <> 1 Then
            ' Not comment line, so lets read the key value pair delimited by "="
            keyValueDelimIdx = InStr(line, "=")
            If keyValueDelimIdx <> 0 AND keyValueDelimIdx <> Len(line) Then
                key = Mid(line, 1, keyValueDelimIdx - 1)
                value = Mid(line, keyValueDelimIdx + 1)
                configurationSettings.Add key, value
            End If
        End If
    Loop
    fileHandler.Close
    Set fso = Nothing
    Set fileHandler = Nothing
End Function
' Get the standard file name. You must call this only after creating employeeRecord
Function GetFileNameToUse()
    Dim dayPortion, monthPortion, employeeName
    'dayPortion = Day(GetEffectiveDate())
    'monthPortion = Month(GetEffectiveDate())
    effectiveDate = GetEffectiveDateUserInput()
    employeeRecord("EffectiveDate") = GetEffectiveDateUserInput()
    effectiveDate = Replace(effectiveDate, "/", "")

    employeeName = Split(employeeRecord("EmployeeName"))

    
    If documentLanguage(1) = "French" Then
        GetFileNameToUse = effectiveDate & configurationSettings("prefix.fr") & employeeName(0) & "_" & Mid(employeeName(1),1,1)
    Else
        GetFileNameToUse = effectiveDate & configurationSettings("prefix.en") & employeeName(0) & "_" & Mid(employeeName(1),1,1)
    End IF
End Function
' Get the case number. You must call this only after creating employeeRecord
Function GetCaseNo()
    If employeeRecord("CaseNumber") <> "" Then
        GetCaseNo = employeeRecord("CaseNumber")
    Else
        GetCaseNo = "CaseNo"
    End If
End Function
' Populate Fields based on Manual Input
Function PopulateFieldsManualInput()
    Set listOfInputs = document.getElementsByTagName("input")
    Set listOfSelects = document.getElementsByTagName("select")

    For Each inputField in listOfInputs
        If InStr(inputField.className, "cadField") Then
            fieldName = inputField.name
            If inputField.type = "checkbox" Then
                If (inputField.getAttribute("Data-Invert-Check") AND inputField.checked=false) OR inputField.checked = true Then
                    employeeRecord(fieldName) = "1"
                Else
                    employeeRecord(fieldName) = ""
                End If
            Else
                employeeRecord(fieldName) = inputField.value
            End If
        End If
    Next
    For Each selectField in listOfSelects
        If InStr(selectField.className, "cadField") Then
            fieldName = selectField.name
            employeeRecord(fieldName) = selectField.value
        End If
    Next
    ' check if the PSHCP level is 1 or 2/3 and set the status in the employee record
    If employeeRecord("PSHCP") = "1" Then
        If instr(employeeRecord("PSHCPLevel"), "1") Then
            employeeRecord("PSHCPLevel1") = "1"
            employeeRecord("PSHCPLevel2Or3") = ""
        Else
            employeeRecord("PSHCPLevel1") = ""
            employeeRecord("PSHCPLevel2Or3") = "1"
        End If
    End If
    ' check if the province is set in manual entry and set the status in the employeeRecord
    If employeeRecord("EmployeeProvince") = "ON" OR employeeRecord("EmployeeProvince") = "QC" Then
        employeeRecord("EmployeeProvinceONOrQC") = "1"
    Else
        employeeRecord("EmployeeProvinceONOrQC") = ""
    End If
    ' Check if the Date is valid
    Dim tempResult
    tempResult = validateAndFormatDate(employeeRecord("EffectiveDate"))
    If tempResult <> "" Then
        employeeRecord("EffectiveDate") = tempResult
    Else
        employeeRecord("EffectiveDate") = GetEffectiveDate()
    End If
    Set PopulateFieldsManualInput = employeeRecord
End Function
' Validate and return dates
Function validateAndFormatDate(dateToBeValidated)
    Dim resultDate, convertedDate
    Dim delimiter, datePartDay, datePartMonth, datePartYear
    if InStr(dateToBeValidated, "/")  Then
        delimiter = "/"
    Else
        delimiter = "-"
    End If
    resultDate = Split(dateToBeValidated, delimiter)
    if (UBound(resultDate) + 1) = 3 Then
        If delimiter = "/" Then
            datePartDay = resultDate(0)
            datePartMonth = resultDate(1)
            datePartYear = resultDate(2)
        Else
            datePartDay = resultDate(2)
            datePartMonth = resultDate(1)
            datePartYear = resultDate(0)
        End If
        validateAndFormatDate = datePartDay & "/" & datePartMonth & "/" & datePartYear
        If Not(IsDate(validateAndFormatDate)) OR datePartMonth > 12 Then
            validateAndFormatDate = ""
        Else
            validateAndFormatDate = datePartDay & "/" & datePartMonth & "/" & datePartYear
        End If
    Else
        validateAndFormatDate = ""
    End If
End Function
' Get the data from CAD
' index is the row number of the PRI we wish to pull out, however doesn't get used anywhere yet.
Function GetCADData(index)
    Set employeeRecord = Nothing
    Dim cadDict
    Set cadDict = CreateObject("Scripting.Dictionary")
    ' Sheets
    Dim sheetCad, sheetJOB
    Set sheetCad = excelWorkbook.Worksheets(CAD_SHEET_CAD)
    Set sheetJOB = excelWorkbook.Worksheets(CAD_SHEET_JOB)
    ' Data from CAD tab
    Dim employeeId, tempResult
    employeeId = sheetCad.Range("B2").Value
    cadDict.Add "PRI", employeeId
    ' Data For Employee
    GetCADDataEmployee index, cadDict, employeeId
    ' Effective Date or default to Nov 15
    tempResult = SimpleFindRowUsingFilter(sheetJOB, "A2", employeeId, 2, 9, "JOB", 10, true)
    tempResult = validateAndFormatDate(tempResult)
    If tempResult <> "" Then
        cadDict.Add "EffectiveDate", tempResult
    Else
        cadDict.Add "EffectiveDate", GetEffectiveDate()
    End If
    ' Get Data Helpers
    GetCADPensionandDisability index, cadDict, employeeId
    GetCADDataServiceBuyback index, cadDict, employeeId
    GetCADDataPSMIP index, cadDict, employeeId
    GetCADDataPSHCP index, cadDict, employeeId
    GetCADDataDCP index, cadDict, employeeId
    GetCADDataUnionInsurance index, cadDict, employeeId
    GetCADDataBilingualBonus index, cadDict, employeeId
    GetCADDataOtherBenefits index, cadDict, employeeId
    GetCADDataLIA index, cadDict, employeeId
    GetCadDataParental index, cadDict, employeeId
    GetCadDataGroup index, cadDict, employeeId
    GetCadFullPartTime index, cadDict, employeeId
    Set sheetCad = Nothing
    Set sheetJOB = Nothing
    Set GetCADData = cadDict
End Function
Function GetCADPensionandDisability(index, cadDict, employeeId)
    Dim sheetDED
    Set sheetDED = excelWorkbook.Worksheets(CAD_SHEET_DED)
    ' Pension and Supplementary Death Benefit: DED Column W SDB001
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 23, "SDB001", 23, true)
    If tempResult <> "" Then
        cadDict.Add "PensionandDSB", "1"
    Else
        cadDict.Add "PensionandDSB", ""
    End If
    'Reset tempResult
    tempResult = ""
    ' Disability Insurance: DED Column R – codes 751, 753, 809, 810
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, Array("751", "753", "809", "810"), 18, true)
    If tempResult <> "" Then
        cadDict.Add "DisabilityInsurance", "1"
    Else
        cadDict.Add "DisabilityInsurance", ""
    End If
End Function
' Get LIA from CAD
Function GetCADDataLIA(index, cadDict, employeeId)
    Dim sheetERN
    Set sheetERN = excelWorkbook.Worksheets(CAD_SHEET_ERN)
    ' LIA: ERN Column L Code UEL on checks between November 15 2020 and November 15 2021
    Dim dateExpression, typeExpression
    ' Check if code UEL
    Set typeExpression = New SearchExpression
    typeExpression.ColumnIndex = 18
    typeExpression.FilterExpression = "UEL"
    ' Check if within 1 year
    Set dateExpression = New SearchExpression
    dateExpression.ColumnIndex = 14
    Dim today, oneYearAgo
    'Don't use date format strings as Excel seems to convert dd/mm/yyyy to mm/dd/yyyy
    today = date()
    oneYearAgo = DateAdd("yyyy",-1,date())
    dateExpression.BetweenExpressionA = ">=" & CDbl(oneYearAgo)
    dateExpression.BetweenExpressionB = "<=" & CDbl(today)
    Dim expressionArray
    expressionArray = Array(typeExpression, dateExpression)
    tempResult = FindRowUsingFilter(sheetERN, "A2", employeeId, 2, expressionArray, 12, true)
    If tempResult <> "" Then
        Dim LIATax
        LIATax = CDbl(tempResult)
        If LIATax > 0 Then
            cadDict.Add "LIA", "1"
        Else
            cadDict.Add "LIA", ""
        End If
    End If
    Set sheetERN = Nothing
End Function
' Get Other Benefits from CAD
Function GetCADDataOtherBenefits(index, cadDict, employeeId)
    ' Sheets
    Dim sheetJOB, sheetDED
    Set sheetJOB = excelWorkbook.Worksheets(CAD_SHEET_JOB)
    Set sheetDED = excelWorkbook.Worksheets(CAD_SHEET_DED)
    ' Get AWW which is number x 2
    tempResult = FindRowUsingNoAdditionalFilter(sheetJOB, "A2", employeeId, 2, 16, true)
    If IsNumeric(tempResult) Then
        tempResult = FormatNumber(Round(tempResult * 2.0, 2), 2)
        cadDict.Add "AWW", tempResult
        cadDict.Add "AnnualandSickLeave", "1"
    Else
        cadDict.Add "AWW", 0
        cadDict.Add "AnnualandSickLeave", ""
    End If
    ' Get Status of Union Deductions
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, Array("5C6", "214", "304", "632", "642", "644", "970"), 18, false)
    cadDict.Add "UnionDues", tempResult
    'Get Recreational Association (RA) Deductions
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, "789", 18, false)
    cadDict.Add "RA", tempResult
    'Get Credit Union Status
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, "786*", 18, false)
    cadDict.Add "CreditUnion", tempResult

    'Get Garnishments Status
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, Array("729", "731" ), 18, false)
    cadDict.Add "Garnishments", tempResult
    'Get Student Loan Status
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, "544*", 18, false)
    cadDict.Add "StudentLoans", tempResult
    'GCWCC Contributions
    Dim expressionArray, dateExpression, typeExpression
    'Check if code is 790*
    Set typeExpression = New SearchExpression
    typeExpression.ColumnIndex = 18
    typeExpression.FilterExpression = "790*"
    Set dateExpression = New SearchExpression
    dateExpression.ColumnIndex = 14
    Dim today, oneMonthAgo
    today = date()
    oneMonthAgo = DateAdd("m", -1, date())
    ' 'Don't use date format strings as Excel seems to convert dd/mm/yyyy to mm/dd/yyyy
    dateExpression.BetweenExpressionA = ">=" & CDbl(oneMonthAgo)
    dateExpression.BetweenExpressionB = "<=" & CDbl(today)
    expressionArray = Array(typeExpression, dateExpression)
    tempResult = FindRowUsingFilter(sheetDED, "A2", employeeId, 2, expressionArray, 18, true)
    If tempResult <> "" Then
        cadDict.Add "GCWCC", tempResult
    Else
        cadDict.Add "GCWCC", ""
    End If
    'Get Self-funded leave status
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, Array("675*", "850*"), 18, false)
    cadDict.Add "Selffunded", tempResult
    Set sheetDED = Nothing
    Set sheetJOB = Nothing
End Function
' Get Union Insurance from CAD
Function GetCADDataUnionInsurance(index, cadDict, employeeId)
    ' Sheets
    Dim sheetDED
    Set sheetDED = excelWorkbook.Worksheets(CAD_SHEET_DED)
    ' Union Insurance
    tempResult = SimpleFindRowUsingFilter(CAD_SHEET_DED, "A2", employeeId, 2, 18, "943", 18, true)
    cadDict.Add "UnionInsurance", tempResult
    Set sheetDED = Nothing
End Function
    'Bilingual Bonus
Function GetCADDataBilingualBonus(index, cadDict, employeeId)
    Dim sheetERN
    Set sheetERN = excelWorkbook.Worksheets(CAD_SHEET_ERN)
    Dim tempResult
    tempResult = SimpleFindRowUsingFilter(sheetERN, "A2", employeeId, 2, 18, "141", 18, true)
    cadDict.Add "BilingualBonus", tempResult
    Set sheetERN = Nothing
End Function
' Get Dental Plan Data from CAD
Function GetCADDataDCP(index, cadDict, employeeId)
    ' Sheets
    Dim sheetBEN
    Set sheetBEN = excelWorkbook.Worksheets(CAD_SHEET_BEN)
    Dim tempResult
    ' Add in empty plan number to be filled in manually
    cadDict.Add "DCPPlanNo", ""
    ' DCP Cert No
    tempResult = SimpleFindRowUsingFilter(sheetBEN, "A2", employeeId, 2, 4, "11", 12, false)
    If tempResult <> "" AND Len(tempResult) = 8 Then
        ' Format the number to ## ### #### if it was ########.
        tempResult = Mid(tempResult, 1,2) & " " & Mid(tempResult, 3,3) & " " & Mid(tempResult, 6,3)
    End If
    cadDict.Add "DCPCertNo", tempResult
    If cadDict("DCPCertNo") <> "" Then
        cadDict.Add "DCPStatus", "1"
    Else
        cadDict.Add "DCPStatus", ""
    End If
    Set sheetBEN = Nothing
End Function
' Get Health Plan Data from CAD
Function GetCADDataPSHCP(index, cadDict, employeeId)
    ' Sheets
    Dim sheetDED, sheetBEN
    Set sheetDED = excelWorkbook.Worksheets(CAD_SHEET_DED)
    Set sheetBEN = excelWorkbook.Worksheets(CAD_SHEET_BEN)
    Dim tempResult
    ' PSHCP No
    tempResult = SimpleFindRowUsingFilter(sheetBEN, "A2", employeeId, 2, 4, "10", 12, false)
    If tempResult <> "" AND Len(tempResult) = 7 Then
        cadDict.Add "PSHCP", "1"
        ' Format the number to # ### #### if it was #######.
        tempResult = Mid(tempResult, 1,1) & " " & Mid(tempResult, 2,3) & " " & Mid(tempResult, 5,3)
    Else
        cadDict.Add "PSHCP", ""
    End If
    cadDict.Add "PSHCPNo", tempResult
    ' PSHCP Level
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 23, "HCP00*", 23, true)
    tempResult = Mid(tempResult, 6)
    If tempResult = "3" Then		' Ugh, clean this up as need to use the level in the condition
        cadDict.Add "PSHCPLevel2Or3", "1"
        cadDict.Add "PSHCPLevel1", ""
    ElseIf tempResult = "2" Then
        cadDict.Add "PSHCPLevel2Or3", "1"
        cadDict.Add "PSHCPLevel1", ""
    Else
        cadDict.Add "PSHCPLevel2Or3", ""
        cadDict.Add "PSHCPLevel1", "1"
    End If
    cadDict.Add "PSHCPLevel", tempResult
    Set sheetDED = Nothing
    Set sheetBEN = Nothing
End Function
' Get PSMIP from CAD
Function GetCADDataPSMIP(index, cadDict, employeeId)
    ' Sheets
    Dim sheetDED, sheetBEN
    Set sheetDED = excelWorkbook.Worksheets(CAD_SHEET_DED)
    Set sheetBEN = excelWorkbook.Worksheets(CAD_SHEET_BEN)
    Dim tempResult
    ' Employee has PSMIP?
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, Array("750","811"), 18, false)
    cadDict.Add "PSMIP", tempResult
    If cadDict("PSMIP") Then
        ' PSMIPNo
        tempResult = SimpleFindRowUsingFilter(sheetBEN, "A2", employeeId, 2, 12, "CL*", 12, false)
        ' Truncate out CL
        cadDict.Add "PSMIPNo", Mid(tempResult, 3)
    End If
    Set sheetDED = Nothing
    Set sheetBEN = Nothing
End Function
' Get Service Buyback from CAD
Function GetCADDataServiceBuyback(index, cadDict, employeeId)
    ' Sheets
    Dim sheetDED
    Set sheetDED = excelWorkbook.Worksheets(CAD_SHEET_DED)
    Dim tempResult
    ' Employee has Service Buyback
    tempResult = SimpleFindRowUsingFilter(sheetDED, "A2", employeeId, 2, 18, Array("5H*","5I*"), 18, false)
    cadDict.Add "ServiceBuyback", tempResult <> ""
    Set sheetDED = Nothing
End Function

' Get Parental Leave from CAD
Function GetCadDataParental(index, cadDict, employeeId)
    ' Sheets
    Dim sheetJOB
    Set sheetJOB = excelWorkbook.Worksheets(CAD_SHEET_JOB)

    ' LIA: Job Column I Code PLA on checks from the last business year
    Dim dateExpression, typeExpression, RFLExpression
    ' Check if code PLA
    Set typeExpression = New SearchExpression
    typeExpression.ColumnIndex = 8
    typeExpression.FilterExpression = "PLA"

    ' Check if within 1 year
    Set dateExpression = New SearchExpression
    dateExpression.ColumnIndex = 10
    Dim today, oneYearAgo
    'Don't use date format strings as Excel seems to convert dd/mm/yyyy to mm/dd/yyyy
    today = date()
    oneYearAgo = DateAdd("yyyy",-1,date())
    dateExpression.BetweenExpressionA = ">=" & CDbl(oneYearAgo)
    dateExpression.BetweenExpressionB = "<=" & CDbl(today)
    Dim expressionArray
    expressionArray = Array(typeExpression, dateExpression)
    tempResult = FindRowUsingFilter(sheetJOB, "A2", employeeId, 2, expressionArray, 9, true)
    If tempResult <> "" Then
        Dim Parental
        Parental = CStr(tempResult)
        If Parental = "R" Then
            cadDict.Add "Parental", "1"
        Else
            cadDict.Add "Parental", ""
        End If
    End If
    Set sheetERN = Nothing
End Function
' Get Employee Info from CAD
Function GetCADDataEmployee(index, cadDict, employeeId)
    ' Sheets
    Dim sheetPER
    Set sheetPER = excelWorkbook.Worksheets(CAD_SHEET_PER)
    ' Data From Employee
    cadDict.Add "EmployeeName", sheetPER.Range("D3").Value
    cadDict.Add "EmployeeStreet1", sheetPER.Range("E3").Value
    cadDict.Add "EmployeeStreet2", sheetPER.Range("F3").Value
    cadDict.Add "EmployeeStreet3", sheetPER.Range("G3").Value
    cadDict.Add "EmployeeCity", sheetPER.Range("H3").Value
    cadDict.Add "EmployeeProvince", sheetPER.Range("J3").Value
    cadDict.Add "EmployeePostalCode", sheetPER.Range("I3").Value

    If cadDict("EmployeeProvince") = "QC" Then
        cadDict.Add "EmployeeProvinceQC", "QC"
    Else
        cadDict.Add "EmployeeProvinceQC", ""
    End If

    If cadDict("EmployeeProvince") = "ON" Then
        cadDict.Add "EmployeeProvinceON", "ON"
    Else
        cadDict.Add "EmployeeProvinceON", ""
    End If

    If cadDict("EmployeeProvince") = "ON" OR cadDict("EmployeeProvince") = "QC" Then
        cadDict("EmployeeProvinceONOrQC") = "1"
    Else
        cadDict("EmployeeProvinceONOrQC") = ""
    End If

    Set sheetPER = Nothing
End Function

Function GetCadDataGroup(index, cadDict, employeeId)
' Sheets
    Dim sheetJob
    Set sheetJob = excelWorkbook.Worksheets(CAD_SHEET_JOB)
    group = sheetJob.Range("D3").Value
    If group <> "CAD" AND group <> "CAS" AND group <> "SSB" AND group <> "STS" Then
        cadDict.add "InvCheck-CasualorStudent", ""
    Else
        cadDict.add "InvCheck-CasualorStudent", "1"
    End If
    Set sheetJob = Nothing
End Function
Function GetCadFullPartTime(index, cadDict, employeeId)
' Sheets
    Dim sheetJob
    Set sheetJob = excelWorkbook.Worksheets(CAD_SHEET_JOB)

    partTime = sheetJob.Range("X3").Value

    If partTime = "P" Then
        cadDict.add "PartTimeEmployee", "1"
    Else
        cadDict.add "PartTimeEmployee", ""
    End If

    Set sheetJob = Nothing
End Function

' Simple filters on one column
Function SimpleFindRowUsingFilter(sheet, rangeStart, employeeId, employeeIdColumnIdx, searchForColumnIdx, searchForExpression, retrieveColumnIdx, isLast)
    Dim expression
    Set expression = New SearchExpression
    expression.ColumnIndex = searchForColumnIdx
    expression.FilterExpression = searchForExpression
    Dim expressionArray
    expressionArray = Array(expression)
    SimpleFindRowUsingFilter = FindRowUsingFilter(sheet, rangeStart, employeeId, employeeIdColumnIdx, expressionArray, retrieveColumnIdx, isLast)
End Function
' Just filter employee ID and apply no additional filter
Function FindRowUsingNoAdditionalFilter(sheet, rangeStart, employeeId, employeeIdColumnIdx, retrieveColumnIdx, isLast)
    FindRowUsingNoAdditionalFilter = FindRowUsingFilter(sheet, rangeStart, employeeId, employeeIdColumnIdx, Array(), retrieveColumnIdx, isLast)
End Function

' Use Filter to grab data
Function FindRowUsingFilter(sheet, rangeStart, employeeId, employeeIdColumnIdx, searchForExpressionArray, retrieveColumnIdx, isLast)
    ' https://docs.microsoft.com/en-us/office/vba/api/excel.range.autofilter

    Dim remainingResults
    Dim areasResults
    Dim rowIdx
    Dim areaIdx

    ' Remove filter and allow error in case no filter.
    On Error Resume Next
    sheet.ShowAllData
    'Apply filter. See https://docs.microsoft.com/en-us/office/vba/api/excel.range.autofilter
    With sheet.Range(rangeStart)
        .AutoFilter employeeIdColumnIdx, "=" & employeeId
        For Each expression In searchForExpressionArray
            If Not IsEmpty(expression.FilterExpression) Then
                ' Filter statement
                If IsArray(expression.FilterExpression) Then
                    .AutoFilter expression.ColumnIndex, expression.FilterExpression, 7
                Else
                    .AutoFilter expression.ColumnIndex, expression.FilterExpression
                End If
            Else
                ' Compare statement
                .AutoFilter expression.ColumnIndex, expression.BetweenExpressionA, 1, expression.BetweenExpressionB
            End If
        Next
    End With

    If debugExcelRowFilter Then
        If isLast Then
            MsgBox sheet.Name & " - Open Excel File to see current filter as we are grabbing last row"
        Else
            MsgBox sheet.Name & " - Open Excel File to see current filter as we are grabbing first row"
        End If
    End If

    ' Select all ranges of visible rows which also include the "header" rows.
    ' This can return multiple sets of rows known as "Areas"
    ' 12 is from https://docs.microsoft.com/en-us/office/vba/api/excel.xlcelltype
    Set remainingResults = sheet.UsedRange.SpecialCells(12)
    ' Backup the areas.
    Set areasResults = remainingResults.Areas

    ' Empty if we only have visible header rows which is area 1 range match.
    isEmptyResults = areasResults(1).Rows.Count = 2 AND areasResults.Count = 1
    If debugExcelRowFilter Then
        MsgBox "Area: " & areasResults.Count
        MsgBox "Area 1 Row:" & areasResults(1).Rows.Count
        MsgBox "Area 2 Row:" & areasResults(2).Rows.Count
    End If
    If isEmptyResults Then
        FindRowUsingFilter = ""
        If debugExcelRowFilter Then
            MsgBox "No matches"
        End If
    ElseIf isLast Then
        ' Grab the last area match and then grab the row count for that area to get the
        ' last row.
        areaIdx = areasResults.Count
        Dim indexRow
        FindRowUsingFilter = ""
        ' Why do we have empty rows at the end of these tables ...
        ' Check first row as only "entry" rows have this populated.
        Do
            rowIdx = areasResults(areaIdx).Rows.Count
            ' Loop over last area match, then work backwards.
            While indexRow = "" AND rowIdx > 0
                indexRow = areasResults(areaIdx).Cells(rowIdx, 1)
                If indexRow <> "" AND NOT (areaIdx = 1 AND rowIdx <= 2) Then
                    FindRowUsingFilter = areasResults(areaIdx).Cells(rowIdx, retrieveColumnIdx)
                End If
                rowIdx = rowIdx - 1
            Wend
            areaIdx = areaIdx - 1
        Loop While indexRow = "" And areaIdx > 0

        If debugExcelRowFilter Then
            MsgBox "Found in Last Area >" & FindRowUsingFilter
        End If
    ElseIf areasResults(1).Rows.Count = 2 Then
        ' Grab first match which should be in Match 2.
        FindRowUsingFilter = areasResults(2).Cells(1, retrieveColumnIdx)
        If debugExcelRowFilter Then
            MsgBox "Found in Area 2 >" & FindRowUsingFilter & "<"
        End If
    Else
        ' Grab first row which should follow match 1 row as the match followed directly
        ' after the header row, therefore same area box.
        FindRowUsingFilter = areasResults(1).Cells(3, retrieveColumnIdx)
        If debugExcelRowFilter Then
            MsgBox "Found in Area 1 >" & FindRowUsingFilter & "<"
        End If
    End If
End Function
' Create temporary file
Function CreateTempLetter(sourceFile)
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Dim fileName
    fileName = GetFileNameToUse()

    workingLetterFile = currentDirectory & "\GeneratedLetters\" & fileName & ".docx"
    FSO.CopyFile sourceFile, workingLetterFile
End Function
' Clean up temporary file
Function DeleteTempLetter()
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFile workingLetterFile
End Function
' See if file exists
Function FileExists(FilePath)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(FilePath) Then
    FileExists=CBool(1)
    Else
    FileExists=CBool(0)
    End If
End Function
' Toggle Work in progress
Function ToggleWorkInProgress(isVisible)
    If isVisible Then
        document.getElementById("WorkInProgress").className = ""
    Else
        document.getElementById("WorkInProgress").className = "hidden"
    End If
End Function

Sub ValidatePage1()
    Set letterSelect = document.getElementById("selectLetter")
    Set languageSelect = document.getElementById("selectLanguage")
    Set dateInput = document.getElementById("inputEffectiveDate")

    If letterSelect.value <> "--" Then
        languageSelect.Disabled = false

        If selectLanguage.value <> "--" Then
            dateInput.Disabled = false

            If Len(dateInput.value) = 10 Then
                TogglePage1StartButton(true)
            Else
                TogglePage1StartButton(false)
            End If
        Else
            dateInput.Disabled = true
            TogglePage1StartButton(false)
        End If

    Else
        languageSelect.Disabled = true
        dateInput.Disabled = true
        TogglePage1StartButton(false)
    End If
End Sub

Function TogglePage1StartButton(isEnabled)
    document.getElementById("page1ButtonStart").Disabled = Not(isEnabled)
End Function

Function ToggleNotCheck(parent, child)
    document.getElementbyID(child).checked = NOT document.getElementbyID(parent).checked
End Function

' Enable or disable the word document generation
Function ToggleWordGenerateButtons(isEnabled)
    document.getElementById("page3ButtonGenerate").Disabled = Not(isEnabled)
End Function
Function EmailSubjectLine()
    EmailSubjectLine = GetFileNameToUse()
End Function
Function EmailBody()
    languageCode = document.getElementById("html").lang
    If languageCode = "fr" Then
        EmailBody = "<ul>" & _
                        "<li>Nom de l’employé: " & employeeRecord("EmployeeName") & "</li>" & _
                        "<li>CIDP: " & employeeRecord("PRI")  & "</li>" &  _
                        "<li>Cas # : " & GetCaseNo()  & "</li>" & _
                        "<li>Adresse courriel personnelle : " & employeeRecord("PersonalEmail")  & "</li>" & _
                        "<li>Langue : " & documentLanguage(0) & "</li>" & _
                        "</ul>"
    ElseIf languageCode = "en" Then
        EmailBody = "<ul>" & _
                        "<li>EE name: " & employeeRecord("EmployeeName") & "</li>" & _
                        "<li>PRI: " & employeeRecord("PRI")  & "</li>" &  _
                        "<li>Case #: " & GetCaseNo()  & "</li>" & _
                        "<li>EE personal Email address: " & employeeRecord("PersonalEmail")  & "</li>" & _
                        "<li>Language: " & documentLanguage(1) & "</li>" & _
                        "</ul>"
    End If
End Function
Function SendEmail()
    Set objOutlook = CreateObject("Outlook.Application")
    Set objEmail = objOutlook.CreateItem(0) '0 is email

    With objEmail
        .To = configurationSettings(CONFIG_GD_EPOST)
        .Subject = EmailSubjectLine()
        .HTMLBody = EmailBody()
        .Attachments.Add pdfFilePath
        If documentLanguage(1) = "English" Then
            For Each attachment in attachmentsEN
            .Attachments.Add attachment
            Next
        ElseIf documentLanguage(1) = "French" Then
            For Each attachment in attachmentsFR
            .Attachments.Add attachment
            Next
        End If
        .Display
    End With
    Set objEmail = Nothing
    Set objOutlook = Nothing
End Function
' Flip the installation status
Function UpdateValidationStatusEntry(LabelId, FilePath)
    Result = FileExists(FilePath)
    If Result Then
        document.getElementById(LabelId).parentElement.className = "success"
    Else
        document.getElementById(LabelId).parentElement.className = "error"
    End If
    UpdateValidationStatusEntry = Result
End Function

' Go to the given page number
Function GoToPage(pageNo, e)
    ' Show page first
    document.getElementById("cad-form-container").scrollTop = 0
    document.body.className = "page-" & pageNo
    ShowProgressOverlay(True)

    Select Case pageNo
        Case 2
            window.setTimeOut "GoToPage2", 1000, "VBScript"
        Case 3
            window.setTimeOut "GoToPage3", 1000, "VBScript"
        Case 4
            window.setTimeOut "GoToPage4", 1000, "VBScript"
        Case 5
            window.setTimeOut "GoToPage5", 1000, "VBScript"
        Case 6
            window.setTimeOut "GoToPage6", 1000, "VBScript"
        Case Else
            ResetData()
            window.setTimeOut "GoToPage1", 0, "VBScript"
    End Select

End Function

' Toggle overlay
Function ShowProgressOverlay(isOn)
    Dim classValue
    If IsOn Then
        classValue = ""
    Else
        classValue = "hidden"
    End If
    ' Turn on the indicator
    document.getElementById("progress-overlay").className = classValue

End Function

' Toggle Error
' All validation of forms in the letterbuilder
Function validateForm()
    Dim targetElement
    ' Validate EffectiveDate
    Set targetElement = document.getElementById("EffectiveDate")
    If validateAndFormatDate(targetElement.value) = "" Then
        ToggleWordGenerateButtons(false)
    Else
        ToggleWordGenerateButtons(true)
    End If
    ' Validate DCP Plan No
    Set targetElement = document.getElementById("DCPPlanNo")
    Set targetCheckbox = document.getElementById("DCPStatus")
    If targetElement.value = "" AND targetCheckbox.checked = true Then
        ToggleWordGenerateButtons(false)
        targetElement.parentElement.className = "form-field form-field-error"
    Else
        ToggleWordGenerateButtons(true)
        targetElement.parentElement.className = "form-field"
    End If
End Function

Function GoToPage1()
    CheckAvailableModules()
    ShowProgressOverlay(False)
    ValidatePage1()
End Function

Function GoToPage2()
    SetDocumentLanguage()
    CheckInstall
    ShowProgressOverlay(False)
End Function

Function GoToPage3()
    OpenCAD()
    ShowProgressOverlay(False)
End Function

Function GoToPage4()
    ShowProgressOverlay(False)
    PopulateScreenCADFields(0)
    PopulateScreenCADFields(1)
    validateForm()
End Function

Function GoToPage5()
    If documentLanguage(1) = "English" Then
        GenerateEnglishLetter()
    Else
        GenerateFrenchLetter()
    End If
    ShowProgressOverlay(False)
End Function

Function GoToPage6()
    CreatePDF()
    document.getElementById("pdfFileName").innerText = pdfFileName
    SendEmail()
    ShowProgressOverlay(False)
End Function

Function GoToPage7()
    ShowProgressOverlay(False)
End Function

' Events
' =======================================
' This procedure is executing on program open
Sub window_onload()
    LoadConfig(configFile)
    'Opens the application on the left side of the screen
    TranslateScreen "en", Nothing
    GoToPage 1, Nothing
End Sub