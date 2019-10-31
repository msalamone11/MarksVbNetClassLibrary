Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.Threading
Imports Excel = Microsoft.Office.Interop.Excel

Namespace PowerTermMethods

    Public Module PowerTermMethods

        Public Sub AddDataToRecollectForm(ExcelFileName As String, PscSiteCode As String,
                                          AccessionNumber As String,
                                          PatientName As String, PatientPhone As String,
                                          DateOfBirth As String, DateOfCollection As String,
                                          ClientAccountNumber As String)

            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing

            Try

                xlApp = GetObject(, "Excel.Application")

                For Each xlWorkBook In xlApp.Workbooks

                    If xlWorkBook.Name = ExcelFileName Then

                        xlWorkBook.Activate()
                        xlWorkBook = xlApp.ActiveWorkbook

                    End If

                Next

                WindowsOSMethods.BringWindowToFront(ExcelFileName)

                xlWorkSheet = xlWorkBook.Worksheets("DATA_ENTRY")

                xlWorkSheet.Cells(7, 4) = PscSiteCode
                xlWorkSheet.Cells(8, 5) = AccessionNumber
                xlWorkSheet.Cells(9, 5) = PatientName
                xlWorkSheet.Cells(10, 5) = PatientPhone
                xlWorkSheet.Cells(11, 5) = DateOfBirth
                xlWorkSheet.Cells(12, 5) = DateOfCollection

                xlWorkSheet.Cells(13, 5).Select

                If ExcelFileName = "SW_Recollect.xlsm" Then
                    xlWorkSheet.Cells(10, 9) = ClientAccountNumber
                End If

            Catch ex As Exception

                MsgBox("    It looks like an error occurred populating the data into Excel. Please close and re-open the Recollect Tool and try again." & vbNewLine & vbNewLine &
                       "If the issue persists, please copy and paste the values from QLS to Excel." & vbNewLine & vbNewLine &
                       "Please note, you need to double click the excel cell before pasting the data into it." & vbNewLine &
                       "If this issue keeps happening in the future, please inform your Tech Rep first, " & vbNewLine & vbNewLine &
                       "and if they aren't able to figure out what is going on, then please inform Mark Salamone by email. Thank you!",
                       MsgBoxStyle.SystemModal, "Unable To Populate Data To Form")

                Exit Sub

            Finally

                releaseObject(xlApp)
                releaseObject(xlWorkBook)
                releaseObject(xlWorkSheet)

            End Try

        End Sub

        Public Function ChangeOrderEntryScreenTextValidation(ChangeOrderEntryScreenText As String)

            Dim regex As Regex = New Regex("Client\s+number\s+\S+")
            Dim match As Match = regex.Match(ChangeOrderEntryScreenText)

            If match.Success Then
                Return True
            Else

                Dim MsgBoxResponse As Int32

                MsgBoxResponse = MsgBox("I don't believe I'm on the Change Order Entry screen. " & vbNewLine & vbNewLine &
                                        "Navigate to the screen by typing 'C2' and pressing enter three times. " & vbNewLine &
                                        "Then click Okay and I should be ableto scrape the information you need. " & vbNewLine & vbNewLine &
                                        "Otherwise, click 'Cancel' to exit the program.",
                                        MsgBoxStyle.OkCancel.SystemModal,
                                        "Change Order Entry Screen")

                If MsgBoxResponse = 2 Then

                    Return False

                End If

                Return True

            End If

        End Function

        Public Function DemographicScreenTextValidation(DemographicsScreenText As String)

            Dim WeAreOnTheDemoScreenCheck1 As Boolean = DemographicsScreenText.Contains("PATIENT INQUIRY") And DemographicsScreenText.Contains("Client number")
            Dim WeAreOnTheDemoScreenCheck2 As Boolean = DemographicsScreenText.Contains("Requisition") And DemographicsScreenText.Contains("Accession")

            If WeAreOnTheDemoScreenCheck1 OrElse WeAreOnTheDemoScreenCheck2 Then
                Return True
            Else
                MsgBox("I don't believe the correct information was captured. Are we on the Patient Inquiry Screen? Please try again.",
                   MsgBoxStyle.SystemModal,
                    "Unable To Capture QLS Text")
                Return False
            End If
        End Function

        Public Function GetAccessionNumber(PowerTermScreenText As String)

            Dim AccessionNumber As String

            Try

                AccessionNumber = Regex.Match(PowerTermScreenText, "Accession\s+[A-Z][A-Z]\d{6}[A-Z]", RegexOptions.Multiline).Value

                AccessionNumber = Regex.Replace(AccessionNumber, "Accession\s+", "").Trim

                If AccessionNumber = "" Then Throw New Exception

                Return AccessionNumber

            Catch ex As Exception

                AccessionNumber = InputBox("I was not able to capture the Accession Number. Could you enter it for me?",
                                       "Input Accession Number", "Insert Accession Number").Trim

                Return AccessionNumber

            End Try

        End Function

        Public Function GetBillTo(PowerTermScreenText As String)

            Dim BillTo As String

            Try

                BillTo = Regex.Match(PowerTermScreenText, "Bill\s+To\s+(?:(?!24).)*").Value 'Capture everything up to but not including "24".

                BillTo = Regex.Replace(BillTo, "Bill\s+To\s+", "").Trim

                If BillTo = "" Then Throw New Exception

                Return BillTo

            Catch ex As Exception

                BillTo = InputBox("I was not able to capture the Bill To. Could you enter it for me?",
                                       "Input Bill To", "Insert Bill To").Trim

                Return BillTo

            End Try

        End Function

        Public Function GetChangeOrderEntryScreenText(PowerTermWindowUserWants As String)

            Dim ChangeOrderScreenText As String

            WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing)

            NavigateToChangeOrderScreen(PowerTermWindowUserWants)

            ChangeOrderScreenText = GetPowerTermScreenText()

            WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing)

            Return ChangeOrderScreenText

        End Function

        Public Function GetClientAccountNumber(PowerTermScreenText As String)

            Dim ClientAccountNumber As String

            Try

                ClientAccountNumber = Regex.Match(PowerTermScreenText, "Client\s+number\s+\S+").Value

                ClientAccountNumber = Regex.Replace(ClientAccountNumber, "Client\s+number\s+", "").Trim

                If ClientAccountNumber = "" Then Throw New Exception

                Return ClientAccountNumber

            Catch ex As Exception

                ClientAccountNumber = InputBox("I was not able to capture the client Account Number. Could you enter it for me?",
                                       "Input Client Account Number", "Insert Client Account Number").Trim

                Return ClientAccountNumber

            End Try

        End Function

        Public Function GetClientAddress(ClientInquiryScreenText As String)

            Dim ClientAddress As String

            Try

                ClientAddress = Regex.Match(ClientInquiryScreenText, "Address\s+1(?:(?!Address\s+2).)*").Value 'Match everything up to but not including "Address 2"

                ClientAddress = Regex.Replace(ClientAddress, "Address\s+1\s+", "").Trim

                If ClientAddress = "" Then Throw New Exception

                Return ClientAddress

            Catch ex As Exception

                ClientAddress = InputBox("I was not able to capture the Client's Address. Could you enter it for me?",
                                       "Input Client's Address", "Insert Client's Address").Trim

                Return ClientAddress

            End Try

        End Function

        Public Function GetClientCity(ClientInquiryScreenText As String)

            Dim ClientCity As String

            Try

                ClientCity = Regex.Match(ClientInquiryScreenText, "City(?:(?!State).)*").Value 'Match everything up to but not including "City"

                ClientCity = Regex.Replace(ClientCity, "City\s+", "").Trim

                If ClientCity = "" Then Throw New Exception

                Return ClientCity

            Catch ex As Exception

                ClientCity = InputBox("I was not able to capture the Client's City. Could you enter it for me?",
                                       "Input Client's City", "Insert Client's City").Trim

                Return ClientCity

            End Try

        End Function

        Public Function GetClientName(PowerTermScreenText As String)

            Dim ClientName As String

            Try

                ClientName = Regex.Match(PowerTermScreenText, "Client\s+number\s+\w+\s+\w+\s+.+", RegexOptions.Multiline).Value

                ClientName = Regex.Replace(ClientName, "Client\s+number\s+\w+\s+\w+\s+", "").Trim

                If ClientName = "" Then Throw New Exception

                Return ClientName

            Catch ex As Exception

                ClientName = InputBox("I was not able to capture the Client Name. Could you enter it for me?",
                                      "Input Client Name", "Insert Client Name").Trim

                Return ClientName

            End Try

        End Function

        Public Function GetClientPhoneNumber(ClientInquiryScreenText As String)

            Dim ClientPhoneNumber As String

            Try

                ClientPhoneNumber = Regex.Match(ClientInquiryScreenText, "Office\s+Phone(?:(?!After\s+Hrs).)*").Value 'Match everything up to but not including "After Hours"

                ClientPhoneNumber = Regex.Replace(ClientPhoneNumber, "Office\s+Phone\s+", "").Trim

                ClientPhoneNumber = ClientPhoneNumber.Replace(" ", "")
                ClientPhoneNumber = ClientPhoneNumber.Replace("-", "")
                ClientPhoneNumber = ClientPhoneNumber.Replace("(", "")
                ClientPhoneNumber = ClientPhoneNumber.Replace(")", "")

                If ClientPhoneNumber = "" Then Throw New Exception

                ClientPhoneNumber = ClientPhoneNumber.Substring(0, 10)

                Return ClientPhoneNumber

            Catch ex As Exception

                ClientPhoneNumber = InputBox("I was not able to capture the Client's Phone Number. Could you enter it for me?",
                                       "Input Client's Phone Number", "Insert Client's Phone Number").Trim

                Return ClientPhoneNumber

            End Try

        End Function

        Public Function GetClientState(ClientInquiryScreenText As String)

            Dim ClientState As String

            Try

                ClientState = Regex.Match(ClientInquiryScreenText, "State(?:(?!Zip).)*").Value 'Match everything up to but not including "Zip"

                ClientState = Regex.Replace(ClientState, "State\s+", "").Trim

                If ClientState = "" Then Throw New Exception

                Return ClientState

            Catch ex As Exception

                ClientState = InputBox("I was not able to capture the Client's State. Could you enter it for me?",
                                       "Input Client's State", "Insert Client's State").Trim

                Return ClientState

            End Try

        End Function

        Public Function GetClientZipCode(ClientInquiryScreenText As String)

            Dim ClientZipCode As String

            Try

                ClientZipCode = Regex.Match(ClientInquiryScreenText, "Zip(?:(?!Std).)*").Value 'Match everything up to but not including "Std"

                ClientZipCode = Regex.Replace(ClientZipCode, "Zip\s+", "").Trim

                If ClientZipCode = "" Then Throw New Exception

                Return ClientZipCode

            Catch ex As Exception

                ClientZipCode = InputBox("I was not able to capture the Client's Zip Code. Could you enter it for me?",
                                       "Input Client's Zip Code", "Insert Client's Zip Code").Trim

                Return ClientZipCode

            End Try

        End Function

        Public Function GetDateOfBirth(PowerTermScreenText As String)

            Dim DateOfBirth As String

            Try

                DateOfBirth = Regex.Match(PowerTermScreenText, "Age,D\.O\.B\..+").Value

                DateOfBirth = Regex.Match(DateOfBirth, "\d\d/\d\d/\d\d\d\d").Value 'Extract just the date

                If DateOfBirth = "" Then Throw New Exception

                Return DateOfBirth

            Catch ex As Exception

                DateOfBirth = InputBox("I was not able to capture the Patient's Date Of Birth. Could you enter it for me?",
                                       "Input Date Of Birth", "Insert Date Of Birth").Trim

                Return DateOfBirth

            End Try

        End Function

        Public Function GetDateOfCollection(PowerTermScreenText As String)

            Dim DateOfCollection As String

            Try

                If PowerTermScreenText.Contains("Date Coll") = False Then

                    Throw New Exception

                Else

                    DateOfCollection = Regex.Match(PowerTermScreenText, "Date\s+Coll\s+\d\d/\d\d/\d\d\d\d", RegexOptions.Multiline).Value
                    DateOfCollection = Regex.Replace(DateOfCollection, "Date\s+Coll\s+", "").Trim

                    If DateOfCollection = "" Then Throw New Exception

                End If

                DateOfCollection = DateOfCollection.Trim

                Return DateOfCollection

            Catch ex As Exception

                DateOfCollection = InputBox("I was not able to capture the Date Of Collection. Could you enter it for me?",
                                            "Input Date Of Collection", "Insert Date Of Collection").Trim

                Return DateOfCollection

            End Try

        End Function

        Public Function GetDemographicsScreenText(PowerTermWindowUserWants As String)

            Dim DemographicsScreenText As String

            NavigateToTopOfDemographicsScreen(PowerTermWindowUserWants)
            DemographicsScreenText = GetPowerTermScreenText()

            If DemographicsScreenText.Contains("Date Coll") = False _
                Or DemographicsScreenText.Contains("Time Coll") = False _
                Or DemographicsScreenText.Contains("Site Code") = False _
                Or DemographicsScreenText.Contains("Patient Phone") = False _
                Or DemographicsScreenText.Contains("Bill To") = False Then

                WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing)
                SendKeys.SendWait("{Enter}")
                Thread.Sleep(1000)

                DemographicsScreenText = DemographicsScreenText & vbCrLf & GetPowerTermScreenText()

                NavigateToTopOfDemographicsScreen(PowerTermWindowUserWants)

            End If

            Return DemographicsScreenText

        End Function

        Public Function GetPatientPhone(PowerTermScreenText As String)

            'Note: The regex of "Patient\s+Phone(?:(?!12\.).)*" will work for both the
            'Demographics Screen as well as the Change Order Entry Screen

            Dim PatientPhone As String

            Try

                PatientPhone = Regex.Match(PowerTermScreenText, "Patient\s+Phone(?:(?!12\.).)*").Value

                PatientPhone = Regex.Replace(PatientPhone, "Patient\s+Phone\s+", "").Trim

                PatientPhone = PatientPhone.Replace(" ", "")
                PatientPhone = PatientPhone.Replace("-", "")
                PatientPhone = PatientPhone.Replace("(", "")
                PatientPhone = PatientPhone.Replace(")", "")

                If PatientPhone = "" Then Throw New Exception

                Return PatientPhone

            Catch ex As Exception

                PatientPhone = InputBox("I was not able to capture the Patient's Phone Number. Could you enter it for me?",
                                       "Input Patient Phone Number", "Insert Patient's Phone Number").Trim

                PatientPhone = PatientPhone.Replace(" ", "")
                PatientPhone = PatientPhone.Replace("-", "")
                PatientPhone = PatientPhone.Replace("(", "")
                PatientPhone = PatientPhone.Replace(")", "")

                Return PatientPhone

            End Try

        End Function

        Public Function GetPatientName(PowerTermScreenText As String)

            Dim PatientName As String

            Try
                'Sometimes "[ONLINE CALL LOG]" will be on the same line as the Patient's name.
                'The below regex will capture what we want up to "[ONLINE"
                PatientName = Regex.Match(PowerTermScreenText, "Patient\s+Name(?:(?!\[ONLINE).)*").Value

                PatientName = Regex.Replace(PatientName, "Patient\s+Name\s+", "").Trim

                If PatientName = "" Then Throw New Exception

                Return PatientName

            Catch ex As Exception

                PatientName = InputBox("I was not able to capture the Patient Name. Could you enter it for me?",
                                       "Input Patient Name", "Insert Patient Name").Trim

                Return PatientName

            End Try

        End Function

        Public Function GetPowerTermScreenText()

            Clipboard.SetText("Cleared text from clipboard")

            SendKeys.SendWait("%{F10}") 'Hold down the Alt key and press F10. PowerTerm auto-copies text to clipboard

            Return Clipboard.GetText

        End Function

        Public Function GetPowerTermWindowUserWants()

            Dim PowerTermWindowTitleUserWantsFullText As String = Nothing

            Dim ProcessList() As Process = Process.GetProcesses()
            Dim PowerTermScreenLetter As String = InputBox("Which QLS Screen Letter do we want to get text from?").ToUpper

            If PowerTermScreenLetter = "" Then 'The cancel button was clicked.
                Return ""
            End If

            While Len(PowerTermScreenLetter) <> 1
                PowerTermScreenLetter = InputBox("The PowerTerm Screen Letter can only be one letter in length. Please try again.")
            End While

            Dim PowerTermWindowTitleUserWantsPartialText As String = "(" & PowerTermScreenLetter & ") "

            For Each process As Process In ProcessList
                If String.IsNullOrEmpty(process.MainWindowTitle) = False Then
                    If process.MainWindowTitle.Contains(PowerTermWindowTitleUserWantsPartialText) Then
                        PowerTermWindowTitleUserWantsFullText = process.MainWindowTitle
                        Exit For
                    End If
                End If
            Next

            Return PowerTermWindowTitleUserWantsFullText

        End Function

        Public Function GetPscSiteCode(PowerTermScreenText As String)

            Dim PscSiteCode As String

            Try

                PscSiteCode = Regex.Match(PowerTermScreenText, "Site\s+Code\s+\w+", RegexOptions.Multiline).Value

                PscSiteCode = Regex.Replace(PscSiteCode, "Site\s+Code\s+", "").Trim

                If PscSiteCode = "" Then Throw New Exception

                Return PscSiteCode

            Catch ex As Exception

                PscSiteCode = InputBox("I was not able to capture the PSC Site Code. Could you enter it for me?",
                                       "Input PSC Site Code", "Insert PSC Site Code").Trim

                Return PscSiteCode

            End Try

        End Function

        Public Function GetRequisitionNumber(PowerTermScreenText As String)

            Dim RequisitionNumber As String

            Try

                RequisitionNumber = Regex.Match(PowerTermScreenText, "Requisition\s+\S+", RegexOptions.Multiline).Value

                RequisitionNumber = Regex.Replace(RequisitionNumber, "Requisition\s+", "").Trim

                If RequisitionNumber = "" Then Throw New Exception

                Return RequisitionNumber

            Catch ex As Exception

                RequisitionNumber = InputBox("I was not able to capture the Requisition Number. Could you enter it for me?",
                                       "Input Requisition Number", "Insert Requisition Number").Trim

                Return RequisitionNumber

            End Try

        End Function

        Public Function GetTimeOfCollectionDemographicsScreen(PowerTermScreenText As String)

            Dim TimeOfCollection As String

            Try

                If PowerTermScreenText.Contains("Time Coll") = False Then

                    Return ""

                Else

                    TimeOfCollection = Regex.Match(PowerTermScreenText, "Time\s+Coll\s+\d\d:\d\d", RegexOptions.Multiline).Value
                    TimeOfCollection = Regex.Replace(TimeOfCollection, "Time\s+Coll\s+", "").Trim

                    Return TimeOfCollection

                End If

            Catch ex As Exception

                Return ""

            End Try

        End Function

        Public Sub NavigateToChangeOrderScreen(PowerTermWindowUserWants As String)

            Dim PowerTermScreenText As String
            Dim DoWhileCounter As Int32

            WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing) : Thread.Sleep(25)

            If Control.IsKeyLocked(Keys.CapsLock) Then

                SendKeys.SendWait("c") : Thread.Sleep(250)

            Else

                SendKeys.SendWait("C") : Thread.Sleep(250)

            End If

            SendKeys.SendWait("2") : Thread.Sleep(250)

            PowerTermScreenText = "Nothing"

            Do While Regex.IsMatch(PowerTermScreenText, "(4\)\s+Sep\s+ID\s+#   (_|\w)|Maxi\s+Completed)") = False
                SendKeys.SendWait("{ENTER}") : Thread.Sleep(250)
                PowerTermScreenText = GetPowerTermScreenText()

                DoWhileCounter += 1

                If DoWhileCounter > 20 Then

                    MsgBox("I wasn't able to get to the Change Order Screen. Please try again.")

                    Exit Sub

                End If

            Loop

            SendKeys.SendWait("{ENTER}") : Thread.Sleep(1000)

        End Sub

        Public Sub NavigateToClientInquiryScreen(PowerTermWindowUserWants As String)

            Dim PowerTermScreenText
            Dim WeAreOnTheClientInquiryScreen As Boolean
            Dim WhileCounter As Int32
            Dim CapsLockIsOn As Boolean = Control.IsKeyLocked(Keys.CapsLock)

            WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing) : Thread.Sleep(250)

            PowerTermScreenText = GetPowerTermScreenText()

            WeAreOnTheClientInquiryScreen = PowerTermScreenText.Contains("Mnemonic") _
                                            Or PowerTermScreenText.Contains("Alpha Name") _
                                            Or PowerTermScreenText.Contains("Address 1")

            While WeAreOnTheClientInquiryScreen = False
                WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing)
                TypeTextPressEnter(PowerTermWindowUserWants, "Y") 'The enter in this method doesn't seem to work.
                SendKeys.SendWait("{Enter}") 'This is an extra enter just in case the above method fails to press enter.
                Thread.Sleep(1000) 'These two delays seem to be necessary when running the program with the 'Get Fr PINQ >' button in Excel
                Thread.Sleep(1000)
                PowerTermScreenText = GetPowerTermScreenText()
                WeAreOnTheClientInquiryScreen = PowerTermScreenText.Contains("Mnemonic") _
                                            Or PowerTermScreenText.Contains("Alpha Name") _
                                            Or PowerTermScreenText.Contains("Address 1")
                WhileCounter += 1

                If WhileCounter > 4 Then
                    Exit Sub
                End If

            End While

            PowerTermScreenText = GetPowerTermScreenText()

            WeAreOnTheClientInquiryScreen = PowerTermScreenText.Contains("Mnemonic") _
                                            Or PowerTermScreenText.Contains("Alpha Name") _
                                            Or PowerTermScreenText.Contains("Address 1")

            If WeAreOnTheClientInquiryScreen = False Then

                MsgBox("I wasn't able to get to the Client Inquiry Screen. " & Environment.NewLine &
                        "Please navigate to the Client Inquiry screen by typing " & Environment.NewLine &
                        "'Y' from inside the patient's Accession. Then press 'Ok'.")

                PowerTermScreenText = GetPowerTermScreenText()

                Exit Sub

            End If

        End Sub

        Public Sub NavigateToDemographicsScreenFromChangeOrderScreen()

            Dim DoWhileCounter As Int32 = 0
            Dim regex As Regex = New Regex("3\)\s+Accession")

            SendKeys.SendWait("{/}") : Thread.Sleep(50) : SendKeys.SendWait("{/}") : Thread.Sleep(50)
            SendKeys.SendWait("{Enter}") : Thread.Sleep(1000)

            Dim PowerTermScreenText = GetPowerTermScreenText()
            Dim match As Match = regex.Match(PowerTermScreenText)

            Do While match.Success = True
                SendKeys.SendWait("{/}") : Thread.Sleep(50) : SendKeys.SendWait("{/}") : Thread.Sleep(50)
                SendKeys.SendWait("{Enter}")
                Thread.Sleep(250)

                PowerTermScreenText = GetPowerTermScreenText()
                match = regex.Match(PowerTermScreenText)

                If DoWhileCounter > 5 Then Exit Do

                DoWhileCounter += 1
            Loop

        End Sub

        Public Sub NavigateToTopOfDemographicsScreen(PowerTermWindowUserWants As String)

            Dim DoWhileCounter As Int32 = 0
            Dim PowerTermScreenText = GetPowerTermScreenText()

            WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing) : Thread.Sleep(25)

            While Regex.IsMatch(PowerTermScreenText, "PATIENT INQUIRY") = False

                Dim CapsLockIsOn As Boolean = Control.IsKeyLocked(Keys.CapsLock)

                If CapsLockIsOn Then

                    SendKeys.SendWait("d") : Thread.Sleep(250)

                Else

                    SendKeys.SendWait("D") : Thread.Sleep(250)

                End If

                SendKeys.SendWait("{Enter}") : Thread.Sleep(1000)

                PowerTermScreenText = GetPowerTermScreenText()

                DoWhileCounter += 1

                If DoWhileCounter > 5 Then

                    MsgBox("I wasn't able to get to the top of the Patient Inquiry screen. Please try again.")

                    Exit Sub

                End If

            End While

        End Sub

        Private Sub releaseObject(ByVal obj As Object)
            Try
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            Catch ex As Exception
                obj = Nothing
            Finally
                GC.Collect()
            End Try
        End Sub

        Public Sub TypeTextPressEnter(PowerTermWindowUserWants As String, TextToType As String)

            Dim CapsLockIsOn As Boolean = Control.IsKeyLocked(Keys.CapsLock)

            WindowsOSMethods.SelectWindow(PowerTermWindowUserWants, Nothing)

            If CapsLockIsOn Then

                SendKeys.SendWait(TextToType.ToLower)

            Else

                SendKeys.SendWait(TextToType.ToUpper)

            End If

            Thread.Sleep(1000)
            SendKeys.SendWait("{Enter}") : Thread.Sleep(1000)

        End Sub

    End Module

End Namespace