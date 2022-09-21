Imports System.Management

Public Class Form1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim searcher As New ManagementObjectSearcher("root\CIMV2", "SELECT * FROM Win32_Printer")

        For Each queryObj As ManagementObject In searcher.[Get]()

            If queryObj("Default") = True Then

                Dim A As New ArrayList

                A.Add("Attributes")
                A.Add("Availability")
                A.Add("AvailableJobSheets")
                A.Add("AveragePagesPerMinute")
                A.Add("Capabilities")
                A.Add("CapabilityDescriptions")
                A.Add("Caption")
                A.Add("CharSetsSupported")
                A.Add("Comment")
                A.Add("ConfigManagerErrorCode")
                A.Add("ConfigManagerUserConfig")
                A.Add("CreationClassName")
                A.Add("CurrentCapabilities")
                A.Add("CurrentCharSet")
                A.Add("CurrentLanguage")
                A.Add("CurrentMimeType")
                A.Add("CurrentNaturalLanguage")
                A.Add("CurrentPaperType")
                A.Add("Default")
                A.Add("DefaultCapabilities")
                A.Add("DefaultCopies")
                A.Add("DefaultLanguage")
                A.Add("DefaultMimeType")
                A.Add("DefaultNumberUp")
                A.Add("DefaultPaperType")
                A.Add("DefaultPriority")
                A.Add("Description")
                A.Add("DetectedErrorState")
                A.Add("DeviceID")
                A.Add("Direct")
                A.Add("DoCompleteFirst")
                A.Add("DriverName")
                A.Add("EnableBIDI")
                A.Add("EnableDevQueryPrint")
                A.Add("ErrorCleared")
                A.Add("ErrorDescription")
                A.Add("ErrorInformation")
                A.Add("ExtendedDetectedErrorState")
                A.Add("ExtendedPrinterStatus")
                A.Add("Hidden")
                A.Add("HorizontalResolution")
                A.Add("InstallDate")
                A.Add("JobCountSinceLastReset")
                A.Add("KeepPrintedJobs")
                A.Add("LanguagesSupported")
                A.Add("LastErrorCode")
                A.Add("Local")
                A.Add("Location")
                A.Add("MarkingTechnology")
                A.Add("MaxCopies")
                A.Add("MaxNumberUp")
                A.Add("MaxSizeSupported")
                A.Add("MimeTypesSupported")
                A.Add("Name")
                A.Add("NaturalLanguagesSupported")
                A.Add("Network")
                A.Add("PaperSizesSupported")
                A.Add("PaperTypesAvailable")
                A.Add("Parameters")
                A.Add("PNPDeviceID")
                A.Add("PortName")
                A.Add("PowerManagementCapabilities")
                A.Add("PowerManagementSupported")
                A.Add("PrinterPaperNames")
                A.Add("PrinterState")
                A.Add("PrinterStatus")
                A.Add("PrintJobDataType")
                A.Add("PrintProcessor")
                A.Add("Priority")
                A.Add("Published")
                A.Add("Queued")
                A.Add("RawOnly")
                A.Add("SeparatorFile")
                A.Add("ServerName")
                A.Add("Shared")
                A.Add("ShareName")
                A.Add("SpoolEnabled")
                A.Add("StartTime")
                A.Add("Status")
                A.Add("StatusInfo")
                A.Add("SystemCreationClassName")
                A.Add("SystemName")
                A.Add("TimeOfLastReset")
                A.Add("UntilTime")
                A.Add("VerticalResolution")
                A.Add("WorkOffline")

                For i = 0 To A.Count - 1
                    Try
                        If queryObj(A(i).ToString) <> Nothing Then
                            ListBox1.Items.Add(A(i).ToString & ": " & queryObj(A(i).ToString).ToString)
                        End If
                    Catch ex As Exception

                    End Try
                Next

            End If
        Next
    End Sub

    ''NAME
    '.Ptr_Name = queryObj("Name").ToString

    ''PORT DETAILS
    '.Ptr_PortName = queryObj("PortName").ToString

    ''DEFAULT PRINTER
    '.Ptr_Default = queryObj("Default").ToString
    'If .Ptr_Default = "False" Then Home.Ptr_Status = 2

    '                    'EXTENDED STATUS
    '                    If queryObj("ExtendedPrinterStatus").ToString <> Nothing Then
    'If queryObj("ExtendedPrinterStatus").ToString = "1" Then
    '.Ptr_StatusMsg = "Other"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "2" Then
    '.Ptr_StatusMsg = "Unknown"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "3" Then
    '.Ptr_StatusMsg = "Idle"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "4" Then
    '.Ptr_StatusMsg = "Printing"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "5" Then
    '.Ptr_StatusMsg = "Warming Up"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "6" Then
    '.Ptr_StatusMsg = "Stopped Printing"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "7" Then
    '.Ptr_StatusMsg = "Offline"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "8" Then
    '.Ptr_StatusMsg = "Paused"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "9" Then
    '.Ptr_StatusMsg = "Error"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "10" Then
    '.Ptr_StatusMsg = "Busy"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "11" Then
    '.Ptr_StatusMsg = "Not Available"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "12" Then
    '.Ptr_StatusMsg = "Waiting"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "13" Then
    '.Ptr_StatusMsg = "Processing"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "14" Then
    '.Ptr_StatusMsg = "Initialization"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "15" Then
    '.Ptr_StatusMsg = "Power Save"
    '                            Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "16" Then
    '.Ptr_StatusMsg = "Pending Deletion"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "17" Then
    '.Ptr_StatusMsg = "I/O Active"
    '                            Home.Ptr_Status = 1
    '                        ElseIf queryObj("ExtendedPrinterStatus").ToString = "18" Then
    '.Ptr_StatusMsg = "Manual Feed"
    '                            Home.Ptr_Status = 1
    '                        End If
    'End If

    ''EXTENDED DETECTED ERROR STATE
    'If queryObj("ExtendedDetectedErrorState").ToString <> Nothing Then
    'If queryObj("ExtendedDetectedErrorState").ToString = "0" Then
    '.Ptr_ErrorMsg = "Unknown"
    'If Home.Ptr_Status <> 1 Then Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "1" Then
    '.Ptr_ErrorMsg = "Other"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "2" Then
    '.Ptr_ErrorMsg = "No Error"
    'If Home.Ptr_Status <> 1 Then Home.Ptr_Status = 0
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "3" Then
    '.Ptr_ErrorMsg = "Low Paper"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "4" Then
    '.Ptr_ErrorMsg = "No Paper"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "5" Then
    '.Ptr_ErrorMsg = "Low Toner"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "6" Then
    '.Ptr_ErrorMsg = "No Toner"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "7" Then
    '.Ptr_ErrorMsg = "Door Open"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "8" Then
    '.Ptr_ErrorMsg = "Jammed"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "9" Then
    '.Ptr_ErrorMsg = "Service Requested"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "10" Then
    '.Ptr_ErrorMsg = "Output Bin Full"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "11" Then
    '.Ptr_ErrorMsg = "Paper Problem"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "12" Then
    '.Ptr_ErrorMsg = "Cannot Print Page"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "13" Then
    '.Ptr_ErrorMsg = "User Intervention Required"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "14" Then
    '.Ptr_ErrorMsg = "Out of Memory"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("ExtendedDetectedErrorState").ToString = "15" Then
    '.Ptr_ErrorMsg = "Server Unknown"
    '                            Home.Ptr_Status = 2
    '                        End If
    'End If

    ''STATE
    'If queryObj("PrinterState").ToString <> Nothing Then
    'If queryObj("PrinterState").ToString = "4" Then
    '.Ptr_ErrorMsg = "Paper Jam"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("PrinterState").ToString = "5" Then
    '.Ptr_ErrorMsg = "Paper Out"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("PrinterState").ToString = "18" Then
    '.Ptr_ErrorMsg = "Toner Low"
    '                            Home.Ptr_Status = 2
    '                        ElseIf queryObj("PrinterState").ToString = "20" Then
    '.Ptr_ErrorMsg = "Page Punt"
    '                            Home.Ptr_Status = 2
    '                        End If
    'End If

End Class