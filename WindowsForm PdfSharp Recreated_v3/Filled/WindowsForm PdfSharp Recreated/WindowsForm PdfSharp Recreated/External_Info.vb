'--------Class Designed for Footer------
Public Class Footer
    Public Property FormNumber As String
    Public Property RevJDate As String
    Public Property CAFNumberDate As String
    'If no new values are entered, these are the values from the original document
    Sub New(ByVal Form As Integer, ByVal Rev As Integer, ByVal CAF As Integer)
        If Form = "" Then
            _FormNumber = "Form-085"
        Else
            _FormNumber = Form
        End If

        If Rev = "" Then
            _RevJDate = "Rev. J (01/29/2013)"
        Else
            _RevJDate = Rev
        End If

        If CAF = "" Then
            _CAFNumberDate = "CAF# 2012-09-08"
        Else
            _CAFNumberDate = CAF
        End If
    End Sub
    Sub New()
        _FormNumber = "Form-085"
        _RevJDate = "Rev. J (01/29/2013)"
        _CAFNumberDate = "CAF# 2012-09-08"
    End Sub
End Class

'-----------END FOOTER CLASS------------
'---------------------------------------
'---------------------------------------
'--------Class Designed for header------

Public Class Header
    Public Property Reference As String
    'If no new values are entered, these are the values from the original document
    Sub New(ByVal value As String)
        _Reference = value
    End Sub
    Sub New()
        _Reference = "Reference WI-2.11.2A"
    End Sub

End Class

'-----------END HEADER CLASS------------
'---------------------------------------
'---------------------------------------
'--------Class Designed for document info and its parts------

Public Class Document_info
    Public Non_Table_Info As Non_Table = New Non_Table
    Public Class Non_Table
        Public Property Sample_Amount As Integer = 3
        Public Property Comment As String = "This document is just a test."
    End Class

    Public Table1 As First_Table = New First_Table
    Public Class First_Table
        Public Property Supplier As String = "Kevin Díaz"
        Public Property Data_Recieved As String = "50"
        Public Property PO_Number As String = "50374"
        Public Property Reciever_Number As String = "20394"
        Public Property ManuExpi_Date As String = "8/25/2017"
        Public Property CareVendor_Number As String = "30453"
        Public Property Part_Number As String = "882932"
        Public Property RMS_Revision As String = "2"
        Public Property Quantity As String = "8"
        Public Property Containers_Recieved As String = "1"
        Public Property Unit_of_Measure As String = "Centimeters"
        Public Property BoxRoll_Number As String = "20134"
        Public Property Resin_Number As String = "991023"
        Public Property Inspection_Level As String = "Reduced  Normal  Tightened "
        Public Property Sampled_By As String = "Kevin Díaz"
        Public Property Data_Sampled As String = "Cheese"
    End Class
    '------

    Public Table2 As Second_Table = New Second_Table
    Public Class Second_Table
        Public Sampling_Level = New String() {"1", "2", "3", "4", "5"}
        Public AQL = New String() {"?", "?", "?", "?", "?"}
        Public Sampling_Size = New String() {"30", "30", "30", "30", "30"}
        Public Accept = New Integer() {11, 12, 13, 14, 15}
        Public Reject = New Integer() {19, 18, 17, 16, 15}
        Public Number_Of_Defects = New String() {"2", "3", "4", "5", "6"}
        Public Type_Of_Defect = New String() {"Wiring Error", "Wiring Error", "Wiring Error", "Wiring Error", "Wiring Error"}
        Public Pass_Fail = New String() {"Fail", "Fail", "Fail", "Fail", "Fail"}
    End Class
    '------

    Public Table3 As Third_Table = New Third_Table("K-Enterprise", "29930421", "7/29/2016", "8/29/2016")
    Public Class Third_Table
        Private ObjectData As New List(Of Third_Table)
        Public Property Is_Empty As Boolean = True
        Public Property Gauge_Employed As String
        Public Property Serial_Number As String
        Public Property Certification_Date As String
        Public Property Next_Certification_Date As String
        Sub New(ByVal Gauge As String, ByVal Serial As String, ByVal Cert As String, ByVal NextC As String)
            _Gauge_Employed = Gauge
            _Serial_Number = Serial
            _Certification_Date = Cert
            _Next_Certification_Date = NextC
            _Is_Empty = False
        End Sub
        Sub New()
            _Gauge_Employed = ""
            _Serial_Number = ""
            _Certification_Date = ""
            _Next_Certification_Date = ""
        End Sub
        Public Function AddData(ByVal Gauge As String, ByVal Serial As String, ByVal Cert As String, ByVal NextC As String)
            Dim Item_ As Third_Table = New Third_Table(Gauge, Serial, Cert, NextC)
            _Is_Empty = False
            ObjectData.Add(Item_)
        End Function
        Public Function GetData() As IEnumerable(Of Third_Table)
            Return ObjectData
        End Function
    End Class

    '------
    Public Table4 As Fourth_Table = New Fourth_Table("30453", "882932")
    Public Class Fourth_Table
        Private ObjectData As New List(Of Fourth_Table)
        Public Property Is_Empty As Boolean = True
        Public Property CareVendorLot_Number As String
        Public Property Part_Number As String
        Public Property Filler_1 As String = ""
        Public Property Filler_2 As String = ""
        Public Property Filler_3 As String = ""
        Public Property Filler_4 As String = ""
        Public Property Filler_5 As String = ""
        Public Property Filler_6 As String = ""
        Public Property Filler_7 As String = ""
        Public Property Filler_8 As String = ""
        Public Property Filler_9 As String = ""
        Sub New(ByVal careLot As String, ByVal part_ As String)
            _CareVendorLot_Number = careLot
            _Part_Number = part_
        End Sub
        Sub New()
            _CareVendorLot_Number = ""
            _Part_Number = ""
        End Sub
        Public Function AddData(ByVal data1 As String, ByVal data2 As String, ByVal data3 As String, ByVal data4 As String,
         ByVal data5 As String, ByVal data6 As String, ByVal data7 As String, ByVal data8 As String, ByVal data9 As String)

            _Is_Empty = False
            Dim Item_ As Fourth_Table = New Fourth_Table()
            Item_.Filler_1 = data1
            Item_.Filler_2 = data2
            Item_.Filler_3 = data3
            Item_.Filler_4 = data4
            Item_.Filler_5 = data5
            Item_.Filler_6 = data6
            Item_.Filler_7 = data7
            Item_.Filler_8 = data8
            Item_.Filler_9 = data9
            ObjectData.Add(Item_)
        End Function
        Public Function GetData() As IEnumerable(Of Fourth_Table)
            Return ObjectData
        End Function
    End Class

    '------
    Public Table5 As Fifth_Table = New Fifth_Table("30453", "882932")
    Public Class Fifth_Table
        Private ObjectQuestion(0) As Boolean
        Private ObjectAnswer(199) As String
        Private Property Count As Integer = -1
        Public Property CareVendorLot_Number As String
        Public Property Part_Number As String
        Sub New(ByVal careLot As String, ByVal part_ As String)
            _CareVendorLot_Number = careLot
            _Part_Number = part_
        End Sub
        Sub New()
            _CareVendorLot_Number = ""
            _Part_Number = ""
        End Sub
        Public Function AddData(ByVal BValue As Boolean)
            If Count = -1 Then
                _Count += 1
                ObjectQuestion(Count) = BValue
            Else
                ReDim Preserve ObjectQuestion(Count)
                ObjectQuestion(Count) = BValue
            End If
            _Count += 1
        End Function
        Public Function GetData()
            For i As Integer = 0 To (Count - 1)
                If ObjectQuestion(i) = True Then
                    ObjectAnswer(i) = "Pass"
                Else
                    ObjectAnswer(i) = "Fail"
                End If
            Next
            For u As Integer = Count To 199
                If 0 > Count Then
                    Count += 1
                Else
                    ObjectAnswer(u) = " pass  fail"
                End If
            Next
            Return ObjectAnswer
        End Function
    End Class
End Class