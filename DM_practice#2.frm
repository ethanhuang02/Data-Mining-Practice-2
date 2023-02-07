VERSION 5.00
Begin VB.Form R76101120 
   Caption         =   "R76101120_HW2"
   ClientHeight    =   11415
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   16725
   LinkTopic       =   "Form2"
   ScaleHeight     =   11415
   ScaleWidth      =   16725
   Begin VB.ListBox List6 
      Height          =   9780
      Left            =   6840
      TabIndex        =   8
      Top             =   1320
      Width           =   3000
   End
   Begin VB.ListBox List5 
      Height          =   9780
      Left            =   13560
      TabIndex        =   7
      Top             =   1320
      Width           =   3000
   End
   Begin VB.ListBox List4 
      Height          =   9780
      Left            =   10200
      TabIndex        =   6
      Top             =   1320
      Width           =   3000
   End
   Begin VB.ListBox List2 
      Height          =   9780
      Left            =   3480
      TabIndex        =   5
      Top             =   1320
      Width           =   3000
   End
   Begin VB.ListBox List1 
      Height          =   9780
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   3000
   End
   Begin VB.TextBox infile 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      TabIndex        =   1
      Text            =   "Breast.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Execute"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label6 
      Caption         =   "Backward"
      Height          =   495
      Left            =   13680
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Forward"
      Height          =   495
      Left            =   10320
      TabIndex        =   11
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Entropy Based"
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Equal Frequency"
      Height          =   495
      Left            =   3600
      TabIndex        =   9
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Equal Width"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Input file :"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "R76101120"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim in_file As String
Dim att(10) As Double, atts(10) As String
Dim data(106, 10) As Double

Dim EWData(106, 10) As Double 'Equal width離散化資料集
Dim EFData(106, 10) As Double 'Equal frequency暫存離散化資料集
Dim EBData(106, 10) As Double 'Entropy based暫存離散化資料集
Dim SortClassValue(106, 10) As Double '存各attribute排序對應的的classvalue
Dim SortAttribute(106, 10) As Double

Dim tempSort(106) As Double '暫存排序過的attribute
Dim tempClass(106) As Double
Dim label(106) As Double
Dim tempCount(10) As Double '暫存各attribute value數量
Dim tempSplitPoint(9) As Double '暫存分割點
Dim splitPointArray(9, 3) As Double
Dim gloryn As Integer
Dim EWUArray(10, 10) As Double
Dim EFUArray(10, 10) As Double
Dim EBUArray(10, 10) As Double

Dim chosen(9) As Integer '紀錄attribute是否被選進subset
Dim attributeValue, counta, index, att1 As Integer
Dim numerator, denominator, tempU, Max, min, upperBound, lowerBound As Double

Public DATASIZE As Integer



Private Sub Partition_click()
    Dim n As Integer
    Dim i As Integer
    Dim j As Integer
    List1.Clear
    List2.Clear
    'List3.Clear
    List4.Clear
    List5.Clear
    List6.Clear
    DATASIZE = 106
    'check whether the file name is empty
    If infile.Text = "" Then
        MsgBox "Please input the file names!", , "File Name"
        infile.SetFocus
    Else
        in_file = App.Path & "\" & infile.Text
        'check whether the data file exists
        If Dir(in_file) = "" Then
            MsgBox "Input file not found!", , "File Name"
            infile.SetFocus
        Else
            Open in_file For Input As #1
            n = 1
            Do While Not EOF(1)
                Input #1, att(1), att(2), att(3), att(4), att(5), att(6), att(7), att(8), att(9), atts(10)
                If atts(10) = "car" Then
                    att(10) = 0
                ElseIf atts(10) = "fad" Then
                    att(10) = 1
                ElseIf atts(10) = "mas" Then
                    att(10) = 2
                ElseIf atts(10) = "gla" Then
                    att(10) = 3
                ElseIf atts(10) = "con" Then
                    att(10) = 4
                Else
                    att(10) = 5
                End If
                
                For i = 1 To 10
                    data(n, i) = att(i)
                Next
                n = n + 1
                'List1.AddItem att(1) & " " & att(2) & " " & att(3) & " " & att(4) & " " & att(5) & " " & att(6) & " " & att(7) & " " & att(8) & " " & att(9) & " " & att(10)
            Loop
            Close #1
        End If
    End If
    
    'copy data
    For i = 1 To DATASIZE
        For j = 1 To 10
            EWData(i, j) = data(i, j)
            EFData(i, j) = data(i, j)
            EBData(i, j) = data(i, j)
        Next j
    Next i
    
    'discretization
    For i = 1 To 9
        EqualWidth i
        EqualFrequency i
        gloryn = 4
        'List3.AddItem "Attribute" & i
        EntorpyBased i
        'List3.AddItem "-------------------------------------------------------"
    Next i
    For i = 1 To 9
        List6.AddItem "Attribute" & i
        If splitPointArray(i, 1) = 0 And splitPointArray(i, 2) = 0 And splitPointArray(i, 3) = 0 Then
            List6.AddItem "(" & SortAttribute(1, i) & "," & SortAttribute(106, i) & "]"
        ElseIf splitPointArray(i, 1) = 0 And splitPointArray(i, 2) = 0 Then
            List6.AddItem "(" & SortAttribute(1, i) & "," & splitPointArray(i, 3) & "]"
            List6.AddItem "(" & splitPointArray(i, 3) & "," & SortAttribute(106, i) & "]"
        ElseIf splitPointArray(i, 1) = 0 Then
            List6.AddItem "(" & SortAttribute(1, i) & "," & splitPointArray(i, 2) & "]"
            List6.AddItem "(" & splitPointArray(i, 2) & "," & splitPointArray(i, 3) & "]"
            List6.AddItem "(" & splitPointArray(i, 3) & "," & SortAttribute(106, i) & "]"
        Else
            List6.AddItem "(" & SortAttribute(1, i) & "," & splitPointArray(i, 1) & "]"
            List6.AddItem "(" & splitPointArray(i, 1) & "," & splitPointArray(i, 2) & "]"
            List6.AddItem "(" & splitPointArray(i, 2) & "," & splitPointArray(i, 3) & "]"
            List6.AddItem "(" & splitPointArray(i, 3) & "," & SortAttribute(106, i) & "]"
        End If
        List6.AddItem "-------------------------------------------------------"
    Next
    '計算U(XY)
    For i = 1 To 10
        For j = 1 To 10
            U EWUArray, EWData, i, j
            'List3.AddItem (m & n & "  |  " & Math.Round(EWUArray(m, n), 4))
            U EFUArray, EFData, i, j
            U EBUArray, EBData, i, j
        Next j
    Next i
    'For i = 1 To 106
        'For j = 1 To 10
            'List5.AddItem EFData(i, j)
        'Next j
    'Next i
    
    List4.Clear
    
    'Forward
    List4.AddItem "【Equal-width】"
    Forward (EWUArray)
    
    List4.AddItem "【Equal-frequency】"
    Forward (EFUArray)
    
    List4.AddItem "【Entropy-based】"
    Forward (EBUArray)
    
    List5.Clear
    
    'Backward
    List5.AddItem "【Equal-width】"
    Backward (EWUArray)
    
    List5.AddItem "【Equal-frequency】"
    Backward (EFUArray)
    
    List5.AddItem "【Entropy-based】"
    Backward (EBUArray)
    
End Sub
Private Sub EqualWidth(att1 As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim width As Double
    Max = -1000000
    min = 1000000
    
    For i = 1 To DATASIZE
        If EWData(i, att1) > Max Then
            Max = EWData(i, att1)
        End If
        If EWData(i, att1) < min Then
            min = EWData(i, att1)
        End If
    Next i

    width = (Max - min) / 10
    List1.AddItem ("Attribute" & att1)
    
    For i = 0 To 9
        upperBound = min + width * i
        lowerBound = min + width * (i + 1)
        tempSplitPoint(i) = lowerBound
        List1.AddItem ("( " & upperBound & " , " & lowerBound & " ]")
    Next i
    List1.AddItem ("-------------------------------------------------------------")
   
    For j = 1 To DATASIZE
        For k = 9 To 1 Step -1
            If EWData(j, att1) > tempSplitPoint(k - 1) And EWData(j, att1) <= tempSplitPoint(k) Then
                EWData(j, att1) = k
                Exit For
            ElseIf EWData(j, att1) <= tempSplitPoint(0) Then
                EWData(j, att1) = 0
            End If
        Next k
    Next j
    'EqualWidth Max, min, att1

End Sub

Private Function EqualFrequency(att1 As Integer)
    Dim i As Integer
    Dim j As Integer
    Dim temp As Double
    Dim discrete As Double
    Dim freq(9) As Integer
    For i = 1 To DATASIZE
        tempSort(i) = data(i, att1)
        label(i) = i
    Next i

    For i = 1 To DATASIZE
        For j = 1 To DATASIZE - 1
            If tempSort(j) > tempSort(j + 1) Then
                temp = tempSort(j)
                tempSort(j) = tempSort(j + 1)
                tempSort(j + 1) = temp
                
                '紀錄label
                temp = label(j)
                label(j) = label(j + 1)
                label(j + 1) = temp
            End If
        Next j
    Next i
    'List5.AddItem tempSort(40) & " " & tempSort(41)
    'List5.AddItem tempSort(51) & " " & tempSort(52)
    
'    List3.AddItem ("Attribute" & att1 & "  |  " & tempSort(1) & "  |  " & tempSort(DATASIZE))
    'EqualFrequency tempSort, att1, label
    freq(0) = 10
    freq(1) = 10
    freq(2) = 10
    freq(3) = 10
    freq(4) = 11
    freq(5) = 11
    freq(6) = 11
    freq(7) = 11
    freq(8) = 11
    freq(9) = 11
    'freq = DATASIZE / 10
    
    List2.AddItem ("Attribute" & att1)
    
    'print interval
    For i = 0 To 3
        If 10 * i = 0 Then
            upperBound = tempSort(10 * i + 1)
            lowerBound = (tempSort(10 * (i + 1)) + tempSort(10 * (i + 1) + 1)) / 2
        'ElseIf freq(i) * (i + 1) = DATASIZE Then
        '    upperBound = (tempSort(freq(i) * i) + tempSort(freq(i) * i + 1)) / 2
        '    lowerBound = tempSort(freq(i) * (i + 1))
        Else
            upperBound = (tempSort(10 * i) + tempSort(10 * i + 1)) / 2
            lowerBound = (tempSort(10 * (i + 1)) + tempSort(10 * (i + 1) + 1)) / 2
        End If
        tempSplitPoint(i) = lowerBound
        List2.AddItem ("( " & upperBound & " , " & lowerBound & " ]")
    Next i
    For i = 4 To 9
        If 11 * i = 0 Then
            upperBound = tempSort(11 * i + 1)
            lowerBound = (tempSort(11 * (i + 1)) + tempSort(11 * (i + 1) + 1)) / 2
        ElseIf 11 * (i + 1) - 4 = DATASIZE Then
            upperBound = (tempSort(11 * i - 4) + tempSort(11 * i + 1 - 4)) / 2
            lowerBound = tempSort(11 * (i + 1) - 4)
        Else
            upperBound = (tempSort(11 * i - 4) + tempSort(11 * i + 1 - 4)) / 2
            lowerBound = (tempSort(11 * (i + 1) - 4) + tempSort(11 * (i + 1) + 1 - 4)) / 2
        End If
        tempSplitPoint(i) = lowerBound
        List2.AddItem ("( " & upperBound & " , " & lowerBound & " ]")
    Next i
    List2.AddItem ("----------------------------------------------------------------")
    
    'discretization
    discrete = 0
    For i = 0 To 3
        For j = 1 To 10
            tempSort(j + 10 * i) = discrete
        Next j
        discrete = discrete + 1
    Next i
    For i = 4 To 9
        For j = 1 To 11
            tempSort(j + 11 * i - 4) = discrete
        Next j
        discrete = discrete + 1
    Next i


    'sort label
    For i = 1 To DATASIZE
        For j = 1 To DATASIZE - 1
            If label(j) > label(j + 1) Then
                temp = label(j)
                label(j) = label(j + 1)
                label(j + 1) = temp
                               
                temp = tempSort(j)
                tempSort(j) = tempSort(j + 1)
                tempSort(j + 1) = temp
                
            End If
        Next j
    Next i
    
    For i = 1 To DATASIZE
        EFData(i, att1) = tempSort(i)
'        List3.AddItem ("col" & att1 & "i" & i & " | " & tempSort(i))
    Next i

End Function

Private Sub EntorpyBased(att1 As Integer)

    Dim i As Integer
    Dim j As Integer
    Dim temp1 As Double
    Dim temp2 As Double
    
    For i = 1 To DATASIZE
        tempSort(i) = data(i, att1)
        tempClass(i) = data(i, 10)
    Next i

    For i = 1 To DATASIZE
        For j = 1 To DATASIZE - 1
            If tempSort(j) > tempSort(j + 1) Then
                temp1 = tempSort(j)
                tempSort(j) = tempSort(j + 1)
                tempSort(j + 1) = temp1
                
                '排序class值
                temp2 = tempClass(j)
                tempClass(j) = tempClass(j + 1)
                tempClass(j + 1) = temp2
            End If
        Next j
    Next i
    
    For i = 1 To DATASIZE
        SortClassValue(i, att1) = tempClass(i)
        SortAttribute(i, att1) = tempSort(i)
        'List4.AddItem ("col" & att1 & " i " & i & " | " & EBData(i, 10))
        'List3.AddItem ("col" & att1 & " i " & i & " | " & SortClassValue(i, att1))
    Next i
    
    Entropy att1, tempClass
    
End Sub

Private Sub Entropy(att1 As Integer, temparray() As Double)

    Dim i, j, k, l, r As Integer
    Dim minEnt, cutpoint, TEnt, size, k0, k1, k2 As Integer
    Dim Tc0, Tc1, Tc2, Tc3, Tc4, Tc5 As Integer
    Dim left() As Double
    Dim right() As Double
    Dim gainValue, threshold, prob, splitPoint, LEnt, REnt, Ent, delta, reject, MinSplitPoint, splitIndex, splitLen, lInd, RInd As Double
    Dim Lc0, Lc1, Lc2, Lc3, Lc4, Lc5 As Double
    Dim Rc0, Rc1, Rc2, Rc3, Rc4, Rc5 As Double
    
    minEnt = 99999999 '紀錄最小的entropy
    Tc0 = 0
    Tc1 = 0
    Tc2 = 0
    Tc3 = 0
    Tc4 = 0
    Tc5 = 0
    TEnt = 0
    size = UBound(temparray) - LBound(temparray)
    cutpoint = 0
    'List4.AddItem att1
    List4.AddItem size
    '計算切割點數量
    If size > 1 Then
        For i = 1 To size - 1
            If temparray(i) <> temparray(i + 1) Then
                cutpoint = cutpoint + 1
            End If
        Next i
    End If
    
    'List3.AddItem att1 & "   " & cutpoint
    
    '不需再切則停止
    If cutpoint = 0 Then
        'List3.AddItem "end"
    Else
        'List3.AddItem "continue"
        For i = 1 To size
            If temparray(i) = 0 Then
                Tc0 = Tc0 + 1
            ElseIf temparray(i) = 1 Then
                Tc1 = Tc1 + 1
            ElseIf temparray(i) = 2 Then
                Tc2 = Tc2 + 1
            ElseIf temparray(i) = 3 Then
                Tc3 = Tc3 + 1
            ElseIf temparray(i) = 4 Then
                Tc4 = Tc4 + 1
            Else
                Tc5 = Tc5 + 1
            End If
        Next i
     
        '計算Ent(S)
        prob = Tc0 / (Tc0 + Tc1 + Tc2 + Tc3 + Tc4 + Tc5)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tc1 / (Tc0 + Tc1 + Tc2 + Tc3 + Tc4 + Tc5)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tc2 / (Tc0 + Tc1 + Tc2 + Tc3 + Tc4 + Tc5)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tc3 / (Tc0 + Tc1 + Tc2 + Tc3 + Tc4 + Tc5)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tc4 / (Tc0 + Tc1 + Tc2 + Tc3 + Tc4 + Tc5)
        TEnt = TEnt - prob * Log2(prob)
        prob = Tc5 / (Tc0 + Tc1 + Tc2 + Tc3 + Tc4 + Tc5)
        TEnt = TEnt - prob * Log2(prob)
        
        'List3.AddItem TEnt
        
        '計算左右個數
        For j = 1 To size - 1
            If temparray(j) <> temparray(j + 1) Then
                splitPoint = (SortAttribute(j, att1) + SortAttribute(j + 1, att1)) / 2
                Lc0 = 0
                Lc1 = 0
                Lc2 = 0
                Lc3 = 0
                Lc4 = 0
                Lc5 = 0
                Rc0 = 0
                Rc1 = 0
                Rc2 = 0
                Rc3 = 0
                Rc4 = 0
                Rc5 = 0
                
                For k = 1 To size
                    If k < j + 1 Then
                        If temparray(k) = 0 Then
                            Lc0 = Lc0 + 1
                        ElseIf temparray(k) = 1 Then
                            Lc1 = Lc1 + 1
                        ElseIf temparray(k) = 2 Then
                            Lc2 = Lc2 + 1
                        ElseIf temparray(k) = 3 Then
                            Lc3 = Lc3 + 1
                        ElseIf temparray(k) = 4 Then
                            Lc4 = Lc4 + 1
                        Else
                            Lc5 = Lc5 + 1
                        End If
                    Else
                        If temparray(k) = 0 Then
                            Rc0 = Rc0 + 1
                        ElseIf temparray(k) = 1 Then
                            Rc1 = Rc1 + 1
                        ElseIf temparray(k) = 2 Then
                            Rc2 = Rc2 + 1
                        ElseIf temparray(k) = 3 Then
                            Rc3 = Rc3 + 1
                        ElseIf temparray(k) = 4 Then
                            Rc4 = Rc4 + 1
                        Else
                            Rc5 = Rc5 + 1
                        End If
                    End If
                Next k
                
                LEnt = 0
                REnt = 0
                
                '計算Ent(s1)
                prob = Lc0 / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5)
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                
                prob = Lc1 / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5)
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                
                prob = Lc2 / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5)
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                
                prob = Lc3 / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5)
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                
                prob = Lc4 / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5)
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                
                prob = Lc5 / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5)
                If prob <> 0 Then
                    LEnt = LEnt - prob * Log2(prob)
                End If
                
                '計算Ent(s2)
                prob = Rc0 / (Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5)
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                
                prob = Rc1 / (Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5)
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                
                prob = Rc2 / (Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5)
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                
                prob = Rc3 / (Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5)
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                
                prob = Rc4 / (Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5)
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                
                prob = Rc5 / (Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5)
                If prob <> 0 Then
                    REnt = REnt - prob * Log2(prob)
                End If
                
                Ent = LEnt * (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5) / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5 + Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5) + REnt * (Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5) / (Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5 + Rc0 + Rc1 + Rc2 + Rc3 + Rc4 + Rc5)
                'List6.AddItem ("col" & att1 & " | " & Ent)
            
            
                '停止條件
                k0 = 6
                k1 = 6
                k2 = 6
                delta = Log2(3 ^ k0 - 2) - k0 * TEnt + k1 * LEnt + k2 * REnt
                reject = TEnt - Ent - (Log2(size - 1) + delta) / size
                'List6.AddItem "reject" & reject

                If minEnt >= Ent And reject > 0 Then
                    minEnt = Ent
                    MinSplitPoint = splitPoint
                    splitIndex = Lc0 + Lc1 + Lc2 + Lc3 + Lc4 + Lc5 + 1
                End If
            End If
        Next j
        
        'List6.AddItem minEnt
        If minEnt <> 99999999 Then
            gloryn = gloryn - 1
            splitLen = splitLen + 1
            splitPointArray(att1, gloryn) = MinSplitPoint
            
'            List4.AddItem att1 & " | " & splitPointArray(splitLen)
'            List3.AddItem "minEnt " & minEnt
'            List3.AddItem "TEnt " & TEnt
            lInd = splitIndex - 1
            RInd = size - lInd
            
            ReDim left(lInd) As Double
            ReDim right(RInd) As Double
            
            l = 1
            r = 1
            For i = 1 To size
                If i < splitIndex Then
                    left(l) = temparray(i)
                    l = l + 1
                Else
                    right(r) = temparray(i)
                    r = r + 1
                End If
            Next i
            
            'print Entropy interval
            
            'List3.AddItem "SplitPoint " & MinSplitPoint
            'List3.AddItem "(" & splitPointArray(splitLen - 1) & "," & splitPointArray(splitLen)
            

            For i = 1 To DATASIZE
                If EBData(i, att1) <= MinSplitPoint Then
                    EBData(i, att1) = 0
                ElseIf EBData(i, att1) > MinSplitPoint Then
                    EBData(i, att1) = 1
                End If
            Next i

            Entropy att1, left
            Entropy att1, right
            
        End If
    End If
End Sub

Private Function H(tempData, att1)
    Dim i, j As Integer
    Dim Hx, p As Double
    counta = 0
    '算categorical各值的數量
    For i = 0 To 9
        For j = 1 To DATASIZE
            If tempData(j, att1) = i Then
                counta = counta + 1
            End If
        Next j
        tempCount(i) = counta
        counta = 0
    Next i

    '計算H(x)
    Hx = 0
    For j = 0 To 9
        p = tempCount(j) / DATASIZE
        Hx = Hx + -p * Log2(p)
    Next j
    
    H = Hx

End Function

Private Function Hxy(tempData, att1, att2)
    Dim x, y, i As Integer
    Dim Pxy, tempHxy As Double
    counta = 0
    For x = 0 To 9
        For y = 0 To 9
            For i = 1 To DATASIZE
                If tempData(i, att1) = x And tempData(i, att2) = y Then
                    counta = counta + 1
                End If
            Next i
            Pxy = counta / DATASIZE
            tempHxy = tempHxy + -Pxy * Log2(Pxy)
            counta = 0
        Next y
    Next x
    Hxy = tempHxy

End Function

Private Function U(tempU, tempData, att1, att2)
    Dim uvalue As Double
    If (H(tempData, att1) + H(tempData, att2) = 0) Then
        tempU(att1, att2) = 1
        tempU(att2, att1) = 1
    Else
    'If (H(tempData, att1) + H(tempData, att2) <> 0) Then
        uvalue = 2 * ((H(tempData, att1) + H(tempData, att2) - Hxy(tempData, att1, att2)) / (H(tempData, att1) + H(tempData, att2)))
    
        tempU(att1, att2) = uvalue
        tempU(att2, att1) = uvalue
    End If
End Function

Private Function Goodness(tempU)
    Dim i, j As Integer
    numerator = 0
    denominator = 0
    
    '計算分子
    For i = 1 To 9
        If chosen(i) = 1 Then '被選擇
            numerator = numerator + tempU(i, 10)
        End If
    Next i
'    List5.AddItem numerator
    '計算分母
    For i = 1 To 9
        For j = 1 To 9
            If chosen(i) = 1 And chosen(j) = 1 Then
                denominator = denominator + tempU(i, j)
            End If
        Next j
    Next i
    denominator = Sqr(denominator)
    'List5.AddItem denominator
    
    '計算goodness
    If denominator <> 0 Then
    Goodness = numerator / denominator
    'List5.AddItem Goodness
    End If
    
End Function

Private Sub Forward(tempU)
    Dim maxGoodness, tempGoodness As Double
    Dim i, j, n, m, r, index As Integer
    maxGoodness = 0
    
    For i = 0 To 9
        chosen(i) = 0
    Next i

    For n = 1 To 9
        index = -1
        For m = 1 To 9
            If chosen(m) <> 1 Then
                chosen(m) = 1
                tempGoodness = Goodness(tempU)
                
                If tempGoodness > maxGoodness Then
                    maxGoodness = tempGoodness
                    index = m
                End If
                chosen(m) = 0
                tempGoodness = 0
            End If
        Next m
        If index = -1 Then Exit For
        chosen(index) = 1
        r = index
        List4.AddItem ("Attribute chosen：A" & r)
        List4.AddItem ("Goodness：" & Math.Round(maxGoodness, 5))
'        List4.AddItem ("-------------------")
    Next n
    List4.AddItem ("     ")
    List4.AddItem ("Attribute subset :")
    For i = 1 To 9
        If chosen(i) = 1 Then
            List4.AddItem ("A" & i)
        End If
    Next i

    List4.AddItem ("-------------------------------------------------")

End Sub

Private Sub Backward(tempU)
    Dim maxGoodness, tempGoodness As Double
    Dim i, j, n, m, r, index As Integer
    '假設全選
    For i = 0 To 9
        chosen(i) = 1
    Next i
    
    maxGoodness = 0
    For n = 1 To 9
        index = -1
        For m = 1 To 9
            If chosen(m) = 1 Then
                chosen(m) = 0
                tempGoodness = Goodness(tempU)
                If tempGoodness > maxGoodness Then
                    maxGoodness = tempGoodness
                    index = m
                End If
                chosen(m) = 1
            End If
        Next m
        If index = -1 Then Exit For 'goodness下降則停止
        chosen(index) = 0
        r = index
        List5.AddItem ("Attribute removed：A" & r)
        List5.AddItem ("Goodness：" & Math.Round(maxGoodness, 5))
'        List5.AddItem ("-------------------")
    Next n
    List5.AddItem ("     ")
    List5.AddItem ("Attribute subset :")
    For i = 1 To 9
        If chosen(i) = 1 Then
            List5.AddItem ("A" & i)
        End If
    Next i
    
    List5.AddItem ("------------------------------------------------------------")
    
End Sub

Static Function Log_2(x) As Single
    If (x = 0) Then
        Log_2 = 0
    Else
        Log_2 = Log(x) / Log(2)
    End If
End Function

Static Function Log2(x) As Double
    If x <> 0 Then
        Log2 = Log(x) / Log(2#)
    ElseIf x = 0 Then
        Log2 = 0
    End If
End Function
