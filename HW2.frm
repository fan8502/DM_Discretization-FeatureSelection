VERSION 5.00
Begin VB.Form Partition 
   Caption         =   "Partition"
   ClientHeight    =   9060
   ClientLeft      =   168
   ClientTop       =   456
   ClientWidth     =   10872
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   10872
   Begin VB.CommandButton backward 
      Caption         =   "backward"
      Height          =   492
      Left            =   8280
      TabIndex        =   9
      Top             =   3240
      Width           =   2172
   End
   Begin VB.CommandButton entropy_base 
      Caption         =   "Entropy-Base discretization"
      Height          =   732
      Left            =   9240
      TabIndex        =   8
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton forward 
      Caption         =   "forward"
      Height          =   492
      Left            =   5880
      TabIndex        =   7
      Top             =   3240
      Width           =   2172
   End
   Begin VB.CommandButton equal_frequency 
      Caption         =   "Equal-Freq discretization"
      Height          =   732
      Left            =   7560
      TabIndex        =   6
      Top             =   1800
      Width           =   1212
   End
   Begin VB.CommandButton equal_width 
      Caption         =   "Equal-Width discretization"
      Height          =   732
      Left            =   5880
      TabIndex        =   5
      Top             =   1800
      Width           =   1212
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8448
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   5412
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
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Text            =   "yeast.txt"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton Partition 
      Caption         =   "Read"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5160
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "請先按Read讀檔"
      Height          =   252
      Left            =   6000
      TabIndex        =   11
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label Label2 
      Caption         =   "若要選擇不同離散化方式請將視窗關閉重新RUN"
      Height          =   252
      Left            =   6000
      TabIndex        =   10
      Top             =   1440
      Width           =   3972
   End
   Begin VB.Label Label5 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   3360
      TabIndex        =   3
      Top             =   1080
      Width           =   852
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
Attribute VB_Name = "Partition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim in_file As String, out_file As String, nstr As String
Dim attribute_name(9) As String '存八個屬性的名字，從index=1開始
Dim i As Integer, j As Integer, k As Integer
Dim a As Integer, b As Integer
Dim fileCount As Integer
Dim fileArray(1484, 10) As Variant
Dim width_Array(1484, 10) As Variant '用來存最後離散化結果
Dim freq_Array(1484, 10) As Variant '用來存最後離散化結果
Dim freq_temp_Array(1484, 10) As Variant
Dim interval_pro(10, 9) As Double '九個屬性的共10個interval的單一P值，從index=0開始，pro(0)存的是interval"1"
Dim HValue(10) As Double '用來存九個H(attribute)值，從index=1開始(包含class)
Dim Pab_Array(11, 11, 10, 10) As Double
Dim Hab_Value(10, 10) As Double
Dim Uab_Value(10, 10) As Double

Private Sub equal_frequency_Click()
    List1.Clear
    equal_width.Enabled = False
    entropy_base.Enabled = False
    
    Dim temp As String
    Dim freq_range(10) As Double
    Dim splitting As Integer
    splitting = 148 '1483/10
    
    For k = 1 To 8
        List1.AddItem "Attribute " & k
        For i = 0 To 1483
            For j = i To 1483 '用原本的陣列做排序(因為排序會覆蓋掉原本陣列順序所以不能再做離散化)
                If CDbl(fileArray(i, k)) > CDbl(fileArray(j, k)) Then '由小到大排序
                    temp = fileArray(i, k)
                    fileArray(i, k) = fileArray(j, k)
                    fileArray(j, k) = temp
                End If
            Next j
        Next i
        For i = 1 To 9
            freq_range(0) = fileArray(0, k)
            freq_range(10) = fileArray(1483, k)
            freq_range(i) = (CDbl(fileArray(splitting * i, k)) + CDbl(fileArray(splitting * i - 1, k))) / 2
            List1.AddItem "splitting point " & i & " : " & freq_range(i)
        Next i
        For i = 0 To 8
        '把資料依照分割點離散化成1-10類
            For j = 0 To 1483 '用原本的temp陣列比大小，用新的陣列存離散化結果
                If CDbl(freq_temp_Array(j, k)) >= freq_range(i) And CDbl(freq_temp_Array(j, k)) < freq_range(i + 1) Then
                    freq_Array(j, k) = i + 1
                ElseIf CDbl(freq_temp_Array(j, k)) >= freq_range(9) And CDbl(freq_temp_Array(j, k)) <= freq_range(10) Then
                    freq_Array(j, k) = 10
                End If
            Next j
        Next i
    Next k
    For k = 1 To 9
        For j = 0 To 1483
            freq_Array(j, k) = Replace(freq_Array(j, k), " ", "")
        Next j
    Next k
    
    PFunction (freq_Array)
    P_abFunction (freq_Array)
    
    For k = 1 To 9 '計算每個屬性H的值，從屬性(1)的H(1)開始計算
        HValue(k) = HFunction(interval_pro, k)
    Next k
    
    For a = 1 To 9 '計算每2個屬性H的值，從H(1,1)
        For b = 1 To 9
            Hab_Value(a, b) = Hab_Function(Pab_Array, a, b)
        Next b
    Next a
    
    For a = 1 To 9 '計算每2個屬性U的值，從U(1,1)
        For b = 1 To 9
            Uab_Value(a, b) = Uab_Function(a, b)
        Next b
    Next a
    forward.Enabled = True
    backward.Enabled = True
End Sub

Private Sub equal_width_Click()
    List1.Clear
    equal_frequency.Enabled = False
    entropy_base.Enabled = False
    
    Dim width_max(8) As Double '八個屬性分別的max
    Dim width_min(8) As Double
    Dim tempmax As Double
    Dim tempmin As Double
    Dim width_w(8) As Double
    Dim width_range(10) As Double '用來紀錄interval的各間隔值 其中0和10分別為min.max
    
    For i = 1 To 8 '從1開始是因為資料的第一列是index 所以陣列的0沒有用處
        tempmax = 0
        tempmin = 1
        For j = 0 To 1483
            If CDbl(fileArray(j, i)) > tempmax Then
                tempmax = fileArray(j, i)
                width_max(i) = tempmax
            End If
            If CDbl(fileArray(j, i)) < tempmin Then
                tempmin = fileArray(j, i)
                width_min(i) = tempmin
            End If
        Next j
        'List1.AddItem width_max(i) & vbTab & width_min(i)
        width_w(i) = (width_max(i) - width_min(i)) / 10 'bin=10
    Next i
     
    For k = 1 To 8
        List1.AddItem "Attribute " & k
        For i = 1 To 9
            width_range(0) = width_min(k)
            width_range(10) = width_max(k)
            width_range(i) = width_min(k) + width_w(k) * i
        Next i
        For i = 0 To 8
            'List1.AddItem "Interval " & i + 1 & " : " & width_range(i) & " ~ " & width_range(i + 1)
            List1.AddItem "splitting point " & i + 1 & " : " & width_range(i + 1)
            '把資料依照分割點離散化成1-10類
            For j = 0 To 1483 '用原本的陣列比大小，用新的陣列存離散化結果
                If CDbl(fileArray(j, k)) >= width_range(i) And CDbl(fileArray(j, k)) < width_range(i + 1) Then
                    width_Array(j, k) = i + 1
                ElseIf CDbl(fileArray(j, k)) >= width_range(9) And CDbl(fileArray(j, k)) <= width_range(10) Then
                    width_Array(j, k) = 10
                End If
            Next j
        Next i
    Next k
    For k = 1 To 9
        For j = 0 To 1483
            width_Array(j, k) = Replace(width_Array(j, k), " ", "")
        Next j
    Next k
  
    PFunction (width_Array)
    P_abFunction (width_Array)
    For k = 1 To 9 '計算每個屬性H的值，從屬性(1)的H(1)開始計算
        HValue(k) = HFunction(interval_pro, k)
    Next k
    
    For a = 1 To 9 '計算每2個屬性H的值，從H(1,1)
        For b = 1 To 9
            Hab_Value(a, b) = Hab_Function(Pab_Array, a, b)
        Next b
    Next a
    
    For a = 1 To 9 '計算每2個屬性U的值，從U(1,1)
        For b = 1 To 9
            Uab_Value(a, b) = Uab_Function(a, b)
        Next b
    Next a
    forward.Enabled = True
    backward.Enabled = True
End Sub
Public Function PFunction(pro_Array)
    Dim pro_count As Double
    For k = 1 To 9 '跑attribute
        For i = 0 To 9 '跑interval，pro(0)存的是interval"1"的次數
            pro_count = 0  '每次跑完一個interval後，count歸零
            For j = 0 To 1483 '跑data
                If pro_Array(j, k) = i + 1 Then
                    pro_count = pro_count + 1
                    interval_pro(i, k) = pro_count / fileCount '計算出各atrribute的每個interval的P值
                End If
            Next j
        Next i
    Next k
End Function
Public Function P_abFunction(disc_Array)
    '計算Pab次數
    For k = 0 To 1483
        For a = 1 To 9
            For b = 1 To 9
                For i = 1 To 10
                    For j = 1 To 10
                        If disc_Array(k, a) = i And disc_Array(k, b) = j Then
                            Pab_Array(i, j, a, b) = Pab_Array(i, j, a, b) + 1
                        End If
                    Next j
                Next i
            Next b
        Next a
    Next k
    '計算Pab值
    For a = 1 To 9
        For b = 1 To 9
            For i = 1 To 10
                For j = 1 To 10
                    Pab_Array(i, j, a, b) = Pab_Array(i, j, a, b) / fileCount
                Next j
            Next i
        Next b
    Next a
End Function
Public Function HFunction(attribute_i_pro, ByVal att_th As Integer) As Double
    Dim cal_HValue As Double
    Dim attribute_pro As Double
    
    For i = 0 To 9
        attribute_pro = attribute_i_pro(i, att_th)
        cal_HValue = cal_HValue - attribute_pro * log2(attribute_pro)
    Next i
    HFunction = cal_HValue
End Function
Public Function Hab_Function(attribute_ab_pro, ByVal att_a As Integer, ByVal att_b As Integer) As Double
    Dim cal_HabValue As Double
    Dim ab_pro As Double
    
    For i = 1 To 10
        For j = 1 To 10
            ab_pro = attribute_ab_pro(i, j, att_a, att_b)
            cal_HabValue = cal_HabValue - ab_pro * log2(ab_pro)
        Next j
    Next i
    Hab_Function = cal_HabValue
    
End Function
Public Function Uab_Function(ByVal att_a As Integer, ByVal att_b As Integer) As Double
    Dim cal_UabValue As Double
    
    If att_a = att_b Then '如果U(a,a)=1
        cal_UabValue = 1
    ElseIf HValue(att_a) = 0 And HValue(att_b) = 0 Then
        cal_UabValue = 0
    Else
        cal_UabValue = 2 * ((HValue(att_a) + HValue(att_b) - Hab_Value(att_a, att_b)) / (HValue(att_a) + HValue(att_b)))
    End If
    Uab_Function = cal_UabValue
    
End Function
Public Function goodnessFunction(att() As Boolean) As Double '拿各屬性跟第九個屬性-class比
    Dim numerator As Double '分子
    Dim denominator As Double '分母
    
    For i = 1 To 8 '前八個屬性跟第九個算Uab的值
        If att(i) Then
            numerator = numerator + Uab_Value(i, 9)
            For j = 1 To 8
                If att(j) Then
                    denominator = denominator + Uab_Value(i, j)
                End If
            Next j
        End If
    Next i
    '計算goodness值
    If denominator = 0 Then
        goodnessFunction = 0
    Else
        goodnessFunction = numerator / (denominator) ^ 0.5
    End If
        
End Function
Public Function final_output(final_index, final_goodness As Double)
    Dim name As Integer
    
    List1.AddItem "select attribute : "
    For i = 1 To 8
        If final_index(i) <> 0 Then
            name = final_index(i)
            List1.AddItem attribute_name(name) & "(" & final_index(i) & ")"
        End If
    Next i
    List1.AddItem "Goodness : " & final_goodness
    List1.AddItem "-----------------------------"
End Function
Public Function back_final_output(final_index, final_goodness As Double)
    Dim name As Integer
    
    List1.AddItem "remove attribute : "
    For i = 1 To 8
        If final_index(i) <> 0 Then
            name = final_index(i)
            List1.AddItem attribute_name(name) & "(" & final_index(i) & ")"
        End If
    Next i
    List1.AddItem "Goodness : " & final_goodness
    List1.AddItem "-----------------------------"
End Function
Public Function set_output(final_select_array)
    Dim final_set As String
    Dim name As Integer
    For i = 1 To 8
        If final_select_array(i) <> 0 Then '表示第i個屬性有被選擇，例如(0,1,0,0)為第2個屬性被選擇
            name = i
            final_set = final_set & attribute_name(name) & "," '將被選的屬性名稱丟到同一個字串內
        End If
    Next i
    
    final_set = Left(final_set, Len(final_set) - 1)
    List1.AddItem "final set: { " & final_set & " }"
End Function

Private Sub Form_Load()
    attribute_name(1) = "mcg"
    attribute_name(2) = "gvh"
    attribute_name(3) = "alm"
    attribute_name(4) = "mit"
    attribute_name(5) = "erl"
    attribute_name(6) = "pox"
    attribute_name(7) = "vac"
    attribute_name(8) = "nuc"
    forward.Enabled = False
    backward.Enabled = False
End Sub

Private Sub forward_Click()
List1.Clear
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim goodness As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant '用來呼叫goodness函式
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = False
    Next i
    goodness = goodnessFunction(select_array)
    temp_max = goodness
    List1.AddItem "initial goodness: " & goodness
    List1.AddItem "-----------------------------"
    
'選1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = True
        select_array(k - 1) = False
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False '把最後一個屬性初始化
    select_array(select_index(1)) = True '把紀錄到的屬性選擇起來掉 (0 0 1 0 0 0 0 0)
    If select_index(1) = 0 Then '等於0時，代表沒有K值輸入，沒有原MAX他更大的值，所以直接結束選取
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'選2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'選3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'選4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'選5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'選6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'選7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'選8個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = True
        ElseIf k = select_index(7) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(8) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False
    select_array(select_index(1)) = True
    select_array(select_index(2)) = True
    select_array(select_index(3)) = True
    select_array(select_index(4)) = True
    select_array(select_index(5)) = True
    select_array(select_index(6)) = True
    select_array(select_index(7)) = True
    select_array(select_index(8)) = True
    If select_index(8) = 0 Then
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
    
line_end:
    List1.AddItem "再向下select已沒有更大Goodness值"
    output = set_output(select_array)
End Sub

Private Sub backward_Click()
List1.Clear
    Dim select_array(9) As Boolean '從index(1)開始存八個attribute
    Dim goodness As Double
    Dim select_index(9) As Integer '紀錄選到哪幾個attribute，從1開始，共8個屬性，故宣告為9
    
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant '用來呼叫goodness函式
    temp_max = 0
    '初始化八個屬性的是否選擇
    For i = 0 To 8
        select_array(i) = True
    Next i
    goodness = goodnessFunction(select_array)
    temp_max = goodness
    List1.AddItem "initial goodness: " & goodness
    List1.AddItem "-----------------------------"
'移除1個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = False
        select_array(k - 1) = True
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(1) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True '把最後一個屬性初始化
    select_array(select_index(1)) = False '把紀錄到的屬性移除掉
    If select_index(1) = 0 Then '等於0時，代表沒有K值輸入，沒有原MAX他更大的值，所以直接結束選取
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'移除2個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(2) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    If select_index(2) = 0 Then
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'移除3個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(3) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    If select_index(3) = 0 Then
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'移除4個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(4) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    If select_index(4) = 0 Then
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'移除5個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(5) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    If select_index(5) = 0 Then
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'移除6個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(6) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    If select_index(6) = 0 Then
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'移除7個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(7) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    select_array(select_index(7)) = False
    If select_index(7) = 0 Then
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'移除8個屬性時的最大G值---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(2) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(3) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(4) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(5) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(6) + 1 Then
            select_array(k) = False
        ElseIf k = select_index(7) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(8) = k '紀錄選擇的attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True
    select_array(select_index(1)) = False
    select_array(select_index(2)) = False
    select_array(select_index(3)) = False
    select_array(select_index(4)) = False
    select_array(select_index(5)) = False
    select_array(select_index(6)) = False
    select_array(select_index(7)) = False
    select_array(select_index(8)) = False
    If select_index(8) = 0 Then
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
     
    
line_end:
    List1.AddItem "再向上select已沒有更好Goodness值"
    output = set_output(select_array)
End Sub
Private Sub Partition_click()
    List1.Clear
    forward.Enabled = False
    backward.Enabled = False
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
            fileCount = 0
            '讀檔並存入二維陣列
            Do While Not EOF(1)
                Dim tmpline As String
                Dim inputdata() As String
                Dim s As Integer
                
                Line Input #1, tmpline
                
                tmpline = Replace(tmpline, "    ", " ")
                tmpline = Replace(tmpline, "   ", " ")
                tmpline = Replace(tmpline, "  ", " ")
                
                inputdata = Split(tmpline, " ")
                For i = 0 To UBound(inputdata)
                    fileArray(s, i) = inputdata(i)
                    freq_Array(s, i) = inputdata(i)
                    freq_temp_Array(s, i) = inputdata(i)
                    width_Array(s, i) = inputdata(i)
                Next i
                s = s + 1 '用s紀錄1484筆，用i記錄每筆的10個資料
                fileCount = fileCount + 1
            Loop
            For i = 0 To 1483
                Select Case freq_Array(i, 9)
                    Case "CYT"
                        freq_Array(i, 9) = 1
                    Case "ERL"
                        freq_Array(i, 9) = 2
                    Case "EXC"
                        freq_Array(i, 9) = 3
                    Case "ME1"
                        freq_Array(i, 9) = 4
                    Case "ME2"
                        freq_Array(i, 9) = 5
                    Case "ME3"
                        freq_Array(i, 9) = 6
                    Case "MIT"
                        freq_Array(i, 9) = 7
                    Case "NUC"
                        freq_Array(i, 9) = 8
                    Case "POX"
                        freq_Array(i, 9) = 9
                    Case "VAC"
                        freq_Array(i, 9) = 10
                End Select
            Next i
            For i = 0 To 1483
                Select Case width_Array(i, 9)
                    Case "CYT"
                        width_Array(i, 9) = 1
                    Case "ERL"
                        width_Array(i, 9) = 2
                    Case "EXC"
                        width_Array(i, 9) = 3
                    Case "ME1"
                        width_Array(i, 9) = 4
                    Case "ME2"
                        width_Array(i, 9) = 5
                    Case "ME3"
                        width_Array(i, 9) = 6
                    Case "MIT"
                        width_Array(i, 9) = 7
                    Case "NUC"
                        width_Array(i, 9) = 8
                    Case "POX"
                        width_Array(i, 9) = 9
                    Case "VAC"
                        width_Array(i, 9) = 10
                End Select
            Next i
            
            List1.AddItem "成功! " & " 共" & fileCount & "筆資料"
            
            Close #1
        End If
    End If
End Sub
Function log2(x As Double) As Double
    If (x = 0) Then
        log2 = 0
    Else
        log2 = Log(x) / Log(2)
    End If
End Function
