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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
      Caption         =   "�Х���ReadŪ��"
      Height          =   252
      Left            =   6000
      TabIndex        =   11
      Top             =   1080
      Width           =   1452
   End
   Begin VB.Label Label2 
      Caption         =   "�Y�n��ܤ��P�����Ƥ覡�бN�����������sRUN"
      Height          =   252
      Left            =   6000
      TabIndex        =   10
      Top             =   1440
      Width           =   3972
   End
   Begin VB.Label Label5 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "�s�ө���"
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
         Name            =   "�s�ө���"
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
Dim attribute_name(9) As String '�s�K���ݩʪ��W�r�A�qindex=1�}�l
Dim i As Integer, j As Integer, k As Integer
Dim a As Integer, b As Integer
Dim fileCount As Integer
Dim fileArray(1484, 10) As Variant
Dim width_Array(1484, 10) As Variant '�ΨӦs�̫������Ƶ��G
Dim freq_Array(1484, 10) As Variant '�ΨӦs�̫������Ƶ��G
Dim freq_temp_Array(1484, 10) As Variant
Dim interval_pro(10, 9) As Double '�E���ݩʪ��@10��interval����@P�ȡA�qindex=0�}�l�Apro(0)�s���Ointerval"1"
Dim HValue(10) As Double '�ΨӦs�E��H(attribute)�ȡA�qindex=1�}�l(�]�tclass)
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
            For j = i To 1483 '�έ쥻���}�C���Ƨ�(�]���ƧǷ|�л\���쥻�}�C���ǩҥH����A��������)
                If CDbl(fileArray(i, k)) > CDbl(fileArray(j, k)) Then '�Ѥp��j�Ƨ�
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
        '���ƨ̷Ӥ����I�����Ʀ�1-10��
            For j = 0 To 1483 '�έ쥻��temp�}�C��j�p�A�ηs���}�C�s�����Ƶ��G
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
    
    For k = 1 To 9 '�p��C���ݩ�H���ȡA�q�ݩ�(1)��H(1)�}�l�p��
        HValue(k) = HFunction(interval_pro, k)
    Next k
    
    For a = 1 To 9 '�p��C2���ݩ�H���ȡA�qH(1,1)
        For b = 1 To 9
            Hab_Value(a, b) = Hab_Function(Pab_Array, a, b)
        Next b
    Next a
    
    For a = 1 To 9 '�p��C2���ݩ�U���ȡA�qU(1,1)
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
    
    Dim width_max(8) As Double '�K���ݩʤ��O��max
    Dim width_min(8) As Double
    Dim tempmax As Double
    Dim tempmin As Double
    Dim width_w(8) As Double
    Dim width_range(10) As Double '�ΨӬ���interval���U���j�� �䤤0�M10���O��min.max
    
    For i = 1 To 8 '�q1�}�l�O�]����ƪ��Ĥ@�C�Oindex �ҥH�}�C��0�S���γB
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
            '���ƨ̷Ӥ����I�����Ʀ�1-10��
            For j = 0 To 1483 '�έ쥻���}�C��j�p�A�ηs���}�C�s�����Ƶ��G
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
    For k = 1 To 9 '�p��C���ݩ�H���ȡA�q�ݩ�(1)��H(1)�}�l�p��
        HValue(k) = HFunction(interval_pro, k)
    Next k
    
    For a = 1 To 9 '�p��C2���ݩ�H���ȡA�qH(1,1)
        For b = 1 To 9
            Hab_Value(a, b) = Hab_Function(Pab_Array, a, b)
        Next b
    Next a
    
    For a = 1 To 9 '�p��C2���ݩ�U���ȡA�qU(1,1)
        For b = 1 To 9
            Uab_Value(a, b) = Uab_Function(a, b)
        Next b
    Next a
    forward.Enabled = True
    backward.Enabled = True
End Sub
Public Function PFunction(pro_Array)
    Dim pro_count As Double
    For k = 1 To 9 '�]attribute
        For i = 0 To 9 '�]interval�Apro(0)�s���Ointerval"1"������
            pro_count = 0  '�C���]���@��interval��Acount�k�s
            For j = 0 To 1483 '�]data
                If pro_Array(j, k) = i + 1 Then
                    pro_count = pro_count + 1
                    interval_pro(i, k) = pro_count / fileCount '�p��X�Uatrribute���C��interval��P��
                End If
            Next j
        Next i
    Next k
End Function
Public Function P_abFunction(disc_Array)
    '�p��Pab����
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
    '�p��Pab��
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
    
    If att_a = att_b Then '�p�GU(a,a)=1
        cal_UabValue = 1
    ElseIf HValue(att_a) = 0 And HValue(att_b) = 0 Then
        cal_UabValue = 0
    Else
        cal_UabValue = 2 * ((HValue(att_a) + HValue(att_b) - Hab_Value(att_a, att_b)) / (HValue(att_a) + HValue(att_b)))
    End If
    Uab_Function = cal_UabValue
    
End Function
Public Function goodnessFunction(att() As Boolean) As Double '���U�ݩʸ�ĤE���ݩ�-class��
    Dim numerator As Double '���l
    Dim denominator As Double '����
    
    For i = 1 To 8 '�e�K���ݩʸ�ĤE�Ӻ�Uab����
        If att(i) Then
            numerator = numerator + Uab_Value(i, 9)
            For j = 1 To 8
                If att(j) Then
                    denominator = denominator + Uab_Value(i, j)
                End If
            Next j
        End If
    Next i
    '�p��goodness��
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
        If final_select_array(i) <> 0 Then '��ܲ�i���ݩʦ��Q��ܡA�Ҧp(0,1,0,0)����2���ݩʳQ���
            name = i
            final_set = final_set & attribute_name(name) & "," '�N�Q�諸�ݩʦW�٥��P�@�Ӧr�ꤺ
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
    Dim select_array(9) As Boolean '�qindex(1)�}�l�s�K��attribute
    Dim goodness As Double
    Dim select_index(9) As Integer '���������X��attribute�A�q1�}�l�A�@8���ݩʡA�G�ŧi��9
    
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant '�ΨөI�sgoodness�禡
    temp_max = 0
    '��l�ƤK���ݩʪ��O�_���
    For i = 0 To 8
        select_array(i) = False
    Next i
    goodness = goodnessFunction(select_array)
    temp_max = goodness
    List1.AddItem "initial goodness: " & goodness
    List1.AddItem "-----------------------------"
    
'��1���ݩʮɪ��̤jG��---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = True
        select_array(k - 1) = False
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(1) = k '������ܪ�attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = False '��̫�@���ݩʪ�l��
    select_array(select_index(1)) = True '������쪺�ݩʿ�ܰ_�ӱ� (0 0 1 0 0 0 0 0)
    If select_index(1) = 0 Then '����0�ɡA�N��S��K�ȿ�J�A�S����MAX�L��j���ȡA�ҥH�����������
        GoTo line_end
    Else
        output = final_output(select_index, temp_max)
    End If
'��2���ݩʮɪ��̤jG��---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = True
        Else
            select_array(k) = True
            select_array(k - 1) = False
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(2) = k '������ܪ�attribute
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
'��3���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(3) = k '������ܪ�attribute
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
'��4���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(4) = k '������ܪ�attribute
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
'��5���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(5) = k '������ܪ�attribute
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
'��6���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(6) = k '������ܪ�attribute
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
'��7���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(7) = k '������ܪ�attribute
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
'��8���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(8) = k '������ܪ�attribute
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
    List1.AddItem "�A�V�Uselect�w�S����jGoodness��"
    output = set_output(select_array)
End Sub

Private Sub backward_Click()
List1.Clear
    Dim select_array(9) As Boolean '�qindex(1)�}�l�s�K��attribute
    Dim goodness As Double
    Dim select_index(9) As Integer '���������X��attribute�A�q1�}�l�A�@8���ݩʡA�G�ŧi��9
    
    Dim temp_max As Double
    Dim i As Integer, k As Integer
    Dim output As Variant '�ΨөI�sgoodness�禡
    temp_max = 0
    '��l�ƤK���ݩʪ��O�_���
    For i = 0 To 8
        select_array(i) = True
    Next i
    goodness = goodnessFunction(select_array)
    temp_max = goodness
    List1.AddItem "initial goodness: " & goodness
    List1.AddItem "-----------------------------"
'����1���ݩʮɪ��̤jG��---------------------------------------------------------------------
    For k = 1 To 8
        select_array(k) = False
        select_array(k - 1) = True
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(1) = k '������ܪ�attribute
            temp_max = goodness
        End If
    Next k
    select_array(8) = True '��̫�@���ݩʪ�l��
    select_array(select_index(1)) = False '������쪺�ݩʲ�����
    If select_index(1) = 0 Then '����0�ɡA�N��S��K�ȿ�J�A�S����MAX�L��j���ȡA�ҥH�����������
        GoTo line_end
    Else
        output = back_final_output(select_index, temp_max)
    End If
'����2���ݩʮɪ��̤jG��---------------------------------------------------------------------
    For k = 1 To 8
        If k = select_index(1) + 1 Then
            select_array(k) = False
        Else
            select_array(k) = False
            select_array(k - 1) = True
        End If
        goodness = goodnessFunction(select_array)
        If goodness > temp_max Then
            select_index(2) = k '������ܪ�attribute
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
'����3���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(3) = k '������ܪ�attribute
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
'����4���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(4) = k '������ܪ�attribute
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
'����5���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(5) = k '������ܪ�attribute
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
'����6���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(6) = k '������ܪ�attribute
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
'����7���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(7) = k '������ܪ�attribute
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
'����8���ݩʮɪ��̤jG��---------------------------------------------------------------------
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
            select_index(8) = k '������ܪ�attribute
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
    List1.AddItem "�A�V�Wselect�w�S����nGoodness��"
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
            'Ū�ɨæs�J�G���}�C
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
                s = s + 1 '��s����1484���A��i�O���C����10�Ӹ��
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
            
            List1.AddItem "���\! " & " �@" & fileCount & "�����"
            
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
