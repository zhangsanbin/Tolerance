Imports System.Text.RegularExpressions

Public Class Tolerance

    ''' <summary>
    ''' 正则表达式 获取字符串中数字
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetNumbers(ByVal str As String) As String
        Return Regex.Replace(str, "[a-z]", "", RegexOptions.IgnoreCase).Trim()
    End Function

    ''' <summary>
    ''' 正则表达式 获取字符串中所有字符，不包括数字
    ''' </summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function GetString(ByVal str As String) As String
        Return Regex.Replace(str, "[0-9]", "", RegexOptions.IgnoreCase).Trim()
    End Function

    ''' <summary>
    ''' 获取IT标准值(IT标准数据完整)
    ''' </summary>
    ''' <param name="Level">IT01/IT0-IT18</param>
    ''' <param name="D">0-10000</param>
    ''' <returns>IT标准值</returns>
    ''' <remarks></remarks>
    Function GetToleranceValue(ByVal Level As String, ByVal D As Decimal) As String
        Dim ds As New DataSet

        ds.ReadXml(Me.GetType.Assembly.GetManifestResourceStream("Tolerance.tbDM_Tolerance_IT_Value_GBT1800p1_2009.xml"))

        Dim dt As System.Data.DataTable = ds.Tables(0)

        Try
            Dim query = From x In dt.AsEnumerable()
                        Where x.Field(Of String)("dia_lower") < D And x.Field(Of String)("dia_upper") >= D
                        Select x
                        Select New With {.Level = x.Field(Of String)(Level)}

            For Each x In query
                If x.Level.ToString <> "" Then Return x.Level.ToString
            Next
        Catch ex As Exception
            Return "-99999"
        End Try
        Return "-99999"
    End Function

    ''' <summary>
    ''' 获取“轴的”公差偏差值ES EI
    ''' </summary>
    ''' <param name="a_zc">a_ac</param>
    ''' <param name="D">0-10000</param>
    ''' <returns>偏差值</returns>
    ''' <remarks></remarks>
    Function GetShaftES_EI(ByVal a_zc As String, ByVal D As Decimal) As String
        Dim ds As New DataSet

        ds.ReadXml(Me.GetType.Assembly.GetManifestResourceStream("Tolerance.tbDM_Tolerance_Shaft_ES_EI_GBT1800p1_2009.xml"))

        Dim dt As System.Data.DataTable = ds.Tables(0)

        Try
            Dim query = From x In dt.AsEnumerable()
                        Where x.Field(Of String)("dia_lower") < D And x.Field(Of String)("dia_upper") >= D
                        Select x
                        Select New With {.a_zc = x.Field(Of String)(a_zc)}

            For Each x In query
                If x.a_zc.ToString <> "" Then Return x.a_zc.ToString
            Next
        Catch ex As Exception
            Return "-99999"
        End Try
        Return "-99999"
    End Function

    ''' <summary>
    ''' 获取“孔的”公差偏差值ES EI
    ''' </summary>
    ''' <param name="A_ZC">A_ZC</param>
    ''' <param name="D">0-10000</param>
    ''' <returns>偏差值</returns>
    ''' <remarks></remarks>
    Function GetHoleES_EI(ByVal A_ZC As String, ByVal D As Decimal) As String
        Dim ds As New DataSet

        ds.ReadXml(Me.GetType.Assembly.GetManifestResourceStream("Tolerance.tbDM_Tolerance_Hole_ES_EI_GBT1800p1_2009.xml"))

        Dim dt As System.Data.DataTable = ds.Tables(0)

        Try
            Dim query = From x In dt.AsEnumerable()
                        Where x.Field(Of String)("dia_lower") < D And x.Field(Of String)("dia_upper") >= D
                        Select x
                        Select New With {.a_zc = x.Field(Of String)(A_ZC)}

            For Each x In query
                If x.a_zc.ToString <> "" Then Return x.a_zc.ToString
            Next
        Catch ex As Exception
            Return "-99999"
        End Try
        Return "-99999"
    End Function

    Function GetP(ByVal s As String) As String
        If (Regex.IsMatch(s, "^[A-Z0-9]+$")) Then
            Return s & ":A-Z"
        ElseIf (Regex.IsMatch(s, "^[a-z0-9]+$")) Then
            Return s & ":a-z"
        End If
    End Function

    Private Sub Tolerance_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Try
            '设置支持命令行启动程序，接受命令参数数值进行动态计算
            Dim arguments As [String]() = Environment.GetCommandLineArgs()
            If (arguments.Length - 1) >= 3 Then
                RadioButton3.Checked = True
                txtD.Text = arguments(1)
                txtH.Text = arguments(2)
                txtS.Text = arguments(3)
            Else
                RadioButton1.Checked = True
                txtD.Text = 100
            End If
        Catch ex As Exception
            Throw ex
        End Try

        GetHoleES_EI()
        GetShaftES_EI()
    End Sub

    Private Sub txtD_LostFocus(sender As Object, e As EventArgs) Handles txtD.LostFocus
        If txtD.Text = "" Then txtD.Focus()
    End Sub

    Private Sub txtD_TextChanged(sender As Object, e As EventArgs) Handles txtD.TextChanged
        '更改名义尺寸
        lbD.Text = txtD.Text
        GetHoleES_EI()
        GetShaftES_EI()
        GetShaft_And_Hole_XR()
    End Sub

    Private Sub txtH_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtH.KeyPress
        e.KeyChar = Convert.ToChar(e.KeyChar.ToString().ToUpper()) '自动转换成大写
    End Sub

    Private Sub txtH_TextChanged(sender As Object, e As EventArgs) Handles txtH.TextChanged
        '更改 孔公差
        LbH.Text = txtH.Text
        GetHoleES_EI()
        GetShaft_And_Hole_XR()
    End Sub

    Private Sub txtS_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtS.KeyPress
        e.KeyChar = Convert.ToChar(e.KeyChar.ToString().ToLower()) '自动转换成小写
    End Sub

    Private Sub txtS_TextChanged(sender As Object, e As EventArgs) Handles txtS.TextChanged
        '更改 轴公差
        LbS.Text = txtS.Text
        GetShaftES_EI()
        GetShaft_And_Hole_XR()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then Me.TopMost = True Else Me.TopMost = False '置顶显示
    End Sub

    ''' <summary>
    ''' 获取并设置“孔”上公差和下公差
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub GetHoleES_EI()

        If Not (Regex.IsMatch(txtH.Text, "^[A-Z0-9]+$")) Then
            'Me.Text = "孔公差有误！"
            LbHV.Text = " ---" & vbCrLf & " ---"
            Exit Sub
        End If

        If txtD.Text = "" Then
            'Me.Text = "公称尺寸有误！"
            LbHV.Text = " ---" & vbCrLf & " ---"
            Exit Sub
        End If

        Dim A_ZC As String

        A_ZC = GetString(txtH.Text)

        Dim ITV As Decimal = GetToleranceValue("IT" & GetNumbers(txtH.Text), txtD.Text) 'IT7

        Dim ES_EI As Decimal = GetHoleES_EI(A_ZC, txtD.Text) 'a

        Dim TRI_IT As Decimal = 0

        Dim retV As String = ""

        If A_ZC = "A" Or A_ZC = "B" Or A_ZC = "C" Or A_ZC = "CD" Or A_ZC = "D" Or A_ZC = "E" Or A_ZC = "EF" Or A_ZC = "F" Or A_ZC = "FG" Or A_ZC = "G" Or A_ZC = "H" Then

            retV = GetSymbol_ES_EI(ES_EI + ITV, ES_EI)

            If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI + ITV, ES_EI)

        ElseIf A_ZC = "K" Or A_ZC = "M" Or A_ZC = "N" Or A_ZC = "P" Or A_ZC = "R" Or A_ZC = "S" Or A_ZC = "T" Or A_ZC = "U" Or A_ZC = "V" Or A_ZC = "X" Or A_ZC = "Y" Or A_ZC = "Z" Or A_ZC = "ZA" Or A_ZC = "ZB" Or A_ZC = "ZC" Then

            '特殊项目算法
            Select Case txtH.Text
                Case "K01", "K0", "K1", "K2", "K3", "K4", "K5", "K6", "K7", "K8"
                    Console.WriteLine("K01~K8")
                    ES_EI = GetHoleES_EI("K01_8", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "K3"
                                Console.WriteLine("K3")
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "K4"
                                Console.WriteLine("K4")
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "K5"
                                Console.WriteLine("K5")
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "K6"
                                Console.WriteLine("K6")
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "K7"
                                Console.WriteLine("K7")
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "K8"
                                Console.WriteLine("K8")
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "K9", "K10", "K11", "K12", "K13", "K14", "K15", "K16", "K17", "K18"
                    Console.WriteLine("K9~K18")
                    ES_EI = GetHoleES_EI("K9_18", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "M01", "M0", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8"
                    Console.WriteLine("M01~M8")
                    ES_EI = GetHoleES_EI("M01_8", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "M3"
                                Console.WriteLine("M3")
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "M4"
                                Console.WriteLine("M4")
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "M5"
                                Console.WriteLine("M5")
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "M6"
                                Console.WriteLine("M6")
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "M7"
                                Console.WriteLine("M7")
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "M8"
                                Console.WriteLine("M8")
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "M9", "M10", "M11", "M12", "M13", "M14", "M15", "M16", "M17", "M18"
                    Console.WriteLine("M9~M18")
                    ES_EI = GetHoleES_EI("M9_18", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "N01", "N0", "N1", "N2", "N3", "N4", "N5", "N6", "N7", "N8"
                    Console.WriteLine("N01~N8")
                    ES_EI = GetHoleES_EI("N01_8", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "N3"
                                Console.WriteLine("N3")
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "N4"
                                Console.WriteLine("N4")
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "N5"
                                Console.WriteLine("N5")
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "N6"
                                Console.WriteLine("N6")
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "N7"
                                Console.WriteLine("N7")
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "N8"
                                Console.WriteLine("N8")
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "N9", "N10", "N11", "N12", "N13", "N14", "N15", "N16", "N17", "N18"
                    Console.WriteLine("N9~N18")
                    ES_EI = GetHoleES_EI("N9_18", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                    '+++++++++++++++++++++++++++++++++++++++++++P-ZC公差带，可以使用函数动态计算，待优化========================================================

                Case "P01", "P0", "P1", "P2", "P9", "P10", "P11", "P12", "P13", "P14", "P15", "P16", "P17", "P18"
                    Console.WriteLine("P01~P2 P9~P18")
                    ES_EI = GetHoleES_EI("P", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "P3", "P4", "P5", "P6", "P7", "P8"
                    Console.WriteLine("P3~P8")
                    ES_EI = GetHoleES_EI("P", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "P3"
                                Console.WriteLine("P3")
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "P4"
                                Console.WriteLine("P4")
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "P5"
                                Console.WriteLine("P5")
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "P6"
                                Console.WriteLine("P6")
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "P7"
                                Console.WriteLine("P7")
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "P8"
                                Console.WriteLine("P8")
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "R01", "R0", "R1", "R2", "R9", "R10", "R11", "R12", "R13", "R14", "R15", "R16", "R17", "R18"
                    ES_EI = GetHoleES_EI("R", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "R3", "R4", "R5", "R6", "R7", "R8"
                    ES_EI = GetHoleES_EI("R", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "R3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "R4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "R5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "R6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "R7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "R8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "S01", "S0", "S1", "S2", "S9", "S10", "S11", "S12", "S13", "S14", "S15", "S16", "S17", "S18"
                    ES_EI = GetHoleES_EI("S", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "S3", "S4", "S5", "S6", "S7", "S8"
                    ES_EI = GetHoleES_EI("S", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "S3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "S4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "S5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "S6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "S7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "S8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "T01", "T0", "T1", "T2", "T9", "T10", "T11", "T12", "T13", "T14", "T15", "T16", "T17", "T18"
                    ES_EI = GetHoleES_EI("T", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "T3", "T4", "T5", "T6", "T7", "T8"
                    ES_EI = GetHoleES_EI("T", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "T3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "T4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "T5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "T6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "T7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "T8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "U01", "U0", "U1", "U2", "U9", "U10", "U11", "U12", "U13", "U14", "U15", "U16", "U17", "U18"
                    ES_EI = GetHoleES_EI("U", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "U3", "U4", "U5", "U6", "U7", "U8"
                    ES_EI = GetHoleES_EI("U", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "U3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "U4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "U5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "U6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "U7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "U8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "V01", "V0", "V1", "V2", "V9", "V10", "V11", "V12", "V13", "V14", "V15", "V16", "V17", "V18"
                    ES_EI = GetHoleES_EI("V", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "V3", "V4", "V5", "V6", "V7", "V8"
                    ES_EI = GetHoleES_EI("V", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "V3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "V4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "V5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "V6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "V7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "V8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "X01", "X0", "X1", "X2", "X9", "X10", "X11", "X12", "X13", "X14", "X15", "X16", "X17", "X18"
                    ES_EI = GetHoleES_EI("X", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "X3", "X4", "X5", "X6", "X7", "X8"
                    ES_EI = GetHoleES_EI("X", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "X3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "X4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "X5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "X6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "X7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "X8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "Y01", "Y0", "Y1", "Y2", "Y9", "Y10", "Y11", "Y12", "Y13", "Y14", "Y15", "Y16", "Y17", "Y18"
                    ES_EI = GetHoleES_EI("Y", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "Y3", "Y4", "Y5", "Y6", "Y7", "Y8"
                    ES_EI = GetHoleES_EI("Y", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "Y3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Y4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Y5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Y6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Y7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Y8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "Z01", "Z0", "Z1", "Z2", "Z9", "Z10", "Z11", "Z12", "Z13", "Z14", "Z15", "Z16", "Z17", "Z18"
                    ES_EI = GetHoleES_EI("Z", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "Z3", "Z4", "Z5", "Z6", "Z7", "Z8"
                    ES_EI = GetHoleES_EI("Z", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "Z3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Z4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Z5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Z6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Z7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "Z8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "ZA01", "ZA0", "ZA1", "ZA2", "ZA9", "ZA10", "ZA11", "ZA12", "ZA13", "ZA14", "ZA15", "ZA16", "ZA17", "ZA18"
                    ES_EI = GetHoleES_EI("ZA", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "ZA3", "ZA4", "ZA5", "ZA6", "ZA7", "ZA8"
                    ES_EI = GetHoleES_EI("ZA", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "ZA3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZA4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZA5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZA6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZA7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZA8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "ZB01", "ZB0", "ZB1", "ZB2", "ZB9", "ZB10", "ZB11", "ZB12", "ZB13", "ZB14", "ZB15", "ZB16", "ZB17", "ZB18"
                    ES_EI = GetHoleES_EI("ZB", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "ZB3", "ZB4", "ZB5", "ZB6", "ZB7", "ZB8"
                    ES_EI = GetHoleES_EI("ZB", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "ZB3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZB4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZB5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZB6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZB7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZB8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end

                Case "ZC01", "ZC0", "ZC1", "ZC2", "ZC9", "ZC10", "ZC11", "ZC12", "ZC13", "ZC14", "ZC15", "ZC16", "ZC17", "ZC18"
                    ES_EI = GetHoleES_EI("ZC", txtD.Text)
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
                Case "ZC3", "ZC4", "ZC5", "ZC6", "ZC7", "ZC8"
                    ES_EI = GetHoleES_EI("ZC", txtD.Text)
                    If txtD.Text > 3 Then
                        Select Case txtH.Text
                            Case "ZC3"
                                TRI_IT = GetHoleES_EI("tri_it3", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZC4"
                                TRI_IT = GetHoleES_EI("tri_it4", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZC5"
                                TRI_IT = GetHoleES_EI("tri_it5", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZC6"
                                TRI_IT = GetHoleES_EI("tri_it6", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZC7"
                                TRI_IT = GetHoleES_EI("tri_it7", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                            Case "ZC8"
                                TRI_IT = GetHoleES_EI("tri_it8", txtD.Text)
                                ES_EI = ES_EI + TRI_IT
                        End Select
                    End If
                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)
                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)
                    'end
            End Select
            'END

        ElseIf A_ZC = "J" Or A_ZC = "JS" Then

            retV = GetSymbol_ES_EI(ITV / 2, 0 - (ITV / 2))

            If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ITV / 2, 0 - (ITV / 2))

            If A_ZC = "JS" Then ES_EI = 0

            '特殊项目算法
            Select Case txtH.Text
                Case "J6"

                    Console.WriteLine("J6")

                    ES_EI = GetHoleES_EI("J6", txtD.Text)

                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)

                Case "J7"

                    Console.WriteLine("J7")

                    ES_EI = GetHoleES_EI("J7", txtD.Text)

                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)

                Case "J8"

                    Console.WriteLine("J8")

                    ES_EI = GetHoleES_EI("J8", txtD.Text)

                    retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetHoleESEI_Img(ES_EI, ES_EI - ITV)

            End Select

        End If

        If ES_EI = -99999 Or ITV = -99999 Then
            LbHV.Text = " ---" & vbCrLf & " ---"
            PH.Visible = False
            LH.Visible = False
        Else
            LbHV.Text = retV
            PH.Visible = True
            LH.Visible = True
        End If

    End Sub

    ''' <summary>
    ''' 获取并设置“轴”上公差和下公差
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub GetShaftES_EI()

        '轴的基本偏差a~h 和 k~zc 及其 “+” 或 “-”。轴的另一个偏差，下极限偏差（ei）或上极限偏差（es）可由轴的“基本偏差”和“标准公差”求得

        If Not (Regex.IsMatch(txtS.Text, "^[a-z0-9]+$")) Then
            LbSV.Text = " ---" & vbCrLf & " ---"
            Exit Sub
        End If

        If txtD.Text = "" Then
            LbSV.Text = " ---" & vbCrLf & " ---"
            Exit Sub
        End If

        Dim A_ZC As String

        A_ZC = GetString(txtS.Text)

        Dim ITV As Decimal = GetToleranceValue("IT" & GetNumbers(txtS.Text), txtD.Text) 'IT7

        Dim ES_EI As Decimal = GetShaftES_EI(A_ZC, txtD.Text) 'a

        Dim es, ei As Decimal

        Dim retV As String = ""

        If A_ZC = "a" Or A_ZC = "b" Or A_ZC = "c" Or A_ZC = "cd" Or A_ZC = "d" Or A_ZC = "e" Or A_ZC = "ef" Or A_ZC = "f" Or A_ZC = "fg" Or A_ZC = "g" Or A_ZC = "h" Then

            es = ES_EI

            ei = ES_EI - ITV

            retV = GetSymbol_ES_EI(ES_EI, ES_EI - ITV)

            If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ES_EI, ES_EI - ITV)

        ElseIf A_ZC = "k" Or A_ZC = "m" Or A_ZC = "n" Or A_ZC = "p" Or A_ZC = "r" Or A_ZC = "s" Or A_ZC = "t" Or A_ZC = "u" Or A_ZC = "v" Or A_ZC = "x" Or A_ZC = "y" Or A_ZC = "z" Or A_ZC = "za" Or A_ZC = "zb" Or A_ZC = "zc" Then

            retV = GetSymbol_ES_EI(ES_EI + ITV, ES_EI)

            If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ES_EI + ITV, ES_EI)

            '特殊项目算法
            Select Case txtS.Text
                Case "k4", "k5", "k6", "k7"

                    Console.WriteLine("k4~k7")

                    ES_EI = GetShaftES_EI("k4_7", txtD.Text)

                    retV = GetSymbol_ES_EI(ITV + ES_EI, ES_EI)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ITV + ES_EI, ES_EI)

                Case "k01", "k0", "k1", "k2", "k3", "k8", "k9", "k10", "k11", "k12", "k13", "k14", "k15", "k116", "k17", "k18"

                    Console.WriteLine("k01~k3 and k8~k18")

                    ES_EI = GetShaftES_EI("k01_3_k8_18", txtD.Text)

                    retV = GetSymbol_ES_EI(ITV + ES_EI, ES_EI)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ITV + ES_EI, ES_EI)

            End Select

        ElseIf A_ZC = "j" Or A_ZC = "js" Then

            retV = GetSymbol_ES_EI(ITV / 2, 0 - (ITV / 2))

            If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ITV / 2, 0 - (ITV / 2))

            If A_ZC = "js" Then ES_EI = 0

            '特殊项目算法
            Select Case txtS.Text
                Case "j5", "j6"

                    Console.WriteLine("j5 or j6")

                    ES_EI = GetShaftES_EI("j6", txtD.Text)

                    retV = GetSymbol_ES_EI(ITV + ES_EI, ES_EI)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ITV + ES_EI, ES_EI)

                Case "j7"

                    Console.WriteLine("j7")

                    ES_EI = GetShaftES_EI("j7", txtD.Text)

                    retV = GetSymbol_ES_EI(ITV + ES_EI, ES_EI)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ITV + ES_EI, ES_EI)

                Case "j8"

                    Console.WriteLine("j8")

                    ES_EI = GetShaftES_EI("j8", txtD.Text)

                    retV = GetSymbol_ES_EI(ITV + ES_EI, ES_EI)

                    If ES_EI <> -99999 Or ITV <> -99999 Then SetShaftESEI_Img(ITV + ES_EI, ES_EI)

            End Select

        End If

        If ES_EI = -99999 Or ITV = -99999 Then
            LbSV.Text = " ---" & vbCrLf & " ---"
            PS.Visible = False
            LS.Visible = False
        Else
            LbSV.Text = retV
            PS.Visible = True
            LS.Visible = True
        End If

    End Sub

    ''' <summary>
    ''' 格式化公差数值，±符号，小数位数
    ''' </summary>
    ''' <param name="es"></param>
    ''' <param name="ei"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetSymbol_ES_EI(ByVal es As Decimal, ByVal ei As Decimal) As String
        Dim _es, _ei As String

        _es = es
        _ei = ei

        If es > 0 Then
            '正数
            If Not Mid(es, 1, 1) = "+" Then
                _es = "+" & es
            End If
        Else
            '负数
            If Not Mid(es, 1, 1) = "-" Then
                _es = "-" & es
            End If
        End If

        If ei > 0 Then
            '正数
            If Not Mid(ei, 1, 1) = "+" Then
                _ei = "+" & ei
            End If
        Else
            '负数
            If Not Mid(ei, 1, 1) = "-" Then
                _ei = "-" & ei
            End If
        End If

retry:  '重试检测

        If _es.Length <> _ei.Length Then
            '位数不一样
            If _es.Length > _ei.Length Then
                'ES 大于 EI 的长度
                If Mid(_es, _es.Length, 1) = "0" Then
                    '尾数是零
                    _es = Mid(_es, 1, _es.Length - 1)
                    GoTo retry
                Else
                    'ES尾数不是零，补充另一个公差EI的尾数
                    _ei = _ei & "0"
                    GoTo retry
                End If
            Else
                'EI 大于 ES 的长度
                If Mid(_ei, _ei.Length, 1) = "0" Then
                    '尾数是零
                    _ei = Mid(_ei, 1, _ei.Length - 1)
                    GoTo retry
                Else
                    'EI尾数不是零，补充另一个公差ES的尾数
                    _es = _es & "0"
                    GoTo retry
                End If
            End If
        Else
            '位数相同, 检测是否两个公差是否都有多余的0
            If Mid(_es, _es.Length, 1) = "0" And Mid(_ei, _ei.Length, 1) = "0" Then
                _es = Mid(_es, 1, _es.Length - 1)
                _ei = Mid(_ei, 1, _ei.Length - 1)
                GoTo retry
            End If
        End If

        If es = 0 Then _es = " 0"

        If ei = 0 Then _ei = " 0"

        Return _es & vbCrLf & _ei

    End Function

    ''' <summary>
    ''' 设置“轴公差带”图示坐标和尺寸
    ''' </summary>
    ''' <param name="es">上公差</param>
    ''' <param name="ei">下公差</param>
    ''' <remarks></remarks>
    Public Sub SetShaftESEI_Img(ByVal es As Decimal, ByVal ei As Decimal)

        Dim t, h, x, c As Decimal

        x = 740 '37/5=7.4
        c = 37  'TOP = 37  中间零位线

        Select Case es - ei
            Case Is > 5.5
                x = 1
            Case Is > 4.5
                x = 50
            Case Is > 3.5
                x = 100
            Case Is > 2.5
                x = 150
            Case Is > 1.5
                x = 200
            Case Is > 0.5
                x = 250
            Case Is > 0.1
                x = 300
            Case Is > 0.09
                x = 350
            Case Is > 0.08
                x = 400
            Case Is > 0.07
                x = 450
            Case Is > 0.06
                x = 500
            Case Is > 0.05
                x = 550
            Case Is > 0.04
                x = 600
            Case Is > 0.03
                x = 650
            Case Is > 0.02
                x = 700
            Case Is > 0.01
                x = 740
        End Select

        If es > 0 Then
            t = es * x
            If t > c Then t = 37
            t = c - t
        Else
            t = (0 - es) * x
            If t > c Then t = ((0 - es) * x) / c
            t = c + t
        End If

        If ei > 0 Then
            h = ei * x
            If h > c Then h = ei * x / c
            h = c - h - t
        ElseIf ei = 0 Then
            h = c - t
        Else
            h = (0 - ei) * x
            If h > c Then h = c
            h = c - t + h
        End If

        PS.Top = t

        PS.Height = h

        LS.Top = (PS.Height / 2) - (LS.Height / 2) + PS.Top

    End Sub

    ''' <summary>
    ''' 设置“孔公差带”图示坐标和尺寸
    ''' </summary>
    ''' <param name="es">上公差</param>
    ''' <param name="ei">下公差</param>
    ''' <remarks></remarks>
    Public Sub SetHoleESEI_Img(ByVal es As Decimal, ByVal ei As Decimal)

        Dim t, h, x, c As Decimal

        x = 740 '37/5=7.4
        c = 37  'TOP = 37  中间零位线

        Select Case es - ei
            Case Is > 5.5
                x = 1
            Case Is > 4.5
                x = 50
            Case Is > 3.5
                x = 100
            Case Is > 2.5
                x = 150
            Case Is > 1.5
                x = 200
            Case Is > 0.5
                x = 250
            Case Is > 0.1
                x = 300
            Case Is > 0.09
                x = 350
            Case Is > 0.08
                x = 400
            Case Is > 0.07
                x = 450
            Case Is > 0.06
                x = 500
            Case Is > 0.05
                x = 550
            Case Is > 0.04
                x = 600
            Case Is > 0.03
                x = 650
            Case Is > 0.02
                x = 700
            Case Is > 0.01
                x = 740
        End Select

        If es > 0 Then
            t = es * x
            If t > c Then t = c
            t = c - t
        Else
            t = (0 - es) * x
            If t > c Then t = ((0 - es) * x) / c
            t = c + t
        End If

        If ei > 0 Then
            h = ei * x
            If h > c Then h = ei * x / c
            h = c - h - t
        ElseIf ei = 0 Then
            h = c - t
        Else
            h = (0 - ei) * x
            If h > c Then h = c
            h = c - t + h
        End If

        PH.Top = t

        PH.Height = h

        LH.Top = (PH.Height / 2) - (LH.Height / 2) + PH.Top

    End Sub

    ''' <summary>
    ''' 获取轴公差与孔公差配合后的，最大最小间隙
    ''' </summary>
    ''' <remarks></remarks>
    Sub GetShaft_And_Hole_XR()
        Try
            Dim hv_es, hv_ei As Decimal
            Dim sv_es, sv_ei As Decimal

            Dim hves_ei() As String
            hves_ei = LbHV.Text.Split(vbCrLf)
            hv_es = hves_ei(0)
            hv_ei = hves_ei(1)

            Dim sves_ei() As String
            sves_ei = LbSV.Text.Split(vbCrLf)
            sv_es = sves_ei(0)
            sv_ei = sves_ei(1)

            If (hv_ei - sv_es) < 0 And (hv_es - sv_ei) < 0 Then
                max_gyl = hv_ei - sv_es
                min_gyl = hv_es - sv_ei

                max_gyl = max_gyl - (max_gyl * 2)
                min_gyl = min_gyl - (min_gyl * 2)

                Label3.Text = "最大过盈：" & max_gyl
                Label4.Text = "最小过盈：" & min_gyl
            Else
                Label3.Text = "最大间隙：" & hv_es - sv_ei
                Label4.Text = "最小间隙：" & hv_ei - sv_es
            End If
        Catch ex As Exception
            Label3.Text = "最大间隙：---"
            Label4.Text = "最小间隙：---"
        End Try

        '判断是否是过盈配合
        If RadioButton6.Checked And txtD.Text <> "" Then
            Dim pz, pzxs, wdpp As Decimal
            pzxs = txtxzxs.Text
            wdpp = 1
rec1:
            Dim SafetyFactor As Decimal = 0.2
            If max_gyl > 0.5 Then
                SafetyFactor = 0.3
            ElseIf max_gyl > 0.3 Then
                SafetyFactor = 0.2
            ElseIf max_gyl > 0.2 Then
                SafetyFactor = 0.1
            ElseIf max_gyl < 0.2 Then
                SafetyFactor = 0
            End If

            pz = 0.000001 * pzxs * txtD.Text * wdpp
            If pz < (max_gyl + SafetyFactor) Then
                wdpp = wdpp + 1
                GoTo rec1
            End If
            If wdpp > 400 Then
                Label13.Text = "出于安全考虑，系统无法给出合理可参照的加热或冷缩温度。"
            Else
                Label13.Text = "孔零件加热至" & wdpp + 25 & "℃，或轴零件冷缩至-" & wdpp - 25 & "℃，温差需满足" & wdpp & "℃"
            End If
        End If

    End Sub

    '最大过盈量：
    Dim max_gyl, min_gyl As Decimal

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked Then
            cbS.Visible = True '1
            cbH.Visible = True '1

            txtS.Visible = False '3
            txtH.Visible = False '3
        End If
        RadioButton4_CheckedChanged(sender, e)

        Dim tmp As String
        tmp = Replace(cbH.Text, "'", "")
        txtH.Text = Replace(tmp, "*", "")

        tmp = Replace(cbS.Text, "'", "")
        txtS.Text = Replace(tmp, "*", "")

        GetHoleES_EI()
        GetShaftES_EI()
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked Then
            cbS.Visible = True
            cbH.Visible = True

            txtS.Visible = False
            txtH.Visible = False
        End If
        RadioButton4_CheckedChanged(sender, e)

        Dim tmp As String
        tmp = Replace(cbH.Text, "'", "")
        txtH.Text = Replace(tmp, "*", "")

        tmp = Replace(cbS.Text, "'", "")
        txtS.Text = Replace(tmp, "*", "")

        GetHoleES_EI()
        GetShaftES_EI()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        If RadioButton3.Checked Then
            cbS.Visible = False
            cbH.Visible = False

            txtS.Visible = True
            txtH.Visible = True
        End If
        txtH.Text = "H7"
        txtS.Text = "k6"
        GetHoleES_EI()
        GetShaftES_EI()
    End Sub

    Private Sub cbH_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbH.SelectedIndexChanged, cbH.TextChanged
        Dim tmp As String
        tmp = Replace(cbH.Text, "'", "")
        txtH.Text = Replace(tmp, "*", "")
    End Sub

    Private Sub cbS_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbS.SelectedIndexChanged, cbS.TextChanged
        Dim tmp As String
        tmp = Replace(cbS.Text, "'", "")
        txtS.Text = Replace(tmp, "*", "")
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged, RadioButton5.CheckedChanged, RadioButton6.CheckedChanged

        If RadioButton1.Checked Then

            cbH.Items.Clear()
            cbS.Items.Clear()
            cbH.Items.AddRange({"H1", "H2", "H3", "H4", "H5", "H6", "H7", "H8", "H9", "H10", "H11", "H12", "H13"})

            If RadioButton4.Checked Then
                GroupBox8.Visible = False
                Me.Height = 335
                cbS.Items.AddRange({"e6", "f6'", "g6*", "h6*"})
                cbH.Text = "H7"
                cbS.Text = "g6*"
            End If

            If RadioButton5.Checked Then
                GroupBox8.Visible = False
                Me.Height = 335
                cbS.Items.AddRange({"j6", "js6'", "k6*", "m6*", "n6*"})
                cbH.Text = "H7"
                cbS.Text = "k6*"
            End If

            If RadioButton6.Checked Then
                GroupBox8.Visible = True
                Me.Height = 480
                cbS.Items.AddRange({"p6*", "r6'", "s6*", "t6'", "u6*", "v6'", "x6'", "z6'"})
                cbH.Text = "H7"
                cbS.Text = "p6*"
            End If

        End If

        If RadioButton2.Checked Then
            cbH.Items.Clear()
            cbS.Items.Clear()
            cbS.Items.AddRange({"h1", "h2", "h3", "h4", "h5", "h6", "h7", "h8", "h9", "h10", "h11", "h12", "h13"})

            If RadioButton4.Checked Then
                GroupBox8.Visible = False
                Me.Height = 335
                cbH.Items.AddRange({"D7", "E7", "F7'", "G7*", "H7*"})
                cbH.Text = "G7*"
                cbS.Text = "h6"
            End If

            If RadioButton5.Checked Then
                GroupBox8.Visible = False
                Me.Height = 335
                cbH.Items.AddRange({"J7", "JS7'", "K7*'", "M7'", "N7*"})
                cbH.Text = "K7"
                cbS.Text = "h6"
            End If

            If RadioButton6.Checked Then
                GroupBox8.Visible = True
                Me.Height = 480
                cbH.Items.AddRange({"P7*", "R7''", "S7*'", "T7''", "U7*", "V7", "X7", "Y7", "Z7"})
                cbH.Text = "P7*"
                cbS.Text = "h6"
            End If
        End If

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        AboutBox.Show()
        AboutBox.TopLevel = True
        AboutBox.Owner = Me
    End Sub

    Private Sub LbHV_Click(sender As Object, e As EventArgs) Handles LbHV.Click
        Clipboard.Clear() ' 清除剪贴板
        Clipboard.SetText(Replace(LbHV.Text, vbCrLf, "^")) ' 拷贝数据到粘贴板
    End Sub

    Private Sub LbSV_Click(sender As Object, e As EventArgs) Handles LbSV.Click
        Clipboard.Clear() ' 清除剪贴板
        Clipboard.SetText(Replace(LbSV.Text, vbCrLf, "^")) ' 拷贝数据到粘贴板
    End Sub

    Private Sub cbcl_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbcl.SelectedIndexChanged
        '        碳钢
        '        紫铜
        '        黄铜
        '        锡青铜
        '        铝合金
        '        铬钢
        '1Cr18Ni9Ti
        If cbcl.Text Like "碳钢" Then
            txtxzxs.Text = "13.0"
        ElseIf cbcl.Text Like "紫铜" Then
            txtxzxs.Text = "17.5"
        ElseIf cbcl.Text Like "黄铜" Then
            txtxzxs.Text = "16.8"
        ElseIf cbcl.Text Like "锡青铜" Then
            txtxzxs.Text = "17.9"
        ElseIf cbcl.Text Like "铝合金" Then
            txtxzxs.Text = "24.8"
        ElseIf cbcl.Text Like "铬钢" Then
            txtxzxs.Text = "11.8"
        ElseIf cbcl.Text Like "1Cr18Ni9Ti" Then
            txtxzxs.Text = "17.0"
        End If

        '判断是否是过盈配合
        If RadioButton6.Checked And txtD.Text <> "" Then
            Dim pz, pzxs, wdpp As Decimal
            pzxs = txtxzxs.Text
            wdpp = 1
rec1:
            Dim SafetyFactor As Decimal = 0.2
            If max_gyl > 0.5 Then
                SafetyFactor = 0.3
            ElseIf max_gyl > 0.3 Then
                SafetyFactor = 0.2
            ElseIf max_gyl > 0.2 Then
                SafetyFactor = 0.1
            ElseIf max_gyl < 0.2 Then
                SafetyFactor = 0
            End If

            pz = 0.000001 * pzxs * txtD.Text * wdpp
            If pz < (max_gyl + SafetyFactor) Then
                wdpp = wdpp + 1
                GoTo rec1
            End If
            If wdpp > 400 Then
                Label13.Text = "出于安全考虑，系统无法给出合理可参照的加热或冷缩温度。"
            Else
                Label13.Text = "孔零件加热至" & wdpp + 25 & "℃，或轴零件冷缩至-" & wdpp - 25 & "℃，温差需满足" & wdpp & "℃"
            End If
        End If

    End Sub

End Class