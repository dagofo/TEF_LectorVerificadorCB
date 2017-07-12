Imports System.Text.RegularExpressions
Imports System.IO

Public Class frmPrincipal


    Private LecturaInicial As String = ""
    Private contadorVeces As Integer = 0

    Private CodeA As New Dictionary(Of String, String)
    Private CodeB As New Dictionary(Of String, String)
    Private CodeB2 As New Dictionary(Of String, String)
    Private CargaIni As Boolean = True
    Private LecturaTipo As String = ""
    Private LecturaCodeCli As String = ""
    Private Printer, ImpresoraPredet As String
    Private ImpresionOK As Boolean = False
    Private EtiquetaCLI, EtiquetaCLI_TIPO, EtiquetaTEF1, EtiquetaTEF2 As String

    Private Sub MemoEdit1_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles MemoEdit1.EditValueChanged
        Try
            If Timer1.Enabled = False Then
                Timer1.Enabled = True
            End If
            If LecturaInicial = "" Then
                Me.TextEdit1.Text = ""
                Me.TextEdit2.Text = ""
                Me.LabelControl3.Text = "-------------------------"
                Me.LabelControl4.Text = "-------------------------"
                LecturaTipo = ""
            End If
            LecturaInicial = Me.MemoEdit1.Text
        Catch ex As Exception

        End Try

    End Sub

    Private Sub MemoEdit1_EditValueChanging(sender As Object, e As DevExpress.XtraEditors.Controls.ChangingEventArgs) Handles MemoEdit1.EditValueChanging
        'If CargaIni Then
        '    Exit Sub
        'End If
    End Sub

    Private Sub Form1_Activate(sender As Object, e As System.EventArgs) Handles MyBase.Activated
        If CargaIni Then
            CargaIni = False

            Me.MemoEdit1.Focus()
            Exit Sub
        End If
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Dim MyFile As StreamReader = Nothing
        Try

            GroupControl1.Visible = False

            Me.MemoEdit1.Text = ""
            LecturaInicial = ""
            contadorVeces = 0
            Dim VarRes As Object

            Me.TextEdit1.Text = ""
            Me.TextEdit2.Text = ""

            MyFile = New StreamReader(Application.StartupPath & "\" & My.Settings.NombreArchivoCSV)

            Dim MyTxt As String

            MyTxt = MyFile.ReadLine()
            While Not MyFile.EndOfStream
                VarRes = MyTxt.Split(";")
                If Not CodeA.ContainsKey(VarRes(1)) Then
                    CodeA.Add(VarRes(1), VarRes(2))
                End If
                If Not CodeB.ContainsKey(VarRes(2)) Then
                    CodeB.Add(VarRes(2), VarRes(1))
                End If
                If Not CodeB2.ContainsKey(Replace(VarRes(2), " ", "")) Then
                    CodeB2.Add(Replace(VarRes(2), " ", ""), VarRes(1))
                End If
                MyTxt = MyFile.ReadLine()
            End While
            If Trim(MyTxt) <> "" Then
                VarRes = MyTxt.Split(";")
                If Not CodeA.ContainsKey(VarRes(1)) Then
                    CodeA.Add(VarRes(1), VarRes(2))
                End If
                If Not CodeB.ContainsKey(VarRes(2)) Then
                    CodeB.Add(VarRes(2), VarRes(1))
                End If
                If Not CodeB2.ContainsKey(Replace(VarRes(2), " ", "")) Then
                    CodeB2.Add(Replace(VarRes(2), " ", ""), VarRes(1))
                End If
                MyTxt = MyFile.ReadLine()
            End If

            MyFile.Close()

            'CodeA.Add("A2C00031675", "54 159 215")
            'CodeB.Add("54 159 215", "A2C00031675")
            'CodeB2.Add("54159215", "A2C00031675")

            Timer1.Enabled = False

        Catch ex As Exception
            Try
                If Not MyFile Is Nothing Then
                    MyFile.Close()
                End If
            Catch ex2 As Exception

            End Try
            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                ex.Source.ToString & Chr(10) & Chr(13) & _
                ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
        End Try

    End Sub

    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Try
            GroupControl1.Visible = False
            Application.DoEvents()

            Me.MemoEdit1.Text = ""
            LecturaInicial = ""
            contadorVeces = 0
            Me.TextEdit1.Text = ""
            Me.TextEdit2.Text = ""
            Me.LabelControl3.Text = "-------------------------"
            Me.LabelControl4.Text = "-------------------------"
            Me.MemoEdit1.Focus()
        Catch ex As Exception
            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                ex.Source.ToString & Chr(10) & Chr(13) & _
                ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
        End Try

    End Sub

    Private Sub Timer1_Tick(sender As System.Object, e As System.EventArgs) Handles Timer1.Tick
        Try
            If Trim(Me.MemoEdit1.Text) = "" Then
                Exit Sub
            End If

            If Me.MemoEdit1.Text = LecturaInicial Then
                contadorVeces = contadorVeces + 1
                If contadorVeces < My.Settings.RepticionesValidacion * IIf(Len(LecturaInicial) > 20, 3, 1) Then
                    Exit Sub
                Else
                    Timer1.Enabled = False

                    procesarDatos(Me.MemoEdit1.Text)
                End If
            Else
                LecturaInicial = Me.MemoEdit1.Text
            End If

        Catch ex As Exception
            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                ex.Source.ToString & Chr(10) & Chr(13) & _
                ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
        End Try

    End Sub

    Private Sub procesarDatos(ByVal CadenaProc As String)

        Try
            Me.MemoEdit2.Text = CadenaProc

            Dim CadenaProcORIG As String = CadenaProc
            Dim pattern As String = "[^a-zA-Z0-9 *-]"

            '"[^a-zA-Z0-9 -]"
            '[^a-zA-Z0-9\s]+$

            Dim reg As New Regex(pattern)

            CadenaProc = reg.Replace(CadenaProc, "")

            '[)>[RS]06[GS]12S0001[GS]1PRK73H2BTTD1400F[GS]PA2C00031675[GS]31PRK73H2BTTD1400F[GS]20P[GS]2PE  [GS]Q5000[GS]16K543931[GS]V8300664[GS]3SS000000363826    [GS]20T0101[GS]15D07122013[GS]9DD20111103[GS]1T21038460-608[GS]Z1 [GS]1Z[RS][EOT][CR]
            '[)>[RS]06[GS]12S0001[GS]1PRK73B2BTTDD152J[GS]P54 158 077[GS]31PRK73B2BTTDD152J[GS]20P[GS]2PE  [GS]Q10000[GS]16K616626[GS]V8300664[GS]3SS000003134524    [GS]20T0101[GS]15D05122016[GS]9DD20141028[GS]1T5628L375-607[GS]Z1 [GS]1Z[RS][EOT][CR]
            Dim VarRes As Object
            Dim i As Integer
            If Len(CadenaProc) > 20 Then

                VarRes = Split(CadenaProc, My.Settings.CaracterParticion1)

                For i = 0 To UBound(VarRes)
                    If Mid(VarRes(i), 1, Len(VarRes(i)) - 2) & My.Settings.CaracterParticion2 = VarRes(i) Then
                        CadenaProc = Mid(VarRes(i), 1, Len(VarRes(i)) - 2)
                        Exit For
                    End If
                Next
                LecturaCodeCli = "PDF147"
            Else
                LecturaCodeCli = "CODE39"
            End If

            Dim LecturaOK As Boolean = False
            'CadenaProc = Replace(CadenaProc, " ", "")
            '  If CodeA.ContainsKey(CadenaProc) Then
            Me.TextEdit1.Text = CadenaProc
            Me.PictureEdit1.Image = createBarCodeSimple(Replace(CadenaProc, " ", ""))
            '  Me.TextEdit2.Text = CodeA.Item(CadenaProc)
            Me.TextEdit2.Text = CadenaProc 'CodeA.Item(CadenaProc)
            '  Me.PictureEdit2.Image = createBarCodeSimple(Replace(CodeA.Item(CadenaProc), " ", ""))
            Me.PictureEdit2.Image = createBarCodeSimple(Replace(CadenaProc, " ", ""))

            Me.LabelControl3.Text = "Código: " & Me.TextEdit1.Text
            Me.LabelControl4.Text = "Código: " & Me.TextEdit2.Text

            Me.Panel1.BackColor = Color.GreenYellow
            Me.MemoEdit1.Text = ""
            contadorVeces = 0
            LecturaOK = True
            LecturaTipo = "CODEA"
            '  End If

            ' Chema, no se ejecuta
            If 1 = 0 Then

                If CodeB.ContainsKey(CadenaProc) Then
                    Me.TextEdit1.Text = CadenaProc
                    Me.PictureEdit1.Image = createBarCodeSimple(Replace(CadenaProc, " ", ""))
                    Me.TextEdit2.Text = CodeB.Item(CadenaProc)
                    Me.PictureEdit2.Image = createBarCodeSimple(Replace(CodeB.Item(CadenaProc), " ", ""))

                    Me.LabelControl3.Text = "Código: " & Me.TextEdit1.Text
                    Me.LabelControl4.Text = "Código: " & Me.TextEdit2.Text

                    Me.Panel1.BackColor = Color.GreenYellow
                    Me.MemoEdit1.Text = ""
                    contadorVeces = 0
                    LecturaOK = True
                    LecturaTipo = "CODEB"
                End If

                If CodeB2.ContainsKey(CadenaProc) Then
                    Me.TextEdit1.Text = CadenaProc
                    Me.PictureEdit1.Image = createBarCodeSimple(Replace(CadenaProc, " ", ""))
                    Me.TextEdit2.Text = CodeB2.Item(CadenaProc)
                    Me.PictureEdit2.Image = createBarCodeSimple(Replace(CodeB2.Item(CadenaProc), " ", ""))

                    Me.LabelControl3.Text = "Código: " & Me.TextEdit1.Text
                    Me.LabelControl4.Text = "Código: " & Me.TextEdit2.Text

                    Me.Panel1.BackColor = Color.GreenYellow
                    Me.MemoEdit1.Text = ""
                    contadorVeces = 0
                    LecturaOK = True
                    LecturaTipo = "CODE2"
                End If

            End If

            If LecturaOK Then
                'comenzamos las verificaciones
                'a) Es correcta la lectura ---> si
                'Dim Resp As MsgBoxResult = MsgBox("¿Es correcta la lectura: " & Me.TextEdit1.Text & " (CODE: " & LecturaCodeCli & " )?" & _
                '                                  Chr(10) & Chr(13) & _
                '                                  Chr(10) & Chr(13) & _
                '                                  "En caso de contestar ""SI"" se iniciara el proceso de IMPRESION, ""CANCELAR IMPRESION"" para continuar sin imprimir de nuevo las etiquetas", MsgBoxStyle.Question + MsgBoxStyle.DefaultButton2 + MsgBoxStyle.YesNoCancel, "Lector Verificación CB - TEF")

                ImpresionOK = False
                Dim frmIni As New InicioLectura
                frmIni.TextEdit1.Text = Me.TextEdit1.Text
                frmIni.ShowDialog()
                If frmIni.Resp = MsgBoxResult.No Then
                    frmIni.Dispose()
                    GoTo EsErronea
                ElseIf frmIni.Resp = MsgBoxResult.Cancel Then
                    frmIni.Dispose()
                    ImpresionOK = True
                    GoTo ContinuarSinImprimir
                End If
                GroupControl1.Visible = True
                Application.DoEvents()
                'b) Imprimir etiquetas
                ImpresionEtiquetas()
                GroupControl1.Visible = False
                Application.DoEvents()

                Dim MyVerOK As New VerifOK
                MyVerOK.LabelControl1.Text = "PROCESO VERIFICACIÓN"
                MyVerOK.TextEdit1.Text = "PEQUE ETIQUETAS PARA VERIFICACIÓN"
                'MyVerOK.TextEdit1.Text = "PEQUE LAS ETIQUETAS ANTES DE CONTINUAR LA VALIDACION"
                MyVerOK.BackColor = Color.LightSteelBlue
                MyVerOK.LabelControl1.BackColor = Color.LightSteelBlue
                MyVerOK.ShowDialog()
                MyVerOK.Dispose()

                'MsgBox("PEGUE LAS ETIQUETAS ANTES DE CONTINUAR CON LA VERIFICACIÓN", MsgBoxStyle.Information, "Lector Verificación CB - TEF")


ContinuarSinImprimir:

                'c) Verificar que las tres lecturas se corresponden.....
                If ImpresionOK = True Then

                    ' chema, solo se hace en dos pasos, por tanto la segunda (paso1) y tercera lectura (paso 2) no se realizan
                    Dim MyFormLEC As LectorEtiquetas

                    If 1 = 0 Then

                        'hacemos la verificación de las etiquetas

                        'a) Marcar la etiqueta del cliente

LeerPaso1:
                        MyFormLEC = New LectorEtiquetas
                        MyFormLEC.LabelControl1.Text = "LEER ETIQUETA DEL CLIENTE DE NUEVO"
                        MyFormLEC.LabelControl1.BackColor = Color.Orange
                        If My.Settings.WordVisible Then
                            MyFormLEC.ShowInTaskbar = True
                        End If
                        MyFormLEC.ShowDialog()
                        If MyFormLEC.EtiquetaLEIDA = Me.TextEdit1.Text And MyFormLEC.EtiquetaLEIDA_TIPO = LecturaCodeCli Then
                            MyFormLEC.Dispose()
                        Else
                            MyFormLEC.Dispose()


                            ''Dim Resp2 As MsgBoxResult = MsgBox("LA VERIFICACIÓN DE LAS ETIQUETAS ES ERRONEA. ¿VOLVER A EMPEZAR LA VERIFICACIÓN?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Lector Verificación CB - TEF")
                            Dim verifEt As New VerifEtiquetas
                            verifEt.TextEdit1.Text = Me.TextEdit1.Text & " (" & LecturaCodeCli & ")"
                            verifEt.TextEdit2.Text = MyFormLEC.EtiquetaLEIDA & " (" & MyFormLEC.EtiquetaLEIDA_TIPO & ")"
                            verifEt.ShowDialog()
                            If verifEt.Resp2 = MsgBoxResult.Yes Then
                                GoTo LeerPaso1
                            Else
                                'MsgBox("VERIFICACION ERRONEA!!. VUELVA A INICIAR EL PROCESO", MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
                                Dim MyVerERR As New Verif_KO
                                MyVerERR.ShowDialog()
                                GoTo EsErronea
                            End If
                        End If

                        'b) Marcar la etiqueta de TEF1
LeerPaso2:
                        MyFormLEC = New LectorEtiquetas
                        '   MyFormLEC.LabelControl1.Text = "LEER LA ETIQUETA DE TEF Nº 1"
                        MyFormLEC.LabelControl1.Text = "LEER LA ETIQUETA TEF DEL CLIENTE"
                        MyFormLEC.LabelControl1.BackColor = Color.Yellow 'MediumTurquoise
                        If My.Settings.WordVisible Then
                            MyFormLEC.ShowInTaskbar = True
                        End If
                        MyFormLEC.ShowDialog()
                        If MyFormLEC.EtiquetaLEIDA = Me.TextEdit1.Text.Replace(" ", "") And MyFormLEC.EtiquetaLEIDA_TIPO = "CODE39" Then
                            MyFormLEC.Dispose()
                        Else
                            MyFormLEC.Dispose()


                            'Dim Resp2 As MsgBoxResult = MsgBox("LA VERIFICACIÓN DE LAS ETIQUETAS ES ERRONEA. ¿VOLVER A EMPEZAR LA VERIFICACIÓN?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Lector Verificación CB - TEF")

                            Dim verifEt As New VerifEtiquetas
                            verifEt.TextEdit1.Text = Me.TextEdit1.Text & " (CODE39)"
                            verifEt.TextEdit2.Text = MyFormLEC.EtiquetaLEIDA & " (" & MyFormLEC.EtiquetaLEIDA_TIPO & ")"
                            verifEt.ShowDialog()
                            If verifEt.Resp2 = MsgBoxResult.Yes Then
                                GoTo LeerPaso2
                            Else
                                'MsgBox("VERIFICACION ERRONEA!!. VUELVA A INICIAR EL PROCESO", MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
                                Dim MyVerERR As New Verif_KO
                                MyVerERR.ShowDialog()
                                GoTo EsErronea
                            End If
                        End If

                    End If

                    'c) Marcar la etiqueta de TEF2
LeerPaso3:
                    MyFormLEC = New LectorEtiquetas
                    ' MyFormLEC.LabelControl1.Text = "LEER LA ETIQUETA DE TEF Nº 2"
                    MyFormLEC.LabelControl1.Text = "LEER LA ETIQUETA DE TEF"
                    MyFormLEC.LabelControl1.BackColor = Color.MediumTurquoise
                    If My.Settings.WordVisible Then
                        MyFormLEC.ShowInTaskbar = True
                    End If
                    MyFormLEC.ShowDialog()
                    If MyFormLEC.EtiquetaLEIDA = Me.TextEdit2.Text.Replace(" ", "") And MyFormLEC.EtiquetaLEIDA_TIPO = "CODE39" Then
                        MyFormLEC.Dispose()
                    Else
                        MyFormLEC.Dispose()


                        'Dim Resp2 As MsgBoxResult = MsgBox("LA VERIFICACIÓN DE LAS ETIQUETAS ES ERRONEA. ¿VOLVER A EMPEZAR LA VERIFICACIÓN?", MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton1, "Lector Verificación CB - TEF")

                        Dim verifEt As New VerifEtiquetas
                        verifEt.TextEdit1.Text = Me.TextEdit2.Text & " (CODE39)"
                        verifEt.TextEdit2.Text = MyFormLEC.EtiquetaLEIDA & " (" & MyFormLEC.EtiquetaLEIDA_TIPO & ")"
                        verifEt.ShowDialog()
                        If verifEt.Resp2 = MsgBoxResult.Yes Then
                            GoTo LeerPaso3
                        Else
                            'MsgBox("VERIFICACION ERRONEA!!. VUELVA A INICIAR EL PROCESO", MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
                            Dim MyVerERR As New Verif_KO
                            MyVerERR.TextEdit1.Text = Me.TextEdit2.Text
                            MyVerERR.ShowDialog()
                            GoTo EsErronea
                        End If
                    End If

                    'MsgBox("VERIFICACIÓN DE ETIQUETAS REALIZADA CORRECTAMENTE!!, PUEDE CONTINUAR", MsgBoxStyle.Information, "Lector Verificación CB - TEF")
                    MyVerOK = New VerifOK
                    MyVerOK.TextEdit1.Text = Me.TextEdit1.Text
                    MyVerOK.ShowDialog()
                    MyVerOK.Dispose()

                    SimpleButton1_Click(Nothing, Nothing)
                Else
                    GoTo EsErronea
                End If
            Else
EsErronea:
                Me.MemoEdit1.Text = ""
                LecturaInicial = ""
                contadorVeces = 0
                'SimpleButton1_Click(Nothing, Nothing)
                Me.LabelControl3.Text = "Código: LECTURA ERRONEA - NO PROCESADO"
                Me.LabelControl4.Text = "Código: " & CadenaProc

                Me.TextEdit1.Text = ""
                Me.PictureEdit1.Image = createBarCodeSimple("")

                Me.TextEdit2.Text = CadenaProc
                Me.PictureEdit2.Image = createBarCodeSimple(Replace(CadenaProc, " ", ""))
                Me.Panel1.BackColor = Color.Red
                LecturaTipo = ""
                LecturaCodeCli = ""
            End If

        Catch ex As Exception
            GroupControl1.Visible = False
            Application.DoEvents()

            Me.MemoEdit1.Text = ""
            LecturaInicial = ""
            contadorVeces = 0

            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                ex.Source.ToString & Chr(10) & Chr(13) & _
                ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
        End Try


    End Sub

    Function createBarCodeSimple(ByVal info As String) As Drawing.Bitmap
        Try
            Dim k As String = info
            Dim stchar As String = String.Empty
            Dim addStar As String = "*"

            Dim full As String = addStar & info & addStar
            info = full
            Dim bc As Drawing.Bitmap = New Drawing.Bitmap(1, 1)
            'Dim myf As Font = New Font("Arial", 12, FontStyle.Regular) ', GraphicsUnit.Point)
            Dim ft As Drawing.Font = New Drawing.Font("3 of 9 Barcode", 40, Drawing.FontStyle.Regular, Drawing.GraphicsUnit.Point)
            Dim g As Drawing.Graphics = Drawing.Graphics.FromImage(bc)
            Dim infoSize As Drawing.SizeF = g.MeasureString(info, ft)

            bc = New Drawing.Bitmap(bc, infoSize.ToSize)
            g = Drawing.Graphics.FromImage(bc)
            g.Clear(Drawing.Color.White)
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.SingleBitPerPixel
            For Each chr As Char In info
                stchar &= chr.ToString '& " "
            Next

            g.DrawString(stchar, ft, New Drawing.SolidBrush(Drawing.Color.Black), 2, 3)
            g.Flush()
            ft.Dispose()
            g.Dispose()

            Return bc
        Catch ex As Exception
            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                 ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                 ex.Source.ToString & Chr(10) & Chr(13) & _
                 ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
            Return Nothing
        End Try

    End Function

    Private Sub MemoEdit1_GotFocus(sender As Object, e As System.EventArgs) Handles MemoEdit1.GotFocus
        Try
            Me.MemoEdit1.BackColor = Color.GreenYellow
        Catch ex As Exception
            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                 ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                 ex.Source.ToString & Chr(10) & Chr(13) & _
                 ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
        End Try
    End Sub

    Private Sub MemoEdit1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles MemoEdit1.KeyPress
        Console.WriteLine(CStr(e.KeyChar))
    End Sub

    Private Sub MemoEdit1_LostFocus(sender As Object, e As System.EventArgs) Handles MemoEdit1.LostFocus
        Try
            Me.MemoEdit1.BackColor = Color.Red
            Me.Panel1.BackColor = Color.DimGray
        Catch ex As Exception
            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                 ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                 ex.Source.ToString & Chr(10) & Chr(13) & _
                 ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
        End Try
    End Sub

    Private Sub SimpleButton2_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton2.Click
        Try
            Me.Close()
        Catch ex As Exception
            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
                ex.StackTrace.ToString & Chr(10) & Chr(13) & _
                ex.Source.ToString & Chr(10) & Chr(13) & _
                ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
        End Try
    End Sub

    Private Sub ImpresionEtiquetas()

        ImpresionOK = False
        Dim objWord As Object = Nothing
        Try
            Dim FinalB As Boolean = False
            Dim i As Integer
            Dim MyFactura As String
            i = 1
            MyFactura = Application.StartupPath & "\" & My.Settings.NombreArchivoDOC.Replace("XX", "_v" & i)
            While File.Exists(MyFactura) And Not FinalB
                Try
                    File.Delete(MyFactura)
                    FinalB = True
                Catch ex As Exception
                    'problemas con el borrado.... pasamos a la siguiente!!!!
                    i = i + 1
                    MyFactura = Application.StartupPath & "\" & My.Settings.NombreArchivoDOC.Replace("XX", "_v" & i)
                End Try
            End While

            'copiamos el archivos
            File.Copy(Application.StartupPath & "\" & My.Settings.NombreArchivoDOC.Replace("XX", ""), MyFactura, True)

            objWord = CreateObject("Word.application")

            objWord.Documents.Open(FileName:=MyFactura, _
                    ConfirmConversions:=False, ReadOnly:=False, AddToRecentFiles:=False, _
                    PasswordDocument:="", PasswordTemplate:="", Revert:=False, _
                    WritePasswordDocument:="", WritePasswordTemplate:="", Format:= _
                    0, XMLTransform:="")

            objWord.visible = False
            If My.Settings.WordVisible Then
                objWord.visible = True
            End If


            objWord.Selection.Find.ClearFormatting()
            objWord.Selection.Find.Replacement.ClearFormatting()

            With objWord.Selection.Find
                .Text = "XXXXXXXXXXXX"
                .Replacement.Text = Me.TextEdit1.Text.Replace(" ", "")
                .Forward = True
                .Wrap = 1 'wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            objWord.Selection.Find.Execute(Replace:=2)  'wdReplaceAll

            With objWord.Selection.Find
                .Text = "YYYYYYYYYYYY"
                .Replacement.Text = Me.TextEdit1.Text.Replace(" ", "")
                .Forward = True
                .Wrap = 1 'wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            objWord.Selection.Find.Execute(Replace:=2)  'wdReplaceAll

            With objWord.Selection.Find
                .Text = "ZZZZZZZZZZZZ"
                .Replacement.Text = Me.TextEdit2.Text.Replace(" ", "")
                .Forward = True
                .Wrap = 1 'wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            objWord.Selection.Find.Execute(Replace:=2)  'wdReplaceAll

            With objWord.Selection.Find
                .Text = "AAAAAAAAAAAA"
                .Replacement.Text = Me.TextEdit2.Text.Replace(" ", "")
                .Forward = True
                .Wrap = 1 'wdFindContinue
                .Format = False
                .MatchCase = False
                .MatchWholeWord = False
                .MatchWildcards = False
                .MatchSoundsLike = False
                .MatchAllWordForms = False
            End With
            objWord.Selection.Find.Execute(Replace:=2)  'wdReplaceAll


            'ponemos esta como impresora seleccionada!!!!!!1
            ImpresoraPredet = ""
            Dim ImpresoraSeleccionada As String
            ImpresoraSeleccionada = Me.CargarImpresoras()

            If ImpresoraSeleccionada <> "" Then
                Try
                    AsignarImpresoraPredeterminada(ImpresoraSeleccionada)

                    objWord.Application.PrintOut(FileName:="", Range:=0, Item:= _
                            0, Copies:=1, Pages:="", PageType:=0, _
                            ManualDuplexPrint:=False, Collate:=True, Background:=True, PrintToFile:= _
                            False, PrintZoomColumn:=0, PrintZoomRow:=0, PrintZoomPaperWidth:=0, _
                            PrintZoomPaperHeight:=0)

                    Threading.Thread.Sleep(1500)
                    AsignarImpresoraPredeterminada(ImpresoraPredet)

                    objWord.Application.NormalTemplate.Saved = True
                    objWord.Application.ActiveDocument.Close(SaveChanges:=False)
                    objWord.Application.Quit()

                    ImpresionOK = True
                Catch ex As Exception
                    AsignarImpresoraPredeterminada(ImpresoraPredet)
                End Try
            Else
                Try
                    objWord.Application.NormalTemplate.Saved = True
                    objWord.Application.ActiveDocument.Close(SaveChanges:=False)
                    objWord.Application.Quit()

                    MsgBox("No se ha encontrado la impresora Predeterminada!.", MsgBoxStyle.Exclamation, "Lector Verificación CB - TEF")

                Catch ex As Exception
                    AsignarImpresoraPredeterminada(ImpresoraPredet)
                End Try
            End If

            '------------liberamos recursos --

            System.Runtime.InteropServices.Marshal.ReleaseComObject(objWord)
            objWord = Nothing

        Catch ex As Exception
            GroupControl1.Visible = False
            Application.DoEvents()
            Try
                If ImpresoraPredet <> "" Then
                    AsignarImpresoraPredeterminada(ImpresoraPredet)
                End If
            Catch ex2 As Exception
                MsgBox("ERROR: " & ex2.Message & Chr(10) & Chr(13) & _
                      ex2.StackTrace.ToString & Chr(10) & Chr(13) & _
                      ex2.Source.ToString & Chr(10) & Chr(13) & _
                      ex2.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
            End Try

            MsgBox("ERROR: " & ex.Message & Chr(10) & Chr(13) & _
               ex.StackTrace.ToString & Chr(10) & Chr(13) & _
               ex.Source.ToString & Chr(10) & Chr(13) & _
               ex.TargetSite.Name, MsgBoxStyle.Critical, "Lector Verificación CB - TEF")
            Try
                objWord.Application.NormalTemplate.Saved = True
            Catch ex2 As Exception

            End Try
            Try

            Catch ex2 As Exception
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objWord)
            End Try
            Try
                objWord.dispose()
            Catch
                'nada
            End Try
            objWord = Nothing

        End Try

    End Sub

    Private Function CargarImpresoras() As String

        Dim prtdoc As System.Drawing.Printing.PrintDocument
        Dim strPrinter As String

        CargarImpresoras = ""
        Try
            prtdoc = New System.Drawing.Printing.PrintDocument

            Dim strDefaultPrinter As String = prtdoc.PrinterSettings.PrinterName
            ImpresoraPredet = strDefaultPrinter
            For Each strPrinter In Printing.PrinterSettings.InstalledPrinters
                If UCase(strPrinter) = UCase(My.Settings.ImpresoraDefecto) Then ' Like "*BULL*ZIP*" Then
                    CargarImpresoras = strPrinter
                End If
            Next

            Try
                prtdoc.Dispose()
            Catch
                'Nada
            End Try
            prtdoc = Nothing

        Catch ex As Exception

            Try
                prtdoc.Dispose()
            Catch
                'Nada
            End Try
            prtdoc = Nothing

            'MsgBox(ex.Message & Chr(10) & Chr(13) & _
            'ex.StackTrace.ToString & Chr(10) & Chr(13) & _
            'ex.Source.ToString & Chr(10) & Chr(13) & _
            'ex.TargetSite.Name, MsgBoxStyle.Critical)

        End Try

    End Function

    ' Para asignar impresora como predeterminada ** Utiliza el modulo mdlImpresoras **
    Private Function AsignarImpresoraPredeterminada(ByVal Impresora As String) As Boolean

        Dim osinfo As OSVERSIONINFO
        Dim retvalue As Integer

        Try
            osinfo.dwOSVersionInfoSize = Len(osinfo)
            GetVersionEx(osinfo)

            If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then    ' Win95/98
                Call Win95SetDefaultPrinter(Impresora)

            ElseIf osinfo.dwPlatformId = VER_PLATFORM_WIN32_NT Then     ' WinNT
                ' This assumes that future versions of Windows use the NT method
                Call WinNTSetDefaultPrinter(Impresora)
            End If

            Return True

        Catch ex As Exception

            'MsgBox(ex.Message & Chr(10) & Chr(13) & _
            'ex.StackTrace.ToString & Chr(10) & Chr(13) & _
            'ex.Source.ToString & Chr(10) & Chr(13) & _
            'ex.TargetSite.Name, MsgBoxStyle.Critical)

            Return False
        End Try

    End Function

    Private Function Win95SetDefaultPrinter(ByVal Impresora As String) As Boolean

        Dim Handle As Long          'handle to printer
        Dim PrinterName As String
        Dim pd As PRINTER_DEFAULTS
        Dim x As Long
        Dim need As Long            ' bytes needed
        Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
        Dim LastError As Long
        Dim t() As Long

        ' determine which printer was selected
        PrinterName = Impresora
        ' none - exit
        If PrinterName = "" Then
            Return False
            'Exit Function
        End If

        Try
            ' set the PRINTER_DEFAULTS members
            pd.pDatatype = 0&
            pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

            ' Get a handle to the printer
            x = OpenPrinter(PrinterName, Handle, pd)
            ' failed the open
            If x = False Then
                'error handler code goes here
                'Exit Function
                Return False
            End If

            ' Make an initial call to GetPrinter, requesting Level 5
            ' (PRINTER_INFO_5) information, to determine how many bytes
            ' you need
            x = GetPrinter(Handle, 5, 0&, 0, need)
            ' don't want to check Err.LastDllError here - it's supposed
            ' to fail
            ' with a 122 - ERROR_INSUFFICIENT_BUFFER
            ' redim t as large as you need
            ReDim t((need \ 4)) 'As Long

            ' and call GetPrinter for keepers this time
            x = GetPrinter(Handle, 5, t(0), need, need)
            ' failed the GetPrinter
            If x = False Then
                'error handler code goes here
                Return False
                'Exit Function
            End If

            ' set the members of the pi5 structure for use with SetPrinter.
            ' PtrCtoVbString copies the memory pointed at by the two string
            ' pointers contained in the t() array into a Visual Basic string.
            ' The other three elements are just DWORDS (long integers) and
            ' don't require any conversion
            pi5.pPrinterName = PtrCtoVbString(t(0))
            pi5.pPortName = PtrCtoVbString(t(1))
            pi5.Attributes = t(2)
            pi5.DeviceNotSelectedTimeout = t(3)
            pi5.TransmissionRetryTimeout = t(4)

            ' this is the critical flag that makes it the default printer
            pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

            ' call SetPrinter to set it
            x = SetPrinter(Handle, 5, pi5, 0)

            If x = False Then   ' SetPrinter failed
                MsgBox("SetPrinter Failed. Error code: " & Err.LastDllError)
                'Exit Function
                Return False
            Else
                If Printer <> Impresora Then
                    ' Make sure Printer object is set to the new printer
                    SelectPrinter(Impresora, Printer)
                End If
            End If

            ' and close the handle
            ClosePrinter(Handle)

            Return True

        Catch ex As Exception

            'MsgBox(ex.Message & Chr(10) & Chr(13) & _
            'ex.StackTrace.ToString & Chr(10) & Chr(13) & _
            'ex.Source.ToString & Chr(10) & Chr(13) & _
            'ex.TargetSite.Name, MsgBoxStyle.Critical)

            Return False
        End Try

    End Function

    Private Function PtrCtoVbString(ByVal Add As Long) As String

        Dim sTemp As String, x As Long

        Try
            x = lstrcpy(sTemp, Add)
            If (InStr(1, sTemp, Chr(0)) = 0) Then
                PtrCtoVbString = ""
            Else
                PtrCtoVbString = sTemp.Substring(0, InStr(1, sTemp, Chr(0)) - 1) 'Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
            End If

        Catch ex As Exception

            'MsgBox(ex.Message & Chr(10) & Chr(13) & _
            'ex.StackTrace.ToString & Chr(10) & Chr(13) & _
            'ex.Source.ToString & Chr(10) & Chr(13) & _
            'ex.TargetSite.Name, MsgBoxStyle.Critical)

            PtrCtoVbString = ""
        End Try

    End Function

    Private Function WinNTSetDefaultPrinter(ByVal Impresora As String) As Boolean

        Dim Buffer As String
        Dim DeviceName As String
        Dim DriverName As String
        Dim PrinterPort As String
        Dim PrinterName As String
        Dim r As Long

        Try
            If Impresora = "" Then
                Return False
                'Exit Function
            End If

            ' Get the printer information for the currently selected
            ' printer in the list. The information is taken from the
            ' WIN.INI file.
            Buffer = Space(1024)
            PrinterName = Impresora
            r = GetProfileString("PrinterPorts", PrinterName, "", _
                Buffer, Len(Buffer))

            ' Parse the driver name and port name out of the buffer
            GetDriverAndPort(Buffer, DriverName, PrinterPort)

            If DriverName <> "" And PrinterPort <> "" Then
                SetDefaultPrinter(Impresora, DriverName, PrinterPort)
                If Printer <> Impresora Then
                    ' Make sure Printer object is set to the new printer
                    SelectPrinter(Impresora, Printer)
                End If
            End If

            Return True

        Catch ex As Exception

            'MsgBox(ex.Message & Chr(10) & Chr(13) & _
            'ex.StackTrace.ToString & Chr(10) & Chr(13) & _
            'ex.Source.ToString & Chr(10) & Chr(13) & _
            'ex.TargetSite.Name, MsgBoxStyle.Critical)

            Return False

        End Try

    End Function

    Private Sub GetDriverAndPort(ByVal Buffer As String, ByRef DriverName As String, ByRef PrinterPort As String)

        Dim iDriver As Integer
        Dim iPort As Integer

        Try
            DriverName = ""
            PrinterPort = ""

            ' The driver name is first in the string terminated by a comma
            iDriver = InStr(Buffer, ",")
            If iDriver > 0 Then

                ' Strip out the driver name
                DriverName = Buffer.Substring(0, iDriver - 1) 'Left(Buffer, iDriver - 1)

                ' The port name is the second entry after the driver name
                ' separated by commas.
                iPort = InStr(iDriver + 1, Buffer, ",")

                If iPort > 0 Then
                    ' Strip out the port name
                    PrinterPort = Mid(Buffer, iDriver + 1, _
                    iPort - iDriver - 1)
                End If
            End If

        Catch ex As Exception

            'MsgBox(ex.Message & Chr(10) & Chr(13) & _
            '       ex.StackTrace.ToString & Chr(10) & Chr(13) & _
            '       ex.Source.ToString & Chr(10) & Chr(13) & _
            '       ex.TargetSite.Name, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub SetDefaultPrinter(ByVal PrinterName As String, ByVal DriverName As String, ByVal PrinterPort As String)

        Dim DeviceLine As String
        Dim r As Long
        Dim l As Long

        Try
            DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
            ' Store the new printer information in the [WINDOWS] section of
            ' the WIN.INI file for the DEVICE= item
            r = WriteProfileString("windows", "Device", DeviceLine)
            ' Cause all applications to reload the INI file:
            l = SendMessage2(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")

        Catch ex As Exception

            'MsgBox(ex.Message & Chr(10) & Chr(13) & _
            '       ex.StackTrace.ToString & Chr(10) & Chr(13) & _
            '       ex.Source.ToString & Chr(10) & Chr(13) & _
            '       ex.TargetSite.Name, MsgBoxStyle.Critical)

        End Try

    End Sub
    ' Para asignar impresora como predeterminada ** Utiliza el modulo mdlImpresoras **


End Class
