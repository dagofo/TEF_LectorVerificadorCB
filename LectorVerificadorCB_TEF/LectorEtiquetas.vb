Imports System.Text.RegularExpressions

Public Class LectorEtiquetas

    Private LecturaInicial As String = ""
    Private contadorVeces As Integer = 0

    Private CargaIni As Boolean = True

    Public EtiquetaLEIDA, EtiquetaLEIDA_TIPO As String

    Private Sub MemoEdit1_EditValueChanged(sender As System.Object, e As System.EventArgs) Handles MemoEdit1.EditValueChanged
        Try
            If Timer1.Enabled = False Then
                Timer1.Enabled = True
            End If
            If LecturaInicial = "" Then
                Me.TextEdit1.Text = ""

                Me.LabelControl3.Text = "-------------------------"

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

        Try
            Me.MemoEdit1.Text = ""
            LecturaInicial = ""
            contadorVeces = 0

            EtiquetaLEIDA = ""
            EtiquetaLEIDA_TIPO = ""
            
            Me.TextEdit1.Text = ""

            Timer1.Enabled = False
            Timer2.Enabled = False
            Me.Timer2.Interval = My.Settings.TiempoEsperaVerif
            Timer2.Enabled = False

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

            If LabelControl1.Text = "LEER LA ETIQUETA DE TEF" Then
                If Mid(CadenaProc, 1, 1) <> "C" Then
                    CadenaProc = CadenaProc & " - Falta la C inicial"
                Else
                    CadenaProc = Mid(CadenaProc, 2, CadenaProc.Length)
                End If
            End If

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
                EtiquetaLEIDA_TIPO = "PDF147"
            Else
                EtiquetaLEIDA_TIPO = "CODE39"
            End If

            Dim LecturaOK As Boolean = False

            Me.TextEdit1.Text = CadenaProc
            Me.PictureEdit1.Image = createBarCodeSimple(Replace(CadenaProc, " ", ""))
            Me.LabelControl3.Text = "Código: " & Me.TextEdit1.Text

            EtiquetaLEIDA = Me.TextEdit1.Text

            Me.Panel1.BackColor = Color.GreenYellow
            Me.MemoEdit1.Text = ""
            contadorVeces = 0
            LecturaOK = True

            Me.Timer2.Enabled = True

        Catch ex As Exception

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

    Private Sub Timer2_Tick(sender As System.Object, e As System.EventArgs) Handles Timer2.Tick
        Try
            Me.Close()
        Catch ex As Exception

        End Try
    End Sub
End Class