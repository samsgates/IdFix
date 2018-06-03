Imports System.IO
Imports System.Text.RegularExpressions
Imports System.ComponentModel



Public Class MForm

    Dim app As InDesign.Application
    Dim tableCol As New ArrayList
    Dim curErrLog As String = ""
    Dim _bw As BackgroundWorker
    Public Delegate Sub update_probar(ByVal text As String, ByVal max As Integer, ByVal cur As Integer)

    Public Sub bw_Dowork(ByVal sender As Object, ByVal e As DoWorkEventArgs)
        Dim live_pdf_info As Updf = CType(e.Argument, Updf)
        convert_epub(live_pdf_info)

    End Sub

    Public Sub bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)
        probar.Visible = False
        lbstatus.Text = ""
        If curErrLog <> "" Then
            gCls.show_error(curErrLog)
            Return
        Else
            gCls.show_message("ePub converted successfully.")
        End If
    End Sub

    Public Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)

    End Sub


    Public Sub probar_update(ByVal Etext As String, ByVal max As Integer, ByVal cur As Integer)
        Try


            If probar.InvokeRequired Then
                Dim up As New update_probar(AddressOf probar_update)
                probar.Invoke(up, New Object() {Etext, max, cur})
            Else
                If Etext <> "" Then
                    lbstatus.Text = Etext
                End If
                probar.Maximum = max
                probar.Value = cur

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub


    Public Sub hide_text(ByVal cdoc As InDesign.Document, ByVal visible As Boolean)
        Try


            tableCol.Clear()

            Dim pageCount As Integer = cdoc.Pages.Count
            'for each page 
            For i = 1 To pageCount
                Dim cPage As InDesign.Page = cdoc.Pages(i)

                'for each text frame
                For Each tFrame As InDesign.TextFrame In cPage.TextFrames
                    tFrame.Locked = False
                    If tFrame.Label = "idfix_img" Then
                        tFrame.Visible = True
                    Else
                        tFrame.Visible = visible
                    End If

                Next

                For h = 1 To cPage.Rectangles.Count
                    Dim iRec As InDesign.Rectangle = cPage.Rectangles(h)
                    If iRec.HtmlItems.Count > 0 Then
                        Dim hItem As InDesign.HtmlItem = iRec.HtmlItems(1)

                        'If hItem.HtmlContent <> "" Then
                        iRec.Visible = visible
                        'End If
                    End If
                Next

            Next ' end page loop




            '//===========get table data=======
            If visible = True Then

                For t As Integer = 1 To cdoc.Stories.Count
                    Dim aStory As InDesign.Story = cdoc.Stories(t)



                    For l = 1 To aStory.Lines.Count
                        Dim tLine As InDesign.Line
                        tLine = aStory.Lines(l)

                        If tLine.Tables.Count > 0 Then
                            For Each mTable As InDesign.Table In tLine.Tables

                                'Dim tabPath As InDesign.Objects = mTable.CreateOutlines(False)
                                'Dim tabTop As Double = tabPath(1).GeometricBounds(0)
                                'Dim tabLeft As Double = tabPath(1).GeometricBounds(1)
                                Dim tabBase As Double = (tLine.Baseline * Convert.ToDouble(cmb_dpi.Text)) / 72
                                Dim tabLeft As Double = (tLine.HorizontalOffset * Convert.ToDouble(cmb_dpi.Text)) / 72
                                Dim ViewHeight As Double = Double.Parse(app.ActiveDocument.DocumentPreferences.PageHeight.ToString())
                                ViewHeight = (ViewHeight * Convert.ToDouble(cmb_dpi.Text)) / 72
                                tabBase = ViewHeight - tabBase

                                mTable.Select()
                                Dim curTf As InDesign.TextFrame = mTable.Parent

                                Dim pNum As String = curTf.ParentPage.Name
                                tableCol.Add(pNum + "|" + tabLeft.ToString() + "|" + tabBase.ToString())
                                app.Select(InDesign.idNothingEnum.idNothing)


                                'tabPath(1).Delete()

                            Next
                        End If

                    Next


                Next
            End If

            '//=======



        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString())
        End Try

    End Sub
    Function get_active_pagenumber(ByVal pageName As String, ByVal aDoc As InDesign.Document) As String
        Dim rtn_pagenum As String = ""

        For p As Integer = 1 To aDoc.Pages.Count
            Dim dPage As InDesign.Page = aDoc.Pages(p)
            If dPage.Name = pageName Then
                rtn_pagenum = p.ToString()
            End If
        Next

        Return rtn_pagenum
    End Function
    Function insert_sp()

        For Each fxt As InDesign.TextFrame In app.ActiveDocument.TextFrames
            If fxt.Label = "b1" Then
                fxt.Contents = ""
                fxt.Contents = InDesign.idSpecialCharacters.idArabicComma  '&#x60c;

                fxt.Contents = InDesign.idSpecialCharacters.idArabicKashida '&#x0640;
                fxt.Contents = InDesign.idSpecialCharacters.idArabicQuestionMark
                fxt.Contents = InDesign.idSpecialCharacters.idArabicSemicolon
                fxt.Contents = InDesign.idSpecialCharacters.idAutoPageNumber
                fxt.Contents = InDesign.idSpecialCharacters.idBulletCharacter
                fxt.Contents = InDesign.idSpecialCharacters.idColumnBreak
                fxt.Contents = InDesign.idSpecialCharacters.idCopyrightSymbol
                fxt.Contents = InDesign.idSpecialCharacters.idDegreeSymbol
                fxt.Contents = InDesign.idSpecialCharacters.idDiscretionaryHyphen
                fxt.Contents = InDesign.idSpecialCharacters.idDiscretionaryLineBreak
                fxt.Contents = InDesign.idSpecialCharacters.idDottedCircle  '◌
                fxt.Contents = InDesign.idSpecialCharacters.idDoubleLeftQuote
                fxt.Contents = InDesign.idSpecialCharacters.idDoubleRightQuote
                fxt.Contents = InDesign.idSpecialCharacters.idDoubleStraightQuote
                fxt.Contents = InDesign.idSpecialCharacters.idEllipsisCharacter
                fxt.Contents = InDesign.idSpecialCharacters.idEmDash
                fxt.Contents = InDesign.idSpecialCharacters.idEmSpace
                fxt.Contents = InDesign.idSpecialCharacters.idEnDash
                fxt.Contents = InDesign.idSpecialCharacters.idEndNestedStyle
                fxt.Contents = InDesign.idSpecialCharacters.idEnSpace
                fxt.Contents = InDesign.idSpecialCharacters.idEvenPageBreak
                fxt.Contents = InDesign.idSpecialCharacters.idFigureSpace
                fxt.Contents = InDesign.idSpecialCharacters.idFixedWidthNonbreakingSpace
                fxt.Contents = InDesign.idSpecialCharacters.idFlushSpace
                fxt.Contents = InDesign.idSpecialCharacters.idFootnoteSymbol
                fxt.Contents = InDesign.idSpecialCharacters.idForcedLineBreak
                fxt.Contents = InDesign.idSpecialCharacters.idFrameBreak
                fxt.Contents = InDesign.idSpecialCharacters.idHairSpace
                fxt.Contents = InDesign.idSpecialCharacters.idHebrewGeresh
                fxt.Contents = InDesign.idSpecialCharacters.idHebrewGershayim
                fxt.Contents = InDesign.idSpecialCharacters.idHebrewMaqaf
                fxt.Contents = InDesign.idSpecialCharacters.idHebrewSofPasuk
                fxt.Contents = InDesign.idSpecialCharacters.idIndentHereTab
                fxt.Contents = InDesign.idSpecialCharacters.idLeftToRightEmbedding
                fxt.Contents = InDesign.idSpecialCharacters.idLeftToRightMark
                fxt.Contents = InDesign.idSpecialCharacters.idLeftToRightOverride
                fxt.Contents = InDesign.idSpecialCharacters.idNextPageNumber
                fxt.Contents = InDesign.idSpecialCharacters.idNonbreakingHyphen
                fxt.Contents = InDesign.idSpecialCharacters.idNonbreakingSpace
                fxt.Contents = InDesign.idSpecialCharacters.idOddPageBreak
                fxt.Contents = InDesign.idSpecialCharacters.idPageBreak
                fxt.Contents = InDesign.idSpecialCharacters.idParagraphSymbol
                fxt.Contents = InDesign.idSpecialCharacters.idPopDirectionalFormatting
                fxt.Contents = InDesign.idSpecialCharacters.idPreviousPageNumber
                fxt.Contents = InDesign.idSpecialCharacters.idPunctuationSpace
                fxt.Contents = InDesign.idSpecialCharacters.idQuarterSpace
                fxt.Contents = InDesign.idSpecialCharacters.idRegisteredTrademark
                fxt.Contents = InDesign.idSpecialCharacters.idRightIndentTab
                fxt.Contents = InDesign.idSpecialCharacters.idRightToLeftEmbedding
                fxt.Contents = InDesign.idSpecialCharacters.idRightToLeftMark
                fxt.Contents = InDesign.idSpecialCharacters.idRightToLeftOverride
                fxt.Contents = InDesign.idSpecialCharacters.idSectionMarker
                fxt.Contents = InDesign.idSpecialCharacters.idSectionSymbol
                fxt.Contents = InDesign.idSpecialCharacters.idSingleLeftQuote
                fxt.Contents = InDesign.idSpecialCharacters.idSingleRightQuote
                fxt.Contents = InDesign.idSpecialCharacters.idSingleStraightQuote
                fxt.Contents = InDesign.idSpecialCharacters.idSixthSpace
                fxt.Contents = InDesign.idSpecialCharacters.idTextVariable
                fxt.Contents = InDesign.idSpecialCharacters.idThinSpace
                fxt.Contents = InDesign.idSpecialCharacters.idThirdSpace
                fxt.Contents = InDesign.idSpecialCharacters.idTrademarkSymbol
                fxt.Contents = InDesign.idSpecialCharacters.idZeroWidthJoiner
                fxt.Contents = InDesign.idSpecialCharacters.idZeroWidthNonjoiner




            End If

        Next
    End Function
    Function get_specialchar(ByVal tChar As InDesign.Character, ByVal curPage As InDesign.Page) As String
        Dim rtnTxt As String = ""
        Try
            Dim intTxt As Integer = Integer.Parse(tChar.Contents.ToString())
            rtnTxt = tChar.Contents.ToString()
        Catch ex As Exception
            Return tChar.Contents.ToString()
        End Try


        If tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idArabicComma).ToString() Then
            rtnTxt = "،"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idArabicKashida).ToString() Then
            rtnTxt = "ـ"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idArabicQuestionMark).ToString() Then
            rtnTxt = "؟"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idArabicSemicolon).ToString() Then
            rtnTxt = "؛"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idAutoPageNumber).ToString() Then
            rtnTxt = curPage.Name

        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idBulletCharacter).ToString() Then
            rtnTxt = "•"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idColumnBreak).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idCopyrightSymbol).ToString() Then
            rtnTxt = "©"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idDegreeSymbol).ToString() Then
            rtnTxt = "°"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idDiscretionaryHyphen).ToString() Then
            rtnTxt = "-"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idDiscretionaryLineBreak).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idDottedCircle).ToString() Then
            rtnTxt = "◌"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idDoubleLeftQuote).ToString() Then
            rtnTxt = """"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idDoubleRightQuote).ToString() Then
            rtnTxt = """"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idDoubleStraightQuote).ToString() Then
            rtnTxt = """"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idEllipsisCharacter).ToString() Then
            rtnTxt = "…"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idEmDash).ToString() Then
            rtnTxt = "—"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idEmSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idEnDash).ToString() Then
            rtnTxt = "–"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idEndNestedStyle).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idEnSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idEvenPageBreak).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idFigureSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idFixedWidthNonbreakingSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idFlushSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idFootnoteSymbol).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idForcedLineBreak).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idFrameBreak).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idHairSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idHebrewGeresh).ToString() Then
            rtnTxt = "׳"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idHebrewGershayim).ToString() Then
            rtnTxt = "״"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idHebrewMaqaf).ToString() Then
            rtnTxt = "־"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idHebrewSofPasuk).ToString() Then
            rtnTxt = "׃"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idIndentHereTab).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idLeftToRightEmbedding).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idLeftToRightMark).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idLeftToRightOverride).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idNextPageNumber).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idNonbreakingHyphen).ToString() Then
            rtnTxt = "-"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idNonbreakingSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idOddPageBreak).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idPageBreak).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idParagraphSymbol).ToString() Then
            rtnTxt = "¶"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idPopDirectionalFormatting).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idPreviousPageNumber).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idPunctuationSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idQuarterSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idRegisteredTrademark).ToString() Then
            rtnTxt = "®"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idRightIndentTab).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idRightToLeftEmbedding).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idRightToLeftMark).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idRightToLeftOverride).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idSectionMarker).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idSectionSymbol).ToString() Then
            rtnTxt = "§"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idSingleLeftQuote).ToString() Then
            rtnTxt = "‘"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idSingleRightQuote).ToString() Then
            rtnTxt = "’"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idSingleStraightQuote).ToString() Then
            rtnTxt = "'"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idSixthSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idTextVariable).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idThinSpace).ToString() Then
            rtnTxt = " "
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idThirdSpace).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idTrademarkSymbol).ToString() Then
            rtnTxt = "™"
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idZeroWidthJoiner).ToString() Then
            rtnTxt = ""
        ElseIf tChar.Contents.ToString() = Convert.ToInt32(InDesign.idSpecialCharacters.idZeroWidthNonjoiner).ToString() Then
            rtnTxt = ""

        End If

        If rtnTxt <> "" Then
            Try
                Dim intTest As Integer = Integer.Parse(rtnTxt)
                If rtnTxt.Length = 1 Then
                    rtnTxt = tChar.Contents.ToString()
                End If
            Catch ex As Exception

            End Try
        End If

        Return rtnTxt
    End Function

    Function add_fontcol(ByVal fntcol As ArrayList, ByVal fntName As String, ByVal fntLocation As String)

        Dim fntFound As Boolean = False
        For i As Integer = 0 To fntcol.Count - 1
            Dim flist As String() = fntcol(i).ToString().Split("|")
            Dim fontName As String = flist(0)
            Dim fontLocation As String = flist(1)
            If fontName = fntName Then
                fntFound = True
            End If

        Next

        If fntFound = False Then
            fntcol.Add(fntName + "|" + fntLocation)
        End If

    End Function
    Function add_table_style(ByVal styCol As ArrayList, ByVal cssStyle As String)
        Dim cFound As Boolean = False
        Dim cssTxt() As String = cssStyle.Split("|")
        Dim cssName As String = cssTxt(0)
        Dim cssClass As String = cssTxt(1)

        For xs As Integer = 0 To styCol.Count - 1
            Dim vlist As String() = styCol(xs).ToString().Split("|")
            If vlist(0) = cssName Then
                cFound = True
            End If
        Next

        If cFound = False Then
            styCol.Add(cssName + "|" + cssClass)

        End If

    End Function



    Function check_style(ByVal styCol As ArrayList, ByVal lineStyle As String) As String
        Dim curClassName As String = ""
        Dim cFound As Boolean = False
        For xs As Integer = 0 To styCol.Count - 1
            Dim vlist As String() = styCol(xs).ToString().Split("|")
            Dim curSty As String = vlist(1) + "|" + vlist(2) + "|" + vlist(3)
            If curSty = lineStyle Then
                cFound = True
                curClassName = vlist(0)
            End If

        Next

        If cFound = False Then
            curClassName = "cls" + (styCol.Count + 1).ToString()
            styCol.Add(curClassName + "|" + lineStyle)

        End If

        Return curClassName
    End Function
    Function convert_rgb(ByVal clr As InDesign.Color) As String

        Dim r As Double
        Dim g As Double
        Dim b As Double
        Dim objclr As Object()

        If clr.Space = InDesign.idColorSpace.idCMYK Then
            objclr = clr.ColorValue
            Dim c As Double = Double.Parse(objclr(0))
            Dim m As Double = Double.Parse(objclr(1))
            Dim y As Double = Double.Parse(objclr(2))
            Dim k As Double = Double.Parse(objclr(3))

            If c <= 0 Then
                c = 0
            End If
            If m <= 0 Then
                m = 0
            End If
            If y <= 0 Then
                y = 0
            End If
            If k <= 0 Then
                k = 0
            End If

            If c > 100 Then
                c = 100
            End If
            If m > 100 Then
                m = 100
            End If
            If y > 100 Then
                y = 100
            End If
            If k > 100 Then
                k = 100
            End If
            c = c / 100
            m = m / 100
            y = y / 100
            k = k / 100

            r = 1 - Math.Min(1, c * (1 - k) + k)
            g = 1 - Math.Min(1, m * (1 - k) + k)
            b = 1 - Math.Min(1, y * (1 - k) + k)

            r = Math.Round(r * 255)
            g = Math.Round(g * 255)
            b = Math.Round(b * 255)

        ElseIf clr.Space = InDesign.idColorSpace.idRGB Then
            objclr = clr.ColorValue
            r = Double.Parse(objclr(0))
            g = Double.Parse(objclr(1))
            b = Double.Parse(objclr(2))
        ElseIf clr.Space = InDesign.idColorSpace.idLAB Then
            objclr = clr.ColorValue
            Dim lab_l As Double = Double.Parse(objclr(0))
            Dim lab_a As Double = Double.Parse(objclr(1))
            Dim lab_b As Double = Double.Parse(objclr(2))
            g = (lab_l * +16) / 116
            r = lab_a * (lab_a / 500 + g)
            b = lab_b * (lab_b / 200 - g)

            If g > 0.008856 Then
                g = g
            Else
                g = (g - 16 / 116) / 7.787
            End If

            If r > 0.008856 Then
                r = r
            Else
                r = (r - 16 / 116) / 7.787
            End If
            If b > 0.008856 Then
                b = b
            Else
                b = (b - 16 / 116) / 7.787
            End If

            r = r * 95.047
            g = g * 100.0
            b = b * 108.883


        ElseIf clr.Space = InDesign.idColorSpace.idMixedInk Then

        End If


        Return "rgb(" + r.ToString() + "," + g.ToString() + "," + b.ToString() + ")"

    End Function
    Public Sub convert_epub(ByVal cPDF As Updf)
        Try


            Try
                ' app = New InDesign.Application
                app = CType(Activator.CreateInstance(Type.GetTypeFromProgID("InDesign.Application"), True), InDesign.Application)

            Catch ex As Exception
                curErrLog += "Open Indesign Document" + Environment.NewLine
                Exit Sub
                
            End Try

            If app.Documents.Count < 1 Then
                curErrLog += "Active document not found" + Environment.NewLine
                Exit Sub
            End If

            Dim docName As String = Path.GetFileNameWithoutExtension(app.ActiveDocument.Name)
            Dim docPath As String = app.ActiveDocument.FilePath

            If app.ActiveDocument.Pages.Count < 1 Then
                curErrLog += "Page not found" + Environment.NewLine

                Exit Sub
            End If

            'create temp directory
            Dim wrkDir As String = Application.StartupPath + "\infile"
            If Directory.Exists(wrkDir) Then
                Directory.Delete(wrkDir, True)
            End If
            Directory.CreateDirectory(wrkDir)

            If Not Directory.Exists(Path.Combine(wrkDir, "export")) Then
                Directory.CreateDirectory(Path.Combine(wrkDir, "export"))
            End If


            If Not Directory.Exists(Path.Combine(wrkDir, "META-INF")) Then
                Directory.CreateDirectory(Path.Combine(wrkDir, "META-INF"))
            End If

            If Not Directory.Exists(Path.Combine(wrkDir, "OEBPS")) Then
                Directory.CreateDirectory(Path.Combine(wrkDir, "OEBPS"))
            End If

            Dim oebDir As String = Path.Combine(wrkDir, "OEBPS")

            If Not Directory.Exists(Path.Combine(oebDir, "html")) Then
                Directory.CreateDirectory(Path.Combine(oebDir, "html"))
            End If

            If Not Directory.Exists(Path.Combine(oebDir, "images")) Then
                Directory.CreateDirectory(Path.Combine(oebDir, "images"))
            End If

            If Not Directory.Exists(Path.Combine(oebDir, "audio")) Then
                Directory.CreateDirectory(Path.Combine(oebDir, "audio"))
            End If

            If Not Directory.Exists(Path.Combine(oebDir, "styles")) Then
                Directory.CreateDirectory(Path.Combine(oebDir, "styles"))
            End If

            Dim imgDir As String = Path.Combine(oebDir, "images")
            Dim styDir As String = Path.Combine(oebDir, "styles")
            Dim htmlDir As String = Path.Combine(oebDir, "html")
            Dim audioDir As String = Path.Combine(oebDir, "audio")
            Dim expDir As String = Path.Combine(wrkDir, "export")


            app.ActiveDocument.ViewPreferences.HorizontalMeasurementUnits = InDesign.idMeasurementUnits.idPixels
            app.ActiveDocument.ViewPreferences.VerticalMeasurementUnits = InDesign.idMeasurementUnits.idPixels
            app.ActiveDocument.ViewPreferences.RulerOrigin = InDesign.idRulerOrigin.idPageOrigin
            'app.ActiveDocument.ViewPreferences.StrokeMeasurementUnits = InDesign.idMeasurementUnits.idPixels
            'app.ActiveDocument.ViewPreferences.TextSizeMeasurementUnits = InDesign.idMeasurementUnits.idPixels
            'app.ActiveDocument.ViewPreferences.TypographicMeasurementUnits = InDesign.idMeasurementUnits.idPixels
            'app.ActiveDocument.ViewPreferences.PrintDialogMeasurementUnits = InDesign.idMeasurementUnits.idPixels

            'app.ActiveDocument.ViewPreferences.GuideSnaptoZone = 1

            app.ActiveDocument.ZeroPoint = {0, 0}
            Dim pageCount As Integer = app.ActiveDocument.Pages.Count

            hide_text(app.ActiveDocument, False)

            'export to image
            app.PNGExportPreferences.PNGQuality = InDesign.idPNGQualityEnum.idHigh
            app.PNGExportPreferences.ExportResolution = Convert.ToDouble(cPDF.b_dpi)

            app.PNGExportPreferences.PageString = "1"
            app.PNGExportPreferences.PNGExportRange = InDesign.idPNGExportRangeEnum.idExportAll


            app.ActiveDocument.Export(InDesign.idExportFormat.idPNGFormat, imgDir + "\image.png", False, , , False)


            hide_text(app.ActiveDocument, True)


            'rename image one
            If File.Exists(imgDir + "\image.png") Then
                File.Move(imgDir + "\image.png", imgDir + "\image1.png")
            End If



            'set objectsytle to html widgets
            For i = 1 To pageCount
                Dim c1Page As InDesign.Page = app.ActiveDocument.Pages(i)
                For h = 1 To c1Page.Rectangles.Count
                    Dim iRec As InDesign.Rectangle = c1Page.Rectangles(h)
                    If iRec.HtmlItems.Count > 0 Then
                        Dim hItem As InDesign.HtmlItem = iRec.HtmlItems(1)
                        If hItem.HtmlContent = "" Then
                            Dim fstylcol As InDesign.ObjectStyles
                            fstylcol = app.ActiveDocument.ObjectStyles
                            Dim styName As String = "idfix_page" + i.ToString() + "_widget" + h.ToString()
                            Dim styFound As Boolean = False
                            Dim styIndex As Integer = 0
                            For Each s As InDesign.ObjectStyle In fstylcol
                                If s.Name = styName Then
                                    styFound = True
                                    styIndex = s.Index
                                End If
                            Next
                            Dim fStyle As InDesign.ObjectStyle
                            If styFound = False Then
                                fStyle = fstylcol.Add()
                                fStyle.Name = styName
                            Else
                                fStyle = app.ActiveDocument.ObjectStyles(styIndex)
                            End If
                            iRec.ApplyObjectStyle(fStyle, True, True)
                        

                        End If
                    End If
                Next
            Next

            'export to html
            app.ActiveDocument.HTMLExportPreferences.CSSExportOption = InDesign.idStyleSheetExportOption.idEmbeddedCSS
            app.ActiveDocument.HTMLExportPreferences.IncludeCSSDefinition = True
            app.ActiveDocument.HTMLExportPreferences.PreserveLocalOverride = True
            app.ActiveDocument.HTMLExportPreferences.ExportOrder = InDesign.idExportOrder.idLayoutOrder
            app.ActiveDocument.HTMLExportPreferences.ExportSelection = False

            app.ActiveDocument.HTMLExportPreferences.ViewDocumentAfterExport = False

            app.ActiveDocument.Export(InDesign.idExportFormat.idHTML, expDir + "\001.html", False, , , True)


            Dim tabcssCol As New ArrayList
            Dim tabhtmlCol As New ArrayList
            'extract table html 
            If tableCol.Count > 0 Then
                Dim tHtml As String = File.ReadAllText(expDir + "\001.html")
                Dim cssTxt As String = File.ReadAllText(expDir + "\001-web-resources\css\001.css")

                tHtml = tHtml.Replace(vbCr, " ").Replace(vbLf, " ")

                cssTxt = cssTxt.Replace(vbCr, " ").Replace(vbLf, "")


                For i As Integer = 0 To tableCol.Count - 1
                    Dim tid As Integer = i + 1
                    Dim getTxt As String = ""
                    Dim rgx As New System.Text.RegularExpressions.Regex("<table\sid=""table-" + tid.ToString() + """(.+?)</table>", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    If rgx.Match(tHtml).Success Then
                        getTxt = rgx.Match(tHtml).Value
                    End If
                    tabhtmlCol.Add(getTxt)
                    'check class and update css
                    rgx = New System.Text.RegularExpressions.Regex("class=""(.+?)""", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                    Dim mCol As System.Text.RegularExpressions.MatchCollection
                    mCol = rgx.Matches(getTxt)
                    For Each m As System.Text.RegularExpressions.Match In mCol
                        Dim mTxt As String = m.Groups(1).Value
                        Dim mtcol() As String = mTxt.Split(" ")
                        For Each cName As String In mtcol
                            Dim crgx As New System.Text.RegularExpressions.Regex("\." + cName + "\s*\{(.+?)\}", System.Text.RegularExpressions.RegexOptions.IgnoreCase)
                            If crgx.Match(cssTxt).Success Then
                                add_table_style(tabcssCol, cName + "|" + crgx.Match(cssTxt).Value)
                            End If
                        Next
                    Next

                Next
            End If


            'get page size

            Dim viewWidth As Double = Double.Parse(app.ActiveDocument.DocumentPreferences.PageWidth.ToString())
            Dim ViewHeight As Double = Double.Parse(app.ActiveDocument.DocumentPreferences.PageHeight.ToString())

            viewWidth = (viewWidth * Convert.ToDouble(cPDF.b_dpi)) / 72
            ViewHeight = (ViewHeight * Convert.ToDouble(cPDF.b_dpi)) / 72



            Dim styCol As New ArrayList
            Dim posCol As New ArrayList
            Dim fontCol As New ArrayList
            'get text frame 
            'for each page 
            Dim pageHtml As String = ""
            Dim pageCss As String = ""

            Dim tFrameHtml As String = ""
            Dim lineHtml As String = ""
            Dim wordHtml As String = ""
            Dim lastClass As String = ""
            Dim wordtagClose As Boolean = False

            For i = 1 To pageCount
                probar_update("Page conversion...", pageCount, i)

                pageHtml = ""
                pageCss = ""
                posCol.Clear()
                Dim cPage As InDesign.Page = app.ActiveDocument.Pages(i)

                'for each text frame
                For f = 1 To cPage.TextFrames.Count
                    tFrameHtml = ""
                    Dim tFrame As InDesign.TextFrame
                    tFrame = cPage.TextFrames(f)

                    tFrame.Locked = False
                    'start line loop
                    For l = 1 To tFrame.Lines.Count
                        lineHtml = ""

                        Dim tLine As InDesign.Line
                        tLine = tFrame.Lines(l)
                        If tLine.Tables.Count > 0 Then
                            Dim mTable As InDesign.Table = tLine.Tables(1)
                        End If

                        Dim tFont As InDesign.Font = tLine.AppliedFont
                        Dim tColor As InDesign.Color = tLine.FillColor
                        Dim rgbValue As String = convert_rgb(tColor)

                        Dim fSize As Double = (tLine.PointSize * Convert.ToDouble(cPDF.b_dpi)) / 72

                        Dim lineStyle As String = fSize.ToString() + "|" + tFont.Name + "|" + rgbValue
                        Dim curClassName As String = check_style(styCol, lineStyle)



                        Dim posStyle As String = tLine.Baseline.ToString() + "|" + tLine.HorizontalOffset.ToString()
                        Dim curClass2 As String = "pls" + (posCol.Count + 1).ToString()
                        posCol.Add(curClass2 + "|" + posStyle)
                        Dim bline As Double = (tLine.Baseline * Convert.ToDouble(cPDF.b_dpi)) / 72
                        Dim hleft As Double = (tLine.HorizontalOffset * Convert.ToDouble(cPDF.b_dpi)) / 72

                        'pageCss += "." + curClass2 + " {position:absolute; top:" + bline.ToString() + "px; left:" + hleft.ToString() + "px; }" + Environment.NewLine
                        'lineHtml = "<div class=""" + curClassName + " " + curClass2 + """>"
                        lastClass = curClassName
                        'word loop
                        For w As Integer = 1 To tLine.Words.Count
                            wordHtml = ""
                            Dim tWord As InDesign.Word = tLine.Words(w)

                            fSize = (tWord.PointSize * Convert.ToDouble(cPDF.b_dpi)) / 72
                            rgbValue = convert_rgb(tWord.FillColor)

                            Dim wordStyle As String = fSize.ToString() + "|" + tWord.AppliedFont.Name + "|" + rgbValue

                            add_fontcol(fontCol, tWord.AppliedFont.Name, tWord.AppliedFont.Location)

                            Dim curWordclassName As String = check_style(styCol, wordStyle)

                            curClass2 = "pls" + (posCol.Count + 1).ToString()

                            bline = (tWord.Baseline * Convert.ToDouble(cPDF.b_dpi)) / 72
                            hleft = (tWord.HorizontalOffset * Convert.ToDouble(cPDF.b_dpi)) / 72
                            bline = ViewHeight - bline
                            'Dim charPath As InDesign.Objects = tWord.CreateOutlines(False)
                            'Dim bSize As Double = charPath(1).GeometricBounds(2) - charPath(1).GeometricBounds(0)
                            'bSize = (bSize * Convert.ToDouble(cPDF.b_dpi)) / 72
                            'bline = bline - bSize
                            'charPath(1).Delete()

                            pageCss += "." + curClass2 + " {position:absolute; bottom:" + bline.ToString() + "px; left:" + hleft.ToString() + "px; }" + Environment.NewLine

                            posStyle = tWord.Baseline.ToString() + "|" + tWord.HorizontalOffset.ToString()
                            posCol.Add(curClass2 + "|" + posStyle)

                            ' If curWordclassName <> curClassName Then
                            wordHtml += "<span class=""" + curWordclassName + " " + curClass2 + """>"
                            lastClass = curWordclassName
                            wordtagClose = True
                            ' Else
                            '  wordtagClose = False
                            ' End If

                            Dim lastBaseline As String = tWord.Baseline.ToString()
                            Dim charSpanStart As Boolean = False
                            'character loop
                            For c As Integer = 1 To tWord.Characters.Count
                                Dim tChar As InDesign.Character = tWord.Characters(c)
                                fSize = (tChar.PointSize * Convert.ToDouble(cPDF.b_dpi)) / 72
                                rgbValue = convert_rgb(tChar.FillColor)

                                Dim charStyle As String = fSize.ToString() + "|" + tChar.AppliedFont.Name + "|" + rgbValue
                                add_fontcol(fontCol, tChar.AppliedFont.Name, tChar.AppliedFont.Location)

                                Dim curCharClassName As String = check_style(styCol, charStyle)


                                'add char position
                                Dim charClass2 As String = ""
                                Dim charTxt As String = ""
                                charTxt = get_specialchar(tChar, cPage)




                                If lastBaseline <> tChar.Baseline.ToString() Or lastClass <> curCharClassName Then
                                    charClass2 = "pls" + (posCol.Count + 1).ToString()
                                    bline = (tChar.Baseline * Convert.ToDouble(cPDF.b_dpi)) / 72
                                    hleft = (tChar.HorizontalOffset * Convert.ToDouble(cPDF.b_dpi)) / 72

                                    bline = ViewHeight - bline
                                    'charPath = tChar.CreateOutlines(False)
                                    'bSize = charPath(1).GeometricBounds(2) - charPath(1).GeometricBounds(0)
                                    'bSize = (bSize * Convert.ToDouble(cPDF.b_dpi)) / 72
                                    'bline = bline - bSize
                                    'charPath(1).Delete()
                                    pageCss += "." + charClass2 + " {position:absolute; bottom:" + bline.ToString() + "px; left:" + hleft.ToString() + "px; }" + Environment.NewLine

                                    Dim charPosStyle As String = tChar.Baseline.ToString() + "|" + tChar.HorizontalOffset.ToString()
                                    posCol.Add(charClass2 + "|" + charPosStyle)
                                End If


                                If curCharClassName <> lastClass Or lastBaseline <> tChar.Baseline.ToString() Then
                                    ' If lastBaseline = tChar.Baseline.ToString() And curCharClassName <> lastClass Then
                                    'lineHtml += "<div class=""" + curCharClassName + """>" + tChar.Contents.ToString() + "</div>"
                                    'ElseIf lastBaseline <> tChar.Baseline.ToString() And curCharClassName = lastClass Then
                                    '  lineHtml += "<div class=""" + charClass2 + """>" + tChar.Contents.ToString() + "</div>"
                                    '  ElseIf lastBaseline <> tChar.Baseline.ToString() And curCharClassName <> lastClass Then
                                    If wordtagClose Then
                                        wordHtml += "</span>"
                                        wordtagClose = False
                                    End If




                                    wordHtml += "<span class=""" + curCharClassName + " " + charClass2 + """>" + charTxt + "</span>"
                                    ' End If
                                    charSpanStart = True
                                Else
                                    If wordHtml.IndexOf("</span>") <> -1 Then


                                        wordHtml = wordHtml.Insert(wordHtml.LastIndexOf("</span>"), charTxt)
                                    Else
                                        wordHtml += charTxt + "</span>"
                                        wordtagClose = False
                                    End If

                                End If
                                lastClass = curCharClassName
                                lastBaseline = tChar.Baseline.ToString()




                            Next 'char loop end





                            lineHtml += wordHtml

                        Next 'word loop end


                        'lineHtml += "</div>"
                        pageHtml += "<div>" + lineHtml + "</div>" 'Environment.NewLine
                    Next 'end line loop


                Next 'end text frame loop


                'search rectable box for inner html box
                Dim recHtml As String = ""
                For h = 1 To cPage.Rectangles.Count
                    Dim iRec As InDesign.Rectangle = cPage.Rectangles(h)
                    If iRec.HtmlItems.Count > 0 Then
                        Dim hItem As InDesign.HtmlItem = iRec.HtmlItems(1)
                        Dim hmTxt As String = hItem.HtmlContent
                        Dim rTop As Double = iRec.GeometricBounds(0)
                        Dim rLeft As Double = iRec.GeometricBounds(1)
                        Dim rWidth As Double = iRec.GeometricBounds(3) - iRec.GeometricBounds(1)
                        Dim rHeight As Double = iRec.GeometricBounds(2) - iRec.GeometricBounds(0)

                        rTop = (rTop * Convert.ToDouble(cPDF.b_dpi)) / 72
                        rLeft = (rLeft * Convert.ToDouble(cPDF.b_dpi)) / 72
                        rWidth = (rWidth * Convert.ToDouble(cPDF.b_dpi)) / 72
                        rHeight = (rHeight * Convert.ToDouble(cPDF.b_dpi)) / 72
                        If hmTxt <> "" Then
                            File.WriteAllText(htmlDir + "\page" + i.ToString() + "r" + h.ToString() + ".html", hmTxt)
                            Dim fmTxt As String = "<iframe style=""width:" + rWidth.ToString() + "px; height:" + rHeight.ToString() + "px;  position:absolute; left:" + rLeft.ToString() + "px; top:" + rTop.ToString() + "px; border:0;"" scrolling=""no"" src=""page" + i.ToString() + "r" + h.ToString() + ".html""></iframe>"
                            recHtml += fmTxt
                        Else 'widget
                            Dim styName As String = iRec.AppliedObjectStyle.Name
                            styName = styName.Replace("_", "-")
                            Dim exHtml As String = File.ReadAllText(expDir + "\001.html")
                            exHtml = exHtml.Replace(vbCr, " ").Replace(vbLf, " ")
                            If exHtml.IndexOf(styName) <> -1 Then
                                Dim sPoint As Integer = exHtml.IndexOf(styName)
                                Dim wgx As New Regex("<iframe(.+?)src=""001-web-resources/html/(.+?)/Assets/(.+?)""", RegexOptions.IgnoreCase)
                                If wgx.Match(exHtml, sPoint).Success Then
                                    Dim mtch As Match = wgx.Match(exHtml, sPoint)
                                    Dim widFld As String = mtch.Groups(2).Value
                                    Dim widHtml As String = mtch.Groups(3).Value
                                    Dim widSrc As String = widFld + "/Assets/" + widHtml

                                    'copy widget folder
                                    If Directory.Exists(htmlDir + "\" + widFld) Then
                                        Directory.Delete(htmlDir + "\" + widFld, True)

                                    End If
                                    gCls.DirectoryCopy(expDir + "\001-web-resources\html\" + widFld, htmlDir + "\" + widFld, True)

                                    Dim fmTxt As String = "<iframe style=""width:" + rWidth.ToString() + "px; height:" + rHeight.ToString() + "px;  position:absolute; left:" + rLeft.ToString() + "px; top:" + rTop.ToString() + "px; border:0;"" scrolling=""no"" src=""" + widSrc + """></iframe>"
                                    recHtml += fmTxt

                                End If
                            End If


                        End If

                    End If
                Next
                If recHtml <> "" Then
                    pageHtml += recHtml
                End If

                Dim pgTxt As String = "<!DOCTYPE HTML>" + Environment.NewLine + "<html xmlns=""http://www.w3.org/1999/xhtml"" xmlns:ibooks=""http://apple.com/ibooks/html-extensions"" xmlns:epub=""http://www.idpf.org/2007/ops"">" + Environment.NewLine
                pgTxt += "<head>" + Environment.NewLine + "<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" + Environment.NewLine
                pgTxt += "<title>" + "Page " + i.ToString() + "</title>" + Environment.NewLine
                pgTxt += "<link rel=""stylesheet"" type=""text/css"" href=""../styles/page.css"" />" + Environment.NewLine
                pgTxt += "<link rel=""stylesheet"" type=""text/css"" href=""../styles/page" + i.ToString() + ".css"" />" + Environment.NewLine
                pgTxt += "<meta name=""viewport"" content=""width=" + viewWidth.ToString() + ", height=" + ViewHeight.ToString() + """ />" + Environment.NewLine


                pgTxt += "</head>" + Environment.NewLine + "<body>" + Environment.NewLine
                pgTxt += "<img src=""../images/image" + i.ToString() + ".png"" width=""" + viewWidth.ToString() + """ height=""" + ViewHeight.ToString() + """ alt="""" />"

                'check hyphen word
                Dim pgRegex As New Regex("<span\sclass=""([^""]+)"">((?:(?!</span>).)*)</span><span\sclass=""([^""]+)"">((?:(?!</span>).)*)</span></div><div><span\sclass=""([^""]+)"">((?:(?!</span>).)*)</span><span\sclass=""([^""]+)"">((?:(?!</span>).)*)</span>", RegexOptions.IgnoreCase)
                Dim pgPoint As Integer = 0
                While pgRegex.Match(pageHtml, pgPoint).Success
                    Dim pgMatch As Match = pgRegex.Match(pageHtml, pgPoint)
                    Dim iPoint As Integer = pgMatch.Index
                    Dim cls1 As String = pgMatch.Groups(1).Value
                    Dim cls2 As String = pgMatch.Groups(3).Value
                    Dim cls3 As String = pgMatch.Groups(5).Value
                    Dim cls4 As String = pgMatch.Groups(7).Value

                    Dim cls1Num As Integer = Integer.Parse(Regex.Match(cls1, "pls(\d+)").Groups(1).Value)
                    Dim cls1Txt() As String = posCol(cls1Num - 1).ToString().Split("|")
                    Dim cls1Sty As String = cls1Txt(1) + "|" + cls1Txt(2)
                    Dim cls2Num As Integer = Integer.Parse(Regex.Match(cls2, "pls(\d+)").Groups(1).Value)
                    Dim cls2Txt() As String = posCol(cls2Num - 1).ToString().Split("|")
                    Dim cls2Sty As String = cls2Txt(1) + "|" + cls2Txt(2)
                    Dim cls3Num As Integer = Integer.Parse(Regex.Match(cls3, "pls(\d+)").Groups(1).Value)
                    Dim cls3Txt() As String = posCol(cls3Num - 1).ToString().Split("|")
                    Dim cls3Sty As String = cls3Txt(1) + "|" + cls3Txt(2)
                    Dim cls4Num As Integer = Integer.Parse(Regex.Match(cls4, "pls(\d+)").Groups(1).Value)
                    Dim cls4Txt() As String = posCol(cls4Num - 1).ToString().Split("|")
                    Dim cls4Sty As String = cls4Txt(1) + "|" + cls4Txt(2)

                    Dim sp1 As String = pgMatch.Groups(2).Value
                    Dim sp2 As String = pgMatch.Groups(4).Value
                    Dim sp3 As String = pgMatch.Groups(6).Value
                    Dim sp4 As String = pgMatch.Groups(8).Value

                    If sp1 = sp3 And sp2 = sp4 And cls1Sty = cls3Sty And cls2Sty = cls4Sty Then
                        pageHtml = pageHtml.Remove(iPoint, pgMatch.Length)
                        Dim sp1SubTxt As String = sp1.Substring(sp1.Length - 1, 1)
                        If sp1SubTxt <> "-" And sp1SubTxt <> "." And sp1SubTxt <> """" And sp1SubTxt <> "'" Then
                            sp1 = sp1 + "-"
                        End If
                        pageHtml = pageHtml.Insert(iPoint, "<span class=""" + cls1 + """>" + sp1 + "</span></div><div><span class=""" + cls2 + """>" + sp2 + "</span>")
                    End If

                    pgPoint = iPoint + 10
                End While

                pageHtml = pageHtml.Replace("&", "&amp;")
                pageHtml = pageHtml.Replace("</span>", "&#160;</span>")
                pageHtml = pageHtml.Replace("</div>", "</div>" + Environment.NewLine)
                pageHtml = gCls.Text2HexaConversion(pageHtml)
                pageHtml = pageHtml.Replace("", "")

                File.WriteAllText(htmlDir + "\page" + i.ToString() + ".html", pgTxt + pageHtml + "</body></html>")
                File.WriteAllText(styDir + "\page" + i.ToString() + ".css", pageCss)
            Next
            'end page 




            'write table content
            For t As Integer = 0 To tableCol.Count - 1
                Dim v1() As String = tableCol(t).ToString().Split("|")
                Dim cPnum As String = get_active_pagenumber(v1(0).ToString(), app.ActiveDocument)

                Dim hPage As String = "page" + cPnum + ".html"
                Dim tLeft As String = v1(1).ToString()
                Dim tBase As String = v1(2).ToString()
                Dim iTxt As String = "<div style=""position:absolute; left:" + tLeft + "px; bottom:" + tBase + "px;"">" + tabhtmlCol(t).ToString() + "</div>"
                Dim pTxt As String = File.ReadAllText(htmlDir + "\" + hPage)
                pTxt = pTxt.Replace("</body>", iTxt + "</body>")
                File.WriteAllText(htmlDir + "\" + hPage, pTxt)

            Next

            probar_update("Style generation...", 100, 90)
            'write common class 
            Dim cmnSty As String = ""
            For xs As Integer = 0 To styCol.Count - 1
                Dim vlist As String() = styCol(xs).ToString().Split("|")
                Dim curSty As String = vlist(1) + "|" + vlist(2) + "|" + vlist(3)
                cmnSty += "." + vlist(0) + " {" + Environment.NewLine
                cmnSty += "font-size: " + vlist(1) + "px;" + Environment.NewLine
                cmnSty += "font-family: " + vlist(2) + ";" + Environment.NewLine
                cmnSty += "color: " + vlist(3) + ";" + Environment.NewLine
                cmnSty += "} " + Environment.NewLine
            Next
            'copy font file
            For xf As Integer = 0 To fontCol.Count - 1
                Dim flist As String() = fontCol(xf).ToString().Split("|")
                Dim fntName As String = flist(0)
                Dim fntLocation As String = flist(1)
                cmnSty += "@font-face {" + Environment.NewLine
                cmnSty += "font-family: " + fntName + ";" + Environment.NewLine
                cmnSty += "src:url(" + Path.GetFileName(fntLocation) + ");" + Environment.NewLine
                cmnSty += "}" + Environment.NewLine

                File.Copy(fntLocation, styDir + "\" + Path.GetFileName(fntLocation), True)


            Next
            cmnSty += "body {" + Environment.NewLine
            cmnSty += "margin: 0;" + Environment.NewLine
            cmnSty += "left: 0;" + Environment.NewLine
            cmnSty += "position: absolute;" + Environment.NewLine
            cmnSty += "width: " + viewWidth.ToString() + "px;" + Environment.NewLine
            cmnSty += "height: " + ViewHeight.ToString() + "px;" + Environment.NewLine
            cmnSty += "}" + Environment.NewLine

            'add table styles
            For ts As Integer = 0 To tabcssCol.Count - 1
                Dim cssTemp As String = tabcssCol(ts).ToString().Split("|")(1).ToString()
                Dim rgx2 As New Regex("(\d+)px;", RegexOptions.IgnoreCase)
                Dim sPoint As Integer = 0
                Do While rgx2.Match(cssTemp, sPoint).Success
                    Dim mTx As Match = rgx2.Match(cssTemp, sPoint)
                    Dim iTx As Integer = Integer.Parse(mTx.Groups(1).Value)
                    iTx = (iTx * Convert.ToDouble(cPDF.b_dpi)) / 72
                    Dim vPoint As Integer = mTx.Index
                    cssTemp = cssTemp.Remove(vPoint, mTx.Value.Length)
                    cssTemp = cssTemp.Insert(vPoint, iTx.ToString() + "px;")
                    sPoint = vPoint + iTx.ToString().Length + 3
                    If sPoint > cssTemp.Length Then
                        sPoint = cssTemp.Length - 1
                    End If
                Loop
                cmnSty += cssTemp + Environment.NewLine
            Next


            File.WriteAllText(styDir + "\page.css", cmnSty)


            'cover image 
            If cPDF.b_coverpath <> "" Then
                If File.Exists(cPDF.b_coverpath) Then
                    File.Copy(cPDF.b_coverpath, imgDir + "\\cover.jpg", True)
                End If
            End If
            If Not File.Exists(imgDir + "\\cover.jpg") Then
                File.Copy(Application.StartupPath + "\\template\\cover.jpg", imgDir + "\\cover.jpg", True)
            End If

            'generate opf metadata
            Dim contX As String = "<manifest>" + vbCrLf
            contX += "<item id=""ncx"" href=""toc.ncx"" media-type=""application/x-dtbncx+xml"" />" + vbCrLf
            contX += "<item id=""pagecommon"" href=""styles/page.css"" media-type=""text/css"" />" + vbCrLf

            For i = 1 To pageCount
                contX += "<item id=""page" + i.ToString() + """ href=""html/page" + i.ToString() + ".html"" media-type=""application/xhtml+xml"" />" + vbCrLf
                contX += "<item id=""pagecss" + i.ToString() + """ href=""styles/page" + i.ToString() + ".css"" media-type=""text/css"" />" + vbCrLf
                contX += "<item id=""image" + i.ToString() + """ href=""images/image" + i.ToString() + ".png"" media-type=""image/png"" />" + vbCrLf
            Next
            contX += "<item id=""cover-image"" href=""images/cover.jpg"" media-type=""image/jpeg"" />" + vbCrLf
            contX += "</manifest>" + vbCrLf
            contX += "<spine toc=""ncx"">" + vbCrLf
            contX += "<itemref idref=""cover"" />" + vbCrLf
            For px = 1 To pageCount
                contX += "<itemref idref=""page" + px.ToString() + """ />" + vbCrLf
            Next
            contX += "</spine>" + vbCrLf

            Dim b_title As String = cPDF.b_title
            Dim b_author As String = cPDF.b_author
            If b_title = "" Then
                b_title = app.ActiveDocument.MetadataPreferences.DocumentTitle
            End If
            If b_author = "" Then
                b_author = app.ActiveDocument.MetadataPreferences.Author
            End If

            Dim rcont As String = File.ReadAllText(Path.Combine(Application.StartupPath, "template\\OEBPS\\content.opf"))
            rcont = rcont.Replace("<dc:title></dc:title>", "<dc:title>" + b_title + "</dc:title>")
            rcont = rcont.Replace("<dc:creator></dc:creator>", "<dc:creator>" + b_author + "</dc:creator>")
            rcont = rcont.Replace("<dc:publisher></dc:publisher>", "<dc:publisher>" + b_author + "</dc:publisher>")
            rcont = rcont.Replace("<dc:rights></dc:rights>", "<dc:rights>" + b_author + "</dc:rights>")
            Dim mdyDate As String = DateTime.Today.ToString("yyyy-MM-dd")
            rcont = rcont.Replace("<dc:date></dc:date>", "<dc:date>" + mdyDate + "</dc:date>")
            rcont = rcont.Replace("</metadata>", "</metadata>" + contX)
            File.WriteAllText(Path.Combine(oebDir, "content.opf"), rcont)


            Dim tocX As String = File.ReadAllText(Path.Combine(Application.StartupPath, "template\\OEBPS\\toc.ncx"))
            tocX = tocX.Replace("<docTitle><text></text></docTitle>", "<docTitle><text>" + b_title + "</text></docTitle>")
            tocX = tocX.Replace("<docAuthor><text></text></docAuthor>", "<docAuthor><text>" + b_author + "</text></docAuthor>")
            tocX = tocX.Replace("<navLabel><text></text></navLabel>", "<navLabel><text>" + b_title + "</text></navLabel>")
            File.WriteAllText(oebDir + "\\toc.ncx", tocX)

            File.Copy(Path.Combine(Application.StartupPath, "template\\META-INF\\container.xml"), Path.Combine(wrkDir, "META-INF\\container.xml"), True)
            File.Copy(Path.Combine(Application.StartupPath, "template\\META-INF\\com.apple.ibooks.display-options.xml"), Path.Combine(wrkDir, "META-INF\\com.apple.ibooks.display-options.xml"), True)
            File.Copy(Path.Combine(Application.StartupPath, "template\\mimetype"), Path.Combine(wrkDir, "mimetype"), True)


            'delete export dir
            Directory.Delete(expDir, True)

            probar_update("Create ePub Package...", 100, 95)
            'epub package
            Dim outepubPath As String = cPDF.b_outpath + "\\" + docName + ".epub"
            Dim psInfo As New ProcessStartInfo()
            Directory.SetCurrentDirectory(wrkDir)
            psInfo.CreateNoWindow = True
            psInfo.UseShellExecute = False
            psInfo.RedirectStandardOutput = True
            psInfo.WindowStyle = ProcessWindowStyle.Hidden
            psInfo.FileName = "ezip.exe"
            psInfo.Arguments = " -Xr9D """ + outepubPath + " "" mimetype *"
            Try
                Dim exeProcess As Process = Process.Start(psInfo)
                exeProcess.WaitForExit()
            Catch ex As Exception
                curErrLog += ex.Message
            End Try
            Directory.SetCurrentDirectory(Application.StartupPath)



            Directory.Delete(wrkDir, True)


        Catch ex As Exception
            curErrLog += ex.Message + Environment.NewLine
            Exit Sub
        End Try
    End Sub

    Private Sub MForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text += " " + Application.ProductVersion

        gCls.update_path_var()
        cmb_dpi.SelectedIndex = 0




    End Sub

    Private Sub kryptonButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kryptonButton2.Click
        Try
            Dim fld As New FolderBrowserDialog

            fld.Description = "Select output folder"
            fld.ShowDialog()
            If fld.SelectedPath <> "" Then
                txt_outfolder.Text = fld.SelectedPath

            End If


        Catch ex As Exception
            gCls.show_error(ex.Message)
            Return
        End Try
    End Sub

    Private Sub kryptonButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kryptonButton3.Click
        Try
            Dim fld As New OpenFileDialog
            fld.Title = "Select cover image file"
            fld.Filter = "JPG File|*.jpg"
            fld.ShowDialog()
            If fld.FileName <> "" Then
                txt_epubcover.Text = fld.FileName
            End If

        Catch ex As Exception
            gCls.show_error(ex.Message)
            Return
        End Try
    End Sub

    Private Sub kryptonButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kryptonButton4.Click
        Try
            Application.Exit()
            Exit Sub
        Catch ex As Exception
            gCls.show_error(ex.Message)
            Exit Sub
        End Try
    End Sub

    Private Sub linkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles linkLabel1.LinkClicked
        Try
            gCls.show_message(Application.ProductName + " " + Application.ProductVersion + Environment.NewLine + "Send your Feedbacks to : vickypatel2020@gmail.com" + Environment.NewLine)

        Catch ex As Exception
            gCls.show_error(ex.Message)
            Return
        End Try


    End Sub

    Private Sub kryptonButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles kryptonButton1.Click
        Try
            If txt_outfolder.Text = "" Then
                gCls.show_error("Select output folder path")
                Return
            End If

            If txt_epubcover.Text <> "" Then
                If Not File.Exists(txt_bookauthor.Text) Then
                    gCls.show_error("Cover image not found")
                    Return
                End If
            End If

            If Not Directory.Exists(txt_outfolder.Text) Then
                gCls.show_error("Output folder not found")
                Return
            End If

            Dim outpath As String = txt_outfolder.Text

            Dim doc_info As New Updf("", txt_booktitle.Text, txt_bookauthor.Text, "", "", outpath, txt_epubcover.Text, cmb_dpi.Text)

            curErrLog = ""
            _bw = New BackgroundWorker()
            _bw.WorkerReportsProgress = True
            _bw.WorkerSupportsCancellation = True
            AddHandler _bw.DoWork, AddressOf bw_Dowork
            AddHandler _bw.ProgressChanged, AddressOf bw_ProgressChanged
            AddHandler _bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted

             topmenu.Visible = True
            
            probar.Visible = True
            probar.Value = 0

            _bw.RunWorkerAsync(doc_info)

        Catch ex As Exception
            gCls.show_error(ex.Message)
            Return

        End Try

    End Sub

    Private Sub ReleaseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReleaseToolStripMenuItem.Click

       


    End Sub

    Private Sub KryptonButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KryptonButton5.Click
        Try
            app = New InDesign.Application


            If app.Selection().Count < 1 Then
                Exit Sub
            End If
            Try

                For i As Integer = 1 To app.Selection().Count
                    Dim tFrame As InDesign.TextFrame = app.Selection(i)
                    tFrame.Label = "idfix_img"
                Next
            Catch ex As Exception

            End Try

        Catch ex As Exception
            gCls.show_error(ex.Message)
            Exit Sub
        End Try
    End Sub

    Private Sub KryptonButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KryptonButton6.Click
        Try
            app = New InDesign.Application


            If app.Selection().Count < 1 Then
                Exit Sub
            End If
            Try

                For i As Integer = 1 To app.Selection().Count
                    Dim tFrame As InDesign.TextFrame = app.Selection(i)
                    tFrame.Label = ""
                Next
            Catch ex As Exception

            End Try

        Catch ex As Exception
            gCls.show_error(ex.Message)
            Exit Sub
        End Try
    End Sub
End Class
