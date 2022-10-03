'@(h) usrMediaList Ver 1.0 2022/08/30 アリエル
'@(s)
''指定された条件（売上期間、企画類、媒体類、サブ媒体類、部門類、売経路類）でSQL文を作ってから売上からの商品企画ごとの商品媒体順位リスト作るプログラム
''企画類、媒体類、サブ媒体類、売経路類は範囲または指定されるIDで検索できる
''金額は税込み/抜き選択できる
''結果行数は多かったらトップまたはワースト商品だけ出力できる
Imports CustomClass
Imports System.Drawing
Imports System.Windows.Forms
Imports Custom.SqlReaders
Imports System.IO
Imports GrapeCity.ActiveReports
Imports System.Threading
Imports System.Data
Imports Syncfusion.Pdf
Imports Syncfusion.Pdf.Grid
Imports System.Drawing

Public Class usrMediaList
    Private StartId As Integer                    '起動システムID
    Private StartUser As String                   '起動ユーザーID
    Private ClearDT As DataTable                  '初期化データテーブル
    Private SQLR As SqlReader                     'SQL分読むため
    Private MediumHD() As Spread.Listheader       '媒体リストヘッダー
    Private MediumExtraHD() As Spread.Listheader    'サブ媒体リストヘッダー
    Private PlanHD() As Spread.Listheader         '企画リストヘッダー
    Private STypeHD() As Spread.Listheader        '売経路リストヘッダー
    Public MediumList As String                   '媒体リスト
    Public MeExtraList As String                  'サブ媒体リスト
    Public PlanList As String                     '企画リスト
    Public STypeList As String                    '売経路リスト
     Private Enum eMei
        ID      '入力・非表示
        名称    '入力・非表示
    End Enum

    Private Sub usrMediaList_Load(sender As Object, e As EventArgs) Handles Me.Load
        StartId = 1
        StartUser = "TestUser"
        alignitem = 0  '右寄せ
        InitSet()
        Dim SP As Spread = New Spread(spdMedium)
        Dim DT As DataTable = ClearDT.Clone
        Dim DT2 As DataTable = ClearDT.Clone
        Dim DT3 As DataTable = ClearDT.Clone
        Dim DT4 As DataTable = ClearDT.Clone
        MediumHD = ListMedium("Medium")
        SP.SetSPREAD(MediumHD, DT, Spread.eSetMode.Edit, False, , , , , False, False)
        spdMedium.ActiveSheet.ColumnHeader.Columns(1).Width = 243
        spdMedium.ActiveSheet.DataAllowAddNew = True
        Dim SP2 As Spread = New Spread(spdMediumType)
        MediumExtraHD = ListMedium("MediumbExtra")
        SP2.SetSPREAD(MediumExtraHD, DT2, Spread.eSetMode.Edit, False, , , , , False, False)
        spdMediumType.ActiveSheet.ColumnHeader.Columns(1).Width = 243
        spdMediumType.ActiveSheet.DataAllowAddNew = True
        Dim SP3 As Spread = New Spread(spdPlan)
        PlanHD = ListMedium("Plan")
        SP3.SetSPREAD(PlanHD, DT3, Spread.eSetMode.Edit, False, , , , , False, False)
        spdPlan.ActiveSheet.ColumnHeader.Columns(1).Width = 243
        spdPlan.ActiveSheet.DataAllowAddNew = True
        Dim SP4 As Spread = New Spread(spdSType)
        STypeHD = ListMedium("SType")
        SP4.SetSPREAD(STypeHD, DT4, Spread.eSetMode.Edit, False, , , , , False, False)
        spdSType.ActiveSheet.ColumnHeader.Columns(1).Width = 243
        spdSType.ActiveSheet.DataAllowAddNew = True
        rdbPlan.Checked = True
        AddHandler txtStrMeId.Enter, AddressOf EnterEvent
        AddHandler txtStrMeNm.Enter, AddressOf EnterEvent
        AddHandler txtEndMeId.Enter, AddressOf EnterEvent
        AddHandler txtEndMeNm.Enter, AddressOf EnterEvent
        AddHandler txtStrPlan.Enter, AddressOf EnterEvent
        AddHandler txtStrPlanNm.Enter, AddressOf EnterEvent
        AddHandler txtEndPlan.Enter, AddressOf EnterEvent
        AddHandler txtEndPlanNm.Enter, AddressOf EnterEvent
        AddHandler txtStrSType.Enter, AddressOf EnterEvent
        AddHandler txtStrSTypeNm.Enter, AddressOf EnterEvent
        AddHandler txtEndSType.Enter, AddressOf EnterEvent
        AddHandler txtEndSTypeNm.Enter, AddressOf EnterEvent
        AddHandler txtStrMedium.Enter, AddressOf EnterEvent
        AddHandler txtStrMediumNm.Enter, AddressOf EnterEvent
        AddHandler txtEndMedium.Enter, AddressOf EnterEvent
        AddHandler txtEndMediumNm.Enter, AddressOf EnterEvent
        AddHandler txtStrMeExtra.Enter, AddressOf EnterEvent
        AddHandler txtStrMeExtraNm.Enter, AddressOf EnterEvent
        AddHandler txtEndMeExtra.Enter, AddressOf EnterEvent
        AddHandler txtEndMeExtraNm.Enter, AddressOf EnterEvent
        AddHandler spdMedium.Enter, AddressOf EnterEvent
        AddHandler spdMediumType.Enter, AddressOf EnterEvent
        AddHandler spdSType.Enter, AddressOf EnterEvent
        AddHandler spdPlan.Enter, AddressOf EnterEvent
        AddHandler dtcStrDate.Enter, AddressOf EnterEvent
        AddHandler dtcEndDate.Enter, AddressOf EnterEvent
        AddHandler dtcStrBDate.Enter, AddressOf EnterEvent
        AddHandler dtcEndBDate.Enter, AddressOf EnterEvent
        AddHandler txtStrSType.Leave, AddressOf LeaveEvent
        AddHandler txtEndSType.Leave, AddressOf LeaveEvent
        AddHandler txtStrMeId.Leave, AddressOf LeaveEvent
        AddHandler txtEndMeId.Leave, AddressOf LeaveEvent
        AddHandler txtStrPlan.Leave, AddressOf LeaveEvent
        AddHandler txtEndPlan.Leave, AddressOf LeaveEvent
        AddHandler txtStrMedium.Leave, AddressOf LeaveEvent
        AddHandler txtEndMedium.Leave, AddressOf LeaveEvent
        AddHandler txtStrMeExtra.Leave, AddressOf LeaveEvent
        AddHandler txtEndMeExtra.Leave, AddressOf LeaveEvent
        AddHandler spdMedium.Leave, AddressOf LeaveEvent
        AddHandler spdMediumType.Leave, AddressOf LeaveEvent
        AddHandler spdPlan.Leave, AddressOf LeaveEvent
        AddHandler spdSType.Leave, AddressOf LeaveEvent
        AddHandler btnConfir.Click, AddressOf Print
        AddHandler btnPrint.Click, AddressOf Print
        AddHandler btnStrMediumPopUp.Click, AddressOf PopUp
        AddHandler btnEndMediumPopUp.Click, AddressOf PopUp
        AddHandler btnStrMePopUp.Click, AddressOf PopUp
        AddHandler btnEndMePopUp.Click, AddressOf PopUp
        AddHandler btnStrMeExtraPopUp.Click, AddressOf PopUp
        AddHandler btnEndMeExtraPopUp.Click, AddressOf PopUp
        AddHandler btnStrPlanPopUp.Click, AddressOf PopUp
        AddHandler btnEndPlanPopUp.Click, AddressOf PopUp
        AddHandler btnStrSTypePopUp.Click, AddressOf PopUp
        AddHandler btnEndSTypePopUp.Click, AddressOf PopUp
    End Sub

    ''' <summary>
    ''' 初期化
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub InitSet()
        dtcStrDate.CustomFormat = "yyyy年MM月dd日 HH時mm分ss秒"
        dtcEndDate.CustomFormat = "yyyy年MM月dd日 HH時mm分ss秒"
        dtcStrDate.Value = Date.Today
        dtcEndDate.Value = Date.Today & " 23:59:59"
        dtcStrBDate.CustomFormat = "yyyy年MM月dd日 HH時mm分ss秒"
        dtcEndBDate.CustomFormat = "yyyy年MM月dd日 HH時mm分ss秒"
        dtcStrBDate.Value = Date.Today
        dtcEndBDate.Value = Date.Today & " 23:59:59"
        txtStrMeId.TextAlign = alignitem
        txtEndMeId.TextAlign = alignitem
        txtStrMedium.TextAlign = alignitem
        txtEndMedium.TextAlign = alignitem
        txtStrMeExtra.TextAlign = alignitem
        txtEndMeExtra.TextAlign = alignitem
        txtStrPlan.TextAlign = alignitem
        txtEndPlan.TextAlign = alignitem
        txtStrSType.TextAlign = alignitem
        txtEndSType.TextAlign = alignitem
        cboRankTarget.SelectedIndex = 1
        cboRank.SelectedIndex = 0
        txtRankCnt.Text = 50
        cboUriKbn.SelectedIndex = 2
        rdbHanniB.Checked = True
        rdbHanniBa.Checked = True
        rdbHanniK.Checked = True
        rdbHanniJ.Checked = True
        txtStrMedium.Text = "0"
        txtEndMedium.Text = "zzzzzz"
        txtStrMeExtra.Text = "0"
        txtEndMeExtra.Text = "zzzzzz"
        txtStrPlan.Text = "0"
        txtEndPlan.Text = "zzzzzz"
        txtStrSType.Text = "0"
        txtEndSType.Text = "zzzz"
        txtStrMeId.Text = "0"
        txtEndMeId.Text = "zzzzzz"
        PlanList = vbNullString
        MediumList = vbNullString
        MediumbunList = vbNullString
        STypeList = vbNullString
        SQLR = New SqlReader
        ClearDT = New DataTable
        ClearDT.Columns.Add("ID", GetType(String))
        ClearDT.Columns.Add("NAME", GetType(String))
    End Sub

    ''' <summary>
    ''' spreadSheet取得時の列情報作成用
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function ListMedium(Type As String) As Spread.Listheader()
        Dim HD(System.Enum.GetNames(GetType(eMei)).Length - 1) As Spread.Listheader
        For Col As Integer = 0 To System.Enum.GetNames(GetType(eMei)).Length - 1
            Dim Item As New Spread.Listheader
            Select Case Col
                Case eMei.ID
                    Select Case Type
                        Case "Medium"
                            Item.Text = " 媒体ID"
                        Case "Plan"
                            Item.Text = " 企画ID"
                        Case "MediumbExtra"
                            Item.Text = " サブ媒体ID"
                        Case "SType"
                            Item.Text = " 売経路ID"
                    End Select
                    Item.FieldName = "class"
                    Item.SQL = "''"
                    Item.Type = Spread.eCelltype.TextBox
                    Item.MaxLength = 10
                    Item.Length = 10
                    Item.TextAlign = FarPoint.Win.Spread.CellHorizontalAlignment.Right
                Case eMei.名称
                    Select Case Type
                        Case "Medium"
                            Item.Text = " 媒体名称"
                        Case "Plan"
                            Item.Text = " 企画名称"
                        Case "Mediumbtype"
                            Item.Text = " サブ媒体名称"
                        Case "SType"
                            Item.Text = " 売経路名称"
                    End Select
                    Item.FieldName = "name"
                    Item.Type = Spread.eCelltype.TextBox
                    Item.MaxLength = 100
                    Item.Length = 32
                    Item.TextAlign = FarPoint.Win.Spread.CellHorizontalAlignment.Left
                    Item.FldSwkbn = Spread.eFldSwkbn.読取専用
            End Select
            HD(Col) = Item
        Next
        Return HD
    End Function

    Private Sub RadioChanged(sender As Object, e As EventArgs) Handles rdbHanniJ.CheckedChanged, rdbHanniK.CheckedChanged, rdbHanniB.CheckedChanged, rdbHanniBa.CheckedChanged, _
                                                                       rdbShiteiB.CheckedChanged, rdbShiteiBa.CheckedChanged, rdbShiteiJ.CheckedChanged, rdbShiteiK.CheckedChanged, _
                                                                       rdbPlan.CheckedChanged, rdbmedium.CheckedChanged
        Select Case sender.name
            Case rdbHanniB.Name
                txtStrMedium.Enabled = (rdbHanniB.Checked)
                btnStrMePopUp.Enabled = (rdbHanniB.Checked)
                txtEndMedium.Enabled = (rdbHanniB.Checked)
                btnEndMePopUp.Enabled = (rdbHanniB.Checked)
                spdMedium.Enabled = Not (rdbHanniB.Checked)
            Case rdbHanniBa.Name
                txtStrMeExtra.Enabled = (rdbHanniBa.Checked)
                btnStrMeExtraPopUp.Enabled = (rdbHanniBa.Checked)
                txtEndMeExtra.Enabled = (rdbHanniBa.Checked)
                btnEndMeExtraPopUp.Enabled = (rdbHanniBa.Checked)
                spdMediumType.Enabled = Not (rdbHanniBa.Checked)
            Case rdbHanniK.Name
                txtStrPlan.Enabled = (rdbHanniK.Checked)
                btnStrPlanPopUp.Enabled = (rdbHanniK.Checked)
                txtEndPlan.Enabled = (rdbHanniK.Checked)
                btnEndPlanPopUp.Enabled = (rdbHanniK.Checked)
                spdPlan.Enabled = Not (rdbHanniK.Checked)
            Case rdbHanniJ.Name
                txtStrSType.Enabled = (rdbHanniJ.Checked)
                btnStrSTypePopUp.Enabled = (rdbHanniJ.Checked)
                txtEndSType.Enabled = (rdbHanniJ.Checked)
                btnEndSTypePopUp.Enabled = (rdbHanniJ.Checked)
                spdSType.Enabled = Not (rdbHanniJ.Checked)
            Case rdbPlan.Name
                cboPrintType.Items.Clear()
                cboPrintType.Items.Add("企画別媒体表")
                cboPrintType.Items.Add("企画のみ表")
                cboPrintType.SelectedIndex = 0
            End Select
    End Sub

    Private Sub LeaveEvent(sender As Object, e As EventArgs)
        Select Case sender.name
            Case txtStrMeId.Name
                txtStrMeNm.Text = getName(sender)
            Case txtEndMeId.Name
                txtEndMeNm.Text = getName(sender)
            Case txtStrMedium.Name
                txtStrMediumNm.Text = getName(sender)
            Case txtEndMedium.Name
                txtEndMediumNm.Text = getName(sender)
            Case txtStrMeExtra.Name
                txtStrMeExtraNm.Text = getName(sender)
            Case txtEndMeExtra.Name
                txtEndMeExtraNm.Text = getName(sender)
            Case txtStrPlan.Name
                txtStrPlanNm.Text = getName(sender)
            Case txtEndPlan.Name
                txtEndPlanNm.Text = getName(sender)
            Case txtStrSType.Name
                txtStrSTypeNm.Text = getName(sender)
            Case txtEndSType.Name
                txtEndSTypeNm.Text = getName(sender)
        End Select
    End Sub

    ''' <summary>
    ''' 名称獲得
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getName(sender As Object) As String
        Dim str As String = vbNullString
        Select Case sender.name
            Case txtStrMeId.Name, txtEndMeId.Name
                SQLR.AddItem("x1") = IIf(sender.name = txtStrMeId.Name, txtStrMeId.Text, txtEndMeId.Text)
                If SQLR.Read("select name from Department where syid=:xid and did=:x1") = 1 Then
                    str = OFtoS(SQLR.DTable.Rows(0), "name")
                End If
            Case txtStrMedium.Name, txtEndMedium.Name
                SQLR.AddItem("x1") = IIf(sender.name = txtStrMedium.Name, txtStrMedium.Text, txtEndMedium.Text)
                If SQLR.Read("select name from medium where syid=:xid and bid=:x1") = 1 Then
                    str = OFtoS(SQLR.DTable.Rows(0), "name")
                End If
            Case txtStrMeExtra.Name, txtEndMeExtra.Name
                SQLR.AddItem("xclass") = 999999
                SQLR.AddItem("x1") = IIf(sender.name = txtStrMeExtra.Name, txtStrMeExtra.Text, txtEndMeExtra.Text)
                If SQLR.Read("select name from extraItem where syid=:xid and class=:xclass and fid=:x1") = 1 Then
                    str = OFtoS(SQLR.DTable.Rows(0), "name")
                End If
            Case txtStrPlan.Name, txtEndPlan.Name
                SQLR.AddItem("xclass") = 999998
                SQLR.AddItem("x1") = IIf(sender.name = txtStrPlan.Name, txtStrPlan.Text, txtEndPlan.Text)
                If SQLR.Read("select name from extraItem where syid=:xid and class=:xclass and fid=:x1") = 1 Then
                    str = OFtoS(SQLR.DTable.Rows(0), "name")
                End If
            Case txtStrSType.Name, txtEndSType.Name
                SQLR.AddItem("x1") = IIf(sender.name = txtStrSType.Name, txtStrSType.Text, txtEndSType.Text)
                If SQLR.Read("select gettypename(:xid,0,:x1) name from dual") = 1 Then
                    str = OFtoS(SQLR.DTable.Rows(0), "name")
                End If
        End Select
        Return str
    End Function

    ''' <summary>
    ''' 印刷
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Public Sub Print(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Dim IsFile As Boolean = False
        If sender.name.ToString.IndexOf("File") > -1 Then IsFile = True
        Dim DT As DataTable = GetData(IsFile)
        If (rdbShiteiB.Checked = True And MediumList <> vbNullString) Or (rdbShiteiBa.Checked = True And MeExtraList <> vbNullString) Or _
           (rdbShiteiJ.Checked = True And STypeList <> vbNullString) Or (rdbShiteiK.Checked = True And PlanList <> vbNullString) Then ListName()
        If DT.Rows.Count < 1 Then
            MessageBox("対象ﾃﾞｰﾀがありません")
        Else
            Select Case sender.name
                Case  btnConfir.Name, btnPrint.Name
                    Dim frm As New frmViewer
                    Dim rpt As New Object
                    If rdbPlan.Checked Then
                            rpt = New rptPlanList(Me)
                    Else
                            rpt = New rptPlanStandings(Me)
                    End If

                    Dim arc As New ActiveReportClass(rpt.Document)
                    arc.DataSource = DT
                    rpt.arc = arc
                    'プリンターの設定--------------------------------------------------------------------
                    rpt.PageSettings.PaperKind = Printing.PaperKind.A4
                    rpt.PageSettings.Orientation = GrapeCity.ActiveReports.Document.Section.PageOrientation.Landscape
                    '余白ととじしろの設定は後で考える
                    rpt.PageSettings.Margins.Top = GrapeCity.ActiveReports.SectionReport.CmToInch(ConfigFile.Instance.PrintMargin.Top * 0.1)
                    rpt.PageSettings.Margins.Bottom = GrapeCity.ActiveReports.SectionReport.CmToInch(ConfigFile.Instance.PrintMargin.Bottom * 0.05)
                    rpt.PageSettings.Margins.Left = GrapeCity.ActiveReports.SectionReport.CmToInch(ConfigFile.Instance.PrintMargin.Left * If(cboPrintType.SelectedIndex = 0, 0.18, 0.55))
                    rpt.PageSettings.Margins.Right = GrapeCity.ActiveReports.SectionReport.CmToInch(ConfigFile.Instance.PrintMargin.Right * 0.07)
                    rpt.PageSettings.Gutter = GrapeCity.ActiveReports.SectionReport.CmToInch(0)
                    Select Case sender.name
                        Case btnConfir.Name
                            rpt.Run()
                            frm.SetDocument = rpt.Document
                            frm.SetFlg = True
                            frm.Show()
                        Case Else
                            rpt.Run()
                            If rdbPlan.Checked Then
                                CType(rpt, rptPlanList).Document.Print(True, True, True)
                            Else
                                CType(rpt, rptPlanStandings).Document.Print(True, True, True)
                            End If
                    End Select
                Case Else
                    If File(sender, DT) = False Then MessageBox("ﾌｧｲﾙ出力中にｴﾗｰが発生しました")
            End Select
        End If
        Me.Cursor = Cursors.Default
    End Sub

    ''' <summary>
    ''' リスト作る
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub ListName()
        If MediumList <> vbNullString Then
            MediumList = Replace(MediumList, "'", "")
            Dim mediumIDs As String() = MediumList.Split(New Char() {","c})
            MediumList = vbNullString
            Dim i As Integer = 0
            For Each ID In mediumIDs
                SQLR.AddItem("xbid") = Trim(ID)
                SQLR.Read("select name from medium where syid =:xid and bid =:xbid")
                MediumList = MediumList + "[" + Trim(ID) + "] " + SQLR.DTable.Rows(0).Item(0) + "、"
                i = i + 1
                If i = 4 Then Exit For
            Next
            MediumList = MediumList.Substring(0, MediumList.Length - 1)
            If mediumIDs.Count > 4 Then MediumList = MediumList + "。。。"
        End If
        If MeExtraList <> vbNullString Then
            MeExtraList = Replace(MeExtraList, "'", "")
            Dim MeExtraIDs As String() = MeExtraList.Split(New Char() {","c})
            MeExtraList = vbNullString
            Dim i As Integer = 0
            For Each ID In MeExtraIDs
                SQLR.AddItem("xmememo") = Trim(ID)
                SQLR.Read("select getextraItemname(:xid, 999998, :xmememo) from dual")
                MeExtraList = MeExtraList + "[" + Trim(ID) + "] " + SQLR.DTable.Rows(0).Item(0) + "、"
                i = i + 1
                If i = 4 Then Exit For
            Next
            MeExtraList = MeExtraList.Substring(0, MeExtraList.Length - 1)
            If MeExtraIDs.Count > 4 Then MeExtraList = MeExtraList + "。。。"
        End If
        If PlanList <> vbNullString Then
            PlanList = Replace(PlanList, "'", "")
            Dim PlanIDs As String() = PlanList.Split(New Char() {","c})
            PlanList = vbNullString
            Dim i As Integer = 0
            For Each ID In PlanIDs
                SQLR.AddItem("xplan") = Trim(ID)
                SQLR.Read("select getextraItemname(:xid, 999999, :xplan) from dual")
                PlanList = PlanList + "[" + Trim(ID) + "] " + SQLR.DTable.Rows(0).Item(0) + "、"
                i = i + 1
                If i = 4 Then Exit For
            Next
            PlanList = PlanList.Substring(0, PlanList.Length - 1)
            If PlanIDs.Count > 4 Then PlanList = PlanList + "。。。"
        End If
        If STypeList <> vbNullString Then
            STypeList = Replace(STypeList, "'", "")
            Dim STypeIDs As String() = STypeList.Split(New Char() {","c})
            STypeList = vbNullString
            Dim i As Integer = 0
            For Each ID In STypeIDs
                SQLR.AddItem("xplan") = Trim(ID)
                SQLR.Read("select gettypename(:xid,0,:xplan) name from dual")
                STypeList = STypeList + "[" + Trim(ID) + "] " + SQLR.DTable.Rows(0).Item(0) + "、"
                i = i + 1
                If i = 4 Then Exit For
            Next
            STypeList = STypeList.Substring(0, STypeList.Length - 1)
            If STypeIDs.Count > 4 Then STypeList = STypeList + "。。。"
        End If
    End Sub

    ''' <summary>
    ''' データテーブル作成
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetData(Optional isFile As Boolean = False) As DataTable
        Dim DT As New DataTable
        Dim str As String = SQLR.ReadSQLNum(1)'リスト作るSQL (修正前)
        SQLR.AddItem("xsday") = Format(dtcStrDate.Value, "yyyy/MM/dd HH:mm:ss")
        SQLR.AddItem("xeday") = Format(dtcEndDate.Value, "yyyy/MM/dd HH:mm:ss")
        SQLR.AddItem("xsmday") = Format(dtcStrBDate.Value, "yyyy/MM/dd HH:mm:ss")
        SQLR.AddItem("xemday") = Format(dtcEndBDate.Value, "yyyy/MM/dd HH:mm:ss")
        SQLR.AddItem("xgarry") = alignitem
        'サブ媒体開始日---------------
        str = str.Replace("@MDAY", "and (select sday from baikanri where syid=:xid and bid=li.bid and bkid=li.bkid) between to_date(:xsmday, 'YYYY/MM/DD HH24:MI:SS') and to_date(:xemday, 'YYYY/MM/DD HH24:MI:SS')")

        '部門-------------------
        SQLR.AddItem("xbflg") = 1
        If alignitem = 0 Then
            str = str.Replace("@DEPARTMENT", " and sl.did between :xsdid and :xedid")
            SQLR.AddItem("xsdid") = Trim(txtStrMeId.Text)
            SQLR.AddItem("xedid") = Trim(txtEndMeId.Text)
        Else
            str = str.Replace("@DEPARTMENT", " and getItems(sl.did,6,1) between :xsdid and :xedid")
            SQLR.AddItem("xsdid") = Trim(txtStrMeId.Text).PadLeft(6)
            SQLR.AddItem("xedid") = Trim(txtEndMeId.Text).PadLeft(6)
        End If

        'サブ媒体-------------------------------------
        MeExtraList = vbNullString
        For i = 0 To spdMediumType.ActiveSheet.Rows.Count - 2
            MeExtraList = MeExtraList + "'" + IIf(alignitem = 0, spdMediumType.ActiveSheet.Cells(i, 0).Text, spdMediumType.ActiveSheet.Cells(i, 0).Text.PadLeft(6)) + "'"
            If i <> spdMediumType.ActiveSheet.Rows.Count - 2 Then MeExtraList = MeExtraList + ","
        Next
        If alignitem = 0 Then
            str = str.Replace("MEDIUMEXTRA", IIf(rdbHanniBa.Checked Or MeExtraList = vbNullString, "and getmemo(:xid,6,li.bid,106001) between :xsmeExtra and :xemeExtra ", "and getmemo(:xid,6,li.bid,106001) in (" + MeExtraList + ") "))
            SQLR.AddItem("xsmeExtra") = Trim(txtStrMeExtra.Text)
            SQLR.AddItem("xemeExtra") = Trim(txtEndMeExtra.Text)
        Else
            str = str.Replace("@MEDIUMEXTRA", IIf(rdbHanniBa.Checked Or MeExtraList = vbNullString, "and getItems(getmemo(:xid,6,li.bid,106001),6,1) between :xsmeExtra and :xemeExtra ", "and getItems(getmemo(:xid,6,li.bid,106001),6,1) in (" + MeExtraList + ") "))
            SQLR.AddItem("xsmeExtra") = Trim(txtStrMeExtra.Text).PadLeft(6)
            SQLR.AddItem("xemeExtra") = Trim(txtEndMeExtra.Text).PadLeft(6)
        End If

        '媒体-------------------
        MediumList = vbNullString
        SQLR.AddItem("xmeflg") = 1

        For i = 0 To spdMedium.ActiveSheet.Rows.Count - 2
            MediumList = MediumList + "'" + IIf(alignitem = 0, spdMedium.ActiveSheet.Cells(i, 0).Text, spdMedium.ActiveSheet.Cells(i, 0).Text.PadLeft(6)) + "'"
            If i <> spdMedium.ActiveSheet.Rows.Count - 2 Then MediumList = MediumList + ","
        Next
        If alignitem = 0 Then
            str = str.Replace("@MEDIUM", IIf(rdbHanniB.Checked Or MediumList = vbNullString, "and li.bid between :xsbid and :xebid  ", "and li.bid in (" + MediumList + ") "))
            If rdbHanniB.Checked Then
                SQLR.AddItem("xsbid") = Trim(txtStrMedium.Text)
                SQLR.AddItem("xebid") = Trim(txtEndMedium.Text)
            End If
        Else
            str = str.Replace("@MEDIUM", IIf(rdbHanniB.Checked Or MediumList = vbNullString, "and getItems(li.bid,6,1) between :xsbid and :xebid  ", "and getItems(li.bid,6,1) in (" + MediumList + ") "))
            If rdbHanniB.Checked Then
                SQLR.AddItem("xsbid") = Trim(txtStrMedium.Text).PadLeft(6)
                SQLR.AddItem("xebid") = Trim(txtEndMedium.Text).PadLeft(6)
            End If

        End If

        '企画-------------------------------------
        PlanList = vbNullString
        For i = 0 To spdPlan.ActiveSheet.Rows.Count - 2
            PlanList = PlanList + "'" + IIf(alignitem = 0, spdPlan.ActiveSheet.Cells(i, 0).Text, spdPlan.ActiveSheet.Cells(i, 0).Text.PadLeft(6)) + "'"
            If i <> spdPlan.ActiveSheet.Rows.Count - 2 Then PlanList = PlanList + ","
        Next
        If alignitem = 0 Then
            str = str.Replace("@PLAN", IIf(rdbHanniK.Checked Or PlanList = vbNullString, "and getMediumMemo(:xid,li.bid,li.bkid,99998) between :xsplan and :xeplan ", "and getMediumMemo(:xid,li.bid,li.bkid,99998) in (" + PlanList + ") "))
            SQLR.AddItem("xsplan") = Trim(txtStrPlan.Text)
            SQLR.AddItem("xeplan") = Trim(txtEndPlan.Text)
        Else
            str = str.Replace("@PLAN", IIf(rdbHanniK.Checked Or PlanList = vbNullString, "and getItems(getMediumMemo(:xid,li.bid,li.bkid,99998),6,1) between :xsplan and :xeplan ", "and getItems( getMediumMemo(:xid,li.bid,li.bkid,99998),6,1) in (" + PlanList + ") "))
            SQLR.AddItem("xsplan") = Trim(txtStrPlan.Text).PadLeft(6)
            SQLR.AddItem("xeplan") = Trim(txtEndPlan.Text).PadLeft(6)
        End If

        '売経路-------------------------------------
        SQLR.AddItem("xkeflg") = 1
        STypeList = vbNullString
        For i = 0 To spdSType.ActiveSheet.Rows.Count - 2
            STypeList = STypeList + "'" + IIf(alignitem = 0, spdSType.ActiveSheet.Cells(i, 0).Text, spdSType.ActiveSheet.Cells(i, 0).Text.PadLeft(4)) + "'"
            If i <> spdSType.ActiveSheet.Rows.Count - 2 Then STypeList = STypeList + ","
        Next
        If alignitem = 0 Then
            str = str.Replace("@SELLTYPE", " and sl.jid between :xsjid and :xejid")
            SQLR.AddItem("xsjid") = Trim(txtStrSType.Text)
            SQLR.AddItem("xejid") = Trim(txtEndSType.Text)
        Else
            str = str.Replace("@SELLTYPE", " and getItems(sl.jid,4,1) between :xsjid and :xejid")
            SQLR.AddItem("xsjid") = Trim(txtStrSType.Text).PadLeft(4)
            SQLR.AddItem("xejid") = Trim(txtEndSType.Text).PadLeft(4)
        End If

        Select Case cboRankTarget.SelectedIndex
            Case 0  '受注数
                str = str.Replace("@TARGET", "suu")
            Case 1  '受注額
                str = str.Replace("@TARGET", "money")
            Case 2  '受注粗利
                str = str.Replace("@TARGET", "arari")
            Case 3  'CPO
                str = str.Replace("sum(@TARGET)", "case sum(promo) when 0 then :xrank else case sum(money) when 0 then :xrank else sum(case cpo when to_number(:xrank) then 0 else cpo end) end end")
                str = str.Replace("@TARGET", "cpo")
            Case 4  'CCPO
                str = str.Replace("sum(@TARGET)", "case sum(promo) when 0 then :xrank else case sum(money) when 0 then :xrank else sum(case ccpo when to_number(:xrank) then 0 else ccpo end) end end")
                str = str.Replace("@TARGET", "ccpo")

        End Select

        Select Case cboRank.SelectedIndex
            Case 0 'ベスト
                If cboRankTarget.SelectedIndex = 3 Or cboRankTarget.SelectedIndex = 4 Then
                    str = str.Replace("@JYUN", " asc")
                Else
                    str = str.Replace("@JYUN", " desc")
                End If
            Case 1 'ワースト
                If cboRankTarget.SelectedIndex = 3 Or cboRankTarget.SelectedIndex = 4 Then
                    str = str.Replace("@JYUN", " desc")
                Else
                    str = str.Replace("@JYUN", vbNullString)
                End If
        End Select

        Select Case cboUriKbn.SelectedIndex
            Case 0  '税抜
                str = str.Replace("li.money", "li.outmoney")
                str = str.Replace("li.tingenka", "li.outgenka ")
                str = str.Replace("@TANKA", "li.outtanka")
            Case 1  '税込
                str = str.Replace("li.money", "li.inmoney")
                str = str.Replace("li.toutgenka", "li.ingenka ")
                str = str.Replace("@TANKA", "li.intanka")
        End Select

        If cboRank.SelectedIndex = 1 Then
            SQLR.AddItem("xrank") = -1
        Else
            SQLR.AddItem("xrank") = 999999
        End If

        If rdbPlan.Checked Then
            str = str.Replace("@PMEDIUM", " null bid, null mname,")
            str = str.Replace("@PPLAN", " plan, pname,")
            str = str.Replace("@TYPE", "a.plan=b.plan")
        Else
            str = str.Replace("@PMEDIUM", " bidn bid, mname,")
            str = str.Replace("@PPLAN", " null plan, null pname,")
            str = str.Replace("@TYPE", "a.bid=b.bid")
        End If

        SQLR.Read(str) 'リスト作るSQL (修正後)
        If SQLR.RowCount > 0 Then
            DT = SQLR.DTable.Clone
            Dim newrow As DataRow = Nothing
            Dim bdate As New Date
            Dim suu As New Decimal
            Dim limoney As New Decimal
            Dim promo As New Decimal
            Dim arari As New Decimal
            Dim ratiototal As New Decimal
            Dim cpo As New Decimal
            Dim ccpo As New Decimal
            Dim oldrank As String = OFtoS(SQLR.DTable.Rows(0), "rank")
            Dim oldncpo As String = OFtoS(SQLR.DTable.Rows(0), "ncpo")
            Dim oldnccpo As String = OFtoS(SQLR.DTable.Rows(0), "nccpo")
            Dim oldbid As String = OFtoS(SQLR.DTable.Rows(0), "bid")
            Dim oldmname As String = OFtoS(SQLR.DTable.Rows(0), "mname")
            Dim oldbkid As String = OFtoS(SQLR.DTable.Rows(0), "bkid")
            Dim oldchannel As String = OFtoS(SQLR.DTable.Rows(0), "channel")
            Dim oldchname As String = OFtoS(SQLR.DTable.Rows(0), "chname")
            Dim oldplan As String = OFtoS(SQLR.DTable.Rows(0), "plan")
            Dim oldpname As String = OFtoS(SQLR.DTable.Rows(0), "pname")
            For Each row In SQLR.DTable.Rows
                If oldplan <> OFtoS(row, "plan") Then 'or _
                    If newrow IsNot Nothing Then
                        newrow("suu") = suu
                        newrow("money") = limoney
                        newrow("arari") = arari
                        newrow("promo") = promo
                        newrow("cpo") = cpo
                        newrow("ccpo") = ccpo
                        newrow("ratiototal") = ratiototal
                        newrow("rank") = oldrank
                        newrow("ncpo") = vbNullString 
                        newrow("nccpo") = vbNullString 
                        newrow("bkid") = vbNullString 
                        newrow("mname") = vbNullString 
                        newrow("bid") = vbNullString 
                        newrow("channel") = vbNullString 
                        newrow("chname") = vbNullString 
                        newrow("plan") = vbNullString 
                        newrow("pname") = vbNullString 
                        DT.Rows.Add(newrow)
                        newrow = Nothing
                    End If
                    oldrank = OFtoS(row, "rank")
                    oldncpo = OFtoS(row, "ncpo")
                    oldnccpo = OFtoS(row, "nccpo")
                    oldbid = OFtoS(row, "bid")
                    oldmname = OFtoS(row, "mname")
                    oldbkid = OFtoS(row, "bkid")
                    oldchannel = OFtoS(row, "channel")
                    oldchname = OFtoS(row, "chname")
                    oldplan = OFtoS(row, "plan")
                    oldpname = OFtoS(row, "pname")
                End If
                If OFtoN(row, "rank") <= StoN(txtRankCnt.Text) Then
                    DT.ImportRow(row)
                    suu += OFtoN(row, "suu")
                    limoney += OFtoN(row, "money")
                    arari += OFtoN(row, "arari")
                    promo += OFtoN(row, "promo")
                    Select Case OFtoN(row, "ccpo")
                        Case Nothing, -1, 999999
                            ccpo += 0
                        Case Else
                            ccpo += OFtoN(row, "ccpo")
                    End Select
                    Select Case OFtoN(row, "cpo")
                        Case Nothing, -1, 999999
                            cpo += 0
                        Case Else
                            cpo += OFtoN(row, "cpo")
                    End Select
                    ratiototal += OFtoN(row, "ratiototal")
                Else
                    If newrow Is Nothing Then
                        newrow = DT.NewRow
                        suu = 0
                        limoney = 0
                        arari = 0
                        promo = 0
                        cpo = 0
                        ccpo = 0
                        ratiototal = 0
                    End If
                    suu += OFtoN(row, "suu")
                    limoney += OFtoN(row, "money")
                    arari += OFtoN(row, "arari")
                    promo = OFtoN(row, "promo")
                    Select Case OFtoN(row, "ccpo")
                        Case Nothing, -1, 999999
                            ccpo += 0
                        Case Else
                            ccpo += OFtoN(row, "ccpo")
                    End Select
                    Select Case OFtoN(row, "cpo")
                        Case Nothing, -1, 999999
                            cpo += 0
                        Case Else
                            cpo += OFtoN(row, "cpo")
                    End Select
                    ratiototal += OFtoN(row, "ratiototal")
                End If
            Next
            If newrow IsNot Nothing Then
                newrow("suu") = suu
                newrow("money") = limoney
                newrow("arari") = arari
                newrow("promo") = promo
                newrow("cpo") = cpo
                newrow("ccpo") = ccpo
                newrow("ratiototal") = ratiototal
                newrow("rank") = oldrank
                newrow("plan") = vbNullString
                If rdbPlan.Checked Then
                    newrow("pname") = "その他"
                    newrow("mname") = vbNullString
                Else
                    newrow("pname") = vbNullString
                    newrow("mname") = "その他"
                End If
                newrow("bid") = vbNullString
                newrow("ncpo") = vbNullString
                newrow("nccpo") = vbNullString
                newrow("bkid") = vbNullString
                newrow("channel") = vbNullString
                newrow("chname") = vbNullString
                DT.Rows.Add(newrow)
            End If
        End If
        For i = 0 To DT.Rows.Count - 1
            If OFtoN(DT.Rows(i), "cpo") = -1 Or OFtoN(DT.Rows(i), "cpo") = 999999 Then
                DT.Rows(i).Item("cpo") = 0

            End If
            If OFtoN(DT.Rows(i), "ccpo") = -1 Or OFtoN(DT.Rows(i), "ccpo") = 999999 Then
                DT.Rows(i).Item("ccpo") = 0
            End If
        Next
        If isFile Then
            DT.Columns.Remove("NCPO")
            DT.Columns.Remove("NCCPO")
            If cboPrintType.SelectedIndex = 1 Then
                DT.Columns.Remove("RATIOTOTAL")
                DT.Columns.Remove("BDATE")
                DT.Columns.Remove("BKID")
                DT.Columns.Remove("CHANNEL")
                DT.Columns.Remove("CHNAME")
                If rdbPlan.Checked Then
                    DT.Columns.Remove("BID")
                    DT.Columns.Remove("BANAME")
                Else
                    DT.Columns.Remove("PLAN")
                    DT.Columns.Remove("KNAME")
                End If
            End If
        End If
        Return DT
    End Function

    Private Function File(sender As Object, DT As DataTable) As Boolean
        File = False
        Dim listName As String = vbNullString
        If cboPrintType.SelectedIndex = 0 Then
                listName = "企画別媒体表【" StartUser "】"
        Else
                listName = "企画のみ表【" StartUser "】"
        End If
        For i As Integer = DT.Columns.Count - 1 To 0 Step -1
            Select Case DT.Columns(i).ColumnName
                Case "CHANNEL"
                    DT.Columns(i).Caption = "チャンネル"
                Case "CHNAME"
                    DT.Columns(i).Caption = "チャンネル名"
                Case "PLAN"
                    DT.Columns(i).Caption = "企画"
                Case "PNAME"
                    DT.Columns(i).Caption = "企画名"
                Case "BDATE"
                    DT.Columns(i).Caption = "媒体開始日"
                Case "RANK"
                    DT.Columns(i).Caption = "順位"
                Case "BANAME"
                    DT.Columns(i).Caption = "媒体名"
                Case "BID"
                    DT.Columns(i).Caption = "媒体"
                Case "BKID"
                    DT.Columns(i).Caption = "媒体管理"
                Case "GDID"
                    DT.Columns(i).Caption = "商品ID"
                Case "GNAME"
                    DT.Columns(i).Caption = "商品名称"
                Case "SUU"
                    DT.Columns(i).Caption = "受注件数"
                Case "MONEY"
                    DT.Columns(i).Caption = "受注金額"
                Case "PROMO"
                    DT.Columns(i).Caption = "販促費"
                Case "CPO"
                    DT.Columns(i).Caption = "CPO"
                Case "CCPO"
                    DT.Columns(i).Caption = "CCPO"
                Case "ARARI"
                    DT.Columns(i).Caption = "受注粗利"
                Case "RATIOTOTAL"
                    DT.Columns(i).Caption = "原価率"
            End Select
        Next
        DT.AcceptChanges()
        Dim doc As PdfDocument = New PdfDocument()
         Dim page As PdfPage = doc.Pages.Add()
         Dim pdfGrid As PdfGrid = New PdfGrid()
         pdfGrid.DataSource = DT
         Dim gridStyle As PdfGridStyle = New PdfGridStyle()
         gridStyle.CellPadding = New PdfPaddings(5, 5, 5, 5)
         PdfGrid.Style = gridStyle
         pdfGrid.Draw(page, New PointF(10, 10))
         doc.Save("Output.pdf")'パス入れ替える
         doc.Close(True)   
    End Function

    'frmPopUp:いろいろ検索画面
    Public Sub PopUp(sender As Object, e As EventArgs, Optional selRow As Integer = 0)
        Dim CtlName As String = GetActiveControl(Me.ParentForm).Name
        Select Case CtlName
            Case btnStrMePopUp.Name, btnEndMePopUp.Name, txtStrMeId.Name, txtEndMeId.Name, txtStrMeNm.Name, txtEndMeNm.Name '部門
                Dim frm As frmPopUp = New frmPopUp(SQLR, StartId, PopUpSQL.Department)
                If frm.ShowDialog(Me) = DialogResult.OK Then
                    If CtlName = btnStrMePopUp.Name Or CtlName = txtStrMeId.Name Then
                        txtStrMeId.Text = frm.GetValue(0)
                        txtStrMeNm.Text = frm.GetValue(1)
                    Else
                        txtEndMeId.Text = frm.GetValue(0)
                        txtEndMeNm.Text = frm.GetValue(1)
                    End If
                End If
            Case btnStrMediumPopUp.Name, btnEndMediumPopUp.Name, txtStrMedium.Name, txtEndMedium.Name, txtStrMediumNm.Name, txtEndMediumNm.Name, spdMedium.Name '媒体
                Dim frm As frmPopUp = New frmPopUp(SQLR, StartId, PopUpSQL.MEDIUM)
                If frm.ShowDialog(Me) = DialogResult.OK Then
                    If CtlName = spdMedium.Name Then
                        spdMedium.ActiveSheet.Cells(selRow, 0).Text = frm.GetValue(0)
                        spdMedium.ActiveSheet.Cells(selRow, 1).Text = frm.GetValue(1)
                    ElseIf CtlName = btnStrMediumPopUp.Name Or CtlName = txtStrMedium.Name Then
                        txtStrMedium.Text = frm.GetValue(0)
                        txtStrMediumNm.Text = frm.GetValue(1)
                    Else
                        txtEndMedium.Text = frm.GetValue(0)
                        txtEndMediumNm.Text = frm.GetValue(1)
                    End If
                End If
            Case btnStrMeExtraPopUp.Name, txtStrMeExtra.Name, txtStrMeExtraNm.Name, btnEndMeExtraPopUp.Name, txtEndMeExtra.Name, txtEndMeExtraNm.Name, spdMediumType.Name
                Dim frm As frmPopUp = New frmPopUp(SQLR, StartID, PopUpSQL.MEMO)
                frm.Title = "サブ"
                frm.Parameter(0).BindValue = 99998 
                If frm.ShowDialog(Me) = DialogResult.OK Then
                    If CtlName = spdMediumType.Name Then
                        spdMediumType.ActiveSheet.Cells(selRow, 0).Text = frm.GetValue(0)
                        spdMediumType.ActiveSheet.Cells(selRow, 1).Text = frm.GetValue(1)
                    ElseIf CtlName = btnStrMeExtraPopUp.Name Or CtlName = txtStrMeExtra.Name Or CtlName = txtStrMeExtraNm.Name Then
                        txtStrMeExtra.Text = frm.GetValue(0)
                        txtStrMeExtraNm.Text = frm.GetValue(1)
                    Else
                        txtEndMeExtra.Text = frm.GetValue(0)
                        txtEndMeExtraNm.Text = frm.GetValue(1)
                    End If
                End If
            Case btnStrPlanPopUp.Name, txtStrPlan.Name, txtStrPlanNm.Name, btnEndPlanPopUp.Name, txtEndPlan.Name, txtEndPlanNm.Name, spdPlan.Name
                Dim frm As frmPopUp = New frmPopUp(SQLR, StartID, PopUpSQL.MEDIUMKANRIPLAN)
                frm.Title = "企画"
                If frm.ShowDialog(Me) = DialogResult.OK Then
                    If CtlName = spdPlan.Name Then
                        spdPlan.ActiveSheet.Cells(selRow, 0).Text = frm.GetValue(0)
                        spdPlan.ActiveSheet.Cells(selRow, 1).Text = frm.GetValue(1)
                    ElseIf CtlName = btnStrPlanPopUp.Name Or CtlName = txtStrPlan.Name Or CtlName = txtStrPlanNm.Name Then
                        txtStrPlan.Text = frm.GetValue(0)
                        txtStrPlanNm.Text = frm.GetValue(1)
                    Else
                        txtEndPlan.Text = frm.GetValue(0)
                        txtEndPlanNm.Text = frm.GetValue(1)
                    End If
                End If
            Case btnStrSTypePopUp.Name, txtStrSType.Name, txtStrSTypeNm.Name, btnEndSTypePopUp.Name, txtEndSType.Name, txtEndSTypeNm.Name, spdSType.Name
                Dim frm As frmPopUp = New frmPopUp(SQLR, StartID, PopUpSQL.EXTRADATA)
                frm.Title = "売経路"
                frm.Parameter(0).BindValue = 0
                If frm.ShowDialog(Me) = DialogResult.OK Then
                    If CtlName = spdSType.Name Then
                        spdSType.ActiveSheet.Cells(selRow, 0).Text = frm.GetValue(0)
                        spdSType.ActiveSheet.Cells(selRow, 1).Text = frm.GetValue(1)
                    ElseIf CtlName = btnStrSTypePopUp.Name Or CtlName = txtStrSType.Name Or CtlName = txtStrSTypeNm.Name Then
                        txtStrSType.Text = frm.GetValue(0)
                        txtStrSTypeNm.Text = frm.GetValue(1)
                    Else
                        txtEndSType.Text = frm.GetValue(0)
                        txtEndSTypeNm.Text = frm.GetValue(1)
                    End If
                End If
        End Select
    End Sub

    'Functionキー(F7:ポップアップ検索画面呼び出す)
    Private Sub spd_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles spdMedium.PreviewKeyDown, spdMediumType.PreviewKeyDown, spdPlan.PreviewKeyDown, spdSType.PreviewKeyDown
        Select Case e.KeyCode
            Case Keys.F7
                If sender.ActiveSheet.GetSelection(0) IsNot Nothing Then
                    If sender.ActiveSheet.GetSelection(0).Row > -1 Then PopUp(sender, e, sender.ActiveSheet.GetSelection(0).Row)
                End If
        End Select
    End Sub

    'Functionキー(F8:次の値、F9:前の値)
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, ByVal keyData As Keys) As Boolean
        If (txtStrMeId.Focused Or txtEndMeId.Focused) _
        And (keyData = Keys.F9 Or keyData = Keys.F8) Then '部門
            Dim Cont As TextboxControl = Nothing
            If txtStrMeId.Focused Then
                Cont = txtStrMeId
            ElseIf txtEndMeId.Focused Then
                Cont = txtEndMeId
            End If
            Dim resCount As Integer = 0
            Dim sqlstr As String = String.Empty
            SQLR.AddItem("xid") = Cont.Text
            SQLR.AddItem("xalign") = CType(Cont.TextAlign, Integer)
            If keyData = Keys.F9 Then
                sqlstr = "select max(id) from (select syid,getid(did,6,:xalign) id from extradata) where syid=:xid and id<getItems(:xid,6,:xalign)"
            ElseIf keyData = Keys.F8 Then
                sqlstr = "select min(id) from (select syid,getid(did,6,:xalign) id from Department) where syid=:xid and id>getItems(:xid,6,:xalign)"
            End If
            resCount = SQLR.Read(sqlstr)
            If resCount > 0 Then
                Dim resSTR = SQLR.DTable(0)(0).ToString()
                If Not resSTR = String.Empty Then '最後、最初だと止める
                    Cont.Text = resSTR.Trim
                    LeaveEvent(Cont, Nothing)
                End If
            End If
            Return True
        ElseIf (txtStrMedium.Focused Or txtEndMedium.Focused) _
        And (keyData = Keys.F9 Or keyData = Keys.F8) Then '媒体
            Dim Cont As TextboxControl = Nothing
            If txtStrMedium.Focused Then
                Cont = txtStrMedium
            ElseIf txtEndMedium.Focused Then
                Cont = txtEndMedium
            End If
            Dim resCount As Integer = 0
            Dim sqlstr As String = String.Empty
            SQLR.AddItem("xid") = Cont.Text
            SQLR.AddItem("xalign") = CType(Cont.TextAlign, Integer)
            If keyData = Keys.F9 Then
                sqlstr = "select max(id) from (select syid,getid(bid,6,:xalign) id from medium) where syid=:xid and id<getItems(:xid,6,:xalign)"
            ElseIf keyData = Keys.F8 Then
                sqlstr = "select min(id) from (select syid,getid(bid,6,:xalign) id from medium) where syid=:xid and id>getItems(:xid,6,:xalign)"
            End If
            resCount = SQLR.Read(sqlstr)
            If resCount > 0 Then
                Dim resSTR = SQLR.DTable(0)(0).ToString()
                If Not resSTR = String.Empty Then '最後、最初だと止める
                    Cont.Text = resSTR.Trim
                    LeaveEvent(Cont, Nothing)
                End If
            End If
            Return True
        ElseIf (txtStrMeExtra.Focused Or txtEndMeExtra.Focused) _
        And (keyData = Keys.F9 Or keyData = Keys.F8) Then 'サブ媒体
            Dim Cont As TextboxControl = Nothing
            If txtStrMeExtra.Focused Then 
                Cont = txtStrMeExtra
            ElseIf txtEndMeExtra.Focused Then 
                Cont = txtEndMeExtra
            End If
            Dim resCount As Integer = 0
            Dim sqlstr As String = String.Empty
            SQLR.AddItem("xid") = Cont.Text
            SQLR.AddItem("xalign") = CType(Cont.TextAlign, Integer)
            SQLR.AddItem("xclass") = 99998
            If keyData = Keys.F9 Then
                sqlstr = "select max(id) from (select syid,getid(fid,6,:xalign) id from extraItem where class=:xclass) where syid=:xid and id<getItems(:xid,6,:xalign)"
            ElseIf keyData = Keys.F8 Then
                sqlstr = "select min(id) from (select syid,getid(fid,6,:xalign) id from extraItem where class=:xclass) where syid=:xid and id>getItems(:xid,6,:xalign)"
            End If
            resCount = SQLR.Read(sqlstr)
            If resCount > 0 Then
                Dim resSTR = SQLR.DataTable(0)(0).ToString()
                If Not resSTR = String.Empty Then '最後、最初だと止める
                    Cont.Text = resSTR.Trim
                    LeaveEvent(Cont, Nothing)
                End If
            End If
            Return True
        ElseIf (txtStrPlan.Focused Or txtEndPlan.Focused) _
        And (keyData = Keys.F9 Or keyData = Keys.F8) Then
            Dim Cont As TextboxControl = Nothing
            If txtStrPlan.Focused Then
                Cont = txtStrPlan
            ElseIf txtEndPlan.Focused Then
                Cont = txtEndPlan
            End If
            Dim resCount As Integer = 0
            Dim sqlstr As String = String.Empty
            SQLR.AddItem("xid") = Cont.Text
            SQLR.AddItem("xalign") = CType(Cont.TextAlign, Integer)
            SQLR.AddItem("xclass") = 99998
            If keyData = Keys.F9 Then
                sqlstr = "select max(id) from (select syid,getid(fid,6,:xalign) id from extraItem where class=:xclass) where syid=:xid and id<getItems(:xid,6,:xalign)"
            ElseIf keyData = Keys.F8 Then
                sqlstr = "select min(id) from (select syid,getid(fid,6,:xalign) id from extraItem where class=:xclass) where syid=:xid and id>getItems(:xid,6,:xalign)"
            End If
            resCount = SQLR.Read(sqlstr)
            If resCount > 0 Then
                Dim resSTR = SQLR.DataTable(0)(0).ToString()
                If Not resSTR = String.Empty Then
                    Cont.Text = resSTR.Trim
                    LeaveEvent(Cont, Nothing)
                End If
            End If
            Return True
        ElseIf (txtStrSType.Focused Or txtEndSType.Focused) _
        And (keyData = Keys.F9 Or keyData = Keys.F8) Then
            Dim Cont As TextboxControl = Nothing
            If txtStrSType.Focused Then
                Cont = txtStrSType
            ElseIf txtEndSType.Focused Then
                Cont = txtEndSType
            End If
            Dim resCount As Integer = 0
            Dim sqlstr As String = String.Empty
            SQLR.AddItem("xid") = Cont.Text
            SQLR.AddItem("xalign") = CType(Cont.TextAlign, Integer)
            SQLR.AddItem("xclass") = 0
            If keyData = Keys.F9 Then
                sqlstr = "select max(id) from (select syid,getid(brid,6,:xalign) id from btype) where syid=:xid and id<getItems(:xid,6,:xalign)"
            ElseIf keyData = Keys.F8 Then
                sqlstr = "select min(id) from (select syid,getid(brid,6,:xalign) id from btype) where syid=:xid and id>getItems(:xid,6,:xalign)"
            End If
            resCount = SQLR.Read(sqlstr)
            If resCount > 0 Then
                Dim resSTR = SQLR.DataTable(0)(0).ToString()
                If Not resSTR = String.Empty Then
                    Cont.Text = resSTR.Trim
                    LeaveEvent(Cont, Nothing)
                End If
            End If
            Return True
        End If
        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    Private Sub spdMedium_Change(sender As Object, e As FarPoint.Win.Spread.ChangeEventArgs) Handles spdMedium.Change, spdMediumType.Change, spdPlan.Change, spdSType.Change
        Dim SV As FarPoint.Win.Spread.SheetView
        Dim Sql As String = vbNullString
        Select Case sender.name
            Case spdMedium.Name
                SV = spdMedium.ActiveSheet
                Sql = "select name from medium where syid=:xid and bid=:x1"
            Case spdMediumType.Name
                SV = spdMediumType.ActiveSheet
                Sql = "select name from extraItem where syid=:xid and class=99998 and fid=:x1"
            Case spdSType.Name
                SV = spdSType.ActiveSheet
                Sql = "select gettypename(:xid,0,:x1) name from dual"
            Case spdPlan.Name
                SV = spdPlan.ActiveSheet
                Sql = "select name from extraItem where syid=:xid and class=99999 and fid=:x1"
            Case Else
                Exit Sub
        End Select
        Try
            If SV.Cells(e.Row, SV.GetViewColumnFromModelColumn(0)).Value IsNot Nothing Then
                SQLR.AddItem("x1") = SV.Cells(e.Row, 0).Value.ToString
                If SQLR.Read(Sql) = 1 Then
                    SV.Cells(e.Row, 1).Value = OFtoS(SQLR.DataTable.Rows(0), "name")
                Else
                    SV.Rows(e.Row).Remove()
                End If
            Else
                SV.Rows(e.Row).Remove()
            End If
        Catch ex As Exception
            MessageBox(Msg.mError, , "表示中エラー")
        Finally
        End Try
    End Sub
 
End Class
