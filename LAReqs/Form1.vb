Imports OAConnection
Imports SAPCOM.RepairsLevels

Public Class Form1
    Public Status As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' EndProcess()
        Dim cn As New OAConnection.Connection
        pgrWorking.Style = ProgressBarStyle.Marquee
        lstStatus.Items.Add("Cleaning Open Requisitions Table")

        cn.ExecuteInServer("Delete From SC_OpenRequis")
        BGOR_L7P.RunWorkerAsync()
        BGOR_G4P.RunWorkerAsync()
        BGOR_GBP.RunWorkerAsync()
        BGOR_L6P.RunWorkerAsync()
    End Sub

    Private Sub BGOR_L7P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOR_L7P.DoWork
        Dim SAP As String = "L7P"
        Dim dtPlants As New DataTable

        Dim POs As New DataTable
        Dim cn As New OAConnection.Connection

        dtPlants = cn.RunSentence("Select distinct Plant From LA_Indirect_Scope Where SAPBox = '" & SAP & "'").Tables(0)

        For Each row As DataRow In dtPlants.Rows
            Dim Rep As New SAPCOM.OpenReqs_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords

            BGOR_L7P.ReportProgress(0, "Getting: " & SAP & " - Plant: " & row("Plant"))

            Rep.IncludePlant(row("Plant"))
            Rep.RepairsLevel = IncludeRepairs

            Rep.ExcludeMatGroup("S731516AW")
            Rep.ExcludeMatGroup("S801416AQ")
            Rep.ExcludeMatGroup("S731516AV")


            Rep.Execute()
            If Rep.Success Then
                If Rep.ErrMessage = Nothing Then
                    POs = Rep.Data

                    Dim TN As New DataColumn
                    Dim SB As New DataColumn
                    Dim CR As New DataColumn
                    Dim PR As New DataColumn

                    TN.ColumnName = "Usuario"
                    TN.Caption = "Usuario"
                    TN.DefaultValue = "BM4691"

                    SB.DefaultValue = SAP
                    SB.ColumnName = "SAPBox"
                    SB.Caption = "SAPBox"

                    CR.DefaultValue = ""
                    CR.Caption = "Currency"
                    CR.ColumnName = "Currency"

                    PR.DefaultValue = 0.0
                    PR.Caption = "Price"
                    PR.ColumnName = "Price"

                    POs.Columns.Add(TN)
                    POs.Columns.Add(SB)
                    POs.Columns.Add(CR)
                    POs.Columns.Add(PR)

                    BGOR_L7P.ReportProgress(0, "Getting Price: " & SAP & " - Plant: " & row("Plant"))

                    Dim Eban As New SAPCOM.EBAN_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                    For Each R As DataRow In Rep.Data.Rows
                        Eban.IncludeDocument(R("Req Number"))
                    Next

                    Eban.AddCustomField("PREIS", "Price")
                    Eban.AddCustomField("WAERS", "Currency")
                    Eban.AddCustomField("PEINH", "Per")
                    Eban.Execute()

                    If Eban.Success Then
                        For Each R As DataRow In POs.Rows

                            Dim rPrice = (From C In Eban.Data.AsEnumerable() _
                                          Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                          Select C.Item("Price"))

                            Dim rCurr = (From C In Eban.Data.AsEnumerable() _
                                         Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                         Select C.Item("Currency"))

                            Dim rPer = (From C In Eban.Data.AsEnumerable() _
                                       Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                       Select C.Item("Per"))

                            Dim xPrice As Double

                            Dim Price_Is_Valid As Boolean
                            Price_Is_Valid = Double.TryParse(rPrice(0), xPrice)

                            If Price_Is_Valid AndAlso xPrice > 0 Then
                                R("Price") = Double.Parse(xPrice) / Double.Parse(rPer(0))
                                R("Currency") = rCurr(0)
                            Else
                                Dim Try2 As New SAPCOM.PRInfo(SAP, "BM4691", "LAT", R("Req Number")) ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                                If Try2.Success Then
                                    Dim I As Double
                                    I = Try2.ItemTotalPrice(R("Item Number"))

                                    R("Price") = I
                                    R("Currency") = rCurr(0)
                                End If
                            End If
                        Next
                    End If

                    cn.AppendTableToSqlServer("SC_OpenRequis", POs)
                Else
                    BGOR_L7P.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
                End If
            Else
                BGOR_L7P.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
            End If
        Next
    End Sub

    Private Sub EndProcess()
        Dim cn As New OAConnection.Connection
        'Colocar código para eliminar las requis que no se encuentran en el scope de PSS
        If Not BGOR_G4P.IsBusy AndAlso Not BGOR_GBP.IsBusy AndAlso Not BGOR_L6P.IsBusy AndAlso Not BGOR_L7P.IsBusy Then
            'lstStatus.Items.Insert(0, "Deleting closed items...")
            'Moved to stored procedure: cn.ExecuteInServer("DELETE FROM dbo.LA_TMP_Open_Req_Distribution Where (NOT EXISTS (SELECT [Req Number] FROM SC_OpenRequis Where (dbo.LA_TMP_Open_Req_Distribution.[Req Number] = [Req Number]) AND (dbo.LA_TMP_Open_Req_Distribution.SAPBox = SAPBox)))")

            'lstStatus.Items.Insert(0, "Deleting items without currency...")
            'Moved to stored procedure: cn.ExecuteInServer("DELETE FROM SC_OpenRequis Where (Currency = '')")

            'lstStatus.Items.Insert(0, "Updating service line...")
            'Moved to stored procedure: cn.ExecuteInServer("Update SC_OpenRequis Set ServiceLine = (Case when (left(Material,1) = 3) Then 'STR' Else 'SS' End)")

            'Agregar las POs que son nuevas a la distribucion temporal
            'lstStatus.Items.Insert(0, "Updating distribution...")
            'Moved to stored procedure: cn.ExecuteInServer("Insert Into LA_TMP_Open_Req_Distribution (SAPBox, [Req Number], Plant) SELECT DISTINCT SAPBox, [Req Number], Plant From SC_OpenRequis WHERE (NOT EXISTS (SELECT [Req Number] From dbo.LA_TMP_Open_Req_Distribution Where (dbo.SC_OpenRequis.[Req Number] = [Req Number]) AND (dbo.SC_OpenRequis.SAPBox = SAPBox)))")

            'Crear una funcion para asignarles el owner a las nuevas.

            cn.ExecuteStoredProcedure("Update_GCT_OpenPR")

            Dim OO As New DataTable
            Dim T As Integer = 0
            OO = cn.RunSentence("Select * From vst_LA_Check_Req_Distribution").Tables(0)

            If OO.Rows.Count > 0 Then
                For Each r As DataRow In OO.Rows
                    Try
                        T += 1
                        lstStatus.Items.Insert(0, "Updating owner: " & T & " of " & OO.Rows.Count)
                        Dim RX As New OAConnection.DMS_User(r("SAPBox"), r("Mat Group"), r("Purch Grp"), r("Purch Org"), r("Plant"), r("Type"))
                        RX.Execute()

                        If RX.Success Then
                            Dim PRStatus As Double
                            PRStatus = cn.RunSentence("SELECT dbo.GetWorkingDatesFn([Release Date], { fn NOW() }) AS Aging FROM dbo.SC_OpenRequis WHERE ([Req Number] = " & r("Req Number") & ") AND (SAPBox = '" & r("SAPBox") & "')").Tables(0).Rows(0).Item(0)
                            cn.ExecuteInServer("Update LA_TMP_Open_Req_Distribution Set SPS = '" & RX.SPS & "', Owner = '" & RX.SPS & "', [Total USD] = dbo._fn_Get_LA_PR_Value('" & r("SAPBox") & "'," & r("Req Number") & "), Aging = " & PRStatus & " Where ((SAPBox = '" & r("SAPBox") & "') And ([Req Number] = '" & r("Req Number") & "'))")
                        Else
                            Dim PRStatus As Double
                            PRStatus = cn.RunSentence("SELECT dbo.GetWorkingDatesFn([Release Date], { fn NOW() }) AS Aging FROM dbo.SC_OpenRequis WHERE ([Req Number] = " & r("Req Number") & ") AND (SAPBox = '" & r("SAPBox") & "')").Tables(0).Rows(0).Item(0)
                            cn.ExecuteInServer("Update LA_TMP_Open_Req_Distribution Set SPS = 'BI5226', Owner = 'BI5226', [Total USD] = dbo._fn_Get_LA_PR_Value('" & r("SAPBox") & "'," & r("Req Number") & "), Aging = " & PRStatus & " Where ((SAPBox = '" & r("SAPBox") & "') And ([Req Number] = '" & r("Req Number") & "'))")
                        End If
                    Catch ex As Exception
                        lstStatus.Items.Insert(0, "Error updating owner: " & T & " of " & OO.Rows.Count)
                    End Try
                Next
            End If

            End
        End If
    End Sub
    Private Sub BGOR_GBP_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOR_GBP.DoWork
        Dim SAP As String = "GBP"
        Dim dtPlants As New DataTable

        Dim POs As New DataTable
        Dim cn As New OAConnection.Connection

        dtPlants = cn.RunSentence("Select distinct Plant From LA_Indirect_Scope Where SAPBox = '" & SAP & "'").Tables(0)

        For Each row As DataRow In dtPlants.Rows
            Dim Rep As New SAPCOM.OpenReqs_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
            BGOR_GBP.ReportProgress(0, "Getting: " & SAP & " - Plant: " & row("Plant"))

            Rep.IncludePlant(row("Plant"))
            Rep.RepairsLevel = IncludeRepairs

            Rep.ExcludeMatGroup("S731516AW")
            Rep.ExcludeMatGroup("S801416AQ")
            Rep.ExcludeMatGroup("S731516AV")

            Rep.Execute()
            If Rep.Success Then
                If Rep.ErrMessage = Nothing Then
                    POs = Rep.Data

                    Dim TN As New DataColumn
                    Dim SB As New DataColumn
                    Dim CR As New DataColumn
                    Dim PR As New DataColumn

                    TN.ColumnName = "Usuario"
                    TN.Caption = "Usuario"
                    TN.DefaultValue = "BM4691"

                    SB.DefaultValue = SAP
                    SB.ColumnName = "SAPBox"
                    SB.Caption = "SAPBox"

                    CR.DefaultValue = ""
                    CR.Caption = "Currency"
                    CR.ColumnName = "Currency"

                    PR.DefaultValue = 0.0
                    PR.Caption = "Price"
                    PR.ColumnName = "Price"

                    POs.Columns.Add(TN)
                    POs.Columns.Add(SB)
                    POs.Columns.Add(CR)
                    POs.Columns.Add(PR)


                    BGOR_GBP.ReportProgress(0, "Getting Price: " & SAP & " - Plant: " & row("Plant"))

                    Dim Eban As New SAPCOM.EBAN_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                    For Each R As DataRow In Rep.Data.Rows
                        Eban.IncludeDocument(R("Req Number"))
                    Next

                    Eban.AddCustomField("PREIS", "Price")
                    Eban.AddCustomField("WAERS", "Currency")
                    Eban.AddCustomField("PEINH", "Per")


                    Eban.Execute()
                    If Eban.Success Then
                        For Each R As DataRow In POs.Rows
                            Dim rPrice = (From C In Eban.Data.AsEnumerable() _
                                          Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                          Select C.Item("Price"))

                            Dim rCurr = (From C In Eban.Data.AsEnumerable() _
                                         Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                         Select C.Item("Currency"))

                            Dim rPer = (From C In Eban.Data.AsEnumerable() _
                                       Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                       Select C.Item("Per"))

                            Dim xPrice As Double

                            Dim Price_Is_Valid As Boolean
                            Price_Is_Valid = Double.TryParse(rPrice(0), xPrice)

                            If Price_Is_Valid AndAlso xPrice > 0 Then
                                R("Price") = Double.Parse(xPrice) / Double.Parse(rPer(0))
                                R("Currency") = rCurr(0)
                            Else
                                Dim Try2 As New SAPCOM.PRInfo(SAP, "BM4691", "LAT", R("Req Number")) ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                                If Try2.Success Then
                                    R("Price") = Try2.ItemTotalPrice(R("Item Number"))
                                    R("Currency") = rCurr(0)
                                End If
                            End If
                        Next
                    End If

                    cn.AppendTableToSqlServer("SC_OpenRequis", POs)
                Else
                    BGOR_GBP.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
                End If
            Else
                BGOR_GBP.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
            End If
        Next
    End Sub
    Private Sub BGOR_L6P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOR_L6P.DoWork
        Dim SAP As String = "L6P"
        Dim dtPlants As New DataTable
        Dim dtPOrg As New DataTable

        Dim POs As New DataTable
        Dim cn As New OAConnection.Connection

        dtPlants = cn.RunSentence("Select distinct Plant From LA_Indirect_Scope Where SAPBox = '" & SAP & "'").Tables(0)
        dtPOrg = cn.RunSentence("Select distinct POrg From LA_Indirect_Scope Where SAPBox = '" & SAP & "'").Tables(0)

        For Each row As DataRow In dtPlants.Rows
            Try

                Dim Rep As New SAPCOM.OpenReqs_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                BGOR_L6P.ReportProgress(0, "Getting: " & SAP & " - Plant: " & row("Plant").ToString.Trim)

                Rep.IncludePlant(row("Plant").ToString.Trim)
                Rep.RepairsLevel = IncludeRepairs

                Rep.ExcludeMatGroup("S731516AW")
                Rep.ExcludeMatGroup("S801416AQ")
                Rep.ExcludeMatGroup("S731516AV")
                Rep.ExcludeMatGroup("FIN PROD")
                Rep.IncludePurchOrg("")

                For Each POrg As DataRow In dtPOrg.Rows
                    Rep.IncludePurchOrg(POrg("POrg"))
                Next

                Rep.Execute()
                If Rep.Success Then
                    If Rep.ErrMessage = Nothing Then
                        POs = Rep.Data

                        Dim D As New DataColumn
                        Dim TN As New DataColumn
                        Dim SB As New DataColumn
                        Dim CR As New DataColumn
                        Dim PR As New DataColumn

                        TN.ColumnName = "Usuario"
                        TN.Caption = "Usuario"
                        TN.DefaultValue = "BM4691"

                        SB.DefaultValue = SAP
                        SB.ColumnName = "SAPBox"
                        SB.Caption = "SAPBox"

                        CR.DefaultValue = ""
                        CR.Caption = "Currency"
                        CR.ColumnName = "Currency"

                        PR.DefaultValue = 0.0
                        PR.Caption = "Price"
                        PR.ColumnName = "Price"

                        POs.Columns.Add(TN)
                        POs.Columns.Add(SB)
                        POs.Columns.Add(CR)
                        POs.Columns.Add(PR)


                        BGOR_L6P.ReportProgress(0, "Getting Price: " & SAP & " - Plant: " & row("Plant"))

                        Dim Eban As New SAPCOM.EBAN_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                        Dim RunEBAN As Boolean = False
                        For Each R As DataRow In Rep.Data.Rows
                            If R("Item Number") = "0" Then
                                R.Delete()
                            Else
                                If (Microsoft.VisualBasic.Left(R("Req Number").ToString, 1) <> "4") And (Microsoft.VisualBasic.Left(R("Req Number").ToString, 1) <> "9") Then
                                    RunEBAN = True
                                    Eban.IncludeDocument(R("Req Number"))
                                End If
                            End If
                        Next

                        POs.AcceptChanges()
                        Eban.AddCustomField("PREIS", "Price")
                        Eban.AddCustomField("WAERS", "Currency")
                        Eban.AddCustomField("PEINH", "Per")

                        If RunEBAN Then
                            Eban.Execute()
                            If Eban.Success Then
                                For Each dr As DataRow In POs.Rows
                                    If dr("Item Number") = "0" Then
                                        dr.Delete()
                                    End If
                                Next

                                For Each R As DataRow In POs.Rows
                                    If (Microsoft.VisualBasic.Left(R("Req Number").ToString, 1) <> "4") And (Microsoft.VisualBasic.Left(R("Req Number").ToString, 1) <> "9") Then

                                        Dim rPrice = (From C In Eban.Data.AsEnumerable() _
                                                      Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                                      Select C.Item("Price"))

                                        Dim rCurr = (From C In Eban.Data.AsEnumerable() _
                                                     Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                                     Select C.Item("Currency"))

                                        Dim rPer = (From C In Eban.Data.AsEnumerable() _
                                                   Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                                   Select C.Item("Per"))

                                        Dim xPrice As Double

                                        Dim Price_Is_Valid As Boolean
                                        Price_Is_Valid = Double.TryParse(rPrice(0), xPrice)

                                        If Price_Is_Valid AndAlso xPrice > 0 Then
                                            R("Price") = Double.Parse(xPrice) / Double.Parse(rPer(0))
                                            R("Currency") = rCurr(0)
                                        Else
                                            BGOR_L6P.ReportProgress(0, "...-> Getting price from SAP: " & SAP & " - Plant: " & row("Plant").ToString.Trim & " - " & R("Req Number").ToString)
                                            Dim Try2 As New SAPCOM.PRInfo(SAP, "BM4691", "LAT", R("Req Number")) ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                                            If Try2.Success Then
                                                R("Price") = Try2.ItemTotalPrice(R("Item Number"))
                                                R("Currency") = rCurr(0)
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If

                        cn.AppendTableToSqlServer("SC_OpenRequis", POs)
                        cn.ExecuteInServer("Delete From SC_OpenRequis Where (left([Req Number],1) = 4) or (left([Req Number],1) = 9)")
                    Else
                        BGOR_L6P.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
                    End If
                Else
                    BGOR_L6P.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
                End If
            Catch ex As Exception
                Status = "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: Unknown error "
                BGOR_L6P.ReportProgress(0)
            End Try

        Next
    End Sub
    Private Sub BGOR_G4P_DoWork(ByVal sender As System.Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGOR_G4P.DoWork
        Dim SAP As String = "G4P"
        Dim dtPlants As New DataTable

        Dim POs As New DataTable
        Dim cn As New OAConnection.Connection
         dtPlants = cn.RunSentence("Select distinct Plant From LA_Indirect_Scope Where SAPBox = '" & SAP & "'").Tables(0)

        For Each row As DataRow In dtPlants.Rows
            Dim Rep As New SAPCOM.OpenReqs_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
            BGOR_G4P.ReportProgress(0, "Getting: " & SAP & " - Plant: " & row("Plant"))

            Rep.IncludePlant(row("Plant"))
            Rep.RepairsLevel = IncludeRepairs

            Rep.ExcludeMatGroup("S731516AW")
            Rep.ExcludeMatGroup("S801416AQ")
            Rep.ExcludeMatGroup("S731516AV")

            Rep.Execute()
            If Rep.Success Then
                If Rep.ErrMessage = Nothing Then
                    POs = Rep.Data

                    Dim TN As New DataColumn
                    Dim SB As New DataColumn
                    Dim CR As New DataColumn
                    Dim PR As New DataColumn
                    Dim PRUS As New DataColumn


                    TN.ColumnName = "Usuario"
                    TN.Caption = "Usuario"
                    TN.DefaultValue = "BM4691"

                    SB.DefaultValue = SAP
                    SB.ColumnName = "SAPBox"
                    SB.Caption = "SAPBox"

                    CR.DefaultValue = ""
                    CR.Caption = "Currency"
                    CR.ColumnName = "Currency"

                    PR.DefaultValue = 0.0
                    PR.Caption = "Price"
                    PR.ColumnName = "Price"

                    POs.Columns.Add(TN)
                    POs.Columns.Add(SB)
                    POs.Columns.Add(CR)
                    POs.Columns.Add(PR)

                    BGOR_G4P.ReportProgress(0, "Getting Price: " & SAP & " - Plant: " & row("Plant"))

                    Dim Eban As New SAPCOM.EBAN_Report(SAP, "BM4691", "LAT") ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                    For Each R As DataRow In Rep.Data.Rows
                        Eban.IncludeDocument(R("Req Number"))
                    Next

                    Eban.AddCustomField("PREIS", "Price")
                    Eban.AddCustomField("WAERS", "Currency")
                    Eban.AddCustomField("PEINH", "Per")


                    Eban.Execute()
                    If Eban.Success Then
                        For Each R As DataRow In POs.Rows
                            Dim rPrice = (From C In Eban.Data.AsEnumerable() _
                                          Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                          Select C.Item("Price"))

                            Dim rCurr = (From C In Eban.Data.AsEnumerable() _
                                         Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                         Select C.Item("Currency"))

                            Dim rPer = (From C In Eban.Data.AsEnumerable() _
                                       Where ((C.Item("Req Number") = R("Req Number")) And (C.Item("Req Item") = R("Item Number"))) _
                                       Select C.Item("Per"))

                            Dim xPrice As Double

                            Dim Price_Is_Valid As Boolean
                            Price_Is_Valid = Double.TryParse(rPrice(0), xPrice)

                            If Price_Is_Valid AndAlso xPrice > 0 Then
                                R("Price") = Double.Parse(xPrice) / Double.Parse(rPer(0))
                                R("Currency") = rCurr(0)
                            Else
                                Dim Try2 As New SAPCOM.PRInfo(SAP, "BM4691", "LAT", R("Req Number")) ' -> Change TNumber; use machine owner, password is taken from LA Tool Password setup @ System menu/Variants/SAP Passwords
                                If Try2.Success Then
                                    R("Price") = Try2.ItemTotalPrice(R("Item Number"))
                                    R("Currency") = rCurr(0)
                                End If
                            End If
                        Next
                    End If

                    cn.AppendTableToSqlServer("SC_OpenRequis", POs)
                Else
                    BGOR_G4P.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
                End If
            Else
                BGOR_G4P.ReportProgress(0, "-> Fail: " & SAP & " - Plant: " & row("Plant") & " :: " & Rep.ErrMessage)
            End If
        Next
    End Sub

    Private Sub BGOR_GBP_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOR_GBP.ProgressChanged
        lblStatus.Text = Status
        lstStatus.Items.Insert(0, Now.ToString & " - " & e.UserState)
        'lstStatus.Items.Add(Status)
    End Sub
    Private Sub BGOR_L6P_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOR_L6P.ProgressChanged
        lblStatus.Text = e.UserState
        lstStatus.Items.Insert(0, Now.ToString & " - " & e.UserState)
    End Sub
    Private Sub BGOR_G4P_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOR_G4P.ProgressChanged
        lblStatus.Text = Status
        lstStatus.Items.Insert(0, Now.ToString & " - " & e.UserState)
        'lstStatus.Items.Add(Status)
    End Sub

    Private Sub BGOR_L7P_ProgressChanged(ByVal sender As Object, ByVal e As System.ComponentModel.ProgressChangedEventArgs) Handles BGOR_L7P.ProgressChanged
        lblStatus.Text = Status
        lstStatus.Items.Insert(0, Now.ToString & " - " & e.UserState)

    End Sub

    Private Sub BGOR_GBP_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOR_GBP.RunWorkerCompleted
        lblStatus.Text = "<<<<<<<<<<<------ Finished: GBP ------>>>>>>>>>>>"
        lstStatus.Items.Insert(0, "<<<<<<<<<<<------ Finished: GBP ------>>>>>>>>>>>")

        EndProcess()
    End Sub
    Private Sub BGOR_L6P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOR_L6P.RunWorkerCompleted
        lblStatus.Text = "<<<<<<<<<<<------ Finished: L6P ------>>>>>>>>>>>"
        lstStatus.Items.Insert(0, "<<<<<<<<<<<------ Finished: L6P ------>>>>>>>>>>>")

        EndProcess()
    End Sub
    Private Sub BGOR_G4P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOR_G4P.RunWorkerCompleted
        lblStatus.Text = "<<<<<<<<<<<------ Finished: G4P ------>>>>>>>>>>>"
        lstStatus.Items.Insert(0, "<<<<<<<<<<<------ Finished: G4P ------>>>>>>>>>>>")

        EndProcess()
    End Sub
    Private Sub BGOR_L7P_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGOR_L7P.RunWorkerCompleted
        lblStatus.Text = "<<<<<<<<<<<------ Finished: L7P ------>>>>>>>>>>>"
        lstStatus.Items.Insert(0, "<<<<<<<<<<<------ Finished: L7P ------>>>>>>>>>>>")

        EndProcess()
    End Sub

    Public Function GetOwner(ByVal pSAP As String, Optional ByVal pPlant As String = Nothing, Optional ByVal pPGrp As String = Nothing, Optional ByVal pPOrg As String = Nothing) As Owner
        Dim cn As New OAConnection.Connection
        Dim Data As DataTable
        Dim Where As String = ""

        Try
            If Not pPlant Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                'Where = Where & "((Plant = '') or (Plant = '" & pPlant & "'))"
                Where = Where & "((Plant = '" & pPlant & "'))"

            End If

            If Not pPGrp Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If
                'Where = Where & "(([P Grp] = '') or ([P Grp] = '" & pPGrp & "'))"
                Where = Where & "(([P Grp] = '" & pPGrp & "'))"
            End If

            If Not pPOrg Is Nothing Then
                If Where <> "" Then
                    Where = Where & " And "
                End If

                'Where = Where & "(([P Org] = '') or ([P Org] = '" & pPOrg & "'))"
                Where = Where & "(([P Org] = '" & pPOrg & "'))"
            End If

            Data = cn.RunSentence("Select *,0 as Value From LA_Indirect_Distribution Where (SAPBox = '" & pSAP & "')" & IIf(Where <> "", " And (" & Where & ")", "")).Tables(0)
            If Data.Rows.Count > 0 Then
                If Data.Rows.Count = 1 Then
                    Dim T As New Owner

                    T.SPS = Data.Rows(0).Item("SPS")
                    T.Owner = Data.Rows(0).Item("Owner")
                    Return T
                Else

                    For Each r As DataRow In Data.Rows
                        Dim val As Integer = 0

                        If (r("SAPBox") = pSAP) Then
                            val += 2
                        Else
                            If r("SAPBox") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("Plant") = pPlant) Then
                            val += 2
                        Else
                            If r("Plant") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("P Org") = pPOrg) Then
                            val += 2
                        Else
                            If r("P Org") = "" Then
                                val += 1
                            End If
                        End If

                        If (r("P Grp") = pPGrp) Then
                            val += 2
                        Else
                            If r("P Grp") = "" Then
                                val += 1
                            End If
                        End If


                        r("Value") = val
                    Next

                    Dim T As New Owner
                    Dim SPS = (From C In Data.AsEnumerable() Order By C.Item("Value") Descending Select C.Item("SPS")).ToList()
                    Dim DOwner = (From C In Data.AsEnumerable() Order By C.Item("Value") Descending Select C.Item("Owner")).ToList()

                    T.SPS = SPS(0)
                    T.Owner = DOwner(0)

                    'MsgBox("Multiple choises for:" & Chr(13) & Chr(13) & "SAPBox: " & pSAP & Chr(13) & "LE: " & pLE & Chr(13) & "Plant:" & pPlant & Chr(13) & "Vendor: " & pVendor & Chr(13) & "Mat. Grp: " & pMatGrp)
                    Return T
                End If
            Else
                ' MsgBox("Rules can't be found")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
End Class

Public Class Owner
    Public SPS
    Public Owner
End Class
