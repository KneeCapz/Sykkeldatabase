Imports MySql.Data.MySqlClient

Public Class Form1
#Region "GlobaleVariabler"
    'Her plasseres globale variabler
    'som gjenbrukes over programmet. De er uten navnekonvensjon
    'Husk å kommentere på kodensfunksjon og hvor i programmet den er tatt i bruk.
    'Globale variabler som må sjekkes ved ønsket tilgang. LogBool settes til TRUE ved inlogging, AdminBool settes til TRUE hvis bruker er admin.

    Dim LogBool As Boolean = False
    Dim AdminBool As Boolean = False
    Dim FornavnString As String = ""
    Dim EtternavnString As String = ""
    Dim Tilkobling As MySqlConnection
    Dim InvAktivtProduktID As String
    Dim SecurityCounter As Integer = 0


#End Region


#Region "GlobaleFunksjonerOgProsedyrer"
    'Her plasseres globale funksjoner og prosdyrer som gjenbrukes over programmet. De er uten navnekonvensjon
    'Husk å kommentere på kodensfunksjon og hvor i programmet den er tatt i bruk.

    Private Sub DBConnect()
        Tilkobling = New MySqlConnection(
        "Server=mysql-ait.stud.idi.ntnu.no;" _
        & "Database=colinft;" _
        & "Uid=colinft;" _
        & "Pwd=BJhYR1HS;")
        Try
            Tilkobling.Open()
        Catch ex As MySqlException
            MsgBox(ex)
        End Try
    End Sub

    Private Sub DBDisconnect()
        Tilkobling.Close()
        Tilkobling.Dispose()
    End Sub

    Private Function SQLWhiteWash(ByVal StartString As String) As String
        Dim EndString As String
        EndString = StartString.Replace("'", "\'")
        Return EndString
    End Function

    Private Function CheckVarChar20(ByVal CheckString As String) As Boolean
        If CheckString.Length > 0 And CheckString.Length < 21 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CheckVarChar30(ByVal CheckString As String) As Boolean
        If CheckString.Length > 0 And CheckString.Length < 31 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CheckVarChar15(ByVal CheckString As String) As Boolean
        If CheckString.Length > 0 And CheckString.Length < 16 Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function CheckIntValue(ByVal CheckString As String) As Boolean
        If CheckString.Length > 0 And CheckString.Length < 11 And IsNumeric(CheckString) Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' SQL SELECT spørring som returnerer DataTable basert på søkekriterier.
    ''' </summary>
    ''' <param name="WhereTable">Tabell det skal søkes i</param>
    ''' <param name="WhereColumn">Kolonner som skal inkluderes i søkeresultat.
    ''' String = "column1, coulmn2, ..."''' </param>
    ''' <param name="WhereCriteria">Kriterier for søk. Angis med alle søkeparametre etter WHERE.
    ''' column1='value1' AND column2='value2' ...</param>
    ''' <returns>DataTabell med søkeresultat fra spørringen</returns>
    Private Function SQLSelect(ByVal WhereTable As String, WhereColumn As String, WhereCriteria As String)
        DBConnect()
        Dim SqlWhereKommando As New MySqlCommand("SELECT " & WhereColumn & " FROM " & WhereTable & " WHERE " _
                & WhereCriteria, Tilkobling)
        Dim SqlWhereAdapter As New MySqlDataAdapter
        Dim SqlWhereTable As New DataTable
        Try
            SqlWhereAdapter.SelectCommand = SqlWhereKommando
            SqlWhereAdapter.Fill(SqlWhereTable)
            Return SqlWhereTable
        Catch SQLex As MySqlException
            MsgBox("Feil ved søk i database:" & SQLex.Message)
            Return Nothing
        End Try
        DBDisconnect()
    End Function

    ''' <summary>
    ''' SQL INSERT spørring som legger til tuppel i valgt tabell
    ''' </summary>
    ''' <param name="InsertTable">Tabell det skal legges inn i</param>
    ''' <param name="InsertColumn">Kolonner som skal legges inn. Angis med parantes.
    ''' (column1, column2, ...)</param>
    ''' <param name="InsertValues">De verdier hver kolonne skal tillegges. Angis med parantes.
    ''' (value1, value2, ...)</param>
    Private Sub SQLInsert(ByVal InsertTable As String, InsertColumn As String, InsertValues As String)
        Try
            DBConnect()
            Dim SqlCom As New MySqlCommand("INSERT INTO " & InsertTable & InsertColumn & " VALUES " &
                InsertValues, Tilkobling)
            SqlCom.ExecuteNonQuery()
            DBDisconnect()
        Catch SQLex As MySqlException
            MsgBox("Feil ved innlegg i database:" & vbNewLine & SQLex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' SQL UDATE spørring som oppdaterer tuppler i tabell
    ''' </summary>
    ''' <param name="UpdateTable">Tabell som skal oppdateres</param>
    ''' <param name="UpdateValues">Kolonner og verdier som skal legges inn i valgte kolonner.
    ''' column1='value1', column2='value2', ...</param>
    ''' <param name="UpdateCriteria">Kriterier for oppdateringen. Alt som kommer etter WHERE.
    ''' column1=ID_number ...</param>
    Public Sub SQLUpdate(ByVal UpdateTable As String, UpdateValues As String, UpdateCriteria As String)
        DBConnect()
        Dim SqlCom As New MySqlCommand("UPDATE " & UpdateTable & " SET " & UpdateValues & " WHERE " _
                                       & UpdateCriteria, Tilkobling)
        Try
            SqlCom.ExecuteNonQuery()
        Catch SQLex As Exception
            MsgBox("Feil ved oppdatering database:" & vbNewLine & SQLex.Message)
        End Try
        DBDisconnect()
    End Sub

    ''' <summary>
    ''' SQL DELETE spørring som sletter data fra tabell ut i fra valgte kriterier
    ''' </summary>
    ''' <param name="DeleteTable">Tabell det skal slettes fra</param>
    ''' <param name="DeleteCriteria">Kriterier for slettingen. Alt etter WHERE.
    ''' coulmn1='ID_number' AND column1='value1' ...</param>
    Public Sub SQLDelete(ByVal DeleteTable As String, DeleteCriteria As String)
        DBConnect()
        Dim SqlCom As New MySqlCommand("DELETE FROM " & DeleteTable & " WHERE " & DeleteCriteria, Tilkobling)
        Try
            SqlCom.ExecuteNonQuery()
        Catch SQLex As Exception
            MsgBox("Feil ved sletting i database:" & vbNewLine & SQLex.Message)
        End Try
        DBDisconnect()
    End Sub

    ''' <summary>
    ''' SQL spørring for å hente ut navn ut i fra ID, og motsatt, for fremmednøkler i tabeller.
    ''' </summary>
    ''' <param name="Tabell">Tabellen det skal FRA</param>
    ''' <param name="KolonneInn">Navn eller ID (kolonne) det skal hentes FRA</param>
    ''' <param name="KolonneUt">Navn eller ID (kolonne) som skal hentes UT</param>
    ''' <param name="VerdiInn">Navn eller ID på tuppel som det gjelder</param>
    ''' <returns>Navn eller ID</returns>
    Private Function SQLHentIDNavn(ByVal Tabell As String, KolonneInn As String, KolonneUt As String, VerdiInn As String)
        Dim SqlCom As MySqlCommand
        Try
            DBConnect()
            Dim Result As String
            SqlCom = New MySqlCommand("SELECT " & KolonneUt & " FROM " & Tabell &
            " WHERE " & KolonneInn & "='" & VerdiInn & "'", Tilkobling)
            Dim SqlRead As MySqlDataReader
            SqlRead = SqlCom.ExecuteReader
            While SqlRead.Read()
                Result = SqlRead(KolonneUt)
            End While
            SqlRead.Close()
            DBDisconnect()
            Return Result
        Catch ex As MySqlException
            MsgBox("Feil ved henting av navn/ID:" & vbNewLine & ex.Message)
            Return Nothing
        End Try
    End Function


    ''' <summary>
    ''' Autopopulasjon av ComboBox. Elementer hentes fra kolonne i database.
    ''' </summary>
    ''' <param name="sender">Combobox som sender forespørsel. ... ComboBox1(_sender_ As Object...</param>
    ''' <param name="Tabell">Tabell kolonnen befinner seg i</param>
    ''' <param name="Kolonne">Kolonnen som skal fylle ComboBoxen</param>
    Private Sub AutoPopCbo(ByVal sender As Object, Tabell As String, Kolonne As String)
        Try
            DBConnect()
            Dim InvSqlCom As New MySqlCommand("SELECT " & Kolonne & " FROM " & Tabell, Tilkobling)
            Dim InvSqlDA As New MySqlDataAdapter
            Dim InvUtstyrComboDaT As New DataTable
            InvSqlDA.SelectCommand = InvSqlCom
            InvSqlDA.Fill(InvUtstyrComboDaT)
            DBDisconnect()
            sender.Items.Clear()
            For Each r In InvUtstyrComboDaT.Rows
                sender.Items.Add(r(Kolonne))
            Next
        Catch ex As MySqlException
            MsgBox("Feil med autoutfylling av " & CStr(sender) & ": " & ex.Message)
        End Try
    End Sub

#End Region


#Region "Form Load og Login"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        With HovedTab.TabPages

            .Remove(StartTab)
            .Remove(UtleieTab)
            .Remove(KDTab)
            .Remove(InventarTab)
            .Remove(LogiTab)
            .Remove(StatTab)
            .Remove(AdminTab)
            .Remove(DBAdminTab)
            HovedTab.SelectedTab = LoginTab

        End With


    End Sub


#End Region


#Region "TabSkifting"
    'Denne regionen er brukt til eventkoding av det som skjer da du skifter tab.
    'Nedenfor er en event som kjoeres hver gang man skifter tab. Den sjekker tabindex verdien og kjorer koden som er under det riktige caset.
    'Dette er til bruk for f.eks tilgangsskjekker og for aa laste inn verdier da man bytter til tabs.
    'Proev aa holde dette region saa ren som mulig, hvis man skal kjoere masse kode er det best aa lage egen prosedyre/funksjon for a derette kalle paa det nedenfor.
    Private Sub HovedTab_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles HovedTab.SelectedIndexChanged
        Dim HovedTabIndex As Integer
        HovedTabIndex = HovedTab.SelectedIndex + 1

        Select Case HovedTabIndex
            Case 1 'Bestemmer det som skjer etter man har valgt startmeny.
                'MsgBox("Startmeny")
                'StartMOTDUpdate()
            Case 2 'Bestemmer det som skjer etter man har valgt Utleiemeny.

                'MsgBox("Utleiemeny")

                Dim dato As Date = Date.Today

                LblUtleieDatoTxt.Text = dato
                'LblUtleieKlokke.Text = TimeOfDay
                Me.CboUtlAvd.SelectedIndex = 2





            Case 3 'Bestemmer det som skjer etter man har valgt Kundedatabasemeny.
                CmbKndSok.Items.Clear()
                Dim innhold = New String() {"ID", "Fornavn", "Etternavn", "Adresse", "Telefon", "Epost", "Rabatt Tier", "Handlet For"}

                For i As Integer = 0 To innhold.Length - 1
                    CmbKndSok.Items.Add(innhold(i))
                Next
                'MsgBox("KDBmeny")
                BtnKndEndre.Enabled = False
            Case 4 'Bestemmer det som skjer etter man har valgt Inventarmeny.
                'MsgBox("Inventarmeny")
            Case 5 'Bestemmer det som skjer etter man har valgt Logistikkmeny.
                MsgBox("Logistikkmeny")
            Case 6 'Bestemmer det som skjer etter man har valgt Statistikkmeny.
                MsgBox("Statistikkmeny")
                CmbStaAvdeling.Items.Clear()

                CmbStaType.Items.Clear()

                Dim StaLokasjoner = New String() {"Haugastøl", "Finse", "Flåm", "Voss", "Myrdal"}
                Dim StaSykkeltype = New String() {"Barnesykker", "Terrengsykkel", "Downhill", "Racer", "Bysykkel", "Elsykkel", "Tandem"}


                For i As Integer = 0 To StaLokasjoner.Length - 1
                    CmbStaAvdeling.Items.Add(StaLokasjoner(i))

                Next
                For j As Integer = 0 To StaSykkeltype.Length - 1
                    CmbStaType.Items.Add(StaSykkeltype(j))
                Next
                StaMestPopulaer()

            Case 7 'Bestemmer det som skjer etter man har valgt Adminmeny.
                MsgBox("AdminMeny")
                AdminAvdelingPopulate()
                AdminBrukerIDCalc()
            Case 8 'Bestemmer det som skjer etter man har valgt AdminDBmeny.
                MsgBox("AdminDBMeny")
        End Select
    End Sub


#End Region


#Region "StartTab"
    'Her plasseres kode som er relevant til Start Tab.
    'Variabler som brukes her skal begynne med Start. Dette er for å unngå klasj.
    'Husk kode kommentarer.

    Private Sub StartMOTDUpdate()
        Try
            DBConnect()
            Dim StartMOTDKommando As New MySqlCommand("Select * FROM message_of_the_day WHERE message_id = 1", Tilkobling)
            Dim StartMOTDAdapter As New MySqlDataAdapter
            Dim StartMOTDTable As New DataTable
            StartMOTDAdapter.SelectCommand = StartMOTDKommando
            StartMOTDAdapter.Fill(StartMOTDTable)
            DBDisconnect()
            Dim StartMOTDRow As DataRow
            For Each StartMOTDRow In StartMOTDTable.Rows
                LblStartMOTD.Text = "Message of the Day: " + StartMOTDRow("message")
            Next

        Catch StartSqlError1 As MySqlException
            MsgBox("Man får ikke koble til databasen: " & StartSqlError1.Message)

        End Try
    End Sub
#End Region


#Region "UtleieTab"
    'Her plasseres kode som er relevant til Utleie Tab.
    'Variabler som brukes her skal begynne med Leie. Dette er for å unngå klasj.
    'Husk kode kommentarer.

    'Utstyr tilgjengelig for valgt sykkel kan legges i combobox som populeres automatisk ?

    'prosedyre for reset
    Private Sub UtleieTomFelt()
        CboUtlKat.SelectedIndex = -1
        CboUtlRabatt.SelectedIndex = -1
        CboUtlSubkat.SelectedIndex = -1
        CboUtlRamme.SelectedIndex = -1
        CboUtlHjulStr.SelectedIndex = -1
        TxtUtlAntall.Text = ""
        TxtUtleieKundeSok.Text = ""
        RdbUtlTimer.Checked = False
        RdbUtlDager.Checked = False
        RdbUtlUke.Checked = False
        DtpUtleieFra.Value = Now
        DtpUtleieTil.Value = Now
        LvUtleieKunde.Items.Clear()
        LvUtleieOrdre.Items.Clear()
        LvUtlVarer.Items.Clear()

    End Sub

    ' tikkende klokke-label
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        LblUtleieKlokke.Text = Format(Now, "HH:mm:ss")
    End Sub

    'kundesøk
    Dim UtlKndSok As Integer
    Private Sub BtnUtleieKundeSok_Click(sender As Object, e As EventArgs) Handles BtnUtleieKundeSok.Click
        LvUtleieKunde.Items.Clear()

        If IsNumeric(TxtUtleieKundeSok.Text) And TxtUtleieKundeSok.Text <> "" And TxtUtleieKundeSok.Text.Length >= 8 Then
            UtlKndSok = SQLWhiteWash(TxtUtleieKundeSok.Text)
        Else
            MsgBox("Vennligst skriv inn et gyldig mobilnummer, 8 siffer")
            TxtUtleieKundeSok.Clear()
        End If

        Dim KndSQLKolonner = New String() {"kunde_id", "kunde_fornavn", "kunde_etternavn", "adresse", "telefon", "epost", "rabatt_id"}


        Try
            DBConnect()
            Dim sporring As New MySqlCommand("SELECT * FROM kunder WHERE telefon =" & UtlKndSok & "", Tilkobling)

            Dim UtleieSøkAdapter As New MySqlDataAdapter
            Dim UtleieSøkTable As New DataTable
            UtleieSøkAdapter.SelectCommand = sporring
            UtleieSøkAdapter.Fill(UtleieSøkTable)

            DBDisconnect()

            Dim KundeRow As DataRow
            Dim Kunde_ID, Kunde_Fornavn, Kunde_Etternavn, Kunde_Adresse, Kunde_Tlf, Kunde_Epost, Kunde_rabatt As String
            LvUtleieKunde.Items.Clear()

            ' sjekker om datatable gir et tomt svar. Gir beskjed om kunde ikke finnes fra før, og hopper til kundemenyen.
            If UtleieSøkTable.Rows.Count = 0 And TxtUtleieKundeSok.Text <> "" Then
                MsgBox("Ingen kunder med dette nummeret er registrert. Registrer ny kunde")
                BtnUtleieNyKunde.PerformClick()
            Else
                For Each KundeRow In UtleieSøkTable.Rows
                    Kunde_ID = KundeRow("kunde_id")
                    Kunde_Fornavn = KundeRow("kunde_fornavn")
                    Kunde_Etternavn = KundeRow("kunde_etternavn")
                    Kunde_Adresse = KundeRow("adresse")

                    Kunde_Tlf = KundeRow("telefon")
                    Kunde_Epost = KundeRow("epost")
                    Kunde_rabatt = KundeRow("rabatt_id")

                    LvUtleieKunde.Items.Add(New ListViewItem({Kunde_ID, Kunde_Fornavn, Kunde_Etternavn, Kunde_Adresse, Kunde_Tlf, Kunde_Epost, Kunde_rabatt}))
                Next

            End If


        Catch feilmelding As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)

        End Try
    End Sub

    Private Sub BtnUtlAddVare_Click(sender As Object, e As EventArgs) Handles BtnUtlAddVare.Click
        LvUtlVarer.Items.Clear()
        varehent()
    End Sub


    Private Sub varehent()

        'fyller varelistv. med data fra valgte bokser.
        '''''''''''''''''''''''''''FROM          WHERE         SELECT    =    
        'String = SQLHentIDNavn("sykkel_typer", "kategori", "type_id", verdi)

        Dim UtlKategori As String = CboUtlKat.SelectedItem
        Dim UtlSubKategori As String = CboUtlSubkat.SelectedItem
        Dim UtlRamme As String = CboUtlRamme.SelectedItem
        Dim UtlHjul As String = CboUtlHjulStr.SelectedItem
        Dim UtlSubkatID, SykkelIdRamme, SykkelIdHjul As String
        Dim SykkelSQLKolonner = New String() {"sykkel_modell", "sykkel_navn", "sykkel_status", "kategori",
                                              "sykkel_kat_timepris", "sykkel_kat_døgnpris", "sykkel_kat_ukepris"}

        'henter typeID fra sykler og utstyr etter data fra subkategori
        'hente varenummer fra sykler med valgt ramme og hjul

        If UtlKategori = "Sykkel" Then
            UtlSubkatID = SQLHentIDNavn("sykkel_typer", "kategori", "type_id", UtlSubKategori)
            SykkelIdRamme = SQLHentIDNavn("sykler", "type_id", "sykkel_id", UtlSubkatID)
            SykkelIdHjul = SQLHentIDNavn("sykler", "type_id", "sykkel_id", UtlSubkatID)

            Try
                DBConnect()
                Dim varer As New MySqlCommand("SELECT sykkel_modell, sykkel_navn, sykkel_status 
                                          FROM sykler s JOIN sykkel_typer st
                                           ON s.type_id = st.type_id WHERE s.type_id =" & UtlSubkatID & " ORDER BY sykkel_status", tilkobling)

                'Dim varer As New MySqlCommand("SELECT sykkel_modell, sykkel_navn, sykkel_status 
                '                           FROM sykler s JOIN sykkel_typer st
                '                           ON s.type_id = st.type_id WHERE s.type_id =" & UtlSubkatID & " AND st.type_id =" & SykkelIdRamme & "
                '                           ORDER BY sykkel_status", tilkobling)

                Dim VaresøkAdapter As New MySqlDataAdapter
                Dim VareSøkTable As New DataTable
                VaresøkAdapter.SelectCommand = varer
                VaresøkAdapter.Fill(VareSøkTable)

                DBDisconnect()

                Dim VareRow As DataRow
                Dim Varenummer, Varenavn, Tilgjengelig As String

                For Each VareRow In VareSøkTable.Rows
                    Varenummer = VareRow("sykkel_modell")
                    Varenavn = VareRow("sykkel_navn")
                    Tilgjengelig = VareRow("sykkel_status")


                    LvUtlVarer.Items.Add(New ListViewItem({Varenummer, Varenavn, Tilgjengelig}))
                Next

            Catch feilmelding As MySqlException
                MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)

            End Try

        Else
            UtlSubkatID = SQLHentIDNavn("utstyr_kategori", "utstyr_kat", "utstyr_kat_id", UtlSubKategori)

            Try
                DBConnect()
                Dim varer As New MySqlCommand("SELECT varenummer, utstyr_navn, utstyr_status 
                                           FROM utstyr WHERE utstyr_id =" & UtlSubkatID & "
                                            ORDER BY utstyr_status", tilkobling)

                Dim VaresøkAdapter As New MySqlDataAdapter
                Dim VareSøkTable As New DataTable
                VaresøkAdapter.SelectCommand = varer
                VaresøkAdapter.Fill(VareSøkTable)

                DBDisconnect()

                Dim VareRow As DataRow
                Dim Varenummer, Varenavn, Tilgjengelig As String

                For Each VareRow In VareSøkTable.Rows
                    Varenummer = VareRow("varenummer")
                    Varenavn = VareRow("utstyr_navn")
                    Tilgjengelig = VareRow("utstyr_status")

                    LvUtlVarer.Items.Add(New ListViewItem({Varenummer, Varenavn, Tilgjengelig}))
                Next

            Catch feilmelding As MySqlException
                MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)

            End Try
        End If









    End Sub





    ' hopp til kundetab. og ta med telefonnr
    Private Sub BtnUtleieNyKunde_Click(sender As Object, e As EventArgs) Handles BtnUtleieNyKunde.Click
        HovedTab.SelectTab(KDTab)
        TxtKndTlf.Text = TxtUtleieKundeSok.Text
    End Sub
    'tøm felt knappen
    Private Sub BtnUtlAbort_Click(sender As Object, e As EventArgs) Handles BtnUtlAbort.Click
        UtleieTomFelt()
    End Sub

    'Fyller combobox med avdelinger som er registrert i database
    Private Sub CboUtlAvd_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboUtlAvd.DropDown
        UtlAutoPopCbo(sender, "avdeling", "avd_navn")
    End Sub

    'autofyll av ramme boks fra database.
    ' denne skal hente ramme fra valgt sykkeltype
    Private Sub CboUtlRamme_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboUtlRamme.DropDown
        UtlAutoPopRamme()
    End Sub

    'autofyll av hjul combobox. 
    'Denne skal hente ramme fra valgt sykkel og ramme kombinasjon.
    Private Sub CboUtlHjulStr_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboUtlHjulStr.DropDown
        UtlAutoPopCbo(sender, "sykler", "hjul_str")
    End Sub

    'hva som skjer når hovedkategori endres
    Private Sub CboUtlKat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboUtlKat.SelectedIndexChanged
        'Dropdownlisten Subkategori endrer kategorier (leses fra databasen) _
        'basert på om det er sykkel eller utstyr som er valgt.
        'ramme og hjul valg deaktiveres om det er valgt utstyr.

        If CboUtlKat.SelectedItem = "Utstyr" Then

            CboUtlRamme.Enabled = False
            CboUtlHjulStr.Enabled = False
            'LvUtlVarer.Columns.Clear()
            CboUtlSubkat.Items.Clear()
            CboUtlRamme.Items.Clear()
            CboUtlHjulStr.Items.Clear()
            UtlAutoPopUtstyr()
            'varehent()
        Else

            'LvUtlVarer.Items.Clear()
            CboUtlRamme.Enabled = True
            CboUtlHjulStr.Enabled = True
            CboUtlSubkat.Items.Clear()
            CboUtlRamme.Items.Clear()
            CboUtlHjulStr.Items.Clear()
            UtlAutoPopSykkel()
            'varehent()
        End If

    End Sub

    'Hva skjer når rammeboks trykkes på.
    Private Sub CboUtlRamme_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles CboUtlRamme.DropDown
        'CboUtlRamme.Items.Clear()
        CboUtlHjulStr.Items.Clear()
        UtlAutoPopRamme()
    End Sub

    'resetter rammeboks om sykkel kategori endres.
    Private Sub CboUtlSubkat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CboUtlSubkat.DropDown

        If CboUtlSubkat.SelectedValue = "Sykkel" Then
            CboUtlRamme.Items.Clear()
            CboUtlHjulStr.Items.Clear()
            UtlAutoPopRamme()
        End If

    End Sub



    'Metode som fyller combobox (som kaller) med data fra database.
    'Tabell og kolonne det skal leses fra må angis som argumenter.
    Private Sub UtlAutoPopCbo(sender, tabell, kolonne)
        Try
            DBConnect()
            Dim UtlSqlCom As New MySqlCommand("SELECT " & kolonne & " FROM " & tabell, Tilkobling)
            Dim UtlSqlDA As New MySqlDataAdapter
            Dim UtlUtstyrComboDaT As New DataTable
            UtlSqlDA.SelectCommand = UtlSqlCom
            UtlSqlDA.Fill(UtlUtstyrComboDaT)
            DBDisconnect()
            sender.Items.Clear()
            Dim UtlUtstyrRow As DataRow
            Dim UtlUtstyrString As String
            For Each UtlUtstyrRow In UtlUtstyrComboDaT.Rows
                UtlUtstyrString = UtlUtstyrRow(kolonne)
                sender.Items.Add(UtlUtstyrString)
            Next
        Catch ex As MySqlException
            MsgBox("Feil med autoutfylling av " & CStr(sender) & ": " & ex.Message)
        End Try
    End Sub

    'Combobox Subkategori styres av indexchanged på combobox kategori. Dermed vil et kall på UtlCboPop fra _
    'combobox for kategori føre til endringer i kategori og ikke subkategori som ønsket.
    'Derfor egne autopopulate for sykkel og utstyr.
    Private Sub UtlAutoPopUtstyr()
        Try
            DBConnect()
            Dim UtlSqlCom As New MySqlCommand("SELECT utstyr_kat FROM utstyr_kategori", Tilkobling)
            Dim UtlSqlDA As New MySqlDataAdapter
            Dim UtlUtstyrComboDaT As New DataTable
            UtlSqlDA.SelectCommand = UtlSqlCom
            UtlSqlDA.Fill(UtlUtstyrComboDaT)
            DBDisconnect()
            CboUtlSubkat.Items.Clear()
            Dim UtlUtstyrRow As DataRow
            Dim UtlUtstyrString As String
            For Each UtlUtstyrRow In UtlUtstyrComboDaT.Rows
                UtlUtstyrString = UtlUtstyrRow("utstyr_kat")
                CboUtlSubkat.Items.Add(UtlUtstyrString)
            Next
        Catch ex As MySqlException
            MsgBox("Feil med autoutfylling av utstyrskategorier: " & ex.Message)
        End Try
    End Sub

    Private Sub UtlAutoPopSykkel()
        Try
            DBConnect()
            Dim UtlSqlCom As New MySqlCommand("SELECT kategori FROM sykkel_typer", Tilkobling)
            Dim UtlSqlDA As New MySqlDataAdapter
            Dim UtlSykkelComboDaT As New DataTable
            UtlSqlDA.SelectCommand = UtlSqlCom
            UtlSqlDA.Fill(UtlSykkelComboDaT)
            DBDisconnect()
            CboUtlSubkat.Items.Clear()
            Dim UtlSykkelRow As DataRow
            Dim UtlSykkelString As String
            For Each UtlSykkelRow In UtlSykkelComboDaT.Rows
                UtlSykkelString = UtlSykkelRow("kategori")
                CboUtlSubkat.Items.Add(UtlSykkelString)
            Next
        Catch ex As MySqlException
            MsgBox("Feil med autoutfylling av sykkelkategorier: " & ex.Message)
        End Try
    End Sub

    Private Sub UtlAutoPopRamme()

        Dim UtlSubKategori As String = CboUtlSubkat.SelectedItem
        Dim UtlSubkatID = SQLHentIDNavn("sykkel_typer", "kategori", "type_id", UtlSubKategori)
        Try
            DBConnect()

            Dim UtlSqlCom As New MySqlCommand("SELECT sykkel_ramme FROM sykler 
                                                WHERE type_id =" & UtlSubkatID & "", tilkobling)
            Dim UtlSqlDA As New MySqlDataAdapter
            Dim UtlRammeComboDaT As New DataTable
            UtlSqlDA.SelectCommand = UtlSqlCom
            UtlSqlDA.Fill(UtlRammeComboDaT)
            DBDisconnect()
            CboUtlRamme.Items.Clear()
            Dim UtlRammeRow As DataRow
            Dim UtlRammeString As String
            For Each UtlRammeRow In UtlRammeComboDaT.Rows
                UtlRammeString = UtlRammeRow("sykkel_ramme")
                CboUtlRamme.Items.Add(UtlRammeString)
            Next
        Catch ex As MySqlException
            MsgBox("Feil med autoutfylling av utstyrskategorier: " & ex.Message)
        End Try
    End Sub





#End Region


#Region "KundeTab"
    'Her plasseres kode som er relevant til Kunde Tab.
    'Variabler som brukes her skal begynne med Kunde. Dette er for å unngå klasj.
    'Husk kode kommentarer.

    'Testing testing 123'
    'Håvard var her >:)'

    Dim KndFornavn As String
    Dim KndEtternavn As String
    Dim KndAdresse As String
    Dim KndFodselsar As Date
    Dim KndTlf As String
    Dim KndEpost As String
    Dim KndTegn As String = ".-@ "

    'Bedre måte å gjøre dette på:    '
    'lage en sub som utfører de repeterende oppgavene.

    '[Morten] Har laget noen korte funksjoner for inputsjekk som kan brukes for de forskjellige textbokser.
    'Funksjonene tar to variabler; keypresseventargs, og en string bestående av ekstra tillatte tegn.
    'Mulig dette ikke er en optimal løsning så kom gjerne med forslag :) (om noen noengang leser dette)

    Private Function KndSjekkInputTall(e)
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Function

    Private Function KndSjekkInputBokstav(e, tegn)
        If Not Char.IsLetter(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) _
            AndAlso Not tegn.Contains(e.KeyChar) Then
            e.Handled = True
        End If
    End Function

    Private Function KndSjekkInputBokstavTall(e, tegn)
        If Not Char.IsLetterOrDigit(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) _
            AndAlso Not tegn.Contains(e.KeyChar) Then
            e.Handled = True
        End If
    End Function

    Private Sub TxtKndFornavn_TextChanged(sender As Object, e As KeyPressEventArgs) Handles TxtKndFornavn.KeyPress
        Dim KndTegn As String = "-"
        KndSjekkInputBokstav(e, KndTegn)
        'If Not (Asc(e.KeyChar) = 8) Then
        '    Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyzæøå -"
        '    If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
        '        e.KeyChar = ChrW(0)
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub TxtKndEtternavn_TextChanged(sender As Object, e As KeyPressEventArgs) Handles TxtKndEtternavn.KeyPress
        Dim KndTegn As String = "-"
        KndSjekkInputBokstav(e, KndTegn)
        'If Not (Asc(e.KeyChar) = 8) Then
        '    Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyzæøå -"
        '    If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
        '        e.KeyChar = ChrW(0)
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub TxtKndAdresse_TextChanged(sender As Object, e As KeyPressEventArgs) Handles TxtKndAdresse.KeyPress
        Dim KndTegn As String = "-."
        KndSjekkInputBokstavTall(e, KndTegn)
        'If Not (Asc(e.KeyChar) = 8) Then
        '    Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyzæøå1234567890"
        '    If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
        '        e.KeyChar = ChrW(0)
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub TxtKndEpost_TextChanged(sender As Object, e As KeyPressEventArgs) Handles TxtKndEpost.KeyPress
        Dim KndTegn As String = "-.@_"
        KndSjekkInputBokstavTall(e, KndTegn)
        'If Not (Asc(e.KeyChar) = 8) Then
        '    Dim allowedChars As String = "abcdefghijklmnopqrstuvwxyzæøå1234567890@!#$%&'*+-/=?^_`{|}~;."
        '    If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
        '        e.KeyChar = ChrW(0)
        '        e.Handled = True
        '    End If
        'End If
    End Sub

    Private Sub TxtKndTlf_TextChanged(sender As Object, e As KeyPressEventArgs) Handles TxtKndTlf.KeyPress
        KndSjekkInputTall(e)
        'If Not (Asc(e.KeyChar) = 8) Then
        '    Dim allowedChars As String = "1234567890"
        '    If Not allowedChars.Contains(e.KeyChar.ToString.ToLower) Then
        '        e.KeyChar = ChrW(0)
        '        e.Handled = True
        '    End If
        'End If
    End Sub


    'LEGG INN NYE KUNDER
    Private Sub BtnKndRegistrer_Click(sender As Object, e As EventArgs) Handles BtnKndRegistrer.Click

        Dim KnDFdato As Date = DateKndReg.Value
        KndFornavn = TxtKndFornavn.Text
        KndEtternavn = TxtKndEtternavn.Text
        KndAdresse = TxtKndAdresse.Text
        KndTlf = TxtKndTlf.Text
        KndEpost = TxtKndEpost.Text

        'Insert new data with SQL

        If KndFornavn = "" Then
            MsgBox("Vennligst fyll inn fornavn")
            Exit Sub
        End If
        If KndEtternavn = "" Then
            MsgBox("Vennligst fyll inn etternavn")
            Exit Sub
        End If
        If KndAdresse = "" Then
            MsgBox("Vennligst fyll inn adresse")
            Exit Sub
        End If
        If KndTlf = "" Then
            MsgBox("Vennligst fyll inn telefonnummer")
            Exit Sub
        End If
        If KndEpost = "" Then
            MsgBox("Vennligst fyll inn epost")
            Exit Sub
        End If

        Try
            DBConnect()
            Dim sporring As New MySqlCommand("INSERT INTO kunder (kunde_fdato, kunde_fornavn, " _
                & "kunde_etternavn, adresse, telefon, epost, rabatt_id, handlet_for) VALUES('" _
                & KnDFdato.ToString("yyyy-MM-dd") & "', '" & KndFornavn & "', '" & KndEtternavn _
                & "', '" & KndAdresse & "', " & KndTlf & ", '" & KndEpost & "', 1, 0)", Tilkobling)
            sporring.ExecuteNonQuery()
            DBDisconnect()
            MsgBox("Innlegging var vellykket")
        Catch feilmelding As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)
        End Try

    End Sub

    Dim KndSokKundeID

    Dim KndFornavnSelected
    Dim KndEtternavnSelected
    Dim KndAdresseSelected
    Dim KndTlfSelected
    Dim KndEpostSelected
    Dim KndRabattSelected
    Dim KndHandletSelected
    Dim KndFdatoSelected As Date


    'ENDRE FANE
    Private Sub BtnKndKundeID_Click(sender As Object, e As EventArgs) Handles BtnKndKundeID.Click

        KndSokKundeID = TxtKndKundeID.Text
        If KndSokKundeID = "" Then
            MsgBox("Vennligst fyll inn ID for søk")
            Exit Sub
        End If

        Try
            DBConnect()
            Dim sporring As New MySqlCommand("SELECT * FROM kunder WHERE kunde_id = " & KndSokKundeID, Tilkobling)
            Dim tempVarSporring = sporring.ExecuteReader

            While tempVarSporring.Read
                KndFornavnSelected = tempVarSporring("kunde_fornavn")
                KndEtternavnSelected = tempVarSporring("kunde_etternavn")
                KndAdresseSelected = tempVarSporring("adresse")
                KndTlfSelected = tempVarSporring("telefon")
                KndEpostSelected = tempVarSporring("epost")
                KndRabattSelected = tempVarSporring("rabatt_id")
                KndHandletSelected = tempVarSporring("handlet_for")
                KndFdatoSelected = tempVarSporring("kunde_fdato")
            End While

            tempVarSporring.Close()
            DBDisconnect()

            TxtKndEndreFN.Text = KndFornavnSelected
            TxtKndEndreEN.Text = KndEtternavnSelected
            TxtKndEndreAdr.Text = KndAdresseSelected
            TxtKndEndreTlf.Text = KndTlfSelected
            TxtKndEndreEP.Text = KndEpostSelected
            TxtKndEndreRbt.Text = KndRabattSelected
            TxtKndEndreHF.Text = KndHandletSelected
            DateKndEndre.Value = KndFdatoSelected

        Catch feilmelding As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)
        End Try

        If KndFornavnSelected <> "" Then
            BtnKndEndre.Enabled = True
        Else
            MsgBox("Beklager, ingen treff på det nummeret")
            BtnKndEndre.Enabled = False
        End If

    End Sub

    Dim KndSokInput
    Dim KndSokSelectedTag

    Private Sub BtnKndSok_Click(sender As Object, e As EventArgs) Handles BtnKndSok.Click
        KndSokInput = TxtKndSok.Text
        Dim KndTempVar = CmbKndSok.SelectedIndex

        If KndSokInput = "" Then
            MsgBox("Vennligst fyll inn gylding søkeord")
            Exit Sub
        End If
        If KndTempVar = -1 Then
            MsgBox("Vennligst velg en søkekategori")
            Exit Sub
        End If

        Dim KndSQLKolonner = New String() {"kunde_id", "kunde_fornavn", "kunde_etternavn",
            "adresse", "telefon", "epost", "rabatt_id", "handlet_for"}
        KndSokSelectedTag = KndSQLKolonner(KndTempVar)

        Try
            DBConnect()
            Dim sporring As New MySqlCommand("SELECT * FROM kunder WHERE " & KndSokSelectedTag _
                & " LIKE '%" & KndSokInput & "%'", Tilkobling)


            Dim KndSearchAdapter As New MySqlDataAdapter
            Dim KndSearchTable As New DataTable
            KndSearchAdapter.SelectCommand = sporring
            KndSearchAdapter.Fill(KndSearchTable)

            DBDisconnect()

            Dim KundeRow As DataRow
            Dim Kunde_ID, Kunde_Fornavn, Kunde_Etternavn, Kunde_Adresse, Kunde_Tlf, Kunde_Epost,
                Kunde_rabatt, Kunde_HF, KundeSok As String
            LvKndSok.Items.Clear()
            For Each KundeRow In KndSearchTable.Rows
                Kunde_ID = KundeRow("kunde_id")
                Kunde_Fornavn = KundeRow("kunde_fornavn")
                Kunde_Etternavn = KundeRow("kunde_etternavn")
                Kunde_Adresse = KundeRow("adresse")

                Kunde_Tlf = KundeRow("telefon")
                Kunde_Epost = KundeRow("epost")
                Kunde_rabatt = KundeRow("rabatt_id")
                Kunde_HF = KundeRow("handlet_for")

                KundeSok = KundeRow(KndSokSelectedTag)
                LvKndSok.Items.Add(New ListViewItem({Kunde_ID, Kunde_Fornavn, Kunde_Etternavn,
                    Kunde_Adresse, Kunde_Tlf, Kunde_Epost, Kunde_rabatt, Kunde_HF, KundeSok}))
            Next



        Catch feilmelding As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)

        End Try

        'SQL for søk 
        'Select * from Kunder where kndsokselectedtag = kndsokInput
        'Skriv ut til listbox i søk
        'Legg kundeID i KndSokKundeID for bruk i endring



    End Sub

    Private Sub BtnKndEndre_Click(sender As Object, e As EventArgs) Handles BtnKndEndre.Click
        Dim KnDFdato As Date = DateKndEndre.Value
        Dim selectedKundeId = TxtKndKundeID.Text
        Dim KndChangeValueFN = TxtKndEndreFN.Text
        Dim KndChangeValueEN = TxtKndEndreEN.Text
        Dim KndChangeValueAdr = TxtKndEndreAdr.Text
        Dim KndChangeValueTlf = TxtKndEndreTlf.Text
        Dim KndChangeValueEP = TxtKndEndreEP.Text
        Dim KndChangeValueRbt = TxtKndEndreRbt.Text
        Dim KndChangeValueHF = TxtKndEndreHF.Text

        If selectedKundeId = "" Then
            MsgBox("Vennligst fyll inn et gyldig ID-nummer")
            Exit Sub
        End If
        If KndChangeValueFN = "" Then
            MsgBox("Vennligst fyll inn et fornavn")
            Exit Sub
        End If
        If KndChangeValueEN = "" Then
            MsgBox("Vennligst fyll inn et etternavn")
            Exit Sub
        End If
        If KndChangeValueAdr = "" Then
            MsgBox("Vennligst fyll inn en adresse")
            Exit Sub
        End If
        If KndChangeValueTlf = "" And KndChangeValueTlf.Length <> 8 Then
            MsgBox("Vennligst fyll inn et gyldig telefonnummer")
            Exit Sub
        End If
        If KndChangeValueEP = "" Then
            MsgBox("Vennligst fyll inn en epostadresse")
            Exit Sub
        End If
        If KndChangeValueRbt = "" Then
            KndChangeValueRbt = "1"
        End If
        If KndChangeValueHF = "" Then
            KndChangeValueHF = "0"
        End If

        Try
            DBConnect()
            Dim sporring As New MySqlCommand("UPDATE kunder SET kunde_fdato='" & KnDFdato.ToString("yyyy-MM-dd") _
                & " , kunde_fornavn = '" & KndChangeValueFN & "', kunde_etternavn = '" & KndChangeValueEN _
                & "', adresse = '" & KndChangeValueAdr & "', telefon = " & KndChangeValueTlf & ", epost = '" _
                & KndChangeValueEP & "', rabatt_id = " & KndChangeValueRbt & ", handlet_for = " & KndChangeValueHF _
                & " WHERE kunde_id = " & selectedKundeId, Tilkobling)
            sporring.ExecuteNonQuery()
            DBDisconnect()
            MsgBox("Endring vellykket")
        Catch feilmelding As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)

        End Try

    End Sub


#End Region


#Region "InventarTab"
    'Her plasseres kode som er relevant til Inventar Tab.
    'Variabler som brukes her skal begynne med Inv. Dette er for å unngå klasj.
    'Husk kode kommentarer.

    'Prosedyre for tømming av alle felt i skjema
    Private Sub InvTomFelt()
        CboInvKategori.SelectedIndex = -1
        CboInvSubkategori.SelectedIndex = -1
        CboInvAvdeling.SelectedIndex = -1
        TxtInvProduktnavn.Text = ""
        TxtInvVareNummer.Text = ""
        TxtInvInnkjopspris.Text = ""
        TxtInvRamme.Text = ""
        TxtInvHjulstorrelse.Text = ""
        TxtInvGirsystem.Text = ""
        CboInvForhandler.SelectedIndex = -1
        CboInvStatus.SelectedIndex = -1
        CboInvSkadet.SelectedIndex = -1
        CboInvSavnet.SelectedIndex = -1
        ChkInvBarneHenger.Checked = False
        ChkInvLastehenger.Checked = False
        ChkInvBarnesete.Checked = False
        ChkInvSykkelveske.Checked = False
    End Sub

    'Tillater kun inntasting av tall i textbox for innkjøpspris
    Private Sub TxtInvInnkjopspris_Keypress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
        Handles TxtInvInnkjopspris.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    'Tillater kun inntasting av tall i textbox for hjulstørrelse
    Private Sub TxtInvHjulstorrelse_Keypress(ByVal sender As Object, ByVal e As KeyPressEventArgs) _
        Handles TxtInvHjulstorrelse.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    'Endrer status på felt i skjema avhengig om det er sykkel eller utstyr som skal registreres.
    'Med kategori "Sykkel" valgt vil alle felt være tilgjengelig.
    'For "Utstyr" vil kun felt som er en kolonne i tabellen for utstyr være aktivert.
    'Innhold i Combobox Subkategori endrer innhold (leses fra databasen) basert på om det er sykkel eller _
    'utstyr som er valgt.
    'Også kolonner i listview (LvInvSok) endres avhengig av valgt kategori (sykkel eller utstyr)
    Private Sub CboInvKategori_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles CboInvKategori.SelectedIndexChanged
        If CboInvKategori.SelectedItem = "Utstyr" Then
            CboInvAvdeling.Enabled = False
            TxtInvRamme.Enabled = False
            TxtInvHjulstorrelse.Enabled = False
            TxtInvGirsystem.Enabled = False
            CboInvStatus.Enabled = False
            CboInvSkadet.Enabled = False
            CboInvSavnet.Enabled = False
            ChkInvBarneHenger.Enabled = False
            ChkInvBarnesete.Enabled = False
            ChkInvLastehenger.Enabled = False
            ChkInvSykkelveske.Enabled = False
            LvInvSok.Items.Clear()
            LvInvSok.Columns.Clear()
            LvInvSok.Columns.AddRange(New ColumnHeader() {ID, Me.Produktnavn, Me.Kategori,
                Me.Varenummer, Me.Innkjøpspris, Me.Forhandler})
            CboInvSubkategori.Items.Clear()
            InvAutoPopUtstyr()
        Else
            CboInvAvdeling.Enabled = True
            TxtInvRamme.Enabled = True
            TxtInvHjulstorrelse.Enabled = True
            TxtInvGirsystem.Enabled = True
            CboInvStatus.Enabled = True
            CboInvSkadet.Enabled = True
            CboInvSavnet.Enabled = True
            ChkInvBarneHenger.Enabled = True
            ChkInvBarnesete.Enabled = True
            ChkInvLastehenger.Enabled = True
            ChkInvSykkelveske.Enabled = True
            LvInvSok.Items.Clear()
            LvInvSok.Columns.Clear()
            'LvInvSok.Columns.AddRange(InvLvKolonnerSykkel())
            LvInvSok.Columns.AddRange(New ColumnHeader() {Me.ID, Me.Produktnavn, Me.Varenummer,
                        Me.Kategori, Me.Ramme, Me.Girsystem, Me.Hjulstørrelse, Me.Innkjøpspris, Me.Avdeling,
                        Me.Forhandler, Me.Status, Me.Skadet, Me.Savnet})

            CboInvSubkategori.Items.Clear()
            InvAutoPopSykkel()
        End If
    End Sub

    'Fyller combobox med forhandlere som er registrert i database
    Private Sub CboInvForhandler_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles CboInvForhandler.DropDown
        AutoPopCbo(sender, "forhandler", "forhandler_navn")
    End Sub

    'Fyller combobox med avdelinger som er registrert i database
    Private Sub CboInvAvdeling_SelectedIndexChanged(sender As Object, e As EventArgs) _
        Handles CboInvAvdeling.DropDown
        AutoPopCbo(sender, "avdeling", "avd_navn")
    End Sub

    'Combobox Subkategori styres av indexchanged på combobox kategori. Dermed vil et kall på InvCboPop fra _
    'combobox for kategori føre til endringer i kategori og ikke subkategori som ønsket.
    'Derfor egne autopopulate for sykkel og utstyr. Disse kalles fra combobox for kategori og endrer _
    'innhold i subkategori.
    Private Sub InvAutoPopUtstyr()
        Dim InvUtstyrComboDaT As New DataTable
        InvUtstyrComboDaT = SQLSelect("utstyr_kategori", "utstyr_kat", "1")
        CboInvSubkategori.Items.Clear()
        Try
            For Each r In InvUtstyrComboDaT.Rows
                CboInvSubkategori.Items.Add(r("utstyr_kat"))
            Next
        Catch ex As Exception
            MsgBox("Feil med autoutfylling av utstyrskategorier fra datatabell: " & ex.Message)
        End Try
    End Sub

    Private Sub InvAutoPopSykkel()
        Dim InvSykkelComboDaT As New DataTable
        InvSykkelComboDaT = SQLSelect("sykkel_typer", "kategori", "1")
        CboInvSubkategori.Items.Clear()
        Try
            For Each r In InvSykkelComboDaT.Rows
                CboInvSubkategori.Items.Add(r("kategori"))
            Next
        Catch ex As Exception
            MsgBox("Feil med autoutfylling av sykkelkategorier fra datatabell: " & ex.Message)
        End Try
    End Sub

    Private Sub InvRegistrerSykkel()
        If CboInvAvdeling.SelectedIndex = -1 Or CboInvForhandler.SelectedIndex = -1 Or
                CboInvSubkategori.SelectedIndex = -1 Or TxtInvProduktnavn.Text.Trim = "" Or
                TxtInvVareNummer.Text.Trim = "" Or TxtInvInnkjopspris.Text.Trim = "" Or
                TxtInvRamme.Text.Trim = "" Or TxtInvHjulstorrelse.Text.Trim = "" Or
                TxtInvGirsystem.Text.Trim = "" Or CboInvStatus.SelectedIndex = -1 Or
                CboInvSkadet.SelectedIndex = -1 Or CboInvSavnet.SelectedIndex = -1 Then

            MsgBox("Vennligst fyll inn alle felt")

        Else

            Dim InvKategori, InvSubkategoriNavn, InvAvdelingNavn, InvProduktnavn, InvVarenummer,
                InvInnkjopspris, InvRamme, InvHjulstorrlese, InvGirsystem, InvForhandlerNavn,
                InvStatus, InvSkadet, InvSavnet, InvForhandlerID, InvAvdelingID, InvSubKategoriID,
                InvSykkelID, InvSisteID, InvRegKolonner, InvRegTabell, InvRegVerdier As String

            Dim InvSqlLeser As MySqlDataReader
            Dim InvBekreftSykkelReg As DialogResult

            InvKategori = CboInvKategori.SelectedItem
            InvSubkategoriNavn = CboInvSubkategori.SelectedItem
            InvAvdelingNavn = CboInvAvdeling.SelectedItem
            InvProduktnavn = SQLWhiteWash(TxtInvProduktnavn.Text.Trim)
            InvVarenummer = SQLWhiteWash(TxtInvVareNummer.Text.Trim)
            InvInnkjopspris = SQLWhiteWash(TxtInvInnkjopspris.Text.Trim)
            InvRamme = SQLWhiteWash(TxtInvRamme.Text.Trim)
            InvHjulstorrlese = SQLWhiteWash(TxtInvHjulstorrelse.Text.Trim)
            InvGirsystem = SQLWhiteWash(TxtInvGirsystem.Text.Trim)
            InvForhandlerNavn = CboInvForhandler.SelectedItem
            InvStatus = CboInvStatus.SelectedItem
            InvSkadet = CboInvSkadet.SelectedIndex
            InvSavnet = CboInvSavnet.SelectedIndex
            InvForhandlerID = SQLHentIDNavn("forhandler", "forhandler_navn",
                    "forhandler_id", InvForhandlerNavn)
            InvSubKategoriID = SQLHentIDNavn("sykkel_typer", "kategori", "type_id", InvSubkategoriNavn)
            InvAvdelingID = SQLHentIDNavn("avdeling", "avd_navn", "avdeling_id", InvAvdelingNavn)

            InvBekreftSykkelReg = MsgBox("Bekreft registrering av sykkel", MsgBoxStyle.OkCancel)
            If InvBekreftSykkelReg = DialogResult.OK Then

                InvRegKolonner = "(forhandler_id, type_id, avdeling_id, sykkel_navn, sykkel_modell, " &
                    "sykkel_pris, sykkel_status, hjul_str, sykkel_ramme, girsystem, savnet, skadet)"
                InvRegTabell = "sykler"
                InvRegVerdier = "('" + InvForhandlerID & "', '" & InvSubKategoriID & "', '" _
                & InvAvdelingID + "', '" & InvProduktnavn & "', '" & InvVarenummer & "', '" _
                & InvInnkjopspris & "', '" & InvStatus & "', '" & InvHjulstorrlese & "', '" _
                & InvRamme & "', '" & InvGirsystem & "', '" & InvSavnet & "', '" & InvSkadet & "')"

                SQLInsert(InvRegTabell, InvRegKolonner, InvRegVerdier)

                Try
                    DBConnect()
                    InvSisteID = "SELECT LAST_INSERT_ID();"
                    Dim InvSqlSisteID As New MySqlCommand(InvSisteID, Tilkobling)
                    InvSqlLeser = InvSqlSisteID.ExecuteReader()
                    While InvSqlLeser.Read()
                        InvSykkelID = InvSqlLeser(0)
                    End While
                    InvSqlLeser.Close()
                    DBDisconnect()
                    InvTomFelt()
                Catch ex As MySqlException
                    MsgBox("Feil ved henting av siste ID:" & vbNewLine & ex.Message)
                End Try

                'Registerer valgt utstyr på siste sykkelID
                If ChkInvBarneHenger.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvSykkelID & ", 4)")
                End If
                If ChkInvBarnesete.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvSykkelID & ", 3)")
                End If
                If ChkInvLastehenger.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvSykkelID & ", 5)")
                End If
                If ChkInvSykkelveske.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvSykkelID & ", 2)")
                End If
            End If
        End If
    End Sub

    Private Sub InvRegistrerUtstyr()
        If CboInvForhandler.SelectedIndex = -1 Or CboInvSubkategori.SelectedIndex = -1 Or
            TxtInvProduktnavn.Text.Trim = "" Or TxtInvVareNummer.Text.Trim = "" Or
            TxtInvInnkjopspris.Text.Trim = "" Then

            MsgBox("Vennligst fyll inn alle felt")

        Else

            Dim InvKategori, InvSubkategoriNavn, InvSubKategoriID, InvForhandlerNavn, InvForhandlerID,
                InvProduktnavn, InvVarenummer, InvInnkjopspris, InvSpForhandlerID, InvSpSubkategoriID,
                InvKolonnerUtstyr, InvVerdierUtstyr As String

            Dim InvBekreftUtstyrReg As DialogResult
            Dim InvSubkategoriIDTabell, InvForhandlerIDTabell As DataTable

            InvKategori = CboInvKategori.SelectedItem
            InvSubkategoriNavn = CboInvSubkategori.SelectedItem
            InvForhandlerNavn = CboInvForhandler.SelectedItem
            InvProduktnavn = SQLWhiteWash(TxtInvProduktnavn.Text.Trim)
            InvVarenummer = SQLWhiteWash(TxtInvVareNummer.Text.Trim)
            InvInnkjopspris = SQLWhiteWash(TxtInvInnkjopspris.Text.Trim)

            InvSpForhandlerID = "forhandler_navn='" & InvForhandlerNavn & "'"
            InvSpSubkategoriID = "utstyr_kat='" & InvSubkategoriNavn & "'"

            InvSubkategoriIDTabell = SQLSelect("utstyr_kategori", "utstyr_kat_id",
                 InvSpSubkategoriID)
            InvForhandlerIDTabell = SQLSelect("forhandler", "forhandler_id",
                InvSpForhandlerID)

            For Each i In InvSubkategoriIDTabell.Rows
                InvSubKategoriID = i("utstyr_kat_id")
            Next
            For Each i In InvForhandlerIDTabell.Rows
                InvForhandlerID = i("forhandler_id")
            Next

            InvBekreftUtstyrReg = MsgBox("Bekreft registrering av utstyr", MsgBoxStyle.OkCancel)
            If InvBekreftUtstyrReg = DialogResult.OK Then

                InvKolonnerUtstyr = "(utstyr_navn, varenummer, utstyr_pris, " _
                    & "forhandler_id, utstyr_kat_id)"
                InvVerdierUtstyr = "('" & InvProduktnavn & "' ,'" & InvVarenummer _
                        & "', '" & InvInnkjopspris & "', '" & InvForhandlerID & "', '" & InvSubKategoriID & "')"
                SQLInsert(InvKolonnerUtstyr, "utstyr", InvVerdierUtstyr)

            End If
        End If
    End Sub

    Private Sub InvSokSykkel()

        'Oppbygging av SQL spørring
        Dim InvSpInit As String = "SELECT s.sykkel_id, s.sykkel_navn, s.sykkel_modell, " _
                & "sykkel_typer.kategori, s.sykkel_ramme, s.girsystem, s.hjul_str, " _
                & "s.sykkel_pris, avdeling.avd_navn, forhandler.forhandler_navn, s.sykkel_status, " _
                & "s.skadet, s.savnet FROM sykler AS s LEFT JOIN avdeling ON " _
                & "s.avdeling_id=avdeling.avdeling_id LEFT JOIN forhandler ON " _
                & "s.forhandler_id=forhandler.forhandler_id LEFT JOIN sykkel_typer ON " _
                & "s.type_id=sykkel_typer.type_id "
        Dim InvSpUtstyrJoin As String = "INNER JOIN sykkel_utstyr AS su ON s.sykkel_id=su.sykkel_id " _
                & "INNER JOIN utstyr_kategori AS uk ON su.utstyr_kat_id=uk.utstyr_kat_id "
        Dim InvSpSykkelNavn As String = "sykkel_navn LIKE '%" & SQLWhiteWash(TxtInvProduktnavn.Text.Trim) & "%'"
        Dim InvSpsykkelModell As String = "kategori LIKE '%" & CboInvSubkategori.SelectedItem & "%'"
        Dim InvSpTypeid As String = "sykkel_modell LIKE '%" & SQLWhiteWash(TxtInvVareNummer.Text.Trim) & "%'"
        Dim InvSpSykkelRamme As String = "sykkel_ramme LIKE '%" & SQLWhiteWash(TxtInvRamme.Text.Trim) & "%'"
        Dim InvSpGirsystem As String = "girsystem LIKE '%" & SQLWhiteWash(TxtInvGirsystem.Text.Trim) & "%'"
        Dim InvSpHjulstorrelse As String = "hjul_str LIKE '%" & SQLWhiteWash(TxtInvHjulstorrelse.Text.Trim) & "%'"
        Dim InvSpSykkelPris As String = "sykkel_pris LIKE '%" & SQLWhiteWash(TxtInvInnkjopspris.Text.Trim) & "%'"
        Dim InvSpAvdeling As String = "avd_navn LIKE '%" & CboInvAvdeling.SelectedItem & "%'"
        Dim InvSpForhandlerID As String = "forhandler_navn LIKE '%" & CboInvForhandler.SelectedItem & "%'"
        Dim InvSpSykkelStatus As String = "sykkel_status LIKE '%" & CboInvStatus.SelectedItem & "%'"
        Dim InvSpSykkelUtstyr As String = ""
        Dim InvSpUtstyrAntall As Integer = -1

        Dim InvSpSkadet As String
        Dim InvSpSavnet As String

        Dim InvSqlSok As MySqlCommand
        Dim invSqlLeser As MySqlDataReader

        'Denne if-else biten sørger for at kun sykler med valgte utstyrskategorier (checkbokser) _
        'tas med i resultatet
        If ChkInvSykkelveske.Checked = True Or ChkInvBarnesete.Checked = True _
                Or ChkInvBarneHenger.Checked = True Or ChkInvLastehenger.Checked = True Then
            InvSpInit = InvSpInit + InvSpUtstyrJoin
            If ChkInvSykkelveske.Checked = True Then
                InvSpSykkelUtstyr = InvSpSykkelUtstyr + " AND(su.utstyr_kat_id=2"
                InvSpUtstyrAntall = InvSpUtstyrAntall + 1
            End If
            If ChkInvBarnesete.Checked = True Then
                If InvSpUtstyrAntall = -1 Then
                    InvSpSykkelUtstyr = InvSpSykkelUtstyr + " AND(su.utstyr_kat_id=3"
                Else
                    InvSpSykkelUtstyr = InvSpSykkelUtstyr + " OR su.utstyr_kat_id=3"
                End If
                InvSpUtstyrAntall = InvSpUtstyrAntall + 1
            End If
            If ChkInvBarneHenger.Checked = True Then
                If InvSpUtstyrAntall = -1 Then
                    InvSpSykkelUtstyr = InvSpSykkelUtstyr + " AND(su.utstyr_kat_id=4"
                Else
                    InvSpSykkelUtstyr = InvSpSykkelUtstyr + " OR su.utstyr_kat_id=4"
                End If
                InvSpUtstyrAntall = InvSpUtstyrAntall + 1
            End If
            If ChkInvLastehenger.Checked = True Then
                If InvSpUtstyrAntall = -1 Then
                    InvSpSykkelUtstyr = InvSpSykkelUtstyr + " AND(su.utstyr_kat_id=5"
                Else
                    InvSpSykkelUtstyr = InvSpSykkelUtstyr + " OR su.utstyr_kat_id=5"
                End If
                InvSpUtstyrAntall = InvSpUtstyrAntall + 1
            End If
            InvSpSykkelUtstyr = InvSpSykkelUtstyr & ")" & " GROUP by s.sykkel_id HAVING COUNT(su.sykkel_id) > " _
                    & InvSpUtstyrAntall
        End If

        'Unngår NULL verdi fra combobox skadet og savnet og dermed søk uten resultat.
        If CboInvSkadet.SelectedIndex = 0 Or CboInvSkadet.SelectedIndex = 1 Then
            InvSpSkadet = "skadet LIKE '%" & CboInvSkadet.SelectedIndex & "%'"
        Else
            InvSpSkadet = "skadet LIKE '%%'"
        End If

        If CboInvSavnet.SelectedIndex = 0 Or CboInvSavnet.SelectedIndex = 1 Then
            InvSpSavnet = "savnet LIKE '%" & CboInvSavnet.SelectedIndex & "%'"
        Else
            InvSpSavnet = "savnet LIKE '%%'"
        End If

        Dim InvSokSporring As String
        InvSokSporring = InvSpInit & " WHERE " & InvSpSykkelNavn & " AND " & InvSpsykkelModell & " AND " &
                InvSpTypeid & " AND " & InvSpSykkelRamme & " AND " & InvSpGirsystem & " AND " &
                InvSpHjulstorrelse & " AND " & InvSpSykkelPris & " AND " & InvSpAvdeling & " AND " &
                InvSpForhandlerID & " AND " & InvSpSykkelStatus & " AND " & InvSpSkadet & " AND " &
                InvSpSavnet & InvSpSykkelUtstyr & ";"

        Dim InvResultatArray(12) As String
        Dim InvResultatObjekt As ListViewItem

        'Sender spørring for søk basert på innlagte data i skjema. Skriver ut returnerte rader til listview 
        Try
            DBConnect()
            InvSqlSok = New MySqlCommand(InvSokSporring, Tilkobling)
            invSqlLeser = InvSqlSok.ExecuteReader()

            'For hver kolonne les inn verdi 0-12 og legg i listview
            While invSqlLeser.Read()
                For i = 0 To invSqlLeser.FieldCount - 1
                    If i = 11 Or i = 12 Then
                        If invSqlLeser(i) = False Then
                            InvResultatArray(i) = "Nei"
                        Else
                            InvResultatArray(i) = "Ja"
                        End If
                    Else
                        InvResultatArray(i) = invSqlLeser(i)
                    End If
                Next
                InvResultatObjekt = New ListViewItem(InvResultatArray)
                LvInvSok.Items.Add(InvResultatObjekt)
            End While
            invSqlLeser.Close()
            DBDisconnect()
        Catch ex As MySqlException
            MsgBox("Feil ved søk i sykkeldatabase:" & vbNewLine & ex.Message)
        End Try

    End Sub

    Private Sub InvSokUtstyr()

        Dim InvSpInit As String = "SELECT utstyr.utstyr_id, utstyr.utstyr_navn, utstyr_kategori.utstyr_kat, " _
               & "utstyr.varenummer, utstyr.utstyr_pris, forhandler.forhandler_navn " _
               & "FROM utstyr LEFT JOIN forhandler ON utstyr.forhandler_id=forhandler.forhandler_id " _
               & "LEFT JOIN utstyr_kategori ON utstyr.utstyr_kat_id=utstyr_kategori.utstyr_kat_id WHERE "

        Dim InvSpUtstyrNavn As String = "utstyr_navn LIKE '%" & SQLWhiteWash(TxtInvProduktnavn.Text.Trim) & "%'"
        Dim InvSpVarenummer As String = "varenummer LIKE '%" & SQLWhiteWash(TxtInvVareNummer.Text.Trim) & "%'"
        Dim InvSpUtstyrPris As String = "utstyr_pris LIKE '%" & SQLWhiteWash(TxtInvInnkjopspris.Text.Trim) & "%'"
        Dim InvSpForhandlerNavn As String = "forhandler_navn LIKE '%" & CboInvForhandler.SelectedItem & "%'"
        Dim InvSpUtstyrkategori As String = "utstyr_kat LIKE '%" & CboInvSubkategori.SelectedItem & "%'"

        Dim InvSokSporring As String
        InvSokSporring = InvSpInit & InvSpUtstyrNavn & " AND " & InvSpUtstyrkategori & " AND " _
            & InvSpVarenummer & " AND " & InvSpUtstyrPris & " AND " & InvSpForhandlerNavn & ";"

        Dim InvResultatArray(5) As String
        Dim InvResultatObjekt As ListViewItem
        Dim InvSqlSok As MySqlCommand
        Dim invSqlLeser As MySqlDataReader

        'Sender spørring for søk basert på innlagte data i skjema. Skriver ut returnerte rader til listview 
        Try
            DBConnect()
            InvSqlSok = New MySqlCommand(InvSokSporring, Tilkobling)
            invSqlLeser = InvSqlSok.ExecuteReader()

            'For hver kolonne les inn verdi 1 til n og legg i listview
            While invSqlLeser.Read()
                For i = 0 To invSqlLeser.FieldCount - 1
                    If TypeOf invSqlLeser(i) Is DBNull Then
                    Else
                        InvResultatArray(i) = invSqlLeser(i)
                    End If
                Next
                InvResultatObjekt = New ListViewItem(InvResultatArray)
                LvInvSok.Items.Add(InvResultatObjekt)
            End While
            invSqlLeser.Close()
            DBDisconnect()

        Catch ex As MySqlException
            MsgBox("Feil ved søk i utstyrdatabase:" & vbNewLine & ex.Message)
        End Try

    End Sub

    Private Sub InvEndreSykkel()

        Dim InvKategori, InvSubkategoriNavn, InvAvdelingNavn, InvProduktnavn, InvVarenummer, InvInnkjopspris,
            InvRamme, InvHjulstorrlese, InvGirsystem, InvForhandlerNavn, InvStatus, InvSkadet, InvSavnet,
            InvForhandlerID, InvAvdelingID, InvSubKategoriID, InvEndreVerdier As String

        InvKategori = CboInvKategori.SelectedItem
        InvSubkategoriNavn = CboInvSubkategori.SelectedItem
        InvAvdelingNavn = CboInvAvdeling.SelectedItem
        InvProduktnavn = SQLWhiteWash(TxtInvProduktnavn.Text.Trim)
        InvVarenummer = SQLWhiteWash(TxtInvVareNummer.Text.Trim)
        InvInnkjopspris = SQLWhiteWash(TxtInvInnkjopspris.Text.Trim)
        InvRamme = SQLWhiteWash(TxtInvRamme.Text.Trim)
        InvHjulstorrlese = SQLWhiteWash(TxtInvHjulstorrelse.Text.Trim)
        InvGirsystem = SQLWhiteWash(TxtInvGirsystem.Text.Trim)
        InvForhandlerNavn = CboInvForhandler.SelectedItem
        InvStatus = CboInvStatus.SelectedItem
        InvSkadet = CboInvSkadet.SelectedIndex
        InvSavnet = CboInvSavnet.SelectedIndex

        If CboInvAvdeling.SelectedIndex = -1 Or CboInvForhandler.SelectedIndex = -1 Or
              CboInvSubkategori.SelectedIndex = -1 Or TxtInvProduktnavn.Text.Trim = "" Or
              TxtInvVareNummer.Text.Trim = "" Or TxtInvInnkjopspris.Text.Trim = "" Or
              TxtInvRamme.Text.Trim = "" Or TxtInvHjulstorrelse.Text.Trim = "" Or
              TxtInvGirsystem.Text.Trim = "" Or CboInvStatus.SelectedIndex = -1 Or
              CboInvSkadet.SelectedIndex = -1 Or CboInvSavnet.SelectedIndex = -1 Then

            MsgBox("Vennligst fyll inn alle felt")

        Else

            Dim InvBekreftSykkelReg As DialogResult
            InvBekreftSykkelReg = MsgBox("Bekreft endring av sykkel", MsgBoxStyle.OkCancel)
            If InvBekreftSykkelReg = DialogResult.OK Then

                InvForhandlerID = SQLHentIDNavn("forhandler", "forhandler_navn",
                    "forhandler_id", InvForhandlerNavn)
                InvSubKategoriID = SQLHentIDNavn("sykkel_typer", "kategori", "type_id", InvSubkategoriNavn)
                InvAvdelingID = SQLHentIDNavn("avdeling", "avd_navn", "avdeling_id", InvAvdelingNavn)
                InvEndreVerdier = "sykkel_navn='" & InvProduktnavn & "', sykkel_modell='" & InvVarenummer _
                    & "', type_id='" & InvSubKategoriID & "', sykkel_ramme='" & InvRamme & "', girsystem='" _
                    & InvGirsystem & "', hjul_str='" & InvHjulstorrlese & "', sykkel_pris='" &
                    InvInnkjopspris & "', avdeling_id='" & InvAvdelingID & "', forhandler_id='" _
                    & InvForhandlerID & "', sykkel_status='" & InvStatus & "', skadet='" & InvSkadet _
                    & "', savnet='" & InvSavnet & "'"

                SQLUpdate("sykler", InvEndreVerdier, "sykkel_id='" & InvAktivtProduktID & "';")
                SQLDelete("sykkel_utstyr", "sykkel_id='" & InvAktivtProduktID & "';")

                If ChkInvBarneHenger.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvAktivtProduktID & ", 4);")
                End If
                If ChkInvBarnesete.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvAktivtProduktID & ", 3);")
                End If
                If ChkInvLastehenger.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvAktivtProduktID & ", 5)")
                End If
                If ChkInvSykkelveske.Checked = True Then
                    SQLInsert("sykkel_utstyr", "(sykkel_id, utstyr_kat_id)", "(" & InvAktivtProduktID & ", 2)")
                End If


            End If
        End If

    End Sub

    Private Sub InvEndreUtstyr()



        Dim InvKategori, InvSubkategoriNavn, InvAvdelingNavn, InvProduktnavn, InvVarenummer, InvInnkjopspris,
            InvForhandlerNavn, InvForhandlerID, InvSubKategoriID, InvEndreVerdier As String

        InvKategori = CboInvKategori.SelectedItem
        InvSubkategoriNavn = CboInvSubkategori.SelectedItem
        InvAvdelingNavn = CboInvAvdeling.SelectedItem
        InvProduktnavn = SQLWhiteWash(TxtInvProduktnavn.Text.Trim)
        InvVarenummer = SQLWhiteWash(TxtInvVareNummer.Text.Trim)
        InvInnkjopspris = SQLWhiteWash(TxtInvInnkjopspris.Text.Trim)
        InvForhandlerNavn = CboInvForhandler.SelectedItem

        If CboInvForhandler.SelectedIndex = -1 Or CboInvSubkategori.SelectedIndex = -1 _
             Or TxtInvProduktnavn.Text.Trim = "" Or TxtInvVareNummer.Text.Trim = "" _
             Or TxtInvInnkjopspris.Text.Trim = "" Then

            MsgBox("Vennligst fyll inn alle felt")

        Else

            Dim InvBekreftSykkelReg As DialogResult
            InvBekreftSykkelReg = MsgBox("Bekreft endring av utstyr", MsgBoxStyle.OkCancel)
            If InvBekreftSykkelReg = DialogResult.OK Then

                InvForhandlerID = SQLHentIDNavn("forhandler", "forhandler_navn", "forhandler_id", InvForhandlerNavn)
                InvSubKategoriID = SQLHentIDNavn("utstyr_kategori", "utstyr_kat",
                    "utstyr_kat_id", InvSubkategoriNavn)
                InvEndreVerdier = "utstyr_navn='" & InvProduktnavn & "', varenummer='" _
                & InvVarenummer & "', utstyr_pris='" & InvInnkjopspris & "', forhandler_id='" _
                & InvForhandlerID & "', utstyr_kat_id='" & InvSubKategoriID & "'"

                SQLUpdate("utstyr", InvEndreVerdier, "utstyr_id='" & InvAktivtProduktID & "'")

            End If
        End If

    End Sub


    'SQLspørring med registrering av nytt produkt basert på innlagt data i skjema.
    Private Sub BtnInvRegistrer_Click(sender As Object, e As EventArgs) Handles BtnInvRegistrer.Click

        If CboInvKategori.SelectedItem = "Sykkel" Then
            InvRegistrerSykkel()
        ElseIf CboInvKategori.SelectedItem = "Utstyr" Then
            InvRegistrerUtstyr()
        Else
            MsgBox("Velg kategori")
        End If

    End Sub

    'SQLspørring med søk på valgte verdier i søkefelt
    Private Sub BtnInvSok_Click(sender As Object, e As EventArgs) Handles BtnInvSok.Click

        LvInvSok.Items.Clear()

        If CboInvKategori.SelectedItem = "Sykkel" Then
            InvSokSykkel()
        ElseIf CboInvKategori.SelectedItem = "Utstyr" Then
            InvSokUtstyr()
        Else
            MsgBox("Velg en kategori")
        End If

    End Sub

    'SQLspørring med endering av data for valgt produkt med ID fra valgt produkt i søkefelt
    Private Sub BtnInvEndre_Click(sender As Object, e As EventArgs) Handles BtnInvEndre.Click

        BtnInvRegistrer.Enabled = True
        BtnInvSok.Enabled = True

        If CboInvKategori.SelectedItem = "Sykkel" Then
            InvEndreSykkel()
        ElseIf CboInvKategori.SelectedItem = "Utstyr" Then
            InvEndreUtstyr()
        Else
            MsgBox("Velg kategori")
        End If

        InvAktivtProduktID = ""
        LblInvAktivProdukt.Text = "Ingen aktive produkt."
        BtnInvEndre.Enabled = False

    End Sub

    'Henter data for objekt med inntastet ID fra database og legger i skjema
    Private Sub BtnInvHent_Click(sender As Object, e As EventArgs) Handles BtnInvHent.Click

        BtnInvRegistrer.Enabled = False
        BtnInvEndre.Enabled = True

        Dim InvForhandlerID, InvAvdelingID, InvSubkategoriID, InvForhandlerNavn, InvAvdelingNavn,
            InvSubkategoriNavn As String

        Dim InvSkadetBol, InvSavnetBol As Boolean

        Dim InvHentDaT As New DataTable
        Dim InvHentUtstyrKatDaT As New DataTable

        If TxtInvHentID.Text = "" Then

            MsgBox("Skriv inn et ID nummer")

            'Spørring for å hente sykkel fra database basert på inntastet ID og legge det inn i skjema
        ElseIf CboInvKategori.SelectedItem = "Sykkel" Then

            InvAktivtProduktID = SQLWhiteWash(TxtInvHentID.Text.Trim)
            LblInvAktivProdukt.Text = "Aktivt produkt ID: " & InvAktivtProduktID
            InvTomFelt()

            InvHentDaT = SQLSelect("sykler", "*", "sykkel_id='" & InvAktivtProduktID & "'")
            InvHentUtstyrKatDaT = SQLSelect("sykkel_utstyr", "utstyr_kat_id", "sykkel_id=" & InvAktivtProduktID)

            Try
                For Each r In InvHentDaT.Rows
                    InvSubkategoriID = r("type_id").ToString
                    InvAvdelingID = r("avdeling_id").ToString
                    TxtInvProduktnavn.Text = r("sykkel_navn")
                    TxtInvVareNummer.Text = r("sykkel_modell")
                    TxtInvInnkjopspris.Text = r("sykkel_pris")
                    TxtInvRamme.Text = r("sykkel_ramme")
                    TxtInvHjulstorrelse.Text = r("hjul_str")
                    TxtInvGirsystem.Text = r("girsystem")
                    InvForhandlerID = r("forhandler_id").ToString
                    CboInvStatus.SelectedItem = r("sykkel_status")
                    InvSkadetBol = r("skadet")
                    InvSavnetBol = r("savnet")
                Next
                For Each r In InvHentUtstyrKatDaT.Rows
                    If r("utstyr_kat_id") = 2 Then
                        ChkInvSykkelveske.Checked = True
                    ElseIf r("utstyr_kat_id") = 3 Then
                        ChkInvBarnesete.Checked = True
                    ElseIf r("utstyr_kat_id") = 4 Then
                        ChkInvBarneHenger.Checked = True
                    ElseIf r("utstyr_kat_id") = 5 Then
                        ChkInvLastehenger.Checked = True
                    End If
                Next
            Catch ex As Exception
                MsgBox("Feil ved innlegging av data i skjema:" & ex.Message)
            End Try

            InvForhandlerNavn = SQLHentIDNavn("forhandler",
                        "forhandler_id", "forhandler_navn", InvForhandlerID)
            InvSubkategoriNavn = SQLHentIDNavn("sykkel_typer", "type_id", "kategori", InvSubkategoriID)
            InvAvdelingNavn = SQLHentIDNavn("avdeling", "avdeling_id", "avd_navn", InvAvdelingID)

            'Fyller inn resterende data i skjema 
            CboInvKategori.SelectedIndex = 0
            CboInvAvdeling.SelectedItem = InvAvdelingNavn
            CboInvForhandler.SelectedItem = InvForhandlerNavn
            CboInvSubkategori.SelectedItem = InvSubkategoriNavn
            If InvSkadetBol = False Then
                CboInvSkadet.SelectedIndex = 0
            Else
                CboInvSkadet.SelectedIndex = 1
            End If
            If InvSavnetBol = False Then
                CboInvSavnet.SelectedIndex = 0
            Else
                CboInvSavnet.SelectedIndex = 1
            End If

            'Spørring å hente utstyr fra database basert på inntastet ID og legge det inn i skjema
        ElseIf CboInvKategori.SelectedItem = "Utstyr" Then

            InvAktivtProduktID = SQLWhiteWash(TxtInvHentID.Text.Trim)
            LblInvAktivProdukt.Text = "Aktivt produkt ID: " & InvAktivtProduktID
            InvTomFelt()

            InvHentDaT = SQLSelect("utstyr", "*", "utstyr_id='" & InvAktivtProduktID & "';")

            Try
                For Each r In InvHentDaT.Rows
                    InvSubkategoriID = r("utstyr_kat_id").ToString
                    TxtInvProduktnavn.Text = r("utstyr_navn")
                    TxtInvVareNummer.Text = r("varenummer")
                    TxtInvInnkjopspris.Text = r("utstyr_pris")
                    InvForhandlerID = r("forhandler_id").ToString
                Next
            Catch ex As Exception
                MsgBox("Feil ved innlegging av data i skjema: " & ex.Message)
            End Try

            InvForhandlerNavn = SQLHentIDNavn("forhandler", "forhandler_id", "forhandler_navn", InvForhandlerID)
            InvSubkategoriNavn = SQLHentIDNavn("utstyr_kategori", "utstyr_kat_id", "utstyr_kat", InvSubkategoriID)

            CboInvKategori.SelectedIndex = 1
            CboInvForhandler.SelectedItem = InvForhandlerNavn
            CboInvSubkategori.SelectedItem = InvSubkategoriNavn

        Else
            MsgBox("Velg kategori")
        End If
    End Sub

    'Avbryter endring av produkt. Ved innhenting av produkt, knappen "Hent", settes "Registrer" _
    'knappen til inaktiv For å unngå feilaktig dobbelregistrering av produkt.
    'Ved å avbryte nullstilles skjema og knappen for registrering aktiveres.
    Private Sub BtnInvAvbrytEndre_Click(sender As Object, e As EventArgs) Handles BtnInvAvbrytEndre.Click
        BtnInvRegistrer.Enabled = True
        BtnInvSok.Enabled = True
        BtnInvEndre.Enabled = False
        InvAktivtProduktID = ""
        InvTomFelt()
        LblInvAktivProdukt.Text = "Ingen aktive produkt"
    End Sub

    'Tømmer alle felt i skjema
    Private Sub BtnInvTom_Click(sender As Object, e As EventArgs) Handles BtnInvTom.Click
        InvTomFelt()
    End Sub


    ' - ----------------------- TEST TEST TEST - --------------------
    'Private Sub BtnInvTest_Click(sender As Object, e As EventArgs) Handles BtnInvTest.Click
    '    Dim testdato As Date
    '    Dim testsp As String = "select kunde_fdato from kunder where kunde_etternavn='faug'"
    '    Dim sql As MySqlCommand
    '    Dim leser As MySqlDataReader
    '    DBConnect()
    '    sql = New MySqlCommand(testsp, tilkobling)
    '    leser = sql.ExecuteReader()
    '    While leser.Read()
    '        testdato = leser(0)
    '    End While
    '    leser.Close()
    '    LblInvAktivProdukt.Text = CStr(testdato)
    'End Sub


#End Region


#Region "LogistikkTab"
    'Her plasseres kode som er relevant til Logistikk Tab.
    'Variabler som brukes her skal begynne med Logi. Dette er for å unngå klasj.
    'Husk kode kommentarer.
#End Region


#Region "StatistikkTab"
    'Her plasseres kode som er relevant til Statistikk Tab.
    'Variabler som brukes her skal begynne med Stat. Dette er for å unngå klasj.
    'Husk kode kommentarer.

    Dim StaValgtAvdTilgj As String
    Dim StaValgtTypTilgj As String

    Private Sub BtnStaSok_Click(sender As Object, e As EventArgs) Handles BtnStaSok.Click
        Dim StaValgtAvdeling = CmbStaAvdeling.SelectedItem
        Dim StaSykkelType = CmbStaType.SelectedItem

        If StaValgtAvdeling = "" Then
            MsgBox("Vennligst velg avdeling")
            Exit Sub
        End If
        If StaSykkelType = "" Then
            MsgBox("Vennligst velg type")
            Exit Sub
        End If


        Dim outputAntallSyklerLedig, outputAntallSyklerUtleid, outputAntallSyklerVerksted As Integer
        Dim StaSykkeltypeValue = New Integer() {9999, 10000, 10001, 10002, 10003, 10004, 10005, 10006}
        Dim StaAvdelingValue = New Integer() {10000, 10001, 10002, 10003, 10004}
        StaValgtAvdTilgj = StaAvdelingValue(CmbStaAvdeling.SelectedIndex)
        StaValgtTypTilgj = StaSykkeltypeValue(CmbStaType.SelectedIndex)


        Try
            DBConnect()

            Dim sporring As New MySqlCommand("SELECT COUNT(sykkel_modell) FROM sykler WHERE avdeling_id = '" _
                                             & StaValgtAvdTilgj & "' AND type_id = '" & StaValgtTypTilgj & "' AND sykkel_status = 'Ledig'", Tilkobling)
            Dim sporring2 As New MySqlCommand("SELECT COUNT(sykkel_modell) FROM sykler WHERE avdeling_id = '" _
                                              & StaValgtAvdTilgj & "' AND type_id = '" & StaValgtTypTilgj & "' AND sykkel_status = 'Utleid'", Tilkobling)
            Dim sporring3 As New MySqlCommand("SELECT COUNT(sykkel_modell) FROM sykler WHERE avdeling_id = '" _
                                              & StaValgtAvdTilgj & "' AND type_id = '" & StaValgtTypTilgj & "' AND sykkel_status = 'Verksted'", Tilkobling)
            outputAntallSyklerLedig = sporring.ExecuteScalar()
            outputAntallSyklerUtleid = sporring2.ExecuteScalar()
            outputAntallSyklerVerksted = sporring3.ExecuteScalar()

            DBDisconnect()
        Catch feilmelding As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)
        End Try

        LvStaTilgjengelig.Items.Clear()
        LvStaTilgjengelig.Items.Add(New ListViewItem({StaValgtAvdeling, StaSykkelType, outputAntallSyklerLedig, outputAntallSyklerUtleid, outputAntallSyklerVerksted}))
    End Sub

    Private Sub StaMestPopulaer()
        Dim StaAntSyklTyp As Integer

        Dim StaSQLFinnAnt = SQLSelect("sykkel_typer", "COUNT(*) AS c", "1")
        For Each row As DataRow In StaSQLFinnAnt.Rows
            StaAntSyklTyp = row.Field(Of Int64)(0)
        Next

        Dim StaAntSykler(StaAntSyklTyp) As Integer
        Dim StaTypeSykkel(StaAntSyklTyp) As String

        Dim StaInputCommand As String = "select sykkel_typer.kategori, count(*) AS c from sykkel_typer join " _
            & "sykler as s on sykkel_typer.type_id=s.type_id join utleie on s.sykkel_id=utleie.sykkel_id group by sykkel_typer.kategori"

        Dim part1 As String = "select sykkel_typer.kategori, count(*) As c"
        Dim part2 As String = "from sykkel_typer join sykler as s on sykkel_typer.type_id=s.type_id join utleie on s.sykkel_id=utleie.sykkel_id group by sykkel_typer.kategori"
        Dim part3 As String = "1"

        ' Dim tretest123 = SQLSelect(part1, part2, part3)

        Try
            DBConnect()
            Dim sporring As New MySqlCommand(StaInputCommand, Tilkobling)
            Dim StatempVarSporring = sporring.ExecuteReader()
            Dim i = 0
            While StatempVarSporring.Read

                StaAntSykler(i) = StatempVarSporring("c")
                StaTypeSykkel(i) = StatempVarSporring("kategori")

                i = i + 1
            End While

            DBDisconnect()


            For j As Integer = 0 To StaAntSykler.Length - 1
                LvStaMestUtleid.Items.Add(New ListViewItem({StaTypeSykkel(j), StaAntSykler(j)}))
            Next

        Catch feilmelding As MySqlException
            MsgBox("Feil ved tilkobling til databasen: " & feilmelding.Message)
        End Try
    End Sub

#End Region


#Region "AdminTab"
    'Her plasseres kode som er relevant til Admin Tab.
    'Variabler som brukes her skal begynne med Admin. Dette er for å unngå klasj.
    'Husk kode kommentarer.


    'Prosedyre som brukes til lagring av ny bruker i MySQLDatabasen.
    'Dette er kjørt når man trykker på ny bruker knappen.
    'Her er det slik at Passord og Avdelinger befinner seg i egne tabeller.
    'Vi henter først AvdelingsIDen til Avdelingen som har blitt valgt vha MySQL som leser Avd_navn og skriver tilsvarende avdelings id til variabel.
    'Da kjører vi insert sql i 3 stadier - insert bruker, insert passord, update bruker. Dette er slik at vi får lenket opp passord fk og bruker fk med hverandre
    'ANB1 : Insert ny bruker ANB2 : Insert ny passord ANB3 : Update bruker for å knytte passord fk ANB4 : laster inn tilsvarende avdelings id.

    Private Sub AdminNyBruker()
        Try
            DBConnect()
            Dim AdminAvdelingNavn As String = ""
            Dim AdminAdminStatus As Integer = 0
            If ChkAdminNBAdmin.Checked = True Then
                AdminAdminStatus = 1
            End If

            TxtAdminNBPassord.Text = SQLWhiteWash(TxtAdminNBPassord.Text)
            TxtAdminNBFornavn.Text = SQLWhiteWash(TxtAdminNBFornavn.Text)
            TxtAdminNBEtternavn.Text = SQLWhiteWash(TxtAdminNBEtternavn.Text)
            TxtAdminNBTime.Text = SQLWhiteWash(TxtAdminNBTime.Text)
            TxtAdminNBEpost.Text = SQLWhiteWash(TxtAdminNBEpost.Text)
            TxtAdminNBTelefon.Text = SQLWhiteWash(TxtAdminNBTelefon.Text)

            Dim AdminNyBruker2 As New MySqlCommand("INSERT INTO passord (passord_id, pwd, bruker_id) VALUES (" & LblBrukerIDNBVis.Text & ", '" & TxtAdminNBPassord.Text & "'," & LblBrukerIDNBVis.Text & ");", Tilkobling)
            Dim AdminNyBruker3 As New MySqlCommand("UPDATE brukere SET passord_id = " & LblBrukerIDNBVis.Text & " WHERE bruker_id = " & LblBrukerIDNBVis.Text & ";", Tilkobling)
            Dim AdminNyBruker4 As New MySqlCommand("SELECT avdeling_id FROM avdeling WHERE avd_navn ='" & CboAdminNBAvdeling.Text & "';", Tilkobling)

            Dim AdminNyBrukerIDAdapter As New MySqlDataAdapter
            Dim AdminNyBrukerIDTable As New DataTable

            AdminNyBrukerIDAdapter.SelectCommand = AdminNyBruker4
            AdminNyBrukerIDAdapter.Fill(AdminNyBrukerIDTable)

            Dim AdminNyBrukerRow As DataRow

            For Each AdminNyBrukerRow In AdminNyBrukerIDTable.Rows
                AdminAvdelingNavn = AdminNyBrukerRow("avdeling_id")
            Next

            Dim AdminNyBruker1 As New MySqlCommand("INSERT INTO brukere (bruker_id, avdeling_id, stilling, fornavn, etternavn, timelonn, telefon, epost, stilling_prosent, admin) VALUES (" & LblBrukerIDNBVis.Text & "," & AdminAvdelingNavn & ",'" & CboAdminNBStilling.Text & "','" & TxtAdminNBFornavn.Text & "','" & TxtAdminNBEtternavn.Text & "'," _
                                                   & TxtAdminNBTime.Text & ",'" & TxtAdminNBTelefon.Text & "', '" & TxtAdminNBEpost.Text & "'," & CboAdminNBSP.Text & "," & AdminAdminStatus & ");", Tilkobling)

            AdminNyBruker1.ExecuteNonQuery()
            AdminNyBruker2.ExecuteNonQuery()
            AdminNyBruker3.ExecuteNonQuery()

            DBDisconnect()
            TxtAdminNBPassord.Text = ""
            TxtAdminNBFornavn.Text = ""
            TxtAdminNBEtternavn.Text = ""
            TxtAdminNBTime.Text = ""
            TxtAdminNBEpost.Text = ""
            TxtAdminNBTelefon.Text = ""
            ChkAdminNBAdmin.Checked = False

            MsgBox("Ny Bruker Opprettet!")

        Catch AdminSqlError6 As MySqlException
            MsgBox("Man får ikke koble til databasen:  " & AdminSqlError6.Message)
        End Try
    End Sub

    'Prosedyre som brukes til endring av ekisterende bruker i SQLdatabasen.
    'Veldig lik den som brukes til ny bruker, men noen forandringer:
    'Som før så hentes tilsvarende avdeling id slik at FK til brukeren kan oppdateres riktig.
    'Da er det SQL update for å bytte bruker detaljene.
    'Hvis passordfeltet er tom, så byttes ikke passordet. AEB1 : Update Bruker AEB3 : Update Passord AEB4 : Hent avdelingid

    Private Sub AdminEndreBruker()

        TxtAdminEBPassord.Text = SQLWhiteWash(TxtAdminEBPassord.Text)
        TxtAdminEBFornavn.Text = SQLWhiteWash(TxtAdminEBFornavn.Text)
        TxtAdminEBEtternavn.Text = SQLWhiteWash(TxtAdminEBEtternavn.Text)
        TxtAdminEBTime.Text = SQLWhiteWash(TxtAdminEBTime.Text)
        TxtAdminEBEpost.Text = SQLWhiteWash(TxtAdminEBEpost.Text)
        TxtAdminEBTelefon.Text = SQLWhiteWash(TxtAdminEBTelefon.Text)

        Try
            DBConnect()
            Dim AdminAvdelingNavn As String = ""
            Dim AdminAdminStatus As Integer = 0
            If ChkAdminNBAdmin.Checked = True Then
                AdminAdminStatus = 1
            End If

            Dim AdminEndreBruker3 As New MySqlCommand("UPDATE passord SET pwd ='" & TxtAdminEBPassord.Text & "' WHERE bruker_id = " & TxtAdminEBBID.Text & ";", Tilkobling)
            Dim AdminEndreBruker4 As New MySqlCommand("SELECT avdeling_id FROM avdeling WHERE avd_navn ='" & CboAdminEBAvdeling.Text & "';", Tilkobling)

            Dim AdminEndreBrukerIDAdapter As New MySqlDataAdapter
            Dim AdminEndreBrukerIDTable As New DataTable

            AdminEndreBrukerIDAdapter.SelectCommand = AdminEndreBruker4
            AdminEndreBrukerIDAdapter.Fill(AdminEndreBrukerIDTable)

            Dim AdminNyBrukerRow As DataRow

            For Each AdminNyBrukerRow In AdminEndreBrukerIDTable.Rows
                AdminAvdelingNavn = AdminNyBrukerRow("avdeling_id")
            Next

            Dim AdminEndreBruker1 As New MySqlCommand("UPDATE brukere SET avdeling_id=" & AdminAvdelingNavn & ", stilling='" & CboAdminEBStilling.Text & "', fornavn='" & TxtAdminEBFornavn.Text & "', etternavn='" & TxtAdminEBEtternavn.Text & "', timelonn=" & TxtAdminEBTime.Text & ", telefon='" & TxtAdminEBTelefon.Text & "', epost='" & TxtAdminEBEpost.Text & "', stilling_prosent=" & CboAdminEBSP.Text & ", admin=" & AdminAdminStatus & " WHERE bruker_id=" & TxtAdminEBBID.Text & ";", Tilkobling)

            AdminEndreBruker1.ExecuteNonQuery()

            If TxtAdminEBPassord.Text = "" Then
            Else
                AdminEndreBruker3.ExecuteNonQuery()
            End If


            DBDisconnect()
            TxtAdminEBPassord.Text = ""
            TxtAdminEBFornavn.Text = ""
            TxtAdminEBEtternavn.Text = ""
            TxtAdminEBTime.Text = ""
            TxtAdminEBEpost.Text = ""
            TxtAdminEBTelefon.Text = ""
            ChkAdminEBAdmin.Checked = False

            MsgBox("Bruker Endret!")

        Catch AdminSqlError8 As MySqlException
            MsgBox("Man får ikke koble til databasen:  " & AdminSqlError8.Message)
        End Try
    End Sub


    'Prosedyre for inn lasting av bruker detaljer. Når man taster inn bruker ID så lastes inn info slik at man vet hva det var før.
    'Her har vi mysql kode som laster in tabellen for tilsvarende brukerID. Den kjøres ved knappetrykk.
    'Da populeres det som befinner seg i formen med verdiene som er i tabellen som vi har laget.
    Private Sub AdminLastInnEndreBruker()

        TxtAdminEBBID.Text = SQLWhiteWash(TxtAdminEBBID.Text)
        Dim EBSPString As String = ""
        Dim AdminEBBIDTable As New DataTable
        Dim AdminEBAvdelingID As String = ""
        AdminEBBIDTable = SQLSelect("*", "brukere", "bruker_id='" & TxtAdminEBBID.Text & "'")
        Dim AdminEBBIDRow As DataRow
        For Each AdminEBBIDRow In AdminEBBIDTable.Rows

            TxtAdminEBFornavn.Text = AdminEBBIDRow("fornavn")
            TxtAdminEBEtternavn.Text = AdminEBBIDRow("etternavn")
            TxtAdminEBTime.Text = AdminEBBIDRow("timelonn")
            TxtAdminEBEpost.Text = AdminEBBIDRow("epost")
            TxtAdminEBTelefon.Text = AdminEBBIDRow("telefon")
            CboAdminEBStilling.SelectedItem = AdminEBBIDRow("stilling")
            EBSPString = AdminEBBIDRow("stilling_prosent")
            CboAdminEBSP.SelectedItem = EBSPString
            If AdminEBBIDRow("admin") = "1" Then
                ChkAdminEBAdmin.Checked = True
            End If
            AdminEBAvdelingID = AdminEBBIDRow("avdeling_id")

        Next

        Dim AdminEBAvdelingTable As New DataTable
        AdminEBAvdelingTable = SQLSelect("*", "avdeling", "avdeling_id='" & AdminEBAvdelingID & "'")
        Dim AdminEBAvdelingRow As DataRow

        For Each AdminEBAvdelingRow In AdminEBAvdelingTable.Rows
            CboAdminEBAvdeling.SelectedItem = AdminEBAvdelingRow("avd_navn")
        Next

        If TxtAdminEBFornavn.Text = "" Then
            MsgBox("Brukeren med ID: " & TxtAdminEBBID.Text & " eksisterer ikke.")
        End If

    End Sub

    'Dette er en prosedyre som kjører på hvert lasting av Adminside og ved lagring av ny bruker.
    'Den kalkulerer neste bruker ID automatisk. Dette er gjort fordi det fjerner muligheten for menneskefeil iht bruker id.
    Private Sub AdminBrukerIDCalc()
        Try
            DBConnect()
            Dim AdminBrukerIDKommando As New MySqlCommand("Select COUNT(bruker_id) FROM brukere", Tilkobling)
            Dim AdminBrukerIDAdapter As New MySqlDataAdapter
            Dim AdminBrukerIDTable As New DataTable
            AdminBrukerIDAdapter.SelectCommand = AdminBrukerIDKommando
            AdminBrukerIDAdapter.Fill(AdminBrukerIDTable)
            DBDisconnect()
            Dim AdminBrukerIDVerdi As Integer
            Dim AdminBrukerRow As DataRow
            For Each AdminBrukerRow In AdminBrukerIDTable.Rows
                AdminBrukerIDVerdi = AdminBrukerRow("COUNT(bruker_id)")
                LblBrukerIDNBVis.Text = (AdminBrukerIDVerdi + 1)
            Next

        Catch AdminSqlError3 As MySqlException
            MsgBox("Man får ikke koble til databasen: " & AdminSqlError3.Message)

        End Try
    End Sub

    'Dette er en prosedyre som populerer ComboBoksene til avdelingene. Dette gir muligheten å velge riktig avdelingsnavn.
    'Comboboksene er tatt i bruk for å ha best mulig information hygiene på plass.
    Private Sub AdminAvdelingPopulate()

        Dim AdminAvdelingFString As String = "avdeling"
        Dim AdminAvdelingTable As New DataTable
        Dim AdminAvdelingRow As DataRow
        Dim AdminAvdelingString As String

        AdminAvdelingTable = SQLSelect(AdminAvdelingFString, "avd_navn", "1")
        CboAdminNBAvdeling.Items.Clear()
        CboAdminEBAvdeling.Items.Clear()

        For Each AdminAvdelingRow In AdminAvdelingTable.Rows
            AdminAvdelingString = AdminAvdelingRow("avd_navn")
            CboAdminNBAvdeling.Items.Add(AdminAvdelingString)
            CboAdminEBAvdeling.Items.Add(AdminAvdelingString)

        Next
    End Sub

    'Dette er søkeboksprosedyren.
    'Den søker vha et select command som gjør at søkefeltet er kolonnen som blir søket på. Vi bruker også LIKE sql commando som gjør at vi får partial matching.
    Private Sub AdminBSSokB_Click(sender As Object, e As EventArgs) Handles AdminBSSokB.Click
        Dim AdminSoekefelt, AdminSoekekategori As String
        AdminSoekefelt = SQLWhiteWash(TxtAdminBSFelt.Text)
        AdminSoekekategori = CboAdminBSEtter.Text



        Try
            DBConnect()
            Dim AdminBrukerSearch As New MySqlCommand("SELECT * FROM brukere WHERE " & AdminSoekekategori & " LIKE '%" & AdminSoekefelt & "%'", Tilkobling)
            Dim AdminSearchAdapter As New MySqlDataAdapter
            Dim AdminSearchTable As New DataTable
            AdminSearchAdapter.SelectCommand = AdminBrukerSearch
            AdminSearchAdapter.Fill(AdminSearchTable)
            DBDisconnect()
            Dim AdminRow As DataRow
            Dim AdminBSbruker_id, AdminBSfornavn, AdminBSetternavn, AdminBSsoekefelt As String
            LvAdminBS.Items.Clear()
            For Each AdminRow In AdminSearchTable.Rows
                AdminBSbruker_id = AdminRow("bruker_id")
                AdminBSfornavn = AdminRow("fornavn")
                AdminBSetternavn = AdminRow("etternavn")
                AdminBSsoekefelt = AdminRow(AdminSoekekategori)
                LvAdminBS.Items.Add(New ListViewItem({AdminBSbruker_id, AdminBSfornavn, AdminBSetternavn, AdminBSsoekefelt}))
            Next
        Catch AdminSqlError2 As MySqlException
            MsgBox("Man får ikke koble til databasen: " & AdminSqlError2.Message)
        End Try

    End Sub

    'Dette oppdaterer MOTD i databasen og reflekterer det som vises på fremsiden. Man klikker på knappen for å kjøre koden.
    'Den kjører også en tilleggsprosydre som oppdaterer på start siden også.
    Private Sub AdminMOTDEndreB_Click(sender As Object, e As EventArgs) Handles AdminMOTDEndreB.Click

        TxtAdminMOTD.Text = SQLWhiteWash(TxtAdminMOTD.Text)

        Try
            DBConnect()
            Dim AdminMOTDSet As New MySqlCommand("UPDATE message_of_the_day SET message = '" & TxtAdminMOTD.Text & "' WHERE message_id = 1;", Tilkobling)
            AdminMOTDSet.ExecuteNonQuery()
            DBDisconnect()
            TxtAdminMOTD.Text = ""
            MsgBox("Message of the Day har blitt endret!")
            StartMOTDUpdate()
        Catch AdminSqlError4 As MySqlException
            MsgBox("Man får ikke koble til databasen: " & AdminSqlError4.Message)
        End Try
    End Sub

    Private Sub AdminNBOpprettB_Click(sender As Object, e As EventArgs) Handles AdminNBOpprettB.Click
        Dim NBInputSjekk As Boolean

        NBInputSjekk = CheckVarChar30(TxtAdminNBPassord.Text)
        If NBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig passord. (Mindre enn 30 tegn)")
            Exit Sub
        End If
        NBInputSjekk = CheckVarChar20(TxtAdminNBFornavn.Text)
        If NBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig fornavn. (Mindre enn 20 tegn)")
            Exit Sub
        End If
        NBInputSjekk = CheckVarChar30(TxtAdminNBEtternavn.Text)
        If NBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig etternavn. (Mindre enn 30 tegn)")
            Exit Sub
        End If
        NBInputSjekk = CheckIntValue(TxtAdminNBTime.Text)
        If NBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig timelønn. Tallformat.")
            Exit Sub
        End If
        NBInputSjekk = CheckVarChar30(TxtAdminNBEpost.Text)
        If NBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig epost. (Mindre enn 30 tegn)")
            Exit Sub
        End If
        NBInputSjekk = CheckVarChar15(TxtAdminNBTelefon.Text)
        If NBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig passord. (Mindre enn 30 tegn)")
            Exit Sub
        End If

        AdminNyBruker()
        AdminBrukerIDCalc()
    End Sub

    Private Sub AdminEBLastInnB_Click(sender As Object, e As EventArgs) Handles AdminEBLastInnB.Click
        AdminLastInnEndreBruker()


    End Sub



    Private Sub AdminEBEndreB_Click(sender As Object, e As EventArgs) Handles AdminEBEndreB.Click
        Dim EBInputSjekk As Boolean

        EBInputSjekk = CheckIntValue(TxtAdminEBBID.Text)
        If EBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig Bruker ID. (Tall mindre enn 12 char)")
            Exit Sub
        End If
        EBInputSjekk = CheckVarChar20(TxtAdminEBFornavn.Text)
        If EBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig fornavn. (Mindre enn 20 char)")
            Exit Sub
        End If
        EBInputSjekk = CheckVarChar30(TxtAdminEBEtternavn.Text)
        If EBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig etternavn. (Mindre enn 30 char)")
            Exit Sub
        End If
        EBInputSjekk = CheckIntValue(TxtAdminEBTime.Text)
        If EBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig timelønn. Tallformat.")
            Exit Sub
        End If
        EBInputSjekk = CheckVarChar30(TxtAdminEBEpost.Text)
        If EBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig epost. (Mindre enn 30 char)")
            Exit Sub
        End If
        EBInputSjekk = CheckVarChar15(TxtAdminEBTelefon.Text)
        If EBInputSjekk = False Then
            MsgBox("Vennligst tast inn gyldig passord. (Mindre enn 30 char)")
            Exit Sub
        End If

        AdminEndreBruker()
    End Sub




















































#End Region

#Region "LoginTab"

    Public Sub LoginCheck()

        Dim LoginBrukerTabell As New DataTable
        Dim LoginPassordTabell As New DataTable
        Dim LoginBrukerCheck As String = ""
        Dim LoginPassordCheck As String = ""
        Dim LoginAdminCheck As Integer = 0
        Dim LoginFornavn As String = ""
        Dim LoginEtternavn As String = ""
        Dim LoginBrukerTabellRow As DataRow
        Dim LoginPassordTabellRow As DataRow

        If SecurityCounter = 3 Then
            MsgBox("Du har blitt låst ut grunnet for mange feilforsøk.")
            Exit Sub
        End If

        If IsNumeric(TxtLoginBrukerID.Text) Then
        Else
            MsgBox("Vennligst tast inn gyldig BrukerID.")
            Exit Sub
        End If
        TxtLoginBrukerID.Text = SQLWhiteWash(TxtLoginBrukerID.Text)
        TxtLoginPassord.Text = SQLWhiteWash(TxtLoginPassord.Text)

        LoginBrukerTabell = SQLSelect("brukere", "*", "bruker_id=" & TxtLoginBrukerID.Text)
        LoginPassordTabell = SQLSelect("passord", "*", "pwd='" & TxtLoginPassord.Text & "' AND passord_id=" & TxtLoginBrukerID.Text)

        For Each LoginBrukerTabellRow In LoginBrukerTabell.Rows
            LoginBrukerCheck = LoginBrukerTabellRow("bruker_id")
            LoginFornavn = LoginBrukerTabellRow("fornavn")
            LoginEtternavn = LoginBrukerTabellRow("etternavn")
            LoginAdminCheck = LoginBrukerTabellRow("admin")
        Next

        For Each LoginPassordTabellRow In LoginPassordTabell.Rows
            LoginPassordCheck = LoginPassordTabellRow("pwd")
        Next

        If TxtLoginBrukerID.Text = LoginBrukerCheck And TxtLoginPassord.Text = LoginPassordCheck Then

        Else
            SecurityCounter = SecurityCounter + 1

            If SecurityCounter = 3 Then
                MsgBox("BrukerID, Passord eller begge er feil. Du har blitt låst ut grunnet for mange feilforsøk.")
            Else
                MsgBox("BrukerID, Passord eller begge er feil. Vennligst prøv igjen. Du har " & (3 - SecurityCounter) & " forsøk igjen.")
            End If
            Exit Sub
        End If

        If LoginAdminCheck <> 0 Then
            AdminBool = True
        End If
        LogBool = True
        FornavnString = LoginFornavn
        EtternavnString = LoginEtternavn
        LoginSuccess()
    End Sub

    Public Sub LoginSuccess()
        With HovedTab.TabPages
            .Insert(0, StartTab)
            .Insert(1, UtleieTab)
            .Insert(2, KDTab)
            .Insert(3, InventarTab)
            .Insert(4, LogiTab)
            .Insert(5, StatTab)
            .Remove(LoginTab)

        End With
        If AdminBool = True Then
            With HovedTab.TabPages
                .Insert(6, AdminTab)
                .Insert(7, DBAdminTab)
            End With
            StartRettigheterLabel.Text = "Brukerrettigheter: Admin"
        Else
            StartRettigheterLabel.Text = "Brukerrettigheter: Bruker"
        End If

        HovedTab.SelectedTab = StartTab

        StartVelkommenLabel.Text = "Velkommen, " & FornavnString & " " & EtternavnString
        StartMOTDUpdate()
    End Sub

    Private Sub BtnLoginLogin_Click(sender As Object, e As EventArgs) Handles BtnLoginLogin.Click
        LoginCheck()
    End Sub

    Private Sub BtnLoginAvslutt_Click(sender As Object, e As EventArgs) Handles BtnLoginAvslutt.Click
        End
    End Sub
#End Region
End Class
