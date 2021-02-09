Imports Microsoft.Office.Interop
Imports Microsoft.SharePoint.Client
Imports Microsoft.SharePoint
Imports Microsoft
Imports System.IO
Imports System.Text
Imports System.Xml
Imports Microsoft.VisualBasic.FileIO



Module Module1
    Dim Cliente(10000, 4) As String
    Dim Macchina(50, 100) As String
    Dim Scelta(0, 100) As String
    Dim Campi(100, 0) As String
    Dim SAM(100, 2) As String
    ReadOnly Indiriz As String = "https://home.intranet.epiroc.com/sites/cc/iyc/"

    Public Property Cliente1 As String(,)
        Get
            Return Cliente
        End Get
        Set(value As String(,))
            Cliente = value
        End Set
    End Property

    Public Property Macchina1 As String(,)
        Get
            Return Macchina
        End Get
        Set(value As String(,))
            Macchina = value
        End Set
    End Property

    Public Property Scelta1 As String(,)
        Get
            Return Scelta2
        End Get
        Set(value As String(,))
            Scelta2 = value
        End Set
    End Property

    Public Property Scelta2 As String(,)
        Get
            Return Scelta3
        End Get
        Set(value As String(,))
            Scelta3 = value
        End Set
    End Property

    Public Property Scelta3 As String(,)
        Get
            Return Scelta
        End Get
        Set(value As String(,))
            Scelta = value
        End Set
    End Property

    Public Property Campi1 As String(,)
        Get
            Return Campi
        End Get
        Set(value As String(,))
            Campi = value
        End Set
    End Property

    Public Property SAM1 As String(,)
        Get
            Return SAM
        End Get
        Set(value As String(,))
            SAM = value
        End Set
    End Property

    Sub Main()
Inizio:
        Try
            Numero_OFF()
            Elenco_Mac()
            SP()
            Apri_SAM()
        Catch Ex As ClientRequestException
            If MsgBox($"Verifica di essere connesso alla rete aziendale {vbCr}{vbCr}Errore: {Ex.Message}", vbRetryCancel, "Errore") = vbRetry Then
                GoTo Inizio
            Else
                Form1.Close()
            End If
        End Try

    End Sub


    Sub Crea_File()
        Dim Codice As String = ""
        Dim Path1 As String = Indiriz & "MRProductsalessupport/SED/"
        Dim I, P As Integer
        Dim context As New ClientContext(Path1)
        Dim testList As List = context.Web.Lists.GetByTitle("Codice")
        Dim query As CamlQuery = CamlQuery.CreateAllItemsQuery(10000)
        Dim items As ListItemCollection = testList.GetItems(query)
        context.Load(items)
        context.ExecuteQuery()
        For Each listITem As ListItem In items
            Codice = listITem("Testo")
        Next



        Dim path As String = My.Computer.FileSystem.SpecialDirectories.Desktop & "\"
        Dim NomeF As String = "Temp.xml"
        Dim fs As FileStream = System.IO.File.Create(path & NomeF)
        Dim info As Byte() = New UTF8Encoding(True).GetBytes(Codice)
        fs.Write(info, 0, info.Length)
        fs.Close()
        Dim mW As Word.Application
        Dim mDoc As Word.Document
        mW = CreateObject("Word.Application")
        mW.Visible = False
        mDoc = mW.Documents.Add(path & NomeF)
        Dim U As Integer
        For U = 1 To mDoc.FormFields.Count
            Campi(U, 0) = mDoc.FormFields(U).Name

        Next

        mDoc.FormFields(Campi(1, 0)).Result = Form1.ComboBox1.Text
        mDoc.FormFields(Campi(2, 0)).Result = Form1.TextBox1.Text
        mDoc.FormFields(Campi(3, 0)).Result = Form1.TextBox2.Text
        mDoc.FormFields(Campi(4, 0)).Result = Form1.TextBox3.Text
        mDoc.FormFields(Campi(5, 0)).Result = Format(Now(), "dd MMMM yyyy")
        mDoc.FormFields(Campi(6, 0)).Result = Form1.Label7.Text
        mDoc.FormFields(Campi(7, 0)).Result = Scelta1(0, 1)
        mDoc.FormFields(Campi(8, 0)).Result = Scelta1(0, 10)
        mDoc.FormFields(Campi(9, 0)).Result = Scelta1(0, 8)
        mDoc.FormFields(Campi(10, 0)).Result = Form1.ComboBox2.Text
        mDoc.FormFields(Campi(11, 0)).Result = Form1.Label8.Text
        mDoc.FormFields(Campi(12, 0)).Result = Scelta1(0, 6)
        mDoc.FormFields(Campi(13, 0)).Result = Scelta1(0, 9)
        mDoc.FormFields(Campi(14, 0)).Result = Scelta1(0, 13)
        mDoc.FormFields(Campi(15, 0)).Result = Scelta1(0, 18)
        mDoc.FormFields(Campi(16, 0)).Result = Scelta1(0, 16)
        mDoc.FormFields(Campi(17, 0)).Result = Scelta1(0, 17)
        mDoc.FormFields(Campi(18, 0)).Result = Scelta1(0, 10)
        mDoc.FormFields(Campi(19, 0)).Result = Scelta1(0, 10)
        mDoc.FormFields(Campi(20, 0)).Result = Scelta1(0, 14)
        mDoc.FormFields(Campi(21, 0)).Result = Scelta1(0, 15)
        mDoc.FormFields(Campi(22, 0)).Result = Scelta1(0, 10)
        For T = 23 To 38
            If Scelta1(0, T + 1) <> "." Then mDoc.FormFields(Campi(T, 0)).Result = Scelta1(0, T + 1)
        Next
        For T = 39 To 45
            If Scelta1(0, T + 1) <> "." Then mDoc.FormFields(Campi(T, 0)).Result = Scelta1(0, T + 1)
        Next
        mDoc.FormFields(Campi(47, 0)).Result = Scelta1(0, 8)
        mDoc.FormFields(Campi(48, 0)).Result = Scelta1(0, 4)
        mDoc.FormFields(Campi(49, 0)).Result = Scelta1(0, 23)
        mDoc.FormFields(Campi(50, 0)).Result = Scelta1(0, 5)
        Dim Cl
        If Right(Form1.ComboBox1.Text, 1) = "." Then
            Cl = Left(Form1.ComboBox1.Text, Len(Form1.ComboBox1.Text) - 1)
        Else
            Cl = Form1.ComboBox1.Text
        End If
        Dim NomeFa As String = "Offerta " & Form1.Label7.Text & " - " & Form1.ListBox1.SelectedItem & "- " & Cl
        mDoc.SaveAs2(path & NomeFa & ".docx", Word.WdSaveFormat.wdFormatDocumentDefault)

        mDoc.SaveAs2(path & NomeFa & ".pdf", Word.WdSaveFormat.wdFormatPDF)
        mDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges)
        mW.Quit()
        Try
            CaricaSP(path & NomeFa & ".docx")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Try
            ScriviSP()
        Catch ex1 As Exception
            MsgBox(ex1.Message)
        End Try

        Dim Scelta
        Scelta = MsgBox("Vuoi il Word?", vbYesNo, "Scelta")
        If Scelta = vbNo Then Kill(path & NomeFa & ".docx")
        MsgBox("Offerta creata in " & My.Computer.FileSystem.SpecialDirectories.Desktop,, "Offerta Creata")

        Form1.Close()
        Form3.Close()
        Try
            Kill(path & NomeF)
        Finally
        End Try

    End Sub



    Sub SP()
        Dim Path As String = Indiriz
        Dim I As Integer
        Dim context As New ClientContext(Path)
        Dim testList As List = context.Web.Lists.GetByTitle("Customers")
        Dim query As CamlQuery = CamlQuery.CreateAllItemsQuery(10000)
        Dim items As ListItemCollection = testList.GetItems(query)
        context.Load(items)
        context.ExecuteQuery()
        I = 0
        Form2.Label1.Text = "Download di: Elenco Clienti"
        Form2.Show()

        Form2.ProgressBar1.Maximum = items.Count

        For Each listItem As ListItem In items
            Form2.ProgressBar1.Value = I

            Cliente1(I, 0) = listItem("Title")
            Cliente1(I, 1) = listItem("CustomerName")
            Cliente1(I, 2) = listItem("Address_x0020_Line_x0020_1")
            Cliente1(I, 3) = listItem("Address_x0020_Line_x0020_2")
            Form1.ComboBox1.Items.Add(listItem("CustomerName"))
            I += 1
        Next
        Form2.Close()
    End Sub


    Sub Testo()

        Dim Indirizzo As String = Cliente1(Form1.ComboBox1.SelectedIndex, 3)
        Indirizzo = Trim(Indirizzo)
        Indirizzo = Trim(Left(Indirizzo, Len(Indirizzo) - 2)) & " (" & Right(Indirizzo, 2) & ")"
        Form1.TextBox1.Text = Cliente1(Form1.ComboBox1.SelectedIndex, 2)
        Form1.TextBox2.Text = Trim(Indirizzo)
    End Sub

    Sub Elenco_Mac()

        Dim Path As String = Indiriz & "MRProductsalessupport/SED/"
        Dim I, P As Integer
        Dim context As New ClientContext(Path)
        Dim testList As List = context.Web.Lists.GetByTitle("Parco Macchine")
        Dim query As CamlQuery = CamlQuery.CreateAllItemsQuery(10000)
        Dim items As ListItemCollection = testList.GetItems(query)
        context.Load(items)
        context.ExecuteQuery()
        I = 0

        Form2.Label1.Text = "Download di: Elenco Macchine"
        Form2.Show()

        Form2.ProgressBar1.Maximum = items.Count

        For Each listItem As ListItem In items
            Form2.ProgressBar1.Value = I


            P = 0

            For Each Fiels In listItem.FieldValues
                Dim Nome As String = Fiels.Key
                On Error Resume Next
                Macchina1(I, P) = listItem("" & Nome & "")
                P += 1
            Next


            If I = 0 Then Form1.ListBox1.Items.Add(listItem("Title"))
            If Form1.ListBox1.Items.Contains(listItem("Title")) = False Then Form1.ListBox1.Items.Add(listItem("Title"))

            I += 1
        Next

        Form2.Close()
    End Sub

    Sub Rock_Drill(Indice As String)
        Form1.ListBox2.Items.Clear()
        Dim I As Integer

        For I = 0 To 50
            If Macchina1(I, 1) = Indice Then Form1.ListBox2.Items.Add(Macchina1(I, 10))
        Next

    End Sub

    Sub Dati(Lista1 As String, Lista2 As String)
        Dim I, O As Integer
        Form1.Label3.Text = ""
        For I = 0 To 50

            If Macchina1(I, 1) = Lista1 And Macchina1(I, 10) = Lista2 Then
                For O = 0 To 100
                    Form1.Label3.Text = Form1.Label3.Text & "; " & Macchina1(I, O)
                    Scelta1(0, O) = Macchina1(I, O)
                Next

            End If
        Next
    End Sub


    Sub Numero_OFF()
        Dim Path As String = Indiriz & "MRProductsalessupport/SED/"
        Dim Numero As String = ""
        Dim Anno
        Dim context As New ClientContext(Path)
        Dim testList As List = context.Web.Lists.GetByTitle("Registro")
        Dim query As CamlQuery = CamlQuery.CreateAllItemsQuery(10000)
        Dim items As ListItemCollection = testList.GetItems(query)
        context.Load(items)
        context.ExecuteQuery()

        For Each listItem As ListItem In items
            Numero = listItem("Title")
        Next
        Anno = Right(Numero, 2)
        Numero = Mid(Numero, 5, 4)
        If Anno = Format(Now(), "yy") Then
            Form1.Label7.Text = "SED-" & Format(Numero + 1, "0000") & "." & Format(Now(), "yy")
        Else
            Form1.Label7.Text = "SED-0001" & "." & Format(Now(), "yy")
        End If
    End Sub


    Sub Controlla()
        With Form1
            If .ListBox1.SelectedItem <> "" _
                And .ListBox2.SelectedItem <> "" _
                And .ComboBox1.Text <> "" _
                And .ComboBox2.Text <> "" _
                And .TextBox1.Text <> "" _
                And .TextBox2.Text <> "" _
                And .TextBox3.Text <> "" Then
                .Button1.Enabled = True
            Else
                .Button1.Enabled = False
            End If
        End With
    End Sub

    Sub Apri_SAM()
        Dim Path As String = Indiriz & "MRProductsalessupport/SED/"
        Dim I As Integer = 0
        Dim context As New ClientContext(Path)
        Dim testList As List = context.Web.Lists.GetByTitle("SAM")
        Dim query As CamlQuery = CamlQuery.CreateAllItemsQuery(10000)
        Dim items As ListItemCollection = testList.GetItems(query)
        context.Load(items)
        context.ExecuteQuery()

        For Each listItem As ListItem In items
            SAM1(I, 0) = listItem("Title")
            SAM1(I, 1) = listItem("Job_x0020_title")
            Form1.ComboBox2.Items.Add(listItem("Title"))
            I += 1
        Next

    End Sub


    Sub JobTitle()
        Dim I As Integer
        For I = 0 To 100
            If SAM1(I, 0) = Form1.ComboBox2.Text Then
                Form1.Label8.Text = SAM1(I, 1)
                Exit For
            End If
        Next
    End Sub

    Sub ScriviSP()
        Dim Path As String = Indiriz & "MRProductsalessupport/SED/"
        Dim context As New ClientContext(Path)
        Dim testList As List = context.Web.Lists.GetByTitle("Registro")
        Dim itemCreateInfo As New ListItemCreationInformation
        Dim oListItem As ListItem
        oListItem = testList.AddItem(itemCreateInfo)
        oListItem("Title") = Form1.Label7.Text
        oListItem("Date") = Format(Now(), "dd/MM/yyyy")
        oListItem("x8vn") = Form1.ComboBox1.Text
        oListItem("Rig") = Form1.ListBox1.SelectedItem
        oListItem("CMP") = Scelta1(0, 4)
        oListItem("Autore") = My.User.Name
        oListItem("hn0t") = Form1.ComboBox2.Text
        oListItem.Update()
        context.ExecuteQuery()
    End Sub

    Sub CaricaSP(fileName As String)
        Dim Path As String = Indiriz & "MRProductsalessupport/SED/"
        Dim Path1 As String = Path & "SED_Offerte_docx"
        Dim context As New ClientContext(Path)
        Dim testList As List = context.Web.Lists.GetByTitle("SED_Offerte_docx")
        context.Load(testList.RootFolder)
        Dim newFile = New FileCreationInformation()
        newFile.Overwrite = True
        newFile.Content = My.Computer.FileSystem.ReadAllBytes(fileName)
        newFile.Url = System.IO.Path.GetFileName(fileName)
        testList.RootFolder.Files.Add(newFile)
        context.ExecuteQuery()
    End Sub


End Module


