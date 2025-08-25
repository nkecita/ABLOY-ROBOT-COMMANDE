Imports System.Net
Imports System.Xml
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.Reflection

Public Structure configRobot
    Dim Folderout As String
    Dim databasefichet As String
    Dim databaseecon As String
    Dim databaseecontest As String
    Dim varftp As String

    Dim databaseeconabloy As String
    Dim databaseeconstremler As String
    Dim databaseeconvachette As String
    Dim databaseeconyale As String
    Dim databaseeconsherlock As String
    Dim databaseeconrehab As String

    Dim modelversion As String
    Dim orderfilter As String
    Dim ftpserver As String
    Dim ftpuser As String
    Dim ftppassword As String
    Dim active As String
End Structure

Module Module1

    Public configGene As configRobot


    Sub Main()
        Try

            ' premier commit
            ' deuxième commit
            ' Lecture du fichier de configuration XML
            Dim i As Integer
            Dim startupPath As String
            Dim dsConfig As New DataSet



            startupPath = Assembly.GetExecutingAssembly().GetName().CodeBase

            startupPath = Path.GetDirectoryName(startupPath)

            dsConfig.ReadXml(startupPath + "\" + "config.xml")

            If dsConfig.Tables.Count = 1 Then
                With dsConfig.Tables(0).Rows(0)


                    configGene.Folderout = .Item("OUTPUT")
                    configGene.databasefichet = .Item("DATABASE-FICHET")
                    configGene.databaseecon = .Item("DATABASE-ECON")
                    configGene.databaseecontest = .Item("DATABASE-ECONTEST")

                    configGene.databaseeconabloy = .Item("DATABASE-ECONABLOY")
                    configGene.databaseeconstremler = .Item("DATABASE-ECONSTREMLER")
                    configGene.databaseeconvachette = .Item("DATABASE-ECONVACHETTE")
                    configGene.databaseeconrehab = .Item("DATABASE-REHAB")
                    configGene.databaseeconyale = .Item("DATABASE-ECONYALE")
                    configGene.databaseeconsherlock = .Item("DATABASE-ECONSHERLOCK")

                    configGene.orderfilter = .Item("ORDER-FILTER")
                    configGene.modelversion = .Item("MODELVERSION")
                    configGene.ftpserver = .Item("FTP-SERVER")
                    configGene.ftpuser = .Item("FTP-USER")
                    configGene.ftppassword = .Item("FTP-PASSWORD")
                    configGene.active = .Item("ACTIVE")
                    configGene.varftp = .Item("FTP")


                End With
            End If

            If is_dayoff() Then
                Return
            End If

            Dim cSql As String
            Dim dsCommande As DataSet = New DataSet()

            cSql = "select cde.[num_commande],cde.[status],cde.[configuration],cde.[environnement],cde.[prodid],econ.prodid as econid,cde.prodid "
            cSql = cSql & "From commandes_portes cde "
            cSql = cSql & "left join econ365 econ on econ.prodid = cde.prodid "
            cSql = cSql & "Where " & configGene.orderfilter
            cSql = cSql & " Order by cde.[num_commande] desc"
            Dim adapter = New OleDb.OleDbDataAdapter(cSql, configGene.databasefichet)
            adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
            adapter.Fill(dsCommande, "list_commande")


            Dim num_cmmOK As String = ""
            Dim num_cmmKO As String = ""

            Dim CurrentEconCde As String = ""
            Dim FileCommandeEcon As String = ""
            Dim FileEconDestination As String = ""
            Dim FileEconArchive As String = ""
            Dim FileEconCorbeille As String = ""

            ' on parcours les commandes à traite
            For i = 0 To dsCommande.Tables(0).Rows.Count - 1


                ' Récupération du numéro de commande courante
                CurrentEconCde = dsCommande.Tables(0).Rows(i).Item("num_commande").ToString

                logInformation("Traitement Commande : " & CurrentEconCde)

                FileCommandeEcon = "D:\clients\fichet\ftpecon\Configurations\eCon365\Nouveau\" & CurrentEconCde & ".xml"
                FileEconDestination = "D:\clients\fichet\ftpecon\Configurations\" & CurrentEconCde & ".xml"
                FileEconArchive = "D:\clients\fichet\ftpecon\Configurations\eCon365\Archive\" & CurrentEconCde & ".xml"
                FileEconCorbeille = "D:\clients\fichet\ftpecon\Configurations\eCon365\Corbeille\" & CurrentEconCde & ".xml"

                If File.Exists(FileCommandeEcon) Then

                    If IsDBNull(dsCommande.Tables(0).Rows(i).Item("econid")) = False Then

                        ' Le fichier exite dans le répertoire Nouveau et le prodid existe dans la table Econ365 -> Processus B
                        logInformation("Formatage Fichier: " & FileCommandeEcon)
                        FormatFileCdeEcon(FileCommandeEcon)  ' Remplacement _LINE_ par &#xa;
                        File.Copy(FileCommandeEcon, FileEconDestination, True)
                        logInformation("Fichier Copié : " & FileEconDestination)
                        File.Move(FileCommandeEcon, FileEconArchive)
                        logInformation("Fichier Archivé : " & FileEconArchive)
                        If configGene.active = "OUI" Then
                            Dim req As String = "UPDATE commandes_portes set status=2 where num_commande=" & dsCommande.Tables(0).Rows(i).Item("num_commande").ToString
                            CreateOleDbCommand(req, configGene.databasefichet)
                            logInformation("Commandes traitee (Processus B) : " & dsCommande.Tables(0).Rows(i).Item("num_commande").ToString)
                        End If
                    Else
                        logInformation("Fichier mis a la corbeille : " & FileEconCorbeille)
                        File.Move(FileCommandeEcon, FileEconCorbeille)
                    End If
                Else  ' On fait le processus Actuel A
                    'logInformation("Fichier non trouvé: " & FileCommandeEcon)

                    Dim dsEcon As New DataSet

                    cSql = "SELECT content FROM econelements WHERE econelements.name='" & _
                            dsCommande.Tables(0).Rows(i).Item("configuration").ToString & _
                           "' AND modelversion=" & configGene.modelversion

                    Dim csqlstring As String

                    Select Case dsCommande.Tables(0).Rows(i).Item("environnement").ToString.ToUpper

                        Case "DEFAULT"
                            csqlstring = configGene.databaseecon
                        Case "TEST"
                            csqlstring = configGene.databaseecontest
                        Case "ABLOY"
                            csqlstring = configGene.databaseeconabloy
                        Case "STREMLER"
                            csqlstring = configGene.databaseeconstremler
                        Case "VACHETTE"
                            csqlstring = configGene.databaseeconvachette
                        Case "YALE"
                            csqlstring = configGene.databaseeconyale
                        Case "SHERLOCK"
                            csqlstring = configGene.databaseeconsherlock
                        Case "REHAB"
                            csqlstring = configGene.databaseeconrehab

                        Case Else
                            csqlstring = configGene.databaseecon
                    End Select

                    Dim adapterEcon = New OleDb.OleDbDataAdapter(cSql, csqlstring)

                    adapterEcon.MissingSchemaAction = MissingSchemaAction.AddWithKey
                    adapterEcon.Fill(dsEcon, "list_configuration")
                    If dsEcon.Tables(0).Rows.Count = 1 Then
                        If ecrire(dsEcon.Tables(0).Rows(0).Item(0).ToString, dsCommande.Tables(0).Rows(i).Item("num_commande").ToString) Then
                            If configGene.active = "OUI" Then
                                Dim req As String = "UPDATE commandes_portes set status=2 where num_commande=" & dsCommande.Tables(0).Rows(i).Item("num_commande").ToString
                                CreateOleDbCommand(req, configGene.databasefichet)
                                num_cmmOK &= dsCommande.Tables(0).Rows(i).Item("num_commande").ToString & " - "
                            End If
                        Else
                            num_cmmKO &= dsCommande.Tables(0).Rows(i).Item("num_commande").ToString & " - "
                        End If
                    End If
                End If
            Next
            logInformation("Commandes traitee : " & num_cmmOK)
            logInformation("Commandes non traitee : " & num_cmmKO)
        Catch ex As Exception
            logError(ex.ToString)
        End Try

    End Sub

    Function ecrire(ByVal wXml As String, ByVal wcommande As String) As Boolean
        Try
            If File.Exists(configGene.Folderout & "\" & wcommande & ".xml") Then
                Return False
            Else
                wXml = wXml.Replace("?>", " encoding='utf-8' ?>")

                Dim monStreamWriter As StreamWriter = New StreamWriter(configGene.Folderout & "\" & wcommande & ".xml", True)

                monStreamWriter.WriteLine(wXml)

                monStreamWriter.Close()
                monStreamWriter = Nothing

                If configGene.varftp.ToString.ToUpper = "OUI" Then

                    ' read in file...

                    Dim di As New IO.DirectoryInfo(configGene.Folderout)
                    Dim aryFi As IO.FileInfo() = di.GetFiles("*.xml")
                    Dim fi As IO.FileInfo

                    For Each fi In aryFi

                        Dim clsRequest As System.Net.FtpWebRequest = DirectCast(System.Net.WebRequest.Create(configGene.ftpserver & "\" & fi.Name), System.Net.FtpWebRequest)
                        clsRequest.Credentials = New System.Net.NetworkCredential(configGene.ftpuser, configGene.ftppassword)
                        clsRequest.Method = System.Net.WebRequestMethods.Ftp.UploadFile


                        Dim bFile() As Byte = System.IO.File.ReadAllBytes(configGene.Folderout & "\" & fi.Name)

                        ' upload file...
                        Dim clsStream As System.IO.Stream = clsRequest.GetRequestStream()
                        clsStream.Write(bFile, 0, bFile.Length)
                        clsStream.Close()
                        clsStream.Dispose()
                        fi.MoveTo(configGene.Folderout & "\BACKUP" & "\" & fi.Name)

                    Next
                End If
                Return True
            End If
        Catch ex As Exception
            logError(ex.ToString)
            Return False
        End Try

    End Function

    Function is_dayoff()
        Dim dsDayOff As DataSet = New DataSet()
        Dim csql As String

        csql = "SELECT dt_ferie FROM jours_feries WHERE dt_ferie = convert(varchar,getdate(),112)"

        Dim adapter = New OleDb.OleDbDataAdapter(csql, configGene.databasefichet)
        adapter.MissingSchemaAction = MissingSchemaAction.AddWithKey
        adapter.Fill(dsDayOff, "list_dayoff")

        If dsDayOff.Tables(0).Rows.Count = 1 Then
            Return True
        Else
            Return False
        End If

    End Function
    Function FormatFileCdeEcon(ByVal CheminFichier As String)
        Try
            If File.Exists(CheminFichier) Then
                Dim contenu As String = File.ReadAllText(CheminFichier)
                Dim contenumodifie As String = contenu.Replace("_line_", "&#xA;")
                File.WriteAllText(CheminFichier, contenumodifie)

            End If
            Return True
        Catch ex As Exception
            logError(ex.ToString)
            Return False
        End Try

    End Function
    Public Sub logError(ByVal str As String)
        EventLog.WriteEntry(System.Reflection.Assembly.GetEntryAssembly().GetName().Name, str, EventLogEntryType.Error)
    End Sub

    Public Sub logInformation(ByVal str As String)
        EventLog.WriteEntry(System.Reflection.Assembly.GetEntryAssembly().GetName().Name, str, EventLogEntryType.Information)
    End Sub

    Private Sub CreateOleDbCommand(ByVal queryString As String, ByVal connectionString As String)
        Using connection As New OleDb.OleDbConnection(connectionString)
            connection.Open()
            Dim command As New OleDb.OleDbCommand(queryString, connection)
            command.ExecuteNonQuery()
        End Using
    End Sub

End Module
