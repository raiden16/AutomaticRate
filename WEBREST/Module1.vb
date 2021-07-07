Imports System
Imports System.Net
Imports Newtonsoft.Json

Module Module1

    Public SBOCompany As SAPbobsCOM.Company
    Dim rateUSD, rateEUR As Decimal

    Sub Main()

        Conectar()
        rateUSD = GetRate("SF43718", "USD")
        rateEUR = GetRate("SF46410", "EUR")
        ORTT()
        Desconectar()

    End Sub


    Public Function Conectar()

        Try

            SBOCompany = New SAPbobsCOM.Company

            SBOCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB
            SBOCompany.Server = My.Settings.Server
            SBOCompany.LicenseServer = My.Settings.LicenseServer
            SBOCompany.DbUserName = My.Settings.DbUserName
            SBOCompany.DbPassword = My.Settings.DbPassword

            SBOCompany.CompanyDB = My.Settings.CompanyDB

            SBOCompany.UserName = My.Settings.UserName
            SBOCompany.Password = My.Settings.Password

            SBOCompany.Connect()

        Catch ex As Exception

            Dim stError As String
            stError = "Error al tratar de hacer conexión con SAP B1. " & ex.Message
            Setlog(stError, " ", " ", " ", " ")

        End Try

    End Function


    Public Function GetRate(ByVal serie As String, ByVal currency As String)

        Dim URL As String
        Dim json As String
        Dim cadena, dato, lenght As Integer
        Dim tcpre, tcpre2, tc As String

        Try

            URL = "https://www.banxico.org.mx/SieAPIRest/service/v1/series/" & serie & "/datos/oportuno?token=1c1363ca530b919937dcb6ed459b278e9fb219d5cab718348fd5839529e00bea"
            json = New WebClient().DownloadString(URL)
            cadena = json.Length
            dato = json.IndexOf("""dato"":""") + 8
            lenght = cadena - dato
            tcpre = ArreglarTexto(json.Substring(dato, lenght), """", " ").ToString.Trim
            tcpre2 = ArreglarTexto(tcpre, "}", " ").ToString.Trim
            tc = ArreglarTexto(tcpre2, "]", "").ToString.Trim

            Return tc

        Catch ex As Exception

            Dim stError As String
            stError = "Error al obtener el tipo de cambio GetRate. " & ex.Message
            Setlog(stError, URL, currency, " ", " ")

        End Try

    End Function


    Public Function ArreglarTexto(ByVal TextoOriginal As String, ByVal QuitarCaracter As String, ByVal PonerCaracter As String)

        TextoOriginal = TextoOriginal.Replace(QuitarCaracter, PonerCaracter)
        Return TextoOriginal

    End Function


    Public Function ORTT()

        Dim oORTT As SAPbobsCOM.SBObob

        oORTT = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)

        Try

            If Now.DayOfWeek <= 4 Then

                oORTT.SetCurrencyRate("USD", Now.Date, rateUSD, True)
                oORTT.SetCurrencyRate("USD", Now.Date.AddDays(1), rateUSD, True)
                oORTT.SetCurrencyRate("EUR", Now.Date, rateEUR, True)
                oORTT.SetCurrencyRate("EUR", Now.Date.AddDays(1), rateEUR, True)

            ElseIf Now.DayOfWeek = 5 Then

                oORTT.SetCurrencyRate("USD", Now.Date, rateUSD, True)
                oORTT.SetCurrencyRate("USD", Now.Date.AddDays(1), rateUSD, True)
                oORTT.SetCurrencyRate("USD", Now.Date.AddDays(2), rateUSD, True)
                oORTT.SetCurrencyRate("USD", Now.Date.AddDays(3), rateUSD, True)
                oORTT.SetCurrencyRate("EUR", Now.Date, rateEUR, True)
                oORTT.SetCurrencyRate("EUR", Now.Date.AddDays(1), rateEUR, True)
                oORTT.SetCurrencyRate("EUR", Now.Date.AddDays(2), rateEUR, True)
                oORTT.SetCurrencyRate("EUR", Now.Date.AddDays(3), rateEUR, True)

            End If

        Catch ex As Exception

            Dim stError As String
            stError = "Error al actualizar el tipo de cambio ORTT. " & ex.Message
            Setlog(stError, " ", " ", rateUSD, rateEUR)

        End Try

    End Function


    Public Function Desconectar()

        Try

            SBOCompany.Disconnect()

        Catch ex As Exception

            Dim stError As String
            stError = "Error al tratar de cerrar la conexión con SAP B1. " & ex.Message
            Setlog(stError, " ", " ", " ", " ")

        End Try

    End Function


    Public Function Setlog(ByVal stError As String, ByVal url As String, ByVal currency As String, ByVal rateUSD As String, ByVal rateEUR As String)

        Dim oRecSettxb As SAPbobsCOM.Recordset
        Dim stQuerytxb As String

        Try

            stError = ArreglarTexto(stError, "'", " ")
            oRecSettxb = SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            stQuerytxb = "Insert Into LOG_Rate values ('" & url & "','" & currency & "','" & rateUSD & "','" & rateEUR & "','" & stError & "',current_date)"
            oRecSettxb.DoQuery(stQuerytxb)

        Catch ex As Exception

            'MsgBox(stError)

        End Try

    End Function


End Module
