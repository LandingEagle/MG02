Imports System.Data.SQLite
Imports System.Text
Imports System.Console


Module Module1

   Const MGdb As String = "C:\Users\Thomas\Dropbox\Developement\VStudio\Projects\MediaGorilla\MediaGorilla.db"

   Public Structure stcArtist
      Dim ArtistID As Int16
      Dim Name As String
      Dim NameSort As String
   End Structure

   Public Structure stcAlbum

   End Structure

   Public Artist As New stcArtist

   Public ds As String = "Data Source=" & MGdb & ";"
   Public MGcnn As New SQLite.SQLiteConnection(ds)


   Sub Main()

      MGcnn.Open()
      Artist = GetArtist(1)
      MsgBox(CrtSrtStr(CStr(Artist.Name)))

      Artist = GetArtist(2)
      MsgBox(CrtSrtStr(CStr(Artist.Name)))

   End Sub

   Function GetArtist(ArtistID As Int16) As stcArtist

      Dim SqlStmt As String = "select * from Artist as Artist where ArtistID = " & Format(ArtistID, "0")

      Dim MGcmd As New SQLiteCommand(SqlStmt, MGcnn)
      Dim MGrdr As SQLiteDataReader = MGcmd.ExecuteReader()

      If Not MGrdr.HasRows Then
         Return Nothing
         Exit Function
      End If

      While MGrdr.Read()

         Artist.ArtistID = MGrdr.Item(0)
         Artist.Name = MGrdr.Item(1)
         Artist.NameSort = MGrdr.Item(2)

         Console.WriteLine("|" & Artist.ArtistID _
                         & ";" & Artist.Name _
                         & ";" & Artist.NameSort _
                         & "|")
      End While

      Return Artist

      MGcmd.Dispose()
      MGcnn.Close()

   End Function

   Function CrtSrtStr(InpVal As String) As String

      Dim ivARA() As String
      Dim OutVal As New StringBuilder

      Dim x As Int16

      Dim indThe As Boolean = False

      InpVal = UCase(Trim(InpVal))
      InpVal = Replace(InpVal, "Ä", "Ae")
      InpVal = Replace(InpVal, "Ö", "Oe")
      InpVal = Replace(InpVal, "Ü", "Ue")
      InpVal = Replace(InpVal, "ß", "Sz")
      ivARA = Split(InpVal, " ")

      If ivARA(0) = "THE" Or ivARA(0) = "DIE" Then
         For x = 1 To ivARA.Count - 1
            If x > 1 Then OutVal.Append(" ")
            OutVal.Append(ivARA(x))
         Next x
         OutVal.Append(", ")
         OutVal.Append(ivARA(0))
      Else
         For x = 0 To ivARA.Count - 1
            If x > 0 And x < ivARA.Count - 1 Then OutVal.Append(" ")
            OutVal.Append(ivARA(x))
         Next x
      End If

      Return Trim(OutVal.ToString)

   End Function

End Module