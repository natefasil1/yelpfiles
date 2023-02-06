Attribute VB_Name = "Module1"
Public contestant1 As String
Public contestant2 As String
Public contestant3 As String
Public finalCategory As String
Public finalQuestion As String
Public cnn As ADODB.Connection


Public Function dbConnection()

   Set cnn = New ADODB.Connection

      With cnn
       .Provider = "Microsoft.Jet.OLEDB.4.0"
       .ConnectionString = "User ID=Admin;password= ;" & " Data Source=" & App.Path & "\jeopardy.mdb;"
       .CursorLocation = adUseClient
       .Open
      End With
End Function

      
      

