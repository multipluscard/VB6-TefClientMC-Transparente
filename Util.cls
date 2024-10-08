VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim strStackTrace As String
Dim fs As New FileSystemObject
Public Sub AdicionaLog(strMsg As String, strStackTrace)
      strStackTrace = Null
      CriarArquivo "tef", strMsg, strStackTrace
End Sub

Private Function CriarArquivo(sNomeArquivo As String, strMsg As String, strStackTrace)
   Dim f As File
   
   Dim sMsg As String
   Dim TextStream As TextStream
   Dim cam
   Dim caminho As String
   Dim aplicacao As String
   Dim assVersin
   Dim assSec
   Dim assAlterna
   aplicacao = CStr(App.Path)
     
   If IsEmpty(strMsg) = True And IsEmpty(strStackTrace) = True Then
   End If
   
   caminho = aplicacao + "\nlog.txt"
   
   If Dir(caminho) <> "" Then
   Else
   End If
   
   If IsEmpty(sNomeArquivo) = True Then
   End If
   
   If ArquivoEmUso(aplicacao + "\" + sNomeArquivo + ".log") Then
   End If

   On Error GoTo ErrHandler
   
   cam = App.Path + "\" + sNomeArquivo + ".log"
   assVersin = CStr(App.Major) & "." & CStr(App.Minor) & "." & CStr(App.Revision)
   assSec = fs.FileExists(cam)
   If assSec = True Then
      Set f = fs.GetFile(cam)
      assAlterna = f.Size
      If assSec = True And assAlterna >= 314572800 Then
      
         On Error GoTo ErrHandler
         Kill App.Path + sNomeArquivo + "_old.log"
         FileCopy App.Path + sNomeArquivo + ".log", App.Path + sNomeArquivo + "_old.log"
         

ErrHandler:
         
      End If
      
      Set TextStream = fs.OpenTextFile(App.Path + "\" + sNomeArquivo + ".log", 8, True, 0)
     
      If strMsg <> "" Then
         TextStream.WriteLine (Format(Date + Time, "dd/mm/yyyy hh:mm:ss") + " v. " + CStr(assVersin) + " :: " + CStr(strMsg))
      End If
      
      If (Not strStackTrace = "") Or (Not IsNull(strStackTrace)) Then
         TextStream.WriteLine (Format(Date + Time, "dd/mm/yyyy hh:mm:ss") + " v. " + CStr(assVersin) + " :: " + CStr(strStackTrace))
      End If
      
      TextStream.Close
      Set TextStream = Nothing
   
   End If
End Function

Public Function ArquivoEmUso(caminhoArquivo As String) As Boolean
   Dim arqtxt As TextStream
   
On Error GoTo ErrHandler
   Set arqtxt = fs.OpenTextFile(caminhoArquivo, 8, True, 0)
   arqtxt.Close
   ArquivoEmUso = False
   Exit Function
   
ErrHandler:
   ArquivoEmUso = True
End Function
