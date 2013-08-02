VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmLoonstroken"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnCancel_Click()
    DoCmd.Close
End Sub

Private Sub btnOK_Click()
    Dim dbs As DAO.Database
    Dim rstPosten As DAO.Recordset
    Dim qryDef As DAO.QueryDef
    
    Dim strSQL As String
    Dim strMdrwNr As String
    Dim strVolgNr As String
    Dim strMaand As String
    
    Dim strStamGegevens As String
    Dim strMutaties As String
    Dim intCount As Integer
    
    intCount = 0
    Me.edtCount.SetFocus
    Me.edtCount.Text = intCount
    
    Set dbs = CurrentDb
                
    Set rstPosten = dbs.OpenRecordset("tblPosten", dbOpenDynaset)
    rstPosten.MoveFirst
    
    Do Until rstPosten.EOF
        intCount = intCount + 1
        Me.edtCount.SetFocus
        Me.edtCount.Text = intCount
        Me.Refresh

        strMdrwNr = Trim(rstPosten.Fields("Veld5").Value)
        strVolgNr = Trim(rstPosten.Fields("Veld2").Value)
        strMaand = Trim(rstPosten.Fields("Veld6").Value)
        
        If Len(strMaand) <> 1 Then
            strMaand = Mid(strMaand, 5, 2)
        Else
            strMaand = "0" & strMaand
        End If
        
        strStamGegevens = "StamGegevens" & strMaand
        strMutaties = "Mutaties" & strMaand
            
        strSQL = "SELECT " & strStamGegevens & ".*, " & strMutaties & ".*, " & strMutaties & ".Sleutel" & " & " & Chr(34) & Chr(95) & Chr(34) & " & " & strMutaties & ".Looncomponent AS Component "
        strSQL = strSQL & ", " & strMutaties & ".Maand AS aMonth, " & strMutaties & ".Sleutel AS aKey "
        strSQL = strSQL & "FROM " & strStamGegevens & ", " & strMutaties & " "
        strSQL = strSQL & "WHERE " & strStamGegevens & ".PersNr = " & strMutaties & ".PersNr AND " & strMutaties & ".PersNr = " & Chr(34) & strMdrwNr & Chr(34)

        dbs.QueryDefs("qryLoonstrook").SQL = strSQL
                
        'DoCmd.OpenReport "rptLoonstrook", acViewPreview
        DoCmd.OutputTo acReport, "rptLoonstrook", acFormatRTF, "F:\Temp\" & strVolgNr & ".rtf", False

        rstPosten.MoveNext
    Loop
    
    DoCmd.Close
End Sub
