' Invoke Code (VB) - dtReportData as In/Out argument (System.Data.DataTable)

' Ensure Status column exists
If Not dtReportData.Columns.Contains("Status") Then
    dtReportData.Columns.Add("Status", GetType(String))
End If

Dim colName As String = "Account Name 1"
Dim colAccNum As String = "Account Number"

' Map: normalized account name -> distinct account numbers
Dim keyToAccNums As New System.Collections.Generic.Dictionary(Of String, System.Collections.Generic.HashSet(Of String))()

' ---------- Pass 1: build dictionary ----------
For Each r As System.Data.DataRow In dtReportData.Rows

    Dim nameRaw As String = If(r.IsNull(colName), "", r(colName).ToString())
    Dim accRaw As String = If(r.IsNull(colAccNum), "", r(colAccNum).ToString())

    Dim nameNorm As String =
        System.Text.RegularExpressions.Regex.Replace(nameRaw.ToLowerInvariant(), "[^a-z0-9]", "")

    Dim acc As String = accRaw.Trim()

    If Not keyToAccNums.ContainsKey(nameNorm) Then
        keyToAccNums(nameNorm) = New System.Collections.Generic.HashSet(Of String)(System.StringComparer.OrdinalIgnoreCase)
    End If

    keyToAccNums(nameNorm).Add(acc)
Next

' ---------- Pass 2: flag rows where same name maps to 2+ different account numbers ----------
For Each r As System.Data.DataRow In dtReportData.Rows

    Dim nameRaw As String = If(r.IsNull(colName), "", r(colName).ToString())
    Dim nameNorm As String =
        System.Text.RegularExpressions.Regex.Replace(nameRaw.ToLowerInvariant(), "[^a-z0-9]", "")

    If keyToAccNums.ContainsKey(nameNorm) AndAlso keyToAccNums(nameNorm).Count >= 2 Then
        r("Status") = "Analysis Required"
    End If
Next
