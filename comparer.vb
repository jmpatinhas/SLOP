Imports System.Text.RegularExpressions
Imports System.Collections.Generic

' Ensure Status column exists
If Not dt.Columns.Contains("Status") Then
    dt.Columns.Add("Status", GetType(String))
End If

Dim colName As String = "Account Name 1"
Dim colAddr As String = "Additional Addresses Address 1"
Dim colAccNum As String = "Account Account Number"

' key -> distinct account numbers
Dim keyToAccNums As New Dictionary(Of String, HashSet(Of String))()

' ---------- Pass 1: build dictionary of keys and distinct account numbers ----------
For Each r As DataRow In dt.Rows

    Dim nameRaw As String = If(r.IsNull(colName), "", r(colName).ToString())
    Dim addrRaw As String = If(r.IsNull(colAddr), "", r(colAddr).ToString())
    Dim accRaw As String = If(r.IsNull(colAccNum), "", r(colAccNum).ToString())

    Dim nameNorm As String = Regex.Replace(nameRaw.ToLowerInvariant(), "[^a-z0-9]", "")
    Dim addrNorm As String = Regex.Replace(addrRaw.ToLowerInvariant(), "[^a-z0-9]", "")
    Dim key As String = nameNorm & "|" & addrNorm

    Dim acc As String = accRaw.Trim()

    If Not keyToAccNums.ContainsKey(key) Then
        keyToAccNums(key) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    End If

    keyToAccNums(key).Add(acc)
Next

' ---------- Pass 2: set Status for rows in groups with 2+ distinct account numbers ----------
For Each r As DataRow In dt.Rows

    Dim nameRaw As String = If(r.IsNull(colName), "", r(colName).ToString())
    Dim addrRaw As String = If(r.IsNull(colAddr), "", r(colAddr).ToString())

    Dim nameNorm As String = Regex.Replace(nameRaw.ToLowerInvariant(), "[^a-z0-9]", "")
    Dim addrNorm As String = Regex.Replace(addrRaw.ToLowerInvariant(), "[^a-z0-9]", "")
    Dim key As String = nameNorm & "|" & addrNorm

    If keyToAccNums.ContainsKey(key) AndAlso keyToAccNums(key).Count >= 2 Then
        r("Status") = "Analysis Required"
    End If
Next
