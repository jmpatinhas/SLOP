Imports System.Text.RegularExpressions
Imports System.Linq

'--- helper: normalize for stable matching
Private Function Norm(s As String) As String
    If s Is Nothing Then Return ""
    s = s.ToLowerInvariant()
    ' remove all non letters/digits (kills spaces + punctuation + accents are kept as-is)
    s = Regex.Replace(s, "[^a-z0-9]", "")
    Return s
End Function

'--- make sure Status column exists
If Not dt.Columns.Contains("Status") Then
    dt.Columns.Add("Status", GetType(String))
End If

'--- columns (adjust if your headers differ)
Dim colName As String = "Account Name 1"
Dim colAddr As String = "Additional Addresses Address 1"
Dim colAccNum As String = "Account Account Number"

'--- key -> set of account numbers seen for that key
Dim keyToAccNums As New Dictionary(Of String, HashSet(Of String))()

' Pass 1: build groups with distinct account numbers
For Each r As DataRow In dt.Rows
    Dim key As String = Norm(r(colName).ToString()) & "|" & Norm(r(colAddr).ToString())
    Dim acc As String = r(colAccNum).ToString().Trim()

    If Not keyToAccNums.ContainsKey(key) Then
        keyToAccNums(key) = New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
    End If

    keyToAccNums(key).Add(acc)
Next

' Pass 2: flag rows whose group has 2+ distinct account numbers
For Each r As DataRow In dt.Rows
    Dim key As String = Norm(r(colName).ToString()) & "|" & Norm(r(colAddr).ToString())
    If keyToAccNums.ContainsKey(key) AndAlso keyToAccNums(key).Count >= 2 Then
        r("Status") = "Analysis Required"
    End If
Next
