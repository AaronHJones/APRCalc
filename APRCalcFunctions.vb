Option Explicit
Function Test_APR_API()
    Dim objRequest As Object
    Dim StrUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    
    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    StrUrl = "https://localhost:7298/api/Calculators"
    blnAsync = True
    
    With objRequest
        .Open "Get", StrUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .ResponseText
    End With
    
    Debug.Print strResponse
End Function
Function MonthlyPymt(ByVal Principal As Double, ByVal Interest As Double, ByVal Term As Integer) As Double
    Dim M, P, i As Double
    Dim n As Integer
    P = Principal
    i = (Interest / 12)
    n = Term
    M = P * (i * (1 + i) ^ n) / ((1 + i) ^ n - 1)
    MonthlyPymt = M
End Function
Function BalloonVal(ByVal PresentValue As Double, ByVal Payment As Double, ByVal Rate As Double, ByVal Term As Integer)
    Dim FV, PV, P, r As Double
    Dim n As Integer
    PV = PresentValue
    P = Round(Payment * 100, 0) / 100
    r = Rate / 12
    n = Term
    FV = (1 + r) ^ n
    FV = PV * FV
    FV = FV - (P * ((((1 + r) ^ n) - 1) / r))
    FV = FV + P
    BalloonVal = FV
End Function
Function GetAPR(ByVal Principal As Double, ByVal Fees As Double, ByVal Rate As Double, ByVal Term As Integer)
    Dim dblAmortPymt, dblActualPymt, dblAPR As Double
    dblAmortPymt = MonthlyPymt((Principal + Fees), Rate, Term)
    dblActualPymt = MonthlyPymt(Principal, Rate, Term)
    dblAPR = Financial.Rate(Term, -dblAmortPymt, Principal)
    GetAPR = dblAPR
End Function
