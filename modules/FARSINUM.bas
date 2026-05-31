Attribute VB_Name = "FARSINUM"
Option Compare Database
Option Explicit


Function ALPHANUM(AANUMBER As Double) As String
 Dim AECHO, ATALAF, STRNUM, MATANUM(10, 3), BNBIST(9), MEGH(3), AAS(4), VA(2), ACOL1, ACOL2, ACOL3, ACOL As String, K, i, ALEN, SETAYEE, j, AM2, AM3, AM, ALINK, ALENT As Integer
                  On Error Resume Next
    AANUMBER = Round(AANUMBER)
    VA(1) = ""
    VA(2) = " و "
    AAS(1) = " هزار"
    AAS(2) = " ميليون"
    AAS(3) = " ميليارد"
    AAS(4) = " ترليون"
    ACOL1 = "    يك  دو  سه  چهارپنج شش  هفت هشت نه"
    K = 1
    For i = 1 To 10
        MATANUM(i, 1) = Trim(MID(ACOL1, K, 4))
        K = K + 4
    Next i
    ACOL2 = "     ده   بيست سي   چهل  پنجاهشصت  هفتادهشتادنود  "
    K = 1
    For i = 1 To 10
        MATANUM(i, 2) = Trim(MID(ACOL2, K, 5))
        K = K + 5
    Next i
    ACOL3 = "      صد    دويست سيصد  چهارصدپانصد ششصد  هفتصد هشتصد نهصد  "
    K = 1
    For i = 1 To 10
        MATANUM(i, 3) = Trim(MID(ACOL3, K, 6))
        K = K + 6
    Next i
    ACOL = "يازده دوازدهسيزده چهاردهپانزدهشانزدههفده  هيجده نوزده "
    K = 1
    For i = 1 To 9
        BNBIST(i) = Trim(MID(ACOL, K, 6))
        K = K + 6
    Next i
    STRNUM = Trim(STR(AANUMBER))
    ATALAF = ""
    ALEN = Len(STRNUM)
    SETAYEE = 0
    K = 0
    Do While K < ALEN
        AECHO = ""
        j = 0
        For i = 1 To 3
            MEGH(i) = ""
        Next i
        Do While j < 3 And j < ALEN
            If ALEN <= K Then
              MEGH(j + 1) = ""
            Else
              MEGH(j + 1) = MID(STRNUM, ALEN - K, 1)
            End If
            j = j + 1
            K = K + 1
        Loop
        If (Val(MEGH(3)) > 0 And Val(MEGH(1)) > 0) Or (Val(MEGH(3)) > 0 And Val(MEGH(2)) > 0) Then
            AM3 = 2
        Else
            AM3 = 1
        End If
        If Val(MEGH(2)) > 0 And Val(MEGH(1)) > 0 Then
            AM2 = 2
        Else
            AM2 = 1
        End If
        SETAYEE = SETAYEE + 1
        AM = Val(MEGH(2) + MEGH(1))
        If AM >= 11 And AM <= 19 Then
            AECHO = MATANUM(Val(MEGH(3)) + 1, 3) + VA(AM3) + BNBIST(AM - 10)
        Else
            AECHO = MATANUM(Val(MEGH(3)) + 1, 3) + VA(AM3) + MATANUM(Val(MEGH(2)) + 1, 2) + VA(AM2) + MATANUM(Val(MEGH(1)) + 1, 1)
        End If
        ALINK = 1
        If SETAYEE > 1 And Not (MEGH(1) = "0" And MEGH(2) = "0" And MEGH(3) = "0") Then
            AECHO = AECHO + AAS(SETAYEE - 1)
            ALINK = 2
        End If
        ATALAF = AECHO + VA(ALINK) + ATALAF
    Loop
    ALENT = Len(Trim(ATALAF))
    If MID(Trim(ATALAF), ALENT - 1, 2) = " و" Then
        ALPHANUM = (MID(Trim(ATALAF), 1, ALENT - 1))
    Else
        ALPHANUM = (Trim(ATALAF))
    End If
End Function

Public Function UCurrentUser()
    UCurrentUser = Forms![BASEKNOW]![UUSER]
End Function


Public Function UCurrentUserIDD()
    UCurrentUserIDD = Forms![BASEKNOW]![USERCOD]
End Function
