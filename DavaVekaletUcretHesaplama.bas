Attribute VB_Name = "DavaVekaletUcretHesaplama"
' Bu fonksiyon, verilen as�l alacak ve mahkeme t�r�ne g�re vekalet �cretini hesaplar.
' Girdi Parametreleri:
'   - AsilAlacak: Vekalet �cretinin hesaplanaca�� as�l alacak tutar� (Double).
'   - MahkemeTuru: "Asliye" veya "T�ketici" olarak belirtilen mahkeme t�r� (String).
' ��kt�:
'   - Vekalet �creti hesaplanm�� de�er (Double) veya bir hata mesaj� (String).
Function VEKALETUCRETHESAPLA2024(AsilAlacak As Double, MahkemeTuru As String) As Variant
Attribute VEKALETUCRETHESAPLA2024.VB_Description = "Bu fonksiyon, verilen as�l alacak ve mahkeme t�r�ne g�re vekalet �cretini hesaplar. Parametreler: AsilAlacak (Double), MahkemeTuru ('Asliye' veya 'Tuketici'). Av. �brahim SANDIKCI"
Attribute VEKALETUCRETHESAPLA2024.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Tablo As Variant
    Dim Tutar As Double
    Dim VekaletUcreti As Double
    Dim KalanTutar As Double
    Dim i As Integer

    ' Tabloyu array olarak tan�mla
    Tablo = Array( _
        Array(400000, 0.16), _
        Array(400000, 0.15), _
        Array(800000, 0.14), _
        Array(1200000, 0.11), _
        Array(1600000, 0.08), _
        Array(2000000, 0.05), _
        Array(2400000, 0.03), _
        Array(2800000, 0.02), _
        Array(11600000, 0.01) _
    )

    ' Mahkeme t�r� kontrol�
    MahkemeTuru = LCase(MahkemeTuru)
    If InStr(MahkemeTuru, "asliye") > 0 Then
        MahkemeTuru = "asliye"
    ElseIf InStr(MahkemeTuru, "t�ketici") > 0 Then
        MahkemeTuru = "t�ketici"
    Else
        VEKALETUCRETHESAPLA2024 = "Ge�ersiz mahkeme t�r�. L�tfen 'Asliye' veya 'T�ketici' olarak girin."
        Exit Function
    End If

    ' Vekalet �cretini hesapla
    VekaletUcreti = 0
    KalanTutar = AsilAlacak

    For i = LBound(Tablo) To UBound(Tablo)
        If KalanTutar > Tablo(i)(0) Then
            Tutar = Tablo(i)(0)
        Else
            Tutar = KalanTutar
        End If

        VekaletUcreti = VekaletUcreti + (Tutar * Tablo(i)(1))
        KalanTutar = KalanTutar - Tutar

        If KalanTutar <= 0 Then Exit For
    Next i

    ' �zel ko�ullar
    If VekaletUcreti > AsilAlacak Then VekaletUcreti = AsilAlacak

    If MahkemeTuru = "asliye" And VekaletUcreti < 30000 Then
        VekaletUcreti = 30000
    ElseIf MahkemeTuru = "t�ketici" And VekaletUcreti < 15000 Then
        VekaletUcreti = 15000
    End If

    VEKALETUCRETHESAPLA2024 = VekaletUcreti
End Function


Sub FonksiyonYardimiEkle()
    Application.MacroOptions _
        Macro:="VEKALETUCRETHESAPLA2024", _
        Description:="Bu fonksiyon, verilen as�l alacak ve mahkeme t�r�ne g�re vekalet �cretini hesaplar. Parametreler: AsilAlacak (Double), MahkemeTuru ('Asliye' veya 'Tuketici'). Av. �brahim SANDIKCI", _
        ArgumentDescriptions:=Array( _
            "Vekalet �cretinin hesaplanaca�� as�l alacak tutar� (�r. 900000).", _
            "Mahkeme t�r� ('Asliye' veya 'T�ketici'). �uan sadece bu iki mahkeme t�r�ne g�re hesaplamalar yap�labilmektedir." _
        )
End Sub

