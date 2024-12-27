Attribute VB_Name = "DavaVekaletUcretHesaplama"
' Bu fonksiyon, verilen asýl alacak ve mahkeme türüne göre vekalet ücretini hesaplar.
' Girdi Parametreleri:
'   - AsilAlacak: Vekalet ücretinin hesaplanacaðý asýl alacak tutarý (Double).
'   - MahkemeTuru: "Asliye" veya "Tüketici" olarak belirtilen mahkeme türü (String).
' Çýktý:
'   - Vekalet ücreti hesaplanmýþ deðer (Double) veya bir hata mesajý (String).
Function VEKALETUCRETHESAPLA2024(AsilAlacak As Double, MahkemeTuru As String) As Variant
Attribute VEKALETUCRETHESAPLA2024.VB_Description = "Bu fonksiyon, verilen asýl alacak ve mahkeme türüne göre vekalet ücretini hesaplar. Parametreler: AsilAlacak (Double), MahkemeTuru ('Asliye' veya 'Tuketici'). Av. Ýbrahim SANDIKCI"
Attribute VEKALETUCRETHESAPLA2024.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim Tablo As Variant
    Dim Tutar As Double
    Dim VekaletUcreti As Double
    Dim KalanTutar As Double
    Dim i As Integer

    ' Tabloyu array olarak tanýmla
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

    ' Mahkeme türü kontrolü
    MahkemeTuru = LCase(MahkemeTuru)
    If InStr(MahkemeTuru, "asliye") > 0 Then
        MahkemeTuru = "asliye"
    ElseIf InStr(MahkemeTuru, "tüketici") > 0 Then
        MahkemeTuru = "tüketici"
    Else
        VEKALETUCRETHESAPLA2024 = "Geçersiz mahkeme türü. Lütfen 'Asliye' veya 'Tüketici' olarak girin."
        Exit Function
    End If

    ' Vekalet ücretini hesapla
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

    ' Özel koþullar
    If VekaletUcreti > AsilAlacak Then VekaletUcreti = AsilAlacak

    If MahkemeTuru = "asliye" And VekaletUcreti < 30000 Then
        VekaletUcreti = 30000
    ElseIf MahkemeTuru = "tüketici" And VekaletUcreti < 15000 Then
        VekaletUcreti = 15000
    End If

    VEKALETUCRETHESAPLA2024 = VekaletUcreti
End Function


Sub FonksiyonYardimiEkle()
    Application.MacroOptions _
        Macro:="VEKALETUCRETHESAPLA2024", _
        Description:="Bu fonksiyon, verilen asýl alacak ve mahkeme türüne göre vekalet ücretini hesaplar. Parametreler: AsilAlacak (Double), MahkemeTuru ('Asliye' veya 'Tuketici'). Av. Ýbrahim SANDIKCI", _
        ArgumentDescriptions:=Array( _
            "Vekalet ücretinin hesaplanacaðý asýl alacak tutarý (ör. 900000).", _
            "Mahkeme türü ('Asliye' veya 'Tüketici'). Þuan sadece bu iki mahkeme türüne göre hesaplamalar yapýlabilmektedir." _
        )
End Sub

