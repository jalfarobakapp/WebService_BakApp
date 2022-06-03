Declare @Idmaeedo Int = #Idmaeedo#


Select TIDO,NUDO,KOPRCT,NOKOPR,KOLTPR,Case UDTRPR When 1 Then CAPRCO1 Else CAPRCO2 End As CANTIDAD,PPPRNELT,
       (Select Top 1 Case UDTRPR When 1 Then PP01UD Else PP02UD End From TABPRE Where KOLT = SUBSTRING(KOLTPR,6,3) And KOPR = KOPRCT) As PRECIO_ACT,
	   Cast(0 As Float) As Diferencia,(Select Top 1 POIVPR From MAEPR Where KOPR = KOPRCT) As POIVPR,PODTGLLI,VADTNELI,VADTBRLI,VANELI,VAIVLI,VABRLI,
	   Cast(0 As Float) As New_VANELI,
	   Cast(0 As Float) As New_VADTNELI,
	   Cast(0 As Float) As New_VADTBRLI,
	   Cast(0 As Float) As New_VAIVLI,
	   Cast(0 As Float) As New_VABRLI,
	   Cast(0 As Float) As DIFERENCIA

Into #Paso
From MAEDDO
WHERE IDMAEEDO = @Idmaeedo

Update #Paso Set Diferencia = PRECIO_ACT - PPPRNELT, 
                 New_VADTNELI = ROUND((CANTIDAD * PRECIO_ACT) * PODTGLLI/100,0)

Update #Paso Set New_VANELI = (CANTIDAD * PRECIO_ACT) - New_VADTNELI
Update #Paso Set New_VAIVLI = Round(New_VANELI*(POIVPR/100),2)
Update #Paso Set New_VABRLI = Round( New_VANELI+New_VAIVLI,0)
Update #Paso Set DIFERENCIA = New_VABRLI - VABRLI

Select Sum(VANELI) As VANEDO,Sum(New_VANELI) As New_VANEDO, Sum(VABRLI) As VABRLI,
       Sum(New_VABRLI) As New_VABRLI,Sum(New_VABRLI-VABRLI) As Diferencia
From #Paso
Select * From #Paso
Drop Table #Paso 