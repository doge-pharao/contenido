REM  *****  BASIC  *****

Sub Main

End Sub


'Función para sumar los valores de un rango
Function SumarRango( tiempoBit, Rango ) As Double 
	Dim oHojaActiva as Object
	Dim Wm_n, Jk, Tk, Ck, Cm, Bm as Object 
	Dim Wm_sig_n, serieTmp, ceilingValue As Double
	Dim co3 As Long, co2 As Long
	Dim args( 1 to 2 ) As Variant
	
	service = createUnoService("com.sun.star.sheet.FunctionAccess")
	
	oHojaActiva = ThisComponent.getCurrentController.getActiveSheet()
	co3 = UBound( Rango,1 )

	Bm = oHojaActiva.getCellByPosition( 11, co3+1)	'L(x-1) Bm
	Wm_n = oHojaActiva.getCellByPosition( 12, co3) 	'M(x-1)	Wm
	Wm_sig_n = Bm.Value								'Wn+1= Bm +...

	serieTmp = 0
	For co2 = UBound( Rango,1)  To LBound(Rango, 1) Step -1			'Todos de prioridad mayor a mi
	    Jk = oHojaActiva.getCellByPosition( 3, co2)			'D(x) J/ms
	    Tk = oHojaActiva.getCellByPosition( 4, co2)			'E(x) T/ms
	    Ck = oHojaActiva.getCellByPosition( 10, co2)			'K(x) Ck
	   
	   	args(1) = ((Wm_n.Value + Jk.Value + tiempoBit) / Tk.Value)
	  	args(2) = 1
	   	ceilingValue = service.callFunction( "CEILING", args() )
	   
	   	serieTmp = serieTmp + ceilingValue * Ck.Value
	     ' MsgBox "co2:" + co2 + "co3:" + co3
	Next co2
	
	'Asignamos el resultado
	SumarRango =  Wm_sig_n + serieTmp
End Function
