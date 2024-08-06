Declare @Idmaeedo As int

Set @Idmaeedo = #Idmaeedo#

--UPDATE MAEEDO SET RUTCONTACT = '13239873' WHERE IDMAEEDO = @Idmaeedo

SELECT 
       -- MAEEDO 
       Edo.IDMAEEDO, Edo.EMPRESA, Edo.TIDO, Edo.NUDO, Edo.TIDO+Edo.NUDO As 'Bk_Tido_Nudo', Edo.ENDO, Edo.SUENDO, Edo.ENDOFI, Edo.TIGEDO, Edo.SUDO,Sucedo.NOKOSU As 'NOMSUDO',
	   Edo.LUVTDO, Edo.FEEMDO, Edo.KOFUDO, Resp.NOKOFU As 'Resp_NOKOFU', Edo.ESDO, Edo.ESPGDO, 
       Edo.CAPRCO, Edo.CAPRAD, Edo.CAPREX, Edo.CAPRNC, Edo.MEARDO, Edo.MODO, Edo.TIMODO, Edo.TAMODO, Edo.NUCTAP, Edo.VACTDTNEDO, Edo.VACTDTBRDO, Edo.NUIVDO, Edo.POIVDO, 
       Edo.VAIVDO, Edo.NUIMDO, Edo.VAIMDO, Edo.VANEDO, Edo.VABRDO, Edo.POPIDO, Edo.VAPIDO, Edo.FE01VEDO, Edo.FEULVEDO, Edo.NUVEDO, Edo.VAABDO, Edo.MARCA, Edo.FEER, 
       Edo.NUTRANSMI, Edo.NUCOCO, Edo.KOTU, Edo.LIBRO, Edo.LCLV, Edo.ESFADO, Edo.KOTRPCVH, Edo.NULICO, Edo.PERIODO, Edo.NUDONODEFI, Edo.TRANSMASI, Edo.POIVARET, Edo.VAIVARET, 
       Edo.RESUMEN, Edo.LAHORA, Edo.KOFUAUDO, Edo.KOOPDO, Edo.ESPRODDO, Edo.DESPACHO, Edo.HORAGRAB, Edo.RUTCONTACT, Edo.SUBTIDO, Edo.TIDOELEC, Edo.ESDOIMP, 
       Edo.CUOGASDIF, Edo.BODESTI, Edo.PROYECTO, Edo.FECHATRIB, Edo.NUMOPERVEN, Edo.BLOQUEAPAG, Edo.VALORRET, Edo.FLIQUIFCV, Edo.VADEIVDO, Edo.KOCANAL, Edo.KOCRYPT, 
       -- MAEEN
       Een.IDMAEEN, Een.KOEN, Een.TIEN, Een.RTEN, Een.SUEN, Een.TIPOSUC, Een.NOKOEN, Een.SIEN, Een.GIEN, Een.PAEN, Een.CIEN, Een.CMEN, Een.DIEN, Een.ZOEN, Een.FOEN, Een.FAEN, 
       Een.CNEN, Een.KOFUEN,Ven.NOKOFU As 'NOMVENDEDOR', Een.LCEN, Een.LVEN, Een.CRSD, Een.CRCH, Een.CRLT, Een.CRPA, Een.CRTO, Een.CREN, Een.FEVECREN, Een.FEULTR, Een.NUVECR, Een.DCCR, Een.INCR, Een.POPICR, 
       Een.KOPLCR, Een.CONTAB, Een.SUBAUXI, Een.CONTABVTA, Een.SUBAUXIVTA, Een.CODCC, Een.NUTRANSMI AS Expr1, Een.RUEN, Een.CPEN, Een.OBEN, Een.DIPRVE, Een.EMAIL, Een.CNEN2, 
       Een.COBRADOR, Een.PROTEACUM, Een.PROTEVIGE, Een.CPOSTAL, Een.HABILITA, Een.CODCONVE, Een.NOTRAEDEUD, Een.NOKOENAMP, Een.BLOQUEADO, Een.DIMOPER, Een.PREFEN, 
       Een.BLOQENCOM, Een.TIPOEN, Een.ACTIEN, Een.TAMAEN, Een.PORPREFEN, Een.CLAVEEN, Een.NVVPIDEPIE, Een.RECEPELECT, Een.ACTECO, Een.DIASVENCI, Een.CATTRIB, Een.AGRETIVA, 
       Een.AGRETIIBB, Een.AGRETGAN, Een.AGPERIVA, Een.AGPERIIBB, Een.TRANSPOEN, Een.FECREEN, Een.FIRMA, Een.MOCTAEN, Een.CTASDELAEN, Een.NACIONEN, Een.DIRPAREN, Een.FECNACEN,
       Een.ESTCIVEN, Een.PROFECEN, Een.CONYUGEN, Een.RUTCONEN, Een.RUTSOCEN, Een.SEXOEN, Een.RELACIEN, Een.ANEXEN1, Een.ANEXEN2, Een.ANEXEN3, Een.ANEXEN4, Een.OCCOBLI, 
       Een.VALIVENPAG, Een.EMAILCOMER, Een.TIPOCONTR, Een.FEREFAUTO, Een.DIACOBRA, Een.CUENTABCO, Een.KOENDPEN, Een.SUENDPEN, Een.RUTALUN, Een.RUTAMAR, Een.RUTAMIE, 
       Een.RUTAJUE, Een.RUTAVIE, Een.RUTASAB, Een.RUTADOM, Een.CATLEGRET, Een.IMPTORET, Een.ENTILIGA, Een.PORCELIGA, Een.ACTECOBCO, Obs.IDMAEEDO AS IDMAEEDO_Obs, Obs.OBDO, Obs.CPDO, 
      -- MAEEDOOB
       Obs.OCDO, Obs.DIENDESP, Obs.TEXTO1, Obs.TEXTO2, Obs.TEXTO3, Obs.TEXTO4, Obs.TEXTO5, Obs.TEXTO6, Obs.TEXTO7, Obs.TEXTO8, Obs.TEXTO9, Obs.TEXTO10, Obs.TEXTO11, 
       Obs.TEXTO12, Obs.TEXTO13, Obs.TEXTO14, Obs.TEXTO15, Obs.CARRIER, Obs.BOOKING, Obs.LADING, Obs.AGENTE, Obs.MEDIOPAGO, Obs.TIPOTRANS, Obs.KOPAE, Obs.KOCIE, Obs.KOCME, 
       Obs.FECHAE, Obs.HORAE, Obs.KOPAD, Obs.KOCID, Obs.KOCMD, Obs.FECHAD, Obs.HORAD, Obs.OBDOEXPO, Obs.MOTIVO,Isnull(Motncv.NOKOCARAC,'') As 'Motncv_NOKOCARAC', 
	   Obs.TEXTO16, Obs.TEXTO17, Obs.TEXTO18, Obs.TEXTO19, 
       Obs.TEXTO20, Obs.TEXTO21, Obs.TEXTO22, Obs.TEXTO23, Obs.TEXTO24, Obs.TEXTO25, Obs.TEXTO26, Obs.TEXTO27, Obs.TEXTO28, Obs.TEXTO29, Obs.TEXTO30, Obs.TEXTO31, Obs.TEXTO32, 
       Obs.TEXTO33, Obs.TEXTO34, Obs.TEXTO35, Obs.PLACAPAT, 
        
        -- MAEENCON
       Isnull(Cont.KOEN,'')       AS 'KOEN_Obs', 
       Isnull(Cont.RUTCONTACT,'') AS 'RUTCONTACT_Obs', 
       Isnull(Cont.SUTRABCON,'')  AS 'SUTRABCON', 
       Isnull(Cont.CLAVECON,'')   AS 'CLAVECON', 
       Isnull(Cont.NOKOCON,'')    AS 'NOKOCON',  
       Isnull(Cont.FONOCON,'')    AS 'FONOCON', 
       Isnull(Cont.FAXCON,'')     AS 'FAXCON',  
       Isnull(Cont.EMAILCON,'')   AS 'EMAILCON',  
       Isnull(Cont.CARGOCON,'')   AS 'CARGOCON',  
       Isnull(Cont.AREACON,'')    AS 'AREACON',  
       Isnull(Cont.AUTORIZADO,'') AS 'AUTORIZADO',  
       Isnull(Cont.DIRECON,'')    AS 'DIRECON', 
       
       -- MAEENCON EN DOCUMENTO
       Isnull(Cont_Ent.KOEN,'')       AS 'Cont_KOEN_Obs', 
       Isnull(Cont_Ent.RUTCONTACT,'') AS 'Cont_RUTCONTACT_Obs', 
       Isnull(Cont_Ent.SUTRABCON,'')  AS 'Cont_SUTRABCON', 
	   Isnull(Cont_Ent.CLAVECON,'')   AS 'Cont_CLAVECON',  
	   Isnull(Cont_Ent.NOKOCON,'')    AS 'Cont_NOKOCON',  
	   Isnull(Cont_Ent.FONOCON,'')    AS 'Cont_FONOCON',  
	   Isnull(Cont_Ent.FAXCON,'')     AS 'Cont_FAXCON',  
       Isnull(Cont_Ent.EMAILCON,'')   AS 'Cont_EMAILCON',  
       Isnull(Cont_Ent.CARGOCON,'')   AS 'Cont_CARGOCON',  
       Isnull(Cont_Ent.AREACON,'')    AS 'Cont_AREACON',  
       Isnull(Cont_Ent.AUTORIZADO,'') AS 'Cont_AUTORIZADO',  
	   Isnull(Cont_Ent.DIRECON,'')    AS 'Cont_DIRECON', 
       
       -- TABRETI Retirador de mercaderia
	   Isnull(Ret.KORETI,'')     As 'Reti_KORETI',
	   Isnull(Ret.RURETI,'')     As 'Reti_RURETI',
	   Isnull(Ret.NORETI,'')     As 'Reti_NORETI',
	   Isnull(Ret.DIRETI,'')     As 'Reti_DIRETI',
	   Isnull(Ret.PARETI,'')     As 'Reti_PARETI',
	   Isnull(Ret.CIRETI,'')     As 'Reti_CIRETI',
	   Isnull(Ret.CMRETI,'')     As 'Reti_CMRETI',
	   Isnull(Ret.RETCLI,'')     As 'Reti_RETCLI',
	   Isnull(Ret.KOENRESP,'')   As 'Reti_KOENRESP',
	   Isnull(Ret.SUENRESP,'')   As 'Reti_SUENRESP',
	   Isnull(Ret.TIPORUC,'')    As 'Reti_TIPORUC',
	   Isnull(Ret.LICENCONDU,'') As 'Reti_LICENCONDU',          
	                
       -- TABPLACA 
	   Isnull(Placa.PLACA,'')    As 'Placa_PLACA',
	   Isnull(Placa.DESCRIP,'')  As 'Placa_DESCRIP',
	   Isnull(Placa.KOENRESP,'') As 'Placa_KOENRESP',
	   Isnull(Placa.SUENRESP,'') As 'Placa_SUENRESP',
	   Isnull(Placa.MARCA,'')    As 'Placa_MARCA',
	   Isnull(Placa.MODELO,'')   As 'Placa_MODELO',
	   Isnull(Placa.ANNO,'')     As 'Placa_ANNO',
	   Isnull(Placa.PADRON,'')   As 'Placa_PADRON',
	   Isnull(Placa2.NOKOEN,'')  As 'Placa_NOKOEN',

       Isnull((Select top 1 NOTIDO From TABTIDO Where TABTIDO.TIDO = Edo.TIDO),'') As 'Bk_Nombre_Documento',

       CAST(0 As Float) As 'Bk_Suma_CantUd1',
	   CAST(0 As Float) As 'Bk_Suma_CantUd2',
	   CAST(0 As Float) As 'Bk_Suma_Reca_Netos',
	   CAST(0 As Float) As 'Bk_Suma_Desc_Netos',
	   CAST(0 As Float) As 'Bk_Suma_Reca_Brutos',
	   CAST(0 As Float) As 'Bk_Suma_Desc_Brutos',
       CAST(0 As Float) As 'Bk_Total_Items',
       CAST(0 As Float) As 'Bk_Sub_Total_Neto',
       CAST(0 As Float) As 'Bk_Sub_Total_Bruto', 
       CAST(0 As Float) As 'Bk_Suma_Dscto_Neto',
       CAST(0 As Float) As 'Bk_Suma_Dscto_Bruto', 
       Isnull((Select top 1 NOKOSU From TABSU Where TABSU.EMPRESA = Edo.EMPRESA And KOSU = Edo.SUDO),'') 
       As 'Bk_Sucursal_Doc_Nombre',  
       Isnull((Select top 1 NOKOFU From TABFU Where KOFU = Edo.KOFUDO),'') 
       As 'Bk_Nom_Responzable',  
       Isnull((Select top 1 NOKOFU From TABFU Where KOFU = Een.KOFUEN),'') 
       As 'Bk_Nom_Vendedor',
       Isnull((Select top 1 NOKOFU From TABFU Where KOFU = Een.COBRADOR),'') 
       As 'Bk_Nom_Cobrador',
       Case Een.TIPOSUC When 'C' Then 'CLIENTE' When 'P' Then 'PROVEEDOR' Else 'AMBOS' End 
       As 'Bk_Tipo_CliProAmb',
       Isnull((Select Top 1 NOKOPA From TABPA Where KOPA = Een.PAEN),'') 
       As 'Bk_Pais', 
       Isnull((Select Top 1 NOKOCI From TABCI Where KOPA = Een.PAEN And KOCI = Een.CIEN),'') 
       As 'Bk_Ciudad', 
       Isnull((Select Top 1 NOKOCM From TABCM Where KOPA = Een.PAEN And KOCI = Een.CIEN And KOCM = Een.CMEN),'') 
       As 'Bk_Comuna',
       Isnull((Select top 1 NOKOZO From TABZO Where KOZO = Een.ZOEN),'') 
       As 'Bk_Zona',
       Isnull((Select top 1 NOKORU From TABRU Where KORU = Een.RUEN),'') 
       As 'Bk_Rubro',
       Isnull((Select top 1 NOKOCARAC From TABCARAC Where KOTABLA = 'ACTIVIDADE' And KOCARAC = Een.ACTIEN),'') 
	   As 'Bk_Activ_Economica',
       Isnull((Select top 1 NOKOCARAC From TABCARAC Where KOTABLA = 'TAMA¥OEMPR' And KOCARAC = Een.TAMAEN),'') 
       As 'Bk_Tam_Empresa',
       Isnull((Select top 1 NOKOCARAC From TABCARAC	Where KOTABLA = 'TIPOENTIDA' And KOCARAC = Een.TIPOEN),'') 
       As 'Bk_Tipo_Entidad',
       convert(nvarchar, convert(datetime, (Edo.HORAGRAB*1.0/3600)/24), 108)
       As 'Bk_Hora_Emision',
       Substring(convert(varchar,Edo.LAHORA,8),1,5) 
	   As 'Bk_Hora_Emision2',
       -- Funciones especiales de BakApp
       CAST('' As Varchar(20))  As 'Bk_Rut', 
       CAST('' As Varchar(200)) As 'Bk_T_Escrito_1_Bruto',
       CAST('' As Varchar(200)) As 'Bk_T_Escrito_2_Bruto',
       CAST('' As Varchar(200)) As 'Bk_Caja_Mod_Codigo', 
       CAST('' As Varchar(200)) As 'Bk_Caja_Mod_Nombre',
       CAST('' As Varchar(200)) As 'Bk_Sucursal_Mod_Codigo',
       CAST('' As Varchar(200)) As 'Bk_Sucursal_Mod_Nombre',
       CAST(0 As Float) As 'Bk_Recargo',
	   Case Een.PREFEN When 1 Then '*** ENTIDAD PREFERENCIAL' Else '' End As 'Bk_Entidad_Preferencial',
	   CAST('' As Varchar(50)) As 'Bk_Nombre_Sudo',
       CAST('' As Varchar(50)) As 'Bk_Direccion_Sudo',
	   CAST('' As Varchar(50)) As 'Bk_1er_Vendedor_Codigo',
	   CAST('' As Varchar(50)) As 'Bk_1er_Vendedor_Nombre',
       Cast(0 As Int) As 'Bk_Cuotas'

Into #Paso_Encabezado       
        
FROM  dbo.MAEEDO AS Edo 

	LEFT JOIN dbo.MAEENCON AS Cont ON Edo.RUTCONTACT = Cont.RUTCONTACT AND Edo.ENDO = Cont.KOEN 
		LEFT JOIN dbo.MAEENCON AS Cont_Ent ON Edo.ENDO = Cont_Ent.KOEN And Cont.RUTCONTACT = Edo.RUTCONTACT
			LEFT JOIN dbo.MAEEDOOB AS Obs ON Edo.IDMAEEDO = Obs.IDMAEEDO 
				LEFT JOIN dbo.MAEEN AS Een ON Edo.ENDO = Een.KOEN AND Edo.SUENDO = Een.SUEN
					LEFT JOIN dbo.TABFU As Ven On Een.KOFUEN = Ven.KOFU
						LEFT JOIN dbo.TABSU Sucedo On Edo.EMPRESA = Sucedo.EMPRESA And Edo.SUDO = Sucedo.KOSU
							LEFT JOIN dbo.TABRETI Ret On Ret.KORETI = Obs.DIENDESP 
								LEFT JOIN dbo.TABPLACA Placa On Placa.PLACA = Obs.PLACAPAT 
									LEFT JOIN dbo.MAEEN Placa2 On Placa.KOENRESP = Placa2.KOEN AND Placa.SUENRESP = Placa2.SUEN
										LEFT JOIN dbo.TABFU Resp On Resp.KOFU = Edo.KOFUDO
											LEFT JOIN dbo.TABCARAC Motncv On Motncv.KOTABLA='MOTIVOSNCV' And Motncv.KOCARAC = Obs.MOTIVO

Where Edo.IDMAEEDO = @Idmaeedo       

If (Select Top 1 DIENDESP From #Paso_Encabezado) <> '' And (Select Top 1 Reti_NORETI From #Paso_Encabezado) = ''
Begin
--Print 'Falta Retirador de Mercaderia'

Update #Paso_Encabezado Set Reti_KORETI = Isnull((Select Top 1 KORETI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_RURETI = Isnull((Select Top 1 RURETI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_NORETI = Isnull((Select Top 1 NORETI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_DIRETI = Isnull((Select Top 1 DIRETI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_PARETI = Isnull((Select Top 1 PARETI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_CIRETI = Isnull((Select Top 1 CIRETI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_CMRETI = Isnull((Select Top 1 CMRETI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_RETCLI = Isnull((Select Top 1 RETCLI From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_KOENRESP = Isnull((Select Top 1 KOENRESP From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_SUENRESP = Isnull((Select Top 1 SUENRESP From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_TIPORUC = Isnull((Select Top 1 TIPORUC From TABRETI Where KORETI Like DIENDESP+'%'),'')
Update #Paso_Encabezado Set Reti_LICENCONDU = Isnull((Select Top 1 LICENCONDU From TABRETI Where KORETI Like DIENDESP+'%'),'')
End	

       
Select Distinct
       -- MAEDDO
       Edd.IDMAEDDO, Edd.IDMAEEDO, Edd.ARCHIRST, Edd.IDRST, Edd.EMPRESA, Edd.TIDO, Edd.NUDO, Edd.ENDO, Edd.SUENDO, Edd.ENDOFI, Edd.LILG, Edd.NULIDO, Edd.SULIDO, Edd.LUVTLIDO, 
       Edd.BOSULIDO, Edd.KOFULIDO, Edd.NULILG, Edd.PRCT, Edd.TICT, Edd.TIPR, Edd.NUSEPR, Edd.KOPRCT AS 'KOPR', Edd.UDTRPR, Edd.RLUDPR, Edd.CAPRCO1, Edd.CAPRAD1, Edd.CAPREX1, Edd.CAPRNC1, 
       Edd.UD01PR, Edd.CAPRCO2, Edd.CAPRAD2, Edd.CAPREX2, Edd.CAPRNC2, Edd.UD02PR, Edd.KOLTPR, Edd.MOPPPR, Edd.TIMOPPPR, Edd.TAMOPPPR, Edd.PPPRNELT, Edd.PPPRNE, 
       Edd.PPPRBRLT, Edd.PPPRBR, Edd.NUDTLI, Edd.PODTGLLI, Edd.VADTNELI, Edd.VADTBRLI, Edd.POIVLI, Edd.VAIVLI, Edd.NUIMLI, Edd.POIMGLLI, Edd.VAIMLI, Edd.VANELI, Edd.VABRLI, Edd.TIGELI, 
       Edd.EMPREPA, Edd.TIDOPA, Edd.NUDOPA, Edd.ENDOPA, Edd.NULIDOPA, Edd.LLEVADESP, Edd.FEEMLI, Edd.FEERLI, Edd.PPPRPM, Edd.OPERACION, Edd.CODMAQ, Edd.ESLIDO, Edd.PPPRNERE1, 
       Edd.PPPRNERE2, Edd.ESFALI, Edd.CAFACO, Edd.CAFAAD, Edd.CAFAEX, Edd.CMLIDO, Edd.NULOTE, Edd.FVENLOTE, Edd.ARPROD, Edd.NULIPROD, Edd.NUCOCO, Edd.NULICO, Edd.PERIODO, 
       Edd.FCRELOTE, Edd.SUBLOTE, 
       Edd.NOKOPR AS 'NOKOPR_Edd', 
       Edd.ALTERNAT, Edd.PRENDIDO, Edd.OBSERVA, Edd.KOFUAULIDO, Edd.KOOPLIDO, Edd.MGLTPR, Edd.PPOPPR, Edd.TIPOMG, Edd.ESPRODLI, 
       Edd.CAPRODCO, Edd.CAPRODAD, Edd.CAPRODEX, Edd.CAPRODRE, Edd.TASADORIG, Edd.CUOGASDIF, Edd.SEMILLA, Edd.PROYECTO, Edd.POTENCIA, Edd.HUMEDAD, Edd.IDTABITPRE, 
       Edd.IDODDGDV, Edd.LINCONDESP, Edd.PODEIVLI, Edd.VADEIVLI, Edd.PRIIDETIQ, Edd.KOLORESCA, Edd.KOENVASE, Edd.PPPRPMSUC, Edd.PPPRPMIFRS, Mp.TIPR AS TIPR_Mp,--, Mp.KOPR, 
       Mp.NOKOPR AS 'NOKOPR_Mp', 
       Mp.KOPRRA, Mp.NOKOPRRA, Mp.KOPRTE, Mp.KOGE, Mp.NMARCA, 
       Mp.UD01PR AS 'UD01PR_Mp', 
       Mp.UD02PR AS 'UD02PR_Mp', Mp.RLUD, Mp.POIVPR, Mp.NUIMPR, Mp.RGPR, 
       Mp.STMIPR, Mp.STMAPR, Mp.MRPR, Mp.ATPR, Mp.RUPR, Mp.STFI1, Mp.STDV1, Mp.STOCNV1, Mp.STFI2, Mp.STDV2, Mp.STOCNV2, Mp.PPUL01, Mp.PPUL02, Mp.MOUL, Mp.TIMOUL, Mp.TAUL, 
       Mp.FEUL, Mp.PM, Mp.FEPM, Mp.FMPR, Mp.PFPR, Mp.HFPR, Mp.VALI, Mp.FEVALI, Mp.TTREPR, Mp.PRRG, Mp.NIPRRG, Mp.NFPRRG, Mp.PMIN, Mp.CAMICO, Mp.CAMIFA, Mp.LOMIFA, Mp.PLANO, 
       Mp.STDV1C, Mp.STOCNV1C, Mp.STDV2C, Mp.STOCNV2C, Mp.METRCO, Mp.DESPNOFAC1, Mp.DESPNOFAC2, Mp.RECENOFAC1, Mp.RECENOFAC2, Mp.TRATALOTE, Mp.DIVISIBLE, Mp.MUDEFA, 
       Mp.EXENTO, Mp.KOMODE, Mp.PRDESRES, Mp.LISCOSTO, Mp.STOCKASEG, Mp.ESACTFI, Mp.CLALIBPR, Mp.KOFUPR, Mp.KOPRDIM, Mp.NODIM1, Mp.NODIM2, Mp.NODIM3, Mp.BLOQUEAPR, 
       Mp.ZONAPR, Mp.CONUBIC, Mp.LTSUBIC, Mp.PESOUBIC, Mp.FUNCLOTE, Mp.LOMAFA, Mp.COLEGPR, Mp.MORGPR, Mp.FECRPR, Mp.FEMOPR, Mp.LOTECAJA, Mp.STTR1, Mp.STTR2, Mp.PODEIVPR, 
       Mp.DIVISIBLE2, Mp.MOVETIQ, Mp.REPUESTO, Mp.VIDAMEDIA, Mp.TRATVMEDIA, Mp.PRESALCLI1, Mp.PRESALCLI2, Mp.PRESDEPRO1, Mp.PRESDEPRO2, Mp.CONSALCLI1, Mp.CONSALCLI2, 
       Mp.CONSDEPRO1, Mp.CONSDEPRO2, Mp.DEVENGNCV1, Mp.DEVENGNCV2, Mp.DEVENGNCC1, Mp.DEVENGNCC2, Mp.DEVSINNCV1, Mp.DEVSINNCV2, Mp.DEVSINNCC1, Mp.DEVSINNCC2, 
       Mp.STENFAB1, Mp.STENFAB2, Mp.STREQFAB1, Mp.STREQFAB2, Mp.PMME, Mp.FEPMME, Mp.VALUNFLEKM, Mp.ANALIZABLE, Mp.TOLELOTE, Mp.PMIFRS, Mp.FEPMIFRS,
       Isnull(Mfch.FICHA,'') As 'MAEFICHA',
       Cast('' As Varchar(1600)) As 'MAEFICHD',
	   CASE UDTRPR WHEN 1 THEN CAPRCO1 ELSE CAPRCO2 END AS 'Bk_Cant_Trans',
	   CASE UDTRPR WHEN 1 THEN Edd.UD01PR ELSE Edd.UD02PR END AS 'Bk_Un_Trans',
	   Tbp.DATOSUBIC As 'UBICACION', 
	   Mst.STFI1 As 'StockUd1',
	   Mst.STFI2 As 'StockUd2',
	   Mst.STDV1 As 'StockdvUd1',
	   Mst.STDV2 As 'StockdvUd2',
	   Mst.STOCNV1 As 'StockncUd1',
	   Mst.STOCNV2 As 'StockncUd2',
	   CASE UDTRPR WHEN 1 THEN Mst.STFI1 ELSE Mst.STFI2 END AS 'STOCK',
	   Mrc.KOMR,
       Mrc.NOKOMR,
       --Isnull(Tcd.KOPRAL,'...') As KOPRAL_PROV,
       --Isnull(Tcd.NOKOPRAL,'...') As NOKOPRAL_PROV,
	   CAST('' As varchar(21)) As 'KOPRAL_PROV',
	   CAST('' As varchar(50)) As 'NOKOPRAL_PROV',
	   CAST('' As varchar(20)) As 'KOPRAL',
	   CAST('' As varchar(50)) As 'NOKOPRAL',
       ISNULL(Mpo.MENSAJE01,'') As 'MENSAJE01',
	   ISNULL(Mpo.MENSAJE02,'') As 'MENSAJE02',
	   ISNULL(Mpo.MENSAJE03,'') As 'MENSAJE03'
       
Into #Paso_Detalle	   
	   
From dbo.MAEDDO Edd 
	LEFT JOIN dbo.MAEPR Mp ON Edd.KOPRCT = Mp.KOPR
		LEFT JOIN dbo.TABBOPR Tbp ON Edd.EMPRESA = Tbp.EMPRESA And Edd.SULIDO = Tbp.KOSU And Tbp.KOBO = Edd.BOSULIDO And Tbp.KOPR = Edd.KOPRCT
			LEFT JOIN dbo.MAEST Mst ON Edd.EMPRESA = Mst.EMPRESA And Edd.SULIDO = Mst.KOSU And Mst.KOBO = Edd.BOSULIDO And Mst.KOPR = Edd.KOPRCT
				 LEFT JOIN dbo.TABMR Mrc ON Mrc.KOMR = Mp.MRPR
                    LEFT JOIN dbo.MAEFICHA Mfch On Mfch.KOPR = Edd.KOPRCT
                        LEFT JOIN dbo.MAEPROBS Mpo On Mpo.KOPR = Edd.KOPRCT
                        
Where IDMAEEDO = @Idmaeedo 
#Condicion_Extra_Maeddo#

Update #Paso_Detalle Set MAEFICHD = Isnull((Select Top 1 FICHA From MAEFICHD Where MAEFICHD.KOPR = #Paso_Detalle.KOPR),'')

Update #Paso_Detalle Set KOPRAL_PROV = (Case When ALTERNAT = '' 
Then 
Isnull((Select Top 1 KOPRAL From TABCODAL Where TABCODAL.KOPR = #Paso_Detalle.KOPR AND TABCODAL.KOEN = (Select ENDO From MAEEDO Where IDMAEEDO = @Idmaeedo)),'...') 
Else 
Isnull((Select Top 1 KOPRAL From TABCODAL Where TABCODAL.KOPR = #Paso_Detalle.KOPR AND TABCODAL.KOEN = (Select ENDO From MAEEDO Where IDMAEEDO = @Idmaeedo) And TABCODAL.KOPRAL = ALTERNAT),'...')
End)

Update #Paso_Detalle Set NOKOPRAL_PROV = (Case When ALTERNAT = '' 
Then 
Isnull((Select Top 1 NOKOPRAL From TABCODAL Where TABCODAL.KOPR = #Paso_Detalle.KOPR AND TABCODAL.KOEN = (Select ENDO From MAEEDO Where IDMAEEDO = @Idmaeedo)),'...') 
Else 
Isnull((Select Top 1 NOKOPRAL From TABCODAL Where TABCODAL.KOPR = #Paso_Detalle.KOPR AND TABCODAL.KOEN = (Select ENDO From MAEEDO Where IDMAEEDO = @Idmaeedo) And TABCODAL.KOPRAL = ALTERNAT),'...')
End)

Update #Paso_Detalle Set KOPRAL = Isnull((Select Top 1 KOPRAL From TABCODAL Where TABCODAL.KOPR = #Paso_Detalle.KOPR AND TABCODAL.KOEN = ''),'..')
Update #Paso_Detalle Set NOKOPRAL = Isnull((Select Top 1 NOKOPRAL From TABCODAL Where TABCODAL.KOPR = #Paso_Detalle.KOPR AND TABCODAL.KOEN = ''),'..')


Update #Paso_Encabezado Set 
    Bk_Suma_CantUd1 = (SELECT SUM(CAPRCO1) From #Paso_Detalle Where PRCT = 0),
	Bk_Suma_CantUd2 = (SELECT SUM(CAPRCO2) From #Paso_Detalle Where PRCT = 0),
	Bk_Suma_Reca_Netos = Isnull((SELECT SUM(VANELI) From #Paso_Detalle Where PRCT = 1 And TICT = 'R'),0),
	Bk_Suma_Desc_Netos = Isnull((SELECT SUM(VANELI) From #Paso_Detalle Where PRCT = 1 And TICT = 'D'),0) *-1,
	Bk_Suma_Reca_Brutos = Isnull((SELECT SUM(VABRLI) From #Paso_Detalle Where PRCT = 1 And TICT = 'R'),0),
	Bk_Suma_Desc_Brutos = Isnull((SELECT SUM(VABRLI) From #Paso_Detalle Where PRCT = 1 And TICT = 'D'),0) *-1,
	Bk_Sub_Total_Neto  = (SELECT SUM(VADTNELI) From #Paso_Detalle) + VANEDO,
	Bk_Sub_Total_Bruto = (SELECT SUM(VADTBRLI) From #Paso_Detalle) + VABRDO,
	Bk_Suma_Dscto_Neto = (SELECT SUM(VADTNELI) From #Paso_Detalle),
	Bk_Suma_Dscto_Bruto = (SELECT SUM(VADTBRLI) From #Paso_Detalle),
	Bk_Total_Items = (SELECT DISTINCT COUNT(NULIDO) From #Paso_Detalle),
    Bk_Recargo = (SELECT ROUND(SUM(POTENCIA*CAPRCO1),0) FROM #Paso_Detalle)

Update #Paso_Encabezado Set Bk_Nombre_Sudo = (Select NOKOSU From TABSU Where EMPRESA = #Paso_Encabezado.EMPRESA And #Paso_Encabezado.SUDO = KOSU )
Update #Paso_Encabezado Set Bk_Direccion_Sudo = (Select DISU From TABSU Where EMPRESA = #Paso_Encabezado.EMPRESA And #Paso_Encabezado.SUDO = KOSU )

Update #Paso_Encabezado Set Bk_1er_Vendedor_Codigo = (Select Top 1 KOFULIDO From #Paso_Detalle)
Update #Paso_Encabezado Set Bk_1er_Vendedor_Nombre = (Select Top 1 NOKOFU From TABFU Where KOFU = (Select Top 1 KOFULIDO From #Paso_Detalle))

Update #Paso_Encabezado Set Bk_Cuotas = Isnull((Select Count(*) From MAEVEN Where IDMAEEDO = @Idmaeedo),0)
Update #Paso_Encabezado Set Bk_Cuotas = 1 Where Bk_Cuotas = 0

-- Encabezado

Select * From #Paso_Encabezado

-- Detalle

Select * From #Paso_Detalle
#Filtro_Productos#
#Orden_Detalle#

-- Referencia

Select Distinct ARCHIRVE,IDRVE,KOFU,FEVENTO,KOTABLA,KOCARAC,NOKOCARAC,ARCHIRSE,IDRSE,HORAGRAB,FECHAREF,LINK,KOFUDEST, 
(Select(Select RTRIM(LTRIM(NOTIDO)) From TABTIDO Td Where Td.TIDO = Edo.TIDO)+' - '+NUDO+' - '+convert(varchar, FEEMDO,103)--+' - '+NOKOCARAC
From MAEEDO Edo Where Edo.IDMAEEDO = IDRSE) As Referencia
From MEVENTO 
Where IDRVE = @Idmaeedo And KOTABLA = 'SET-FE'

-- Productos Agrupados 

Select EMPRESA,TIDO,NUDO,ENDO,SUENDO,ENDOFI,LILG,SULIDO,
       RIGHT('00000' + CAST(ROW_NUMBER() OVER (ORDER BY (SELECT NULL)) AS VARCHAR( 5)), 5) AS NULIDO,
       --LUVTLIDO,
       BOSULIDO,KOFULIDO,NULILG,PRCT,TICT,TIPR,KOPR,NOKOPR_Edd,UDTRPR,RLUDPR,
	   Sum(CAPRCO1) As CAPRCO1,Sum(CAPRAD1) As CAPRAD1,Sum(CAPREX1) As CAPREX1,Sum(CAPRNC1) As CAPRNC1,
	   UD01PR,
	   Sum(CAPRCO2) As CAPRCO2,Sum(CAPRAD2) As CAPRAD2,Sum(CAPREX2) As CAPREX2,Sum(CAPRNC2) As CAPRNC2,
	   UD02PR,--KOLTPR,MOPPPR,TIMOPPPR,TAMOPPPR,
	   PPPRNELT,PPPRBRLT,
	   PPPRNE,PPPRBR,	
	   --NUDTLI,
       Round(PODTGLLI,0) As PODTGLLI,
	   Sum(VADTNELI) As VADTNELI,Sum(VADTBRLI) As VADTBRLI,
	   --POIVLI,
	   --VAIVLI,
	   --NUIMLI,POIMGLLI,	
	   Sum(VAIMLI) As VAIMLI,Sum(VANELI) As VANELI,Sum(VABRLI) As VABRLI,
	   --TIGELI,EMPREPA,TIDOPA,NUDOPA,ENDOPA,NULIDOPA,LLEVADESP,
	   --FEEMLI,FEERLI,--PPPRPM,OPERACION,CODMAQ,ESLIDO,PPPRNERE1,PPPRNERE2,ESFALI,CAFACO,CAFAAD,CAFAEX,CMLIDO,NULOTE,FVENLOTE,ARPROD,NULIPROD,NUCOCO,NULICO,PERIODO,
	   --FCRELOTE,SUBLOTE,
	   --ALTERNAT,PRENDIDO,OBSERVA,KOFUAULIDO,KOOPLIDO,MGLTPR,PPOPPR,TIPOMG,ESPRODLI,CAPRODCO,CAPRODAD,CAPRODEX,CAPRODRE,TASADORIG,CUOGASDIF,SEMILLA,PROYECTO,POTENCIA
	   --HUMEDAD,IDTABITPRE,IDODDGDV,LINCONDESP,PODEIVLI,VADEIVLI,PRIIDETIQ,KOLORESCA,KOENVASE,PPPRPMSUC,PPPRPMIFRS,Expr1,
	   KOPR,NOKOPR_Mp,KOPRRA,NOKOPRRA,KOPRTE,
	   --KOGE,NMARCA,UD01PR_Mp,UD02PR_Mp,RLUD,POIVPR,NUIMPR,RGPR,STMIPR,STMAPR,MRPR,ATPR,RUPR,STFI1,STDV1,STOCNV1,STFI2,STDV2,STOCNV2,
	   --PPUL01,PPUL02,MOUL,TIMOUL,TAUL,FEUL,PM,FEPM,FMPR,PFPR,HFPR,VALI,FEVALI,TTREPR,PRRG,NIPRRG,NFPRRG,PMIN,CAMICO,CAMIFA,LOMIFA,PLANO,STDV1C,STOCNV1C,STDV2C,STOCNV2C,	
	   --METRCO,DESPNOFAC1,DESPNOFAC2,RECENOFAC1,RECENOFAC2,TRATALOTE,DIVISIBLE,MUDEFA,EXENTO,KOMODE,PRDESRES,LISCOSTO,STOCKASEG,ESACTFI,CLALIBPR,KOFUPR,KOPRDIM,NODIM1,NODIM2,
	   --NODIM3,BLOQUEAPR,ZONAPR,CONUBIC,LTSUBIC,PESOUBIC,FUNCLOTE,LOMAFA,COLEGPR,MORGPR,FECRPR,FEMOPR,LOTECAJA,STTR1,STTR2,PODEIVPR,DIVISIBLE2,MOVETIQ,REPUESTO,VIDAMEDIA,
	   --TRATVMEDIA,PRESALCLI1,PRESALCLI2,PRESDEPRO1,PRESDEPRO2,CONSALCLI1,CONSALCLI2,CONSDEPRO1,CONSDEPRO2,DEVENGNCV1,DEVENGNCV2,DEVENGNCC1,DEVENGNCC2,DEVSINNCV1,DEVSINNCV2,	
	   --DEVSINNCC1,DEVSINNCC2,STENFAB1,STENFAB2,STREQFAB1,STREQFAB2,PMME,FEPMME,VALUNFLEKM,ANALIZABLE,TOLELOTE,PMIFRS,FEPMIFRS,
	   Sum(Bk_Cant_Trans) As Bk_Cant_Trans,Bk_Un_Trans--,UBICACION,StockUd1,StockUd2,STOCK
From #Paso_Detalle
Group By EMPRESA,TIDO,NUDO,ENDO,SUENDO,ENDOFI,LILG,SULIDO,BOSULIDO,KOFULIDO,NULILG,PRCT,TICT,TIPR,KOPR,UDTRPR,RLUDPR,
		 UD01PR,UD02PR,NOKOPR_Edd,KOPR,NOKOPR_Mp,KOPRRA,NOKOPRRA,KOPRTE,
		 PPPRNE,PPPRBR,PPPRNELT,PPPRBRLT,Round(PODTGLLI,0),Bk_Un_Trans

-- Documentos relacionados con recargos
Select Distinct Edo.IDMAEEDO,Edo.TIDO,Edo.NUDO,Edo.FEEMDO
From MAEDCR Rec
Inner Join MAEDDO Ddo On Ddo.IDMAEDDO = Rec.IDDDODCR
Inner Join MAEEDO Edo On Edo.IDMAEEDO = Ddo.IDMAEEDO
Where Rec.IDMAEEDO = @Idmaeedo


Drop Table #Paso_Encabezado
Drop Table #Paso_Detalle

--SELECT TOP 10 * FROM MAEEDO WHERE TIDO = 'BLV' ORDER BY FEEMDO DESC 