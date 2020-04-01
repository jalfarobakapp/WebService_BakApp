Declare @Codigo as varchar(20),
		@Empresa As char(2),
		@Sucursal As varchar(3),
        @Bodega As varchar(3),
		@Lista As Varchar(3)

Select @Codigo = '#Codigo#',@Empresa = '#Empresa#',@Sucursal = '#Sucursal#',@Bodega = '#Bodega#',@Lista = '#Lista#'

Select Mp.KOPR,Mp.KOPRRA,Mp.KOPRTE,NOKOPR,Mp.RLUD,Mp.UD01PR,Mp.UD02PR,Mp.STFI1 As STFI1_Cons,Mp.STFI2 As STFI2_Cons,Mp.POIVPR,Mp.LISCOSTO,Mp.TIPR,
Tp.PP01UD,Tp.PPUL02,Tp.MG01UD,Tp.MG02UD,Tp.DTMA01UD,Tp.DTMA02UD,
Rtrim(Ltrim(Isnull(Tp.ECUACION,''))) As ECUACION,Rtrim(Ltrim(Isnull(Tp.ECUACIONU2,''))) As ECUACIONU2,
Ms.STFI1,Ms.STFI2,
Mpm.PM As PmLinea,Mpm.PPUL01 As Precio_UC1,Mpm.PPUL02 As Precio_UC2,Mpm.PMIFRS As PmIFRS,
Mps.PMSUC As PmSucLinea,
Tbpp.DATOSUBIC As UbicacionBod

From MAEPR Mp
	Left Join TABPRE Tp On Tp.KOLT = @Lista And Tp.KOPR = Mp.KOPR 
		Left Join MAEST Ms On Ms.EMPRESA = @Empresa And Ms.KOSU = @Sucursal And Ms.KOBO = @Bodega And Ms.KOPR = Mp.KOPR
			Left Join MAEPREM Mpm On Ms.KOPR = Mpm.KOPR And Mpm.EMPRESA = @Empresa
				Left Join MAEPMSUC Mps On Mps.EMPRESA = @Empresa And Mps.KOSU = @Sucursal And Mps.KOPR = Mp.KOPR
					Left Join TABBOPR Tbpp On Tbpp.EMPRESA = @Empresa And Tbpp.KOSU = @Sucursal And Tbpp.KOBO = @Bodega And Tbpp.KOPR = Mp.KOPR

Where Mp.KOPR = @Codigo