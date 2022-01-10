Declare @Codigo as varchar(20),
		@Empresa As char(2),
		@Sucursal As varchar(3),
        @Bodega As varchar(3),
		@Lista As Varchar(3),
		@UnTrans As Int

Select @Codigo = '#Codigo#',@Empresa = '#Empresa#',@Sucursal = '#Sucursal#',@Bodega = '#Bodega#',@Lista = '#Lista#',@UnTrans = #UnTrans#

Select  Cast(0 As Int) As 'Id_DocEnc',
		@Empresa As 'Empresa',
		@Sucursal As 'Sucursal',
		@Bodega As 'Bodega',
		Mp.KOPR As 'Codigo',
		--Mp.KOPRRA As '',
		--Mp.KOPRTE As '',
		NOKOPR As 'Descripcion',
		@UnTrans as 'UnTrans',
	    Case @UnTrans When 1 Then Mp.UD01PR When 2 Then Mp.UD02PR End As 'UdTrans',
		Mp.RLUD As 'Rtu',
		Mp.UD01PR As 'Ud01PR',
		Mp.UD02PR As 'Ud02PR',
		Mp.POIVPR As 'PorIva',
		Cast(0 As Float) As 'PorIla',
		Case @UnTrans When 1 Then Isnull(Ms.STFI1,0) When 2 Then Isnull(Ms.STFI2,0) End As 'StockBodega',
		--Mp.STFI2 As 'STFI2_Cons',
		Mp.LISCOSTO As 'CodLista',
		Cast(0 as Bit) as 'Prct',
		Cast('' As Varchar(1)) As 'Tict',
		Mp.TIPR As 'Tipr',
		--Tp.PP01UD As 'Precio',
		Cast(0 As Float) As 'Precio',
		Cast(0 As Float) As 'PrecioListaUd1',
		Cast(0 As Float) As 'PrecioListaUd2',
		Cast(0 As Float) As 'DescuentoPorc',
		--Tp.MG01UD As 'DescuentoPorc',
		--Tp.MG02UD,
		--Tp.DTMA01UD As 'DescMaximo',
		Cast(0 As Float) As 'DescMaximo',
		--Tp.DTMA02UD,
		Cast('' As Varchar(242)) As 'Ecuacion',
		--Rtrim(Ltrim(Isnull(Tp.ECUACION,''))) As 'Ecuacion',
		--Rtrim(Ltrim(Isnull(Tp.ECUACIONU2,''))) As 'ECUACIONU2',
		--Ms.STFI1 As StockUd1,
		--Ms.STFI2 As StockUd2,
		Isnull(Mpm.PM,0) As 'PmLinea',
		--Mpm.PPUL01 As Precio_UC1,
		--Mpm.PPUL02 As Precio_UC2,
		Isnull(Mps.PMSUC,0) As 'PmSucLinea',
		Isnull(Mpm.PMIFRS,0) As 'PmIFRS',
		Tbpp.DATOSUBIC As 'UbicacionBod',
		'' As 'Moneda',
		'' As 'Tipo_Moneda',
		Cast(0 As Float) As 'Tipo_Cambio'

From MAEPR Mp
		Left Join MAEST Ms On Ms.EMPRESA = @Empresa And Ms.KOSU = @Sucursal And Ms.KOBO = @Bodega And Ms.KOPR = Mp.KOPR
			Left Join MAEPREM Mpm On Ms.KOPR = Mpm.KOPR And Mpm.EMPRESA = @Empresa
				Left Join MAEPMSUC Mps On Mps.EMPRESA = @Empresa And Mps.KOSU = @Sucursal And Mps.KOPR = Mp.KOPR
					Left Join TABBOPR Tbpp On Tbpp.EMPRESA = @Empresa And Tbpp.KOSU = @Sucursal And Tbpp.KOBO = @Bodega And Tbpp.KOPR = Mp.KOPR

Where Mp.KOPR = @Codigo



