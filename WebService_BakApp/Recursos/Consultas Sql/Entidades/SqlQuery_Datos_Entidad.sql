
Declare @CodEntidad Char(10) = '#CodEntidad#',
        @SucEntidad Char(10) = '#SucEntidad#'

Select CAST( '' AS Varchar(15)) As 'Rut',
       *, 
       Isnull((Select top 1 NOKOFU From TABFU Where KOFU = KOFUEN),'') As 'VENDEDOR',
       Isnull((Select top 1 NOKOFU From TABFU Where KOFU = COBRADOR),'') As 'NOMCOBRADOR',
       Case TIPOSUC When 'C' Then 'CLIENTE' When 'P' Then 'PROVEEDOR' Else 'AMBOS' End As 'TIPOSUCURSAL',
       (Select Top 1 NOKOEN From MAEEN Where KOEN = @CodEntidad And SUEN = @SucEntidad) As 'RAZON',
       Isnull((Select Top 1 NOKOPA From TABPA Where KOPA = PAEN),'') As 'PAIS', 
       Isnull((Select Top 1 NOKOCI From TABCI Where KOPA = PAEN And KOCI = CIEN),'') As 'CIUDAD', 
       Isnull((Select Top 1 NOKOCM From TABCM Where KOPA = PAEN And KOCI = CIEN And KOCM = CMEN),'') As 'COMUNA',
       Isnull((Select top 1 NOKOZO From TABZO Where KOZO = ZOEN),'') As 'ZONA',
       Isnull((Select top 1 NOKORU From TABRU Where KORU = RUEN),'') As 'RUBRO',
       Isnull((Select top 1 NOKOCARAC From TABCARAC Where KOTABLA = 'ACTIVIDADE' And KOCARAC = ACTIEN),'') As 'ACTECONOMICA',
       Isnull((Select top 1 NOKOCARAC From TABCARAC Where KOTABLA = 'TAMA¥OEMPR' And KOCARAC = TAMAEN),'') As 'TAMAEMPRESA',
       Isnull((Select top 1 NOKOCARAC From TABCARAC Where KOTABLA = 'TIPOENTIDA' And KOCARAC = TIPOEN),'') As 'TIPOENTIDAD'
       
From MAEEN Where KOEN = @CodEntidad And SUEN = @SucEntidad