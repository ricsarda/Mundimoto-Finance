import sys
import pandas as pd
import openpyxl

# Parámetros
MES = 12
AÑO = 2024

# Argumentos:
# argv[1] = inf_usu_FC
# argv[2] = inf_usu_AB
# argv[3] = inf_usu_FT
# argv[4] = comp_alb
# argv[5] = ruta_archivo_final_excel

inf_usu_FC = sys.argv[1]  # Ej: "FC.xlsx"
inf_usu_AB = sys.argv[2]  # Ej: "AB.xlsx"
inf_usu_FT = sys.argv[3]  # Ej: "FT.xlsx"
comp_alb   = sys.argv[4]  # Ej: "Compras.xlsx"
ruta_archivo_final_excel = sys.argv[5]  # Ej: "DAILY.xlsx"

Limpiar_FC = pd.read_excel(inf_usu_FC)
Limpiar_FC = Limpiar_FC.loc[Limpiar_FC['SerieFactura'].isin(['FC','FP','FI','FL','AC'])]

Limpiar_AB = pd.read_excel(inf_usu_AB)
Limpiar_FT = pd.read_excel(inf_usu_FT)

DAILY = pd.concat([Limpiar_FC, Limpiar_AB, Limpiar_FT])

DAILY['FechaFactura'] = pd.to_datetime(DAILY['FechaFactura'])
DAILY = DAILY[(DAILY['FechaFactura'].dt.month == MES) & (DAILY['FechaFactura'].dt.year == AÑO)]
columnas_seleccionadasDAILY = ['SerieFactura' , 'IdDelegacion' , 'NumeroFactura' , 'FechaFactura', 'RazonSocial' , 'Unidades' , 'CodigoArticulo' , 'CodigoFamilia' , 'DescripcionArticulo', 'PrecioCompra', 'BaseImponible1' , 'ImporteCoste' , 'MargenBeneficio']
DAILY = DAILY[columnas_seleccionadasDAILY]
DAILY['PrecioCompra'] = DAILY['PrecioCompra'].astype(float)
DAILY['BaseImponible1'] = DAILY['BaseImponible1'].astype(float)
DAILY['Unidades'] = DAILY['Unidades'].astype(float)
DAILY['Margen'] = DAILY['BaseImponible1'] - DAILY['PrecioCompra']
DAILY['MargenAC'] = DAILY['BaseImponible1'] - DAILY['ImporteCoste']

def margen_FC(x):
    if x <= 0:
        return float(x)
    else:
        return float(x/1.21)

def cuota_FC(x,y):
    if x - y > 0:
        return float(x - y)
    else:
        return float(0)

def margen_AC(x):
    return float(x/1.21)

def cuota_AC(x,y):
    return float(x - y)

Desguace = DAILY.loc[DAILY['CodigoFamilia'].isin(['DESGUACE' , 'desguace'])]
DesguaceAC = Desguace.loc[Desguace['SerieFactura'].isin(['AC'])]
DesguaceF = Desguace.loc[Desguace['SerieFactura'].isin(['FC','FP','FI','FL'])]
Desguaceunits = DesguaceF['Unidades'].sum()
DesguaceACunits = DesguaceAC['Unidades'].sum()
Desguaceunits = DesguaceACunits + Desguaceunits

DesguaceF['MargenIVA'] = DesguaceF['Margen'].apply(margen_FC)
DesguaceAC['MargenIVA'] = DesguaceAC['MargenBeneficio'].apply(margen_AC)

DesguacemargenFC = DesguaceF['MargenIVA'].sum()
DesguacemargenAC = DesguaceAC['MargenIVA'].sum()
Desguacemargen = DesguacemargenFC + DesguacemargenAC

DesguaceAC['Cuota'] = [cuota_AC(val1, val2) for val1, val2 in zip(DesguaceAC['MargenBeneficio'], DesguaceAC['MargenIVA'])]
DesguaceF['Cuota'] = [cuota_FC(val1, val2) for val1, val2 in zip(DesguaceF['Margen'], DesguaceF['MargenIVA'])]

DesguaceAC['Venta sin IVA'] = DesguaceAC['BaseImponible1'] - DesguaceAC['Cuota']
DesguaceF['Venta sin IVA'] = DesguaceF['BaseImponible1'] - DesguaceF['Cuota']

DesguaceVentaAC = DesguaceAC['Venta sin IVA'].sum()
DesguaceVentaFC = DesguaceF['Venta sin IVA'].sum()
DesguaceVenta = DesguaceVentaAC + DesguaceVentaFC
DesguaceACunits = -DesguaceACunits

B2CFC = DAILY.loc[DAILY['SerieFactura'].isin(['FC'])]
B2CFC = B2CFC.loc[B2CFC['CodigoFamilia'].isnull()]
B2CFCunidades = B2CFC['Unidades'].sum()
B2CFC['MargenIVA'] = B2CFC['Margen'].apply(margen_FC)
B2CmargenFC = B2CFC['MargenIVA'].sum()
B2CFC['Cuota'] = [cuota_FC(val1, val2) for val1, val2 in zip(B2CFC['Margen'], B2CFC['MargenIVA'])]
B2CFC['Venta sin IVA'] = B2CFC['BaseImponible1'] - B2CFC['Cuota']
B2CVentaFC = B2CFC['Venta sin IVA'].sum()

B2CAC = DAILY.loc[DAILY['SerieFactura'].isin(['AC'])]
B2CAC = B2CAC.loc[B2CAC['CodigoFamilia'].isnull()]
B2CAC = B2CAC.loc[B2CAC['IdDelegacion'].isnull()]
B2CACunidades = B2CAC['Unidades'].sum()
B2CAC['MargenIVA'] = B2CAC['MargenBeneficio'].apply(margen_AC)
B2CmargenAC = B2CAC['MargenIVA'].sum()
B2CAC['Cuota'] = [cuota_AC(val1, val2) for val1, val2 in zip(B2CAC['MargenBeneficio'], B2CAC['MargenIVA'])]
B2CAC['Venta sin IVA'] = B2CAC['BaseImponible1'] - B2CAC['Cuota']
B2CVentaAC = B2CAC['Venta sin IVA'].sum()

B2CVN = DAILY.loc[DAILY['SerieFactura'].isin(['FC','FL','FP','FI'])]
B2CVN = B2CVN.loc[B2CVN['CodigoFamilia'].isin(['VN','vn'])]
B2CVN = B2CVN.loc[B2CVN['IdDelegacion'].isnull()]
B2CVNunidades = B2CVN['Unidades'].sum()

B2CVN['MargenIVA'] = B2CVN['MargenBeneficio'].apply(margen_FC)
B2CmargenVN = B2CVN['MargenIVA'].sum()
B2CVN['Cuota'] = [cuota_FC(val1, val2) for val1, val2 in zip(B2CVN['Margen'], B2CVN['MargenIVA'])]
B2CVN['Venta sin IVA'] = B2CVN['BaseImponible1'] - B2CVN['Cuota']
B2CVentaVN = B2CVN['Venta sin IVA'].sum()

B2CVNAC = DAILY.loc[DAILY['SerieFactura'].isin(['AC'])]
B2CVNAC = B2CVNAC.loc[B2CVNAC['CodigoFamilia'].isin(['VN','vn'])]
B2CVNAC = B2CVNAC.loc[B2CVNAC['IdDelegacion'].isnull()]
B2CVNACunidades = B2CVNAC['Unidades'].sum()
B2CVNAC['MargenIVA'] = B2CVNAC['MargenBeneficio'].apply(margen_AC)
B2CVNACmargen = B2CVNAC['MargenIVA'].sum()
B2CVNAC['Cuota'] = [cuota_AC(val1, val2) for val1, val2 in zip(B2CVNAC['MargenBeneficio'], B2CVNAC['MargenIVA'])]
B2CVNAC['Venta sin IVA'] = B2CVNAC['BaseImponible1'] - B2CVNAC['Cuota']
B2CVNACVenta = B2CVNAC['Venta sin IVA'].sum()

CAMBIODENOMBRE = DAILY[DAILY['CodigoArticulo'].isin(['CAMBIO DE NOMBRE','SUPLIDO CN'])]
CAMBIODENOMBRE = CAMBIODENOMBRE.loc[CAMBIODENOMBRE['CodigoFamilia'].isnull()]

CAMBIODENOMBREFT = CAMBIODENOMBRE.loc[CAMBIODENOMBRE['SerieFactura'].isin(['FT'])]
CAMBIONOMBREMARGENFT = CAMBIODENOMBREFT['BaseImponible1'].sum()

CAMBIODENOMBREAB = CAMBIODENOMBRE.loc[CAMBIODENOMBRE['SerieFactura'].isin(['AB'])]
CAMBIONOMBREMARGENAB = CAMBIODENOMBREAB['BaseImponible1'].sum()

B2Cunidades = B2CACunidades + B2CFCunidades + B2CVNunidades + B2CVNACunidades
B2Cmargen = B2CmargenAC + B2CmargenFC + B2CmargenVN + B2CVNACmargen + CAMBIONOMBREMARGENFT + CAMBIONOMBREMARGENAB
B2CVenta = B2CVentaFC + B2CVentaAC + B2CVentaVN + B2CVNACVenta + CAMBIONOMBREMARGENFT + CAMBIONOMBREMARGENAB
B2CACunidades = - B2CACunidades - B2CVNACunidades

B2BAC = DAILY.loc[DAILY['SerieFactura'].isin(['AC'])]
B2BAC = B2BAC.loc[B2BAC['IdDelegacion'].isin(['B2B'])]
B2BAC = B2BAC.loc[B2BAC['CodigoFamilia'].isnull()]
B2BACunidades = B2BAC['Unidades'].sum()
B2BACventa = B2BAC['BaseImponible1'].sum()
B2BmargenAC = B2BAC['MargenBeneficio'].sum()

B2BFP = DAILY.loc[DAILY['SerieFactura'].isin(['FP','FL'])]
B2BFP = B2BFP.loc[B2BFP['IdDelegacion'].isin(['B2B'])]
B2BFP = B2BFP.loc[B2BFP['CodigoFamilia'].isnull()]
B2BFPunidades = B2BFP['Unidades'].sum()
B2BFPventa = B2BFP['BaseImponible1'].sum()
B2BFP['MargenIVA'] = B2BFP['Margen'].apply(margen_FC)
B2BmargenFP = B2BFP['MargenIVA'].sum()

B2BFI = DAILY.loc[DAILY['SerieFactura'].isin(['FI'])]
B2BFI = B2BFI.loc[B2BFI['IdDelegacion'].isin(['B2B'])]
B2BFI = B2BFI.loc[B2BFI['CodigoFamilia'].isnull()]
B2BFIunidades = B2BFI['Unidades'].sum()
B2BFIventa = B2BFI['BaseImponible1'].sum()
B2BmargenFI = B2BFI['Margen'].sum()

B2Bunidades = B2BACunidades + B2BFPunidades + B2BFIunidades
B2Bmargen = B2BmargenAC + B2BmargenFP + B2BmargenFI
B2BVenta = B2BACventa + B2BFPventa + B2BFIventa
B2BACunidades = - B2BACunidades

BASIC = DAILY[DAILY['CodigoArticulo'].isin(['TRANSPORTE NACIONAL'])]
BASICUNITS = BASIC['Unidades'].sum()
BASICINVO = BASIC['BaseImponible1'].sum()

SPORT = DAILY[DAILY['CodigoArticulo'].isin(['SPORT PLUS','SPORT 500','SPORT 300'])]
SPORTUNITS = SPORT['Unidades'].sum()
SPORTINVO = SPORT['BaseImponible1'].sum()

COMPLETO = DAILY[DAILY['CodigoArticulo'].isin(['PACK COMPLETO','PACK'])]
COMPLETOUNITS = COMPLETO['Unidades'].sum()
COMPLETOINVO = COMPLETO['BaseImponible1'].sum()

PREMIUM = DAILY[DAILY['CodigoArticulo'].isin(['PACK PREMIUM'])]
PREMIUMUNITS = PREMIUM['Unidades'].sum()
PREMIUMINVO = PREMIUM['BaseImponible1'].sum()

STREET = DAILY[DAILY['CodigoArticulo'].isin(['STREET PLUS','STREET 125','STREET 300','STREET 500'])]
STREETUNITS = STREET['Unidades'].sum()
STREETINVO = STREET['BaseImponible1'].sum()

seguro = DAILY[DAILY['CodigoArticulo'].isin(['SEGURO'])]
segurounits = seguro['Unidades'].sum()
seguroinvo = seguro['BaseImponible1'].sum()

Purchaces = pd.read_excel(comp_alb)
Purchaces['Fecha albarán'] = pd.to_datetime(Purchaces['Fecha albarán'])
Purchaces = Purchaces[(Purchaces['Fecha albarán'].dt.month == MES) & (Purchaces['Fecha albarán'].dt.year == AÑO)]

Purchacesimporte = Purchaces['Base imponible'].sum()
PurchacesCV = Purchaces.loc[Purchaces['Serie albarán'] == 'CV']
PurchacesCV = PurchacesCV['Nº líneas'].sum()
PurchacesAB = Purchaces.loc[Purchaces['Serie albarán'] == 'AB']
PurchacesAB = PurchacesAB['Nº líneas'].sum()
Purchacesunidades = PurchacesCV - PurchacesAB

DesguaceACimporte = DesguaceAC['BaseImponible1'].sum()
B2CACimorte = B2CAC['BaseImponible1'].sum()
B2BACimorte = B2BAC['BaseImponible1'].sum()
ValortotalAC = - DesguaceACimporte - B2BACimorte - B2CACimorte

Report = [
    'UNITS B2C', 'INVOICED B2C', 'MARGIN B2C',
    'UNITS B2B', 'INVOICED B2B', 'MARGIN B2B',
    'UNITS SCRAP', 'INVOICED SCRAP', 'MARGIN SCRAP',
    'FINANCING UNITS', 'TOTAL FINANCED', 'COMISSION',
    'BASIC UNITS', 'BASIC INVOICED',
    'COMPLETE UNITS', 'COMPLETE INVOICED',
    'PREMIUM UNITS', 'PREMIUM INVOICED',
    'STREET UNITS', 'STREET INVOICED',
    'SPORT UNITS', 'SPORT INVOICED',
    'SEGURO UNITS', 'SEGURO INVOICED',
    'PURCHASES UNITS', 'PURCHASES INVOICED',
    'STOCK UNITS', 'STOCK VALUE',
    'RETURNED B2C', 'RETURNED B2B', 'RETURNED SCRAPS', 'RETURNED NV', 'RETURNED'
]
Resultados = [
    B2Cunidades, B2CVenta, B2Cmargen,
    B2Bunidades, B2BVenta, B2Bmargen,
    Desguaceunits, DesguaceVenta, Desguacemargen,
    0, 0, 0,
    BASICUNITS, BASICINVO,
    COMPLETOUNITS, COMPLETOINVO,
    PREMIUMUNITS, PREMIUMINVO,
    STREETUNITS, STREETINVO,
    SPORTUNITS, SPORTINVO,
    segurounits, seguroinvo,
    Purchacesunidades, Purchacesimporte,
    0, 0,
    B2CACunidades, B2BACunidades, DesguaceACunits, B2CVNACunidades, ValortotalAC
]

tablafinal = {
    'Report': Report,
    'Resultados': Resultados
}

Reportdaily = pd.DataFrame(tablafinal)
Reportdaily['Resultados'] = Reportdaily['Resultados'].round(2)

Reportdaily.to_excel(ruta_archivo_final_excel, index=False)
