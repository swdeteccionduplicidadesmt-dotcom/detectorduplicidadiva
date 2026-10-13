# Detector de duplicidad de IVA

Script en Python que cruza el **Libro de Ventas de Sofnet** contra los cobros de
**Transbank** y **Mercado Libre / Mercado Pago** para detectar ventas cuyo IVA
podría declararse dos veces: una vez en la boleta o factura emitida y otra en
el Comprobante de Pago Electrónico (CPE).

## Qué hace

1. Busca los Excel de entrada en las carpetas de Sofnet, Transbank y Mercado Libre.
   Detecta solo la fila de encabezados (por ejemplo, la fila 19 del export de Transbank).
2. Concilia cada cobro con una venta Sofnet (1 a 1) con estas reglas, en orden:
   - **R1:** mismo día y mismo monto.
   - **R2:** mismo día y monto con tolerancia de ±$5.
   - **R3:** mismo mes y mismo monto.
   - **R4:** mismo mes y monto con tolerancia de ±$5.
3. Genera un Excel por mes, `Conciliacion_AAAA_MM.xlsx`, con estas pestañas:

| Pestaña | Contenido |
|---|---|
| **Resumen** | Registros y monto bruto por tipo de documento y operador. Los cobros no conciliados se informan como tipo **48** (CPE). |
| **Conciliacion** | Detalle de cada cobro y la venta Sofnet asociada (regla aplicada, diferencias). |
| **Sofnet_sin_cobro** | Ventas Sofnet del mes sin ningún cobro asociado. |
| **CPE Detalle** | Cobros agrupados por día, operador y canal, con los folios asociados y las columnas de corrección J a N. |
| **Informe CPE** | Formato del formulario SII de Comprobantes de Pago Electrónico (Presencial e Internet por día), con montos corregidos y motivo del cambio. |

### Canal de venta
- **Mercado Libre:** siempre Internet.
- **Transbank:** según la columna *N° de boleta*.
  - `0000000000`, vacío o `-` → **Internet**.
  - Cualquier número mayor que 0 → **Presencial**.

## Instalación

```bash
pip install -r requirements.txt
```

## Uso

```bash
python Detector_de_coincidencias.py --base "C:\ruta\a\Ventas"
```

Estructura de carpetas esperada dentro de `--base`:

```
Ventas/
├── Inputs/
│   ├── Sofnet/Libro de ventas/Libro de ventas 2026/   (*.xlsx)
│   ├── Transbank/Ventas/                              (*.xlsx)
│   └── Mercado libre/                                 (*.xlsx)
└── Output/Detector_de_coincidencias/                  (se crea solo)
```

> Los archivos Excel están excluidos del repositorio (`.gitignore`) porque
> contienen información de ventas real.
