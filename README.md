# Generador de formatos de facturación especial

Aplicación en Python para cargar el libro de Excel, previsualizar las columnas principales y generar PDFs imprimibles con el formato `FORMATO ENTREGA FACTURACIÓN ESPECIAL`.

## Instalación

```powershell
pip install -r requirements.txt
```

## Uso

```powershell
python app.py
```

1. Seleccione el archivo Excel.
2. Pulse `Cargar`.
3. Revise la grilla paginada.
4. Pulse `Generar PDFs`.

Los PDFs se crean en `Salidad Impresiones` por defecto.

## Reglas implementadas

- Agrupa por `DIRECCION DE ENTREGA REAL`, `SEGMENTO EJECUTIVO` y titular de pago.
- Usa `TITULAR PAGO REAL`; si no existe, usa `TITULAR DE PAGO`.
- Genera una tabla de máximo 39 registros por página.
- Si un grupo supera 39 registros, continúa en páginas adicionales del mismo PDF.
- Calcula `DEUDA` como `SALDO VENCIDO - IMPORTE`.
- Totaliza `IMPORTE MES`, `DEUDA` y `TOTAL`.
- Usa el logo ubicado en `Logo/logo-afinia.png` si existe.
