# Parámetros del año gravable — referencia AG2025 y ritual de actualización anual

**Regla:** ningún parámetro se usa sin verificar su vigencia para el año gravable que se declara.
Esta tabla trae los valores del **año gravable 2025** (declarado en 2026) como referencia trabajada.
Para otro año: busque en la web cada fuente, actualice `scripts/motor_renta.py` (clase `Params`) y la
lista `params` + calendario de `scripts/generar_excel.py` (hoja 16_PARAMETROS), y deje constancia en
la memoria.

## Valores AG2025 (verificados en su momento)

| Parámetro | Valor AG2025 | Fuente / norma |
|---|---|---|
| UVT del año gravable | $49.799 | Resolución DIAN 000193 de 2024 |
| UVT del año de presentación (2026) | $52.374 | Resolución DIAN 000238 de 2025 |
| Límite conjunto exentas+deducciones | 40% de la base y 1.340 UVT | Art. 336 num. 3 E.T. |
| Renta exenta laboral 25% | tope 790 UVT | Art. 206 num. 10 |
| AFC + FVP + AVC | 30% del ingreso y 3.800 UVT | Arts. 126-1 y 126-4 |
| Intereses de vivienda | 1.200 UVT | Art. 119 |
| Intereses ICETEX | 100 UVT | Art. 119 inc. 3 |
| Medicina prepagada | 16 UVT/mes | Art. 387 |
| Dependientes art. 387 | 32 UVT/mes (y 10% del ingreso laboral) | Art. 387 inc. 2 |
| Dependiente adicional (fuera del límite) | 72 UVT c/u, máx. 4 | Art. 336 num. 3 inc. 2 |
| 1% compras con factura electrónica | máx. 240 UVT, fuera del límite | Art. 336 num. 5 |
| RAIS voluntario (INCRNGO) | 25% del ingreso y 2.500 UVT | Art. 55 |
| Componente inflacionario rendimientos | 55,43% no gravado | Decreto 898 de 2026; arts. 38-41 |
| Componente inflacionario costos financieros | 28,35% NO deducible | Decreto 898 de 2026; arts. 81, 118 |
| GMF deducible | 50% de lo pagado | Art. 115 |
| Tarifa general | tabla art. 241 (0/19/28/33/35/37/39%; primer tramo hasta 1.090 UVT) | Art. 241 |
| Ganancia ocasional | 15% general; 20% loterías | Arts. 314, 317 |
| Renta presuntiva | 0% desde 2021 | Art. 188 |
| Sanción mínima | 10 UVT del año de presentación | Art. 639 |
| Anticipo | 25%/50%/75% según años declarando; opción promedio 2 años | Art. 807 |
| Firmeza general / beneficio de auditoría | 3 años / 6-12 meses si el impuesto crece 35%/25% | Arts. 714, 689-3 |

## Obligación de declarar (topes AG2025, en UVT del año gravable)

Obligado si supera CUALQUIERA: patrimonio bruto 4.500 UVT ($224,1M) · ingresos brutos 1.400 UVT
($69,7M) · consumos con tarjeta 1.400 UVT · compras 1.400 UVT · consignaciones/depósitos/inversiones
1.400 UVT · o responsable de IVA. (Arts. 592-594-3 E.T.) El reporte de exógena de la DIAN trae los
cinco "topes" ya medidos — úselo como primera verificación.

## Ritual de actualización anual (qué buscar y dónde)

1. **UVT del año gravable y del año de presentación** — resolución DIAN de noviembre-diciembre
   ("por la cual se fija el valor de la UVT"). Buscar: `UVT <año> resolución DIAN`.
2. **Componente inflacionario** (rendimientos y costos financieros) — decreto reglamentario de
   mediados del año siguiente. Buscar: `componente inflacionario año gravable <año> decreto`.
3. **Calendario de plazos** por dos últimos dígitos del NIT — decreto de plazos (Minhacienda,
   diciembre). Buscar: `calendario tributario <año+1> personas naturales decreto`. Actualice la tabla
   `cal` en generar_excel.py.
4. **Formulario e instructivo del 210** — resolución DIAN que prescribe el formulario del año.
   Verifique numeración de casillas antes de asumir que no cambió.
5. **Reformas tributarias**: si hubo ley nueva, verifique tabla 241, topes de los arts. 206, 336,
   387 y tarifas de ganancia ocasional.

Fuentes oficiales: Estatuto Tributario en `secretariasenado.gov.co/senado/basedoc/estatuto_tributario.html`
· micrositio DIAN "Renta personas naturales AG <año>" · normograma DIAN. Prefiera SIEMPRE el texto de
la norma sobre artículos de prensa o blogs contables.
