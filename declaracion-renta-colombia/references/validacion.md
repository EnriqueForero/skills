# Validación — cómo se garantiza la calidad (los tres candados)

La tesis: un Excel que solo se revisa "a ojo" no está validado. La calidad se garantiza con TRES
mecanismos independientes, y el veredicto "validado" exige los tres en verde.

## Candado 1 — Controles vivos (hoja 12_CONTROLES)

Fórmulas dentro del propio libro: residencia, plazo, coherencia patrimonial (arts. 236-239),
exclusión mutua de saldo a pagar/a favor, conteo de ámbar, y los CRUCES EXACTOS contra certificados
que usted define en `CONTROLES_EXTRA` del caso (suma de retenciones, de ingresos laborales, de
rendimientos — con tolerancia de $1). Producen el DICTAMEN automático:

- `NO LISTO PARA PRESENTAR` — hay BLOQUEO/FALTA/ERROR/INVESTIGAR (algo crítico sin resolver).
- `CALCULADO — CIERRE PENDIENTES` — matemática lista; quedan ámbar/revisiones (estado honesto típico).
- `LISTO PARA REVISIÓN FINAL` — sin pendientes: transcribir al MUISCA.

## Candado 2 — El libro se prueba a sí mismo (hoja 13_QA)

23 bases de prueba en los bordes exactos de la tabla del art. 241 (±0,01 UVT alrededor de cada
umbral), con el valor esperado calculado en Python con `Decimal` embebido como constante. Si
cualquier fila da FALLA, el archivo está corrupto: no se usa. La convención de la tarifa es la del
AyudaRenta oficial: base en UVT truncada a 2 decimales, épsilon 1E-9 contra los errores binarios de
coma flotante de Excel, resultado redondeado a miles (art. 577).

## Candado 3 — Recomputación independiente (scripts/validar.py)

```bash
python3 scripts/validar.py caso_X.py Declaracion_X.xlsx
```

Qué hace, en orden: (1) recalcula el archivo DE VERDAD con LibreOffice headless (un xlsx recién
generado no tiene valores en caché); (2) reconstruye el caso desde el archivo de caso; (3) liquida
las 114 casillas + auxiliares con el motor `Decimal` en DOS capas (oficial redondeada a miles y
exacta al centavo); (4) compara TODO contra el Excel con tolerancia $0,5 (oficial) y $0,01 (exacta);
(5) lee el dictamen y las fronteras QA; (6) imprime JSON con paridad, fallas, casillas clave y
SHA-256 del archivo. Sale con código 0 solo si la paridad es 100%.

Interpretación de fallas: el reporte da casilla, capa y ambos valores. Causa típica: una entrada
difiere entre el caso .py y el Excel (se editó uno y no el otro) — corrija la ENTRADA y regenere;
nunca "cuadre" una fórmula azul a mano. Si el caso .py y el Excel se editaron por caminos distintos,
el caso .py es la fuente de verdad: regenere el Excel desde él.

## Pruebas de estrés (recomendadas antes de entregar; obligatorias si hubo cambios de fórmulas)

Duplique el caso y valide variantes — el motor y el Excel deben moverse JUNTOS en todas:

| Escenario | Qué demuestra |
|---|---|
| ±$1 en un ingreso | La capa exacta responde al peso (sensibilidad) |
| AFC/vivienda/medicina muy por encima del tope | Los topes cortan igual en ambos motores |
| Dependientes 0 / 2 | Las dos deducciones y la c.139 se activan bien |
| Años declarando 1 / 2 / 3 | El anticipo cambia 25/50/75% con opción promedio |
| Todo en "No" (caso vacío) | Sin residuos: todas las casillas en 0 |
| Un ingreso reclasificado (c.32→c.43) | La cascada del límite se reparte distinto y cuadra |
| Mutación deliberada (escriba `=0` en una casilla no nula del F210) | validar.py DEBE detectarla — prueba que el detector detecta |

## Reglas de honestidad del veredicto

- "Paridad 100%" habla de MATEMÁTICA. Diga siempre el alcance: no certifica hechos ni soportes.
- Cualquier cambio en el archivo invalida la validación anterior: re-corra y re-registre en 13_QA.
- Sin LibreOffice no hay validación real. Ofrezca la alternativa (abrir y guardar en Excel y
  re-correr) pero jamás rebaje el estándar en silencio.
- El nivel de confianza final separa: matemática (alta si 100%), clasificación (según evidencia
  leída) y hechos del declarante (lo que solo él puede confirmar — lístelo).
