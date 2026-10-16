---
name: declaracion-renta-colombia
description: >
  Prepara de principio a fin la declaración de renta de una persona natural RESIDENTE en Colombia
  (Formulario 210): recolecta evidencia (exógena DIAN, borrador sugerido, certificados), entrevista al
  declarante sobre las cifras que no se explican solas, clasifica por cédulas, genera un Excel auditado
  de 18 hojas con controles y dictamen, lo valida con un motor de cálculo independiente (paridad $0) y
  entrega memoria de decisiones + veredicto de presentabilidad. USE ESTA SKILL SIEMPRE que el usuario
  mencione: declaración de renta, impuesto de renta, DIAN, Formulario 210, información exógena, borrador
  sugerido, certificado de ingresos y retenciones (220), UVT, renta líquida, saldo a favor, MUISCA,
  patrimonio a declarar, o adjunte certificados tributarios, exógena o declaraciones de años anteriores
  — incluso si solo pregunta "¿estoy obligado a declarar?", "¿cuánto me toca pagar de renta?" o pide
  "ayuda con mis impuestos" en contexto colombiano. NO aplica a empresas ni al Formulario 110.
---

# Declaración de renta — persona natural residente (Colombia, Formulario 210)

## Qué produce esta skill y qué no

Cada uso completo entrega: (1) un **Excel auditado** de 18 hojas numeradas con las 114 casillas del
Formulario 210 calculadas, controles vivos y dictamen automático; (2) una **validación independiente**
en Python que recalcula todo desde cero y exige coincidencia exacta; (3) una **memoria de decisiones**
en Word con la evidencia de cada cifra; y (4) un **veredicto de presentabilidad** honesto. La promesa
de calidad es doble: ninguna cifra sin origen documentado, y ninguna fórmula sin verificación
independiente.

**Fuera de alcance — detectar y derivar a contador, nunca improvisar:** obligados a llevar
contabilidad, régimen ECE, activos/rentas del exterior significativos (Formulario 160/110), no
residentes, sucesiones ilíquidas, y correcciones litigiosas. Diga por qué se deriva y qué preparó
hasta ese punto. Esta skill es una hoja de trabajo rigurosa, no asesoría legal: la responsabilidad de
lo declarado es del declarante (arts. 571, 596, 647-648 E.T.) — dígalo en cada entrega.

## El flujo maestro (el orden importa)

| Fase | Qué se hace | Regla de oro |
|---|---|---|
| 0 | Verificar vigencia normativa del año gravable | Nunca calcule con parámetros sin verificar |
| 1 | Recolectar evidencia (exógena, borrador DIAN, certificados) | Sin soporte no se declara (art. 786) |
| 2 | Entrevistar al declarante | Toda cifra inexplicada = una pregunta explícita |
| 3 | Construir el caso y generar el Excel | Verde = entrada; ÁMBAR = estimado por reemplazar |
| 4 | Contrastar contra borrador DIAN y exógena | El borrador es sugerencia, no verdad |
| 5 | Validar con el motor independiente | Sin paridad 100% no existe "validado" |
| 6 | Veredicto + memoria + entrega | El dictamen manda: no se presenta con bloqueos |

Cree una lista de tareas con estas fases al empezar. Si el usuario solo quiere una parte (p. ej.
"¿estoy obligado a declarar?"), resuelva esa parte con el mismo rigor y ofrezca el resto.

## FASE 0 · Vigencia normativa

Los parámetros cambian TODOS los años: UVT, componente inflacionario, calendario de plazos, y a veces
tablas y topes por reforma tributaria. Antes de calcular:

1. Determine el **año gravable** (el año de los ingresos, normalmente el anterior al actual).
2. Lea `references/parametros.md`: trae los valores del AG2025 como referencia trabajada y la lista
   exacta de qué verificar cada año con sus fuentes oficiales.
3. **Busque en la web** los valores del año que aplica (UVT en la resolución DIAN de diciembre,
   componente inflacionario en el decreto de la ley de plazos, calendario en el decreto de plazos).
   Si algún parámetro difiere de la referencia, actualice `scripts/motor_renta.py` (clase `Params`) y
   la hoja 16_PARAMETROS del caso, y dígalo en la memoria.

Si no puede verificar un parámetro, márquelo como riesgo en el veredicto — no lo esconda.

## FASE 1 · Evidencia (pedir antes de calcular)

Pida al usuario, en este orden de valor probatorio (la **jerarquía de evidencia** que resuelve todo
conflicto): 1.º instructivo oficial del Formulario 210 · 2.º certificados de terceros del año ·
3.º declaraciones presentadas de años anteriores · 4.º información exógena · 5.º estimaciones
(SIEMPRE en ámbar). Cuando dos fuentes pelean, gana la de mayor jerarquía y la diferencia se anota.

**Documentos y cómo obtenerlos** (guía paso a paso para el usuario en `references/fuentes_dian.md`):
- **Información exógena** ("lo que terceros le contaron a la DIAN de usted"): MUISCA → Consulta de
  información reportada por terceros. Es la columna vertebral de la conciliación.
- **Borrador sugerido DIAN** (declaración prellenada): MUISCA → Diligenciar 210 → opción "sugerida".
  Se usa para CONTRASTAR (Fase 4), jamás como fuente primaria.
- Declaración del año anterior (PDF presentado), RUT, certificado 220 del empleador, certificados
  tributarios de cada banco/fondo (rendimientos, retenciones, saldos, GMF), certificado del crédito
  hipotecario (saldo + intereses + % de cada codeudor), predial (avalúo catastral), medicina
  prepagada, certificados AFC/FVP, soportes de dependientes.

Registre cada documento en la hoja 14_SOPORTES con estado (OK/PARCIAL/FALTA) y criticidad. Primero
verifique la **obligación de declarar** (topes en UVT de patrimonio, ingresos, consumos, consignaciones
— en `references/parametros.md`): la exógena trae los topes ya calculados por la DIAN.

## FASE 2 · Entrevista — las preguntas que los documentos no responden

Los errores graves no vienen de fórmulas: vienen de cifras heredadas que nadie explicó. En un caso
real, una fila del año anterior rotulada "Movimientos" resultó ser $155M de FLUJOS declarados como
saldo; en otro, un "apartamento de $565M" eran en realidad dos inmuebles compartidos con porcentajes
distintos. Pregunte SIEMPRE (con AskUserQuestion cuando esté disponible) lo que aplique:

- **Efectivo declarado antes**: ¿de qué se compone? ¿incluía cuentas que ahora aparecen aparte?
- **Inmuebles**: ¿% de propiedad según ESCRITURA? ¿avalúo del año (predial)? ¿costo fiscal? ¿quiénes
  son los otros propietarios y qué declaran ellos? (las declaraciones de una familia deben contar la
  misma historia).
- **Créditos compartidos**: ¿% de cada codeudor? ¿quién paga realmente los intereses?
- **Cuentas que desaparecen o aparecen**: ¿se cerró? ¿a dónde fue el dinero? (retiros grandes sin
  destino documentado = pregunta obligatoria; si fue donación a un familiar, el receptor tiene
  ganancia ocasional).
- **Dependientes**: ¿quiénes, edad, meses, tienen rentas propias? (la dependencia debe ser real).
- **Honorarios**: ¿hubo costos asociados? (define casilla 32 vs 43 y el derecho al 25%).
- **Retención sin ingreso reportado** en exógena: ¿de qué cuenta? — estime el ingreso con la tarifa
  de retención típica (p. ej. rendimientos ≈ retención ÷ 7%), en ÁMBAR, y pida el certificado.
- Saldo a favor del año anterior: ¿se pidió en devolución o se arrastra?

Registre cada respuesta con fecha en la celda de observación correspondiente ("SU RESPUESTA dd/mm").

## FASE 3 · Construcción del Excel

No construya el libro a mano: use el generador probado del paquete.

1. Copie `scripts/caso_demo.py` → `caso_<nombre>.py`. Ese archivo ES el esquema documentado: perfil,
   ingresos (una fila por tercero), beneficios (lista canónica de 36 filas con reglas y topes),
   patrimonio (con columna del año anterior), dependientes, y opcionales (HISTORICO, EXOGENA,
   DECISIONES, CONTROLES_EXTRA con cruces exactos contra certificados).
2. Llene el caso desde la evidencia. Valores en pesos EXACTOS (el redondeo legal a miles del art. 577
   lo aplica el formulario). Estados: `Confirmado` (certificado o dos fuentes), `Provisional` (una
   fuente), `ESTIMADO` (ámbar — inferencia documentada que el usuario DEBE reemplazar),
   `Pendiente/Falta` (no se declara).
3. Genere y recalcule (las fórmulas necesitan un recálculo real):
   ```bash
   python3 scripts/generar_excel.py caso_<nombre>.py Declaracion_Renta_<nombre>.xlsx
   soffice --headless --convert-to xlsx --outdir . Declaracion_Renta_<nombre>.xlsx  # o el recalc del entorno
   ```
4. Reglas de clasificación que más se prestan a error — el detalle completo, casilla por casilla, con
   las 20 trampas conocidas, está en `references/reglas_f210.md` (léalo al clasificar):
   honorarios sin costos → c.32 y participan del 25%; cesantías → fuera de la base del 25% (exención
   propia); arriendos → capital (c.58) aunque la DIAN "sugiera" no laboral; componente inflacionario
   → % del año sobre rendimientos BRUTOS de entidades vigiladas; inmuebles → MAYOR entre costo fiscal
   y avalúo×% (art. 277); deudas compartidas → la parte real de cada codeudor; dependientes → doble
   deducción (387 + 72 UVT) solo con renta laboral; límite 40%/1.340 UVT → casi todo beneficio compite
   por ese techo (fuera: 72 UVT, 1% factura electrónica, RAIS voluntario que es INCRNGO).

La plantilla `assets/Plantilla_Renta_F210_Ejemplo.xlsx` (datos 100% ficticios de "María Ejemplo
Pérez") muestra el resultado esperado: colores, estados, controles, dictamen. Ofrézcala al usuario
como referencia navegable, y úsela usted como patrón de calidad visual.

## FASE 4 · Contraste con el borrador DIAN y la exógena

Construya en la hoja 15_EXOGENA la copia de lo reportado por terceros, y concilie POR TERCERO (nunca
por totales): cada valor del Excel debe (a) coincidir con exógena/borrador, (b) diferir con
explicación escrita, o (c) no estar reportado y tener soporte propio. Diferencias típicas que NO son
errores suyos — el borrador/exógena se equivoca o sugiere mal:

- Ingresos por mandato (inmobiliarias) sugeridos como "no laborales" cuando son arriendos = capital.
- Retenciones reportadas sin el ingreso correspondiente (estime en ámbar, pida certificado).
- Saldos ausentes (entidades bajo el umbral de reporte) o duplicados entre certificado y exógena
  (NUNCA sume dos fuentes del mismo hecho).
- El borrador no incluye deducciones que sí proceden (dependientes, medicina, GMF) ni conoce
  porcentajes de copropiedad.

Todo lo que el borrador traiga y usted excluya, o viceversa, va a la hoja 17_DECISIONES con su
evidencia. Los cruces numéricos exactos (suma de retenciones, de ingresos laborales, de rendimientos
contra certificados) van como CONTROLES_EXTRA para que queden vivos en 12_CONTROLES.

## FASE 5 · Validación (no negociable)

```bash
python3 scripts/validar.py caso_<nombre>.py Declaracion_Renta_<nombre>.xlsx
```

El validador recalcula el libro con LibreOffice, reconstruye el caso y compara las 115 casillas × 2
capas (oficial redondeada + motor exacto) + auxiliares contra `scripts/motor_renta.py` (aritmética
`Decimal`, convención oficial de la tarifa: base en UVT truncada a 2 decimales + épsilon 1E-9).
Exige **diferencia $0**. Además la hoja 13_QA trae 23 pruebas de frontera de la tarifa embebidas que
el propio libro evalúa. Registre fecha y resultado en 13_QA (celda B4).

- Si falla: el reporte dice casilla, capa y ambos valores. En el 95% de los casos es una entrada mal
  digitada — corrija la ENTRADA, jamás una fórmula azul, y vuelva a correr.
- Robustez adicional (recomendada si hay tiempo): repita la validación mutando entradas (±$1 en un
  ingreso, topes excedidos, dependientes 0/2, años de anticipo 1/2/3) — el método completo está en
  `references/validacion.md`.
- Nunca escriba "validado", "verificado" o "100%" sin haber corrido validar.py con éxito EN ESTA
  versión del archivo. Si LibreOffice no está disponible, dígalo y trátelo como NO validado.

## FASE 6 · Veredicto, memoria y entrega

1. **Dictamen** (lo calcula 12_CONTROLES): `NO LISTO` (bloqueos/errores/investigar) → decir qué
   documento o decisión lo desbloquea; `CALCULADO — CIERRE PENDIENTES` (ámbar/revisiones abiertas) →
   listar pendientes priorizados por materialidad; `LISTO PARA REVISIÓN FINAL` → guiar transcripción
   al MUISCA casilla por casilla.
2. **Memoria de decisiones** (Word breve): resumen ejecutivo con cifras clave, composición de cada
   casilla relevante con su fuente, decisiones tomadas con evidencia, pendientes con impacto y
   acción, resultado de la validación (comparaciones, SHA-256), advertencia legal. Sin la memoria el
   trabajo no es auditable.
3. Entregue Excel + memoria + JSON de validación. Termine SIEMPRE con: dictamen en una línea,
   los 3 pendientes más importantes, la fecha límite del declarante (por los dos últimos dígitos del
   NIT) y su nivel de confianza separando matemática (verificada) de hechos (dependen del usuario).

## Dos ejemplos de calibración

**"¿Estoy obligada a declarar? Gané como $60 millones el año pasado."** → Fase 0 + topes: compare
ingresos brutos, patrimonio, consumos, consignaciones contra los topes UVT del año (y recomiende
descargar la exógena, que ya trae los topes medidos por la DIAN). Respuesta directa con norma y
números; ofrezca continuar con la declaración completa.

**"Adjunto mi exógena y mi 220, ayúdame con la renta"** (asalariado, algo de bancos) → flujo completo:
extraer ambos archivos, conciliar 220 vs exógena al peso, entrevistar (dependientes, vivienda,
medicina, otras cuentas, saldo a favor anterior), construir caso, generar, contrastar, validar,
memoria + veredicto. Resultado típico: archivo en `CALCULADO — CIERRE PENDIENTES` con los ámbar
listados — eso es un ÉXITO honesto, no un fracaso: solo el usuario puede cerrar sus hechos.

## Errores reales que este método existe para cazar

| Síntoma | Lo que era en realidad | Defensa del método |
|---|---|---|
| Fila "Movimientos" de $155M en patrimonio | Flujos del año declarados como saldo | Conciliación saldo-vs-movimiento por cuenta + HISTORICO |
| "Mi apartamento vale $565M" | Dos inmuebles, 50% y 18%, con costos distintos | Pregunta de escritura + art. 277 por fila |
| Retención de $64.007 sin ingreso | Cuenta nueva no reportada en exógena | Estimación ámbar (ret ÷ 7%) — el certificado real difirió solo 0,39% |
| Saldos "del año pasado" arrastrados | Cuentas ya vaciadas o cerradas | Certificado de saldo por cuenta, cada año |
| AFC del empleador ≠ AFC del banco | Dos cuentas AFC sumaban el total exacto | Conciliar antes de elegir; nunca sumar fuentes duplicadas |
| Más certificados "para pagar menos" | Límite 40%/1.340 UVT ya copado: recorte | Hoja 11_LIMITE_40 muestra el recorte y las palancas reales |

## Recursos del paquete

| Archivo | Cuándo usarlo |
|---|---|
| `scripts/caso_demo.py` | SIEMPRE como esquema: cópielo y llénelo con el caso real |
| `scripts/generar_excel.py` | Generar el libro de 18 hojas desde el caso |
| `scripts/motor_renta.py` | Motor Decimal (única fuente de verdad del cálculo); actualice `Params` cada año |
| `scripts/validar.py` | Validación obligatoria de paridad Excel↔motor |
| `assets/Plantilla_Renta_F210_Ejemplo.xlsx` | Referencia visual navegable con datos ficticios |
| `references/parametros.md` | Valores del año, ritual de actualización anual, fuentes oficiales |
| `references/reglas_f210.md` | Mapa de casillas, clasificación cedular y las 20 trampas |
| `references/validacion.md` | Método de validación completo y pruebas de estrés |
| `references/fuentes_dian.md` | Guías paso a paso para el usuario: exógena, borrador, MUISCA, devoluciones |
