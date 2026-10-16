# -*- coding: utf-8 -*-
"""Motor de liquidación Renta AG2025 v3 — fuente única de verdad (Decimal).

Replica la norma (E.T. arts. 26, 55-56, 103, 119, 126-1/4, 206, 241, 336, 387,
577, 807; Decreto 898/2026; instructivo Res. DIAN 000044/2024) y la convención
de redondeo del AyudaRenta oficial (TRUNC 2 decimales en UVT + épsilon 1E-9).
Lo usan: el generador del Excel (valores esperados de QA), el verificador y el
notebook auditor. Dos capas: oficial (casillas a miles) y exacta (sin redondeo).
"""
from __future__ import annotations

from dataclasses import dataclass, field
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP, getcontext
from typing import Any, Optional

getcontext().prec = 40
EPS = Decimal("0.000000001")


def D(x: Any) -> Decimal:
    if x is None or x == "":
        return Decimal(0)
    if isinstance(x, Decimal):
        return x
    if isinstance(x, bool):
        return Decimal(int(x))
    return Decimal(str(x))


def r1000(x: Any) -> Decimal:
    """Art. 577 E.T.: aproximación al múltiplo de mil más cercano."""
    return (D(x) / 1000).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * 1000


def trunc2(x: Any) -> Decimal:
    return D(x).quantize(Decimal("0.01"), rounding=ROUND_DOWN)


def si(v: Any) -> bool:
    return str(v or "").strip().lower() in {"sí", "si", "yes", "true", "1"}


@dataclass(frozen=True)
class Params:
    """PARAMETROS C5..C27 (mismos códigos de la hoja)."""

    uvt: Decimal = Decimal(49_799)            # C5  UVT_AG
    uvt_pres: Decimal = Decimal(52_374)       # C6  UVT_PRES
    lim_global: Decimal = Decimal(1_340)      # C7
    lim_25: Decimal = Decimal(790)            # C8
    lim_afc: Decimal = Decimal(3_800)         # C9
    lim_viv: Decimal = Decimal(1_200)         # C10
    lim_icetex: Decimal = Decimal(100)        # C11
    lim_salud_mes: Decimal = Decimal(16)      # C12
    lim_dep_mes: Decimal = Decimal(32)        # C13
    dep_adic: Decimal = Decimal(72)           # C14
    max_dep: Decimal = Decimal(4)             # C15
    lim_fe: Decimal = Decimal(240)            # C16
    sanc_min: Decimal = Decimal(10)           # C17
    tope_pat: Decimal = Decimal(4_500)        # C18
    tope_ing: Decimal = Decimal(1_400)        # C19
    tope_cont: Decimal = Decimal(100_000)     # C20
    comp_inf: Decimal = Decimal("0.5543")     # C21 Decreto 898/2026
    tar_go: Decimal = Decimal("0.15")         # C22
    tar_lot: Decimal = Decimal("0.20")        # C23
    tar_div2: Decimal = Decimal("0.35")       # C24
    ren_pre: Decimal = Decimal(0)             # C25
    comp_inf_gastos: Decimal = Decimal("0.2835")  # C26 Decreto 898/2026 (art. 81/118)
    lim_rais: Decimal = Decimal(2_500)        # C27
    tabla241: tuple = (
        (Decimal(0), Decimal(1090), Decimal(0), Decimal(0)),
        (Decimal(1090), Decimal(1700), Decimal("0.19"), Decimal(0)),
        (Decimal(1700), Decimal(4100), Decimal("0.28"), Decimal(116)),
        (Decimal(4100), Decimal(8670), Decimal("0.33"), Decimal(788)),
        (Decimal(8670), Decimal(18970), Decimal("0.35"), Decimal(2296)),
        (Decimal(18970), Decimal(31000), Decimal("0.37"), Decimal(5901)),
        (Decimal(31000), Decimal("999999999"), Decimal("0.39"), Decimal(10352)),
    )
    tabla_div2016: tuple = (
        (Decimal(0), Decimal(1090), Decimal(0), Decimal(0)),
        (Decimal(1090), Decimal(1700), Decimal("0.19"), Decimal(0)),
        (Decimal(1700), Decimal(4100), Decimal("0.28"), Decimal(116)),
        (Decimal(4100), Decimal("999999999"), Decimal("0.33"), Decimal(788)),
    )
    # Instructivo casilla 36 num. 4: parte NO gravada de cesantías según salario
    # mensual promedio de los últimos 6 meses (en UVT).
    tabla_cesantias: tuple = (
        (Decimal(0), Decimal(350), Decimal(1)),
        (Decimal(350), Decimal(410), Decimal("0.9")),
        (Decimal(410), Decimal(470), Decimal("0.8")),
        (Decimal(470), Decimal(530), Decimal("0.6")),
        (Decimal(530), Decimal(590), Decimal("0.4")),
        (Decimal(590), Decimal(650), Decimal("0.2")),
        (Decimal(650), Decimal("999999999"), Decimal(0)),
    )


@dataclass
class Fila:
    categoria: str = ""
    bruto: Decimal = Decimal(0)
    costos: Decimal = Decimal(0)
    retencion: Decimal = Decimal(0)
    incluir: bool = False


@dataclass
class Beneficio:
    base: Decimal = Decimal(0)
    d: Decimal = Decimal(0)     # meses / % / cantidad
    aplicar: bool = False


@dataclass
class PartidaPat:
    tipo: str = "Activo"        # Activo | Deuda
    reportado: Decimal = Decimal(0)
    participacion: Decimal = Decimal(1)
    alterno: Decimal = Decimal(0)
    incluir: bool = False


@dataclass
class Dependiente:
    meses: int = 0
    art387: bool = False
    u72: bool = False
    laboral: bool = False       # DUR 1.2.1.20.3: concurrencia exige renta laboral


@dataclass
class Caso:
    """Entradas completas (espejo de PERFIL / INGRESOS / BENEFICIOS /
    DEPENDIENTES / PATRIMONIO)."""

    ingresos: list = field(default_factory=list)          # list[Fila]
    ben: dict = field(default_factory=dict)               # id(5..40) -> Beneficio
    esp: dict = field(default_factory=dict)               # id(45..65) -> (valor, aplicar)
    patrimonio: list = field(default_factory=list)        # list[PartidaPat]
    dependientes: list = field(default_factory=list)      # list[Dependiente]
    salario_prom_6m_uvt: Decimal = Decimal(0)             # para tabla cesantías
    contabilidad: bool = False                            # PERFIL B18 == "Sí"
    anios: int = 3                                        # PERFIL B19
    neto_2024: Decimal = Decimal(0)                       # PERFIL B20
    saldo_favor_ant: Decimal = Decimal(0)                 # PERFIL B21
    anticipo_ant: Decimal = Decimal(0)                    # PERFIL B22
    usar_faltante: bool = False                           # PATRIMONIO toggle
    pat_bruto_2024: Decimal = Decimal(0)                  # PERFIL B23

    def suma(self, cats: Optional[set] = None, campo: str = "bruto") -> Decimal:
        t = Decimal(0)
        for f in self.ingresos:
            if f.incluir and (cats is None or f.categoria in cats):
                t += D(getattr(f, campo))
        return t


def _tarifa(base_pesos: Decimal, tabla: tuple, uvt: Decimal, oficial: bool) -> Decimal:
    """Tarifa marginal por tabla (base a UVT TRUNC 2 dec, épsilon anti-flotante)."""
    b = trunc2(max(Decimal(0), D(base_pesos)) / uvt)
    for desde, hasta, tar, fija in tabla:
        if b <= hasta:
            raw = Decimal(0) if tar == 0 else (trunc2((b - desde) * tar + EPS) + fija) * uvt
            return r1000(raw) if oficial else raw
    raise AssertionError("base fuera de tabla")


def beneficios_aceptados(caso: Caso, p: Params) -> dict:
    """Columna F de BENEFICIOS: valor aceptado por regla, antes de límites cedulares."""
    b: dict[int, Decimal] = {}
    g = lambda i: caso.ben.get(i, Beneficio())
    ing_total = caso.suma()
    laboral = caso.suma({"Trabajo - laboral"})
    rend = caso.suma({"Capital - rendimientos"})

    b[5] = g(5).base if g(5).aplicar else Decimal(0)                       # salud
    b[6] = g(6).base if g(6).aplicar else Decimal(0)                       # pensión
    b[7] = min(g(7).base, Decimal("0.25") * ing_total, p.lim_rais * p.uvt) if g(7).aplicar else Decimal(0)
    b[8] = g(8).base if g(8).aplicar else Decimal(0)                       # otros INCR lab
    b[9] = g(9).base if g(9).aplicar else Decimal(0)                       # INCR con costos
    # Componente inflacionario: 55,43% de los rendimientos BRUTOS elegibles
    # (arts. 38 y 40-1 E.T.; Decreto 898/2026). Solo no obligados a contabilidad.
    b[10] = rend * g(10).d if (g(10).aplicar and not caso.contabilidad) else Decimal(0)
    b[13] = g(13).base if g(13).aplicar else Decimal(0)                    # AFC seleccionado
    b[14] = g(14).base if g(14).aplicar else Decimal(0)                    # FVP
    b[15] = g(15).base if g(15).aplicar else Decimal(0)                    # AVC
    b[16] = g(16).base if g(16).aplicar else Decimal(0)                    # otras exentas lab
    b[17] = min(g(17).base, p.lim_viv * p.uvt) if g(17).aplicar else Decimal(0)
    b[18] = min(g(18).base, p.lim_icetex * p.uvt) if g(18).aplicar else Decimal(0)
    b[19] = min(g(19).base, g(19).d * p.lim_salud_mes * p.uvt) if g(19).aplicar else Decimal(0)
    b[20] = g(20).base * g(20).d if g(20).aplicar else Decimal(0)          # GMF 50%
    deps387 = [x for x in caso.dependientes if x.art387 and x.laboral]
    deps72 = [x for x in caso.dependientes if x.u72 and x.laboral]
    meses387 = max([Decimal(x.meses) for x in deps387] or [Decimal(0)])
    b[21] = min(Decimal("0.10") * laboral, meses387 * p.lim_dep_mes * p.uvt) if deps387 else Decimal(0)
    b[22] = min(Decimal(len(deps72)), p.max_dep) * p.dep_adic * p.uvt if deps72 else Decimal(0)
    b[23] = min(g(23).base * g(23).d, p.lim_fe * p.uvt) if g(23).aplicar else Decimal(0)
    for i in range(24, 41):
        b[i] = g(i).base if g(i).aplicar else Decimal(0)
    b[37] = (Decimal("0.25") * g(37).base) if g(37).aplicar else Decimal(0)  # donaciones 25%
    meta = {
        "cesantias": caso.suma({"Trabajo - cesantías"}),
        "dep138": Decimal(sum(1 for x in caso.dependientes if (x.art387 or x.u72) and x.laboral)),
        "rend": rend,
        "costos_fin": caso.suma({"Capital - rendimientos"}, "costos"),
    }
    return {"b": b, "meta": meta}


def liquidar(caso: Caso, p: Params = Params(), oficial: bool = True) -> dict:
    """Devuelve dict casilla->valor. oficial=True redondea a miles (art. 577)."""
    R = (lambda x: r1000(x)) if oficial else (lambda x: D(x))
    acc = beneficios_aceptados(caso, p)
    b, meta = acc["b"], acc["meta"]
    e = lambda i: caso.esp.get(i, (Decimal(0), False))
    ev = lambda i: e(i)[0] if e(i)[1] else Decimal(0)
    c: dict[Any, Any] = {}

    # --- patrimonio -------------------------------------------------------
    act = sum((max(x.reportado * x.participacion, x.alterno)
               for x in caso.patrimonio if x.incluir and x.tipo == "Activo"), Decimal(0))
    if caso.usar_faltante:
        act += max(Decimal(0), caso.pat_bruto_2024 - act)
    deu = sum((x.reportado * x.participacion if x.alterno == 0 else x.alterno)
              for x in caso.patrimonio if x.incluir and x.tipo == "Deuda")
    c[29], c[30] = R(act), R(deu)
    c[31] = max(Decimal(0), c[29] - c[30])

    # --- cédula general: trabajo (32-42) ---------------------------------
    lab_total = caso.suma({"Trabajo - laboral", "Trabajo - cesantías", "Trabajo - sin costos"})
    c[32] = R(lab_total)
    c[33] = R(b[5] + b[6] + b[7] + b[8])
    c[34] = max(Decimal(0), c[32] - c[33])
    ing_bruto_total = caso.suma()
    tope_afc_total = min(b[13] + b[14] + b[15] + b[24] + b[25] + b[26],
                         Decimal("0.30") * ing_bruto_total, p.lim_afc * p.uvt)
    c[35] = R(min(b[13] + b[14] + b[15], tope_afc_total))
    c[38] = R(min(c[34], b[17]))
    c[39] = R(min(c[34], b[18] + b[19] + b[20] + b[21]))
    c[40] = c[38] + c[39]
    # Cesantías: parte no gravada por tabla num. 4 (instructivo casilla 36);
    # el 25% NO se otorga sobre cesantías (instructivo, num. 9).
    ces = meta["cesantias"]
    pct_ces = next(pc for d0, h0, pc in p.tabla_cesantias
                   if caso.salario_prom_6m_uvt <= h0)
    ces_exenta = min(ces, ces * pct_ces)
    fuera = ev(65)                                            # exentas no sometidas al límite
    base25 = max(Decimal(0), c[34] - c[35] - c[38] - c[39] - b[16] - fuera - R(ces))
    c[36] = R(min(c[34] - c[35],
                  max(Decimal(0), b[16] + fuera + R(ces_exenta))
                  + min(p.lim_25 * p.uvt, Decimal("0.25") * base25)))
    c[37] = c[35] + c[36]

    # --- límite 40% / 1.340 UVT (art. 336-3) ------------------------------
    con_costos = caso.suma({"Trabajo - con costos"})
    cap_total = caso.suma({"Capital - rendimientos", "Capital - otros"})
    nolab = caso.suma({"No laboral"})
    c[43] = R(con_costos)
    c[44] = R(min(b[9], con_costos))
    c[45] = R(caso.suma({"Trabajo - con costos"}, "costos"))
    c[46] = max(Decimal(0), c[43] - c[44] - c[45])
    c[58] = R(cap_total)
    c[59] = R(b[10])
    # Costos financieros: no deducible el componente inflacionario de gastos
    # (28,35%) para no obligados a contabilidad (art. 81/118 E.T.; Dcto. 898/2026)
    cf = meta["costos_fin"] + caso.suma({"Capital - otros"}, "costos")
    cf_deducible = cf * (1 - p.comp_inf_gastos) if (not caso.contabilidad and cf > 0) else cf
    c[60] = R(cf_deducible)
    c[61] = max(Decimal(0), c[58] - c[59] - c[60])
    c[62] = R(ev(49))
    c[74] = R(nolab)
    c[75] = R(min(c[74], ev(54)))
    c[76] = R(min(max(Decimal(0), c[74] - c[75]), ev(55)))
    c[77] = R(caso.suma({"No laboral"}, "costos"))
    c[78] = max(Decimal(0), c[74] - c[75] - c[76] - c[77])
    c[79] = R(ev(56))

    base_lim = max(Decimal(0), c[32] + c[43] + c[58] + c[74] + c[62] + c[79]
                   - c[33] - c[44] - c[59] - c[76])
    lim = r1000(min(Decimal("0.40") * base_lim, p.lim_global * p.uvt))
    c["L_base"], c["L_lim"] = base_lim, lim

    # asignación en cascada del límite y de los topes AFC / vivienda
    c[41] = R(min(c[34], fuera + min(max(Decimal(0), c[34] - fuera),
                                     max(Decimal(0), c[37] + c[40] - fuera), lim)))
    c[42] = max(Decimal(0), c[34] - c[41])
    usado_t = max(Decimal(0), c[41] - fuera)
    c[47] = R(min(b[24], max(Decimal(0), tope_afc_total - c[35])))
    c[48] = R(ev(45))
    c[49] = c[47] + c[48]
    c[50] = R(min(c[46], ev(46), max(Decimal(0), p.lim_viv * p.uvt - c[38])))
    c[51] = R(min(c[46], ev(47)))
    c[52] = c[50] + c[51]
    c[53] = R(min(c[46], c[49] + c[52], max(Decimal(0), lim - usado_t)))
    c[54] = max(Decimal(0), c[43] - c[44] - c[45] - c[53])
    c[55] = max(Decimal(0), c[44] + c[45] - c[43])
    c[56] = R(min(c[54], ev(48)))
    c[57] = max(Decimal(0), c[54] - c[56])
    c[63] = R(min(b[25], max(Decimal(0), tope_afc_total - c[35] - c[47])))
    c[64] = R(ev(50))
    c[65] = c[63] + c[64]
    c[66] = R(min(c[61] + c[62], ev(51), max(Decimal(0), p.lim_viv * p.uvt - c[38] - c[50])))
    c[67] = R(min(c[61] + c[62], ev(52)))
    c[68] = c[66] + c[67]
    c[69] = R(min(c[61] + c[62], c[65] + c[68], max(Decimal(0), lim - usado_t - c[53])))
    c[70] = max(Decimal(0), c[58] + c[62] - c[59] - c[60] - c[69])
    c[71] = max(Decimal(0), c[59] + c[60] - c[58] - c[62])
    c[72] = R(min(c[70], ev(53)))
    c[73] = max(Decimal(0), c[70] - c[72])
    c[80] = R(min(b[26], max(Decimal(0), tope_afc_total - c[35] - c[47] - c[63])))
    c[81] = R(ev(57))
    c[82] = c[80] + c[81]
    c[83] = R(min(c[78] + c[79], ev(58), max(Decimal(0), p.lim_viv * p.uvt - c[38] - c[50] - c[66])))
    c[84] = R(min(c[78] + c[79], ev(59)))
    c[85] = c[83] + c[84]
    c[86] = R(min(c[78] + c[79], c[82] + c[85], max(Decimal(0), lim - usado_t - c[53] - c[69])))
    c[87] = max(Decimal(0), c[74] - c[75] - c[76] - c[77] + c[79] - c[86])
    c[88] = max(Decimal(0), c[75] + c[76] + c[77] - c[74] - c[79])
    c[89] = R(min(c[87], ev(60)))
    c[90] = max(Decimal(0), c[87] - c[89])

    # --- resumen cédula general ------------------------------------------
    c[138] = meta["dep138"]
    c[139] = R(min(max(Decimal(0), c[41] + c[42] + c[53] + c[57]), b[22]))
    rlo = c[42] + c[57] + c[73] + c[90]
    c[28] = R(b[23]) if rlo > 0 else Decimal(0)               # 336-5: sin pérdidas
    c[91] = max(Decimal(0), c[41] + c[42] + c[53] + c[57] + c[69] + c[73] + c[86] + c[90])
    c[92] = min(c[91], c[28] + c[41] + c[53] + c[69] + c[86] + c[139])
    c[93] = max(Decimal(0), c[91] - c[92])
    c[94], c[95], c[96] = R(b[27]), R(b[28]), R(b[29])
    c[97] = max(Decimal(0), c[93] - c[94] - c[95] + c[96])
    c[98] = R(max(Decimal(0), c[31] * p.ren_pre + ev(64)))

    # --- pensiones y dividendos ------------------------------------------
    c[99] = R(caso.suma({"Pensión"}))
    c[100] = R(b[30])
    c[101] = max(Decimal(0), c[99] - c[100])
    c[102] = R(min(c[101], b[31] * 1000 * p.uvt))
    c[103] = max(Decimal(0), c[101] - c[102])
    c[104] = R(b[32])
    c[105] = R(min(c[104], ev(61)))
    c[106] = max(Decimal(0), c[104] - c[105])
    c[107] = R(caso.suma({"Dividendo 1"}))
    c[108] = R(caso.suma({"Dividendo 2"}))
    c[109] = R(ev(62))
    c[110] = R(min(c[109], ev(63)))
    c[118] = R(c[108] * p.tar_div2)
    c[111] = R(max(c[97], c[98]) + c[103] + c[107] + c[108] - c[118])

    # --- ganancias ocasionales -------------------------------------------
    lot = caso.suma({"Lotería"})
    c[112] = R(caso.suma({"Ganancia ocasional"}) + lot)
    c[113] = R(b[33])
    c[114] = R(min(max(Decimal(0), c[112] - c[113]), b[34]))
    c[115] = max(Decimal(0), c[112] - c[113] - c[114])

    # --- liquidación ------------------------------------------------------
    c[116] = _tarifa(c[111], p.tabla241, p.uvt, oficial) if c[97] >= c[98] else Decimal(0)
    c[117] = _tarifa(c[111], p.tabla241, p.uvt, oficial) if c[98] > c[97] else Decimal(0)
    c[119] = _tarifa(c[106], p.tabla_div2016, p.uvt, oficial)
    c[120] = R(max(Decimal(0), c[109] - c[110]) * Decimal("0.35"))
    c[121] = c[116] + c[117] + c[118] + c[119] + c[120]
    c[122] = R(b[36])
    c[123] = R(b[37])
    bd = trunc2(max(Decimal(0), c[107] + c[108] - c[118]) / p.uvt)
    d254 = Decimal(0) if bd <= 1090 else trunc2((bd - 1090) * Decimal("0.19") + EPS) * p.uvt
    c[124] = R(d254 + b[38])
    c[125] = min(c[122] + c[123] + c[124], c[121])
    c[126] = max(Decimal(0), c[121] - c[125])
    lot_grav = min(c[115], R(lot) if oficial else lot)
    c[127] = R((c[115] - lot_grav) * p.tar_go + lot_grav * p.tar_lot)
    c[128] = R(min(c[127], b[35]))
    c[129] = max(Decimal(0), c[126] + c[127] - c[128])
    c[130] = R(caso.anticipo_ant)
    c[131] = R(caso.saldo_favor_ant)
    c[132] = R(caso.suma(None, "retencion"))
    pct = Decimal("0.25") if caso.anios == 1 else Decimal("0.50") if caso.anios == 2 else Decimal("0.75")
    op1 = max(Decimal(0), r1000(pct * c[126] - c[132]))
    op2 = op1 if caso.anios == 1 else max(Decimal(0), r1000(pct * (c[126] + caso.neto_2024) / 2 - c[132]))
    c["A_op1"], c["A_op2"] = op1, op2
    c[133] = min(op1, op2)
    c[134] = max(Decimal(0), c[129] + c[133] - c[130] - c[131] - c[132])
    c[135] = R(b[39])
    c[136] = max(Decimal(0), c[129] + c[133] + c[135] - c[130] - c[131] - c[132])
    c[137] = max(Decimal(0), c[130] + c[131] + c[132] - c[129] - c[133] - c[135])
    c[140] = "" if c[43] == 0 else ("X" if c[45] / c[43] > Decimal("0.60") else "")
    c[141] = R(b[40])
    c[980] = c[136] + c[141]
    return c


# ---------------------------------------------------------------- fila por casilla
def fila_casilla(casilla: int) -> int:
    """Mapa casilla -> fila en FORMULARIO_210 y MOTOR_EXACTO (compatible v2)."""
    t = [(28, 31, -22), (32, 42, -20), (43, 57, -18), (58, 73, -16), (74, 90, -14),
         (91, 98, -10), (99, 103, -8), (104, 111, -6), (112, 115, -4),
         (116, 129, -2), (130, 132, 0), (133, 141, 2)]
    for lo, hi, off in t:
        if lo <= casilla <= hi:
            return casilla + off
    if casilla == 980:
        return 145
    raise ValueError(casilla)


FILAS_AUX = {"L_base": 79, "L_lim": 80, "A_op1": 133, "A_op2": 134}
FILAS_TODAS = sorted({fila_casilla(x) for x in list(range(28, 142)) + [980]}
                     | set(FILAS_AUX.values()))


# ============================================================================
# Puente genérico: construye un Caso del motor desde el dict CASO de un
# archivo de caso (véase caso_demo.py). Lo usan validar.py y los auditores.
# ============================================================================
def caso_desde_config(cfg):
    from decimal import Decimal as _D
    caso = Caso()
    for r in cfg["INGRESOS"]:
        caso.ingresos.append(Fila(r[0], _D(str(r[4])), _D(str(r[5])), _D(str(r[6])), r[7] == "Sí"))
    for r in cfg["BENEFICIOS"]:
        i = r[0]
        base, dd = r[2], r[3]
        if i == 10:
            base = caso.suma({"Capital - rendimientos"})
        if i == 13:
            base = cfg.get("AFC_SELECCION", 0); dd = 0
        if i in (21, 22):
            continue
        caso.ben[i] = Beneficio(_D(str(base or 0)), _D(str(dd or 0)), r[4] == "Sí")
    esp = cfg.get("ESPECIALES", {})
    for i in range(45, 66):
        v, ap = esp.get(i, (0, False))
        caso.esp[i] = (_D(str(v)), bool(ap))
    for r in cfg["PATRIMONIO"]:
        caso.patrimonio.append(PartidaPat(r[0], _D(str(r[3])), _D(str(r[4])), _D(str(r[5])), r[6] == "Sí"))
    for r in cfg["DEPENDIENTES"]:
        caso.dependientes.append(Dependiente(int(r[2]), r[3] == "Sí", r[4] == "Sí", r[5] == "Sí"))
    caso.salario_prom_6m_uvt = _D(str(cfg.get("salario_uvt", 0)))
    caso.contabilidad = False
    caso.anios = cfg["anios"]
    caso.neto_2024 = _D(str(cfg["neto_2024"]))
    caso.saldo_favor_ant = _D(str(cfg["saldo_favor_ant"]))
    caso.anticipo_ant = _D(str(cfg["anticipo_ant"]))
    caso.pat_bruto_2024 = _D(str(cfg["pat_bruto_2024"]))
    return caso
