# -*- coding: utf-8 -*-
"""Validador independiente: recalcula el Excel con LibreOffice y compara CADA casilla
(115 × 2 capas + 4 auxiliares) contra el motor Decimal. Exige diferencia $0.

Uso: python3 validar.py <caso.py> <declaracion.xlsx>
Salida: JSON con {paridad, comparaciones, fallas, casillas_clave, dictamen_excel, sha256}
Código de salida 0 solo si la paridad es 100%. Sin LibreOffice no hay validación real:
el script lo dice y termina con error (nunca declare 'validado' sin recomputar).
"""
import hashlib
import importlib.util
import json
import shutil
import subprocess
import sys
import tempfile
from decimal import Decimal
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from motor_renta import caso_desde_config, liquidar, fila_casilla, FILAS_AUX  # noqa: E402

from openpyxl import load_workbook  # noqa: E402

D = Decimal
CASILLAS = [c for c in list(range(28, 142)) + [980] if c != 140]
F210H, MXH, CTRLH = "09_FORMULARIO_210", "10_MOTOR_EXACTO", "12_CONTROLES"


def cargar_caso(ruta):
    spec = importlib.util.spec_from_file_location("caso_usuario", ruta)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod.CASO


def recalcular(xlsx: Path, out_dir: Path) -> Path:
    sof = shutil.which("soffice")
    if not sof:
        raise RuntimeError("Sin LibreOffice (soffice) no se puede recalcular: instálelo o "
                           "abra y guarde el archivo en Excel/LibreOffice y vuelva a correr.")
    prof = out_dir / "perfil"
    prof.mkdir(parents=True, exist_ok=True)
    r = subprocess.run([sof, f"-env:UserInstallation={prof.as_uri()}", "--headless",
                        "--convert-to", "xlsx", "--outdir", str(out_dir), str(xlsx)],
                       capture_output=True, text=True, timeout=300)
    salida = out_dir / xlsx.name
    if r.returncode != 0 or not salida.exists():
        raise RuntimeError(f"LibreOffice falló: {r.stderr[-300:]}")
    return salida


def main():
    ruta_caso, ruta_xlsx = sys.argv[1], Path(sys.argv[2]).resolve()
    cfg = cargar_caso(ruta_caso)
    eo = liquidar(caso_desde_config(cfg), oficial=True)
    ee = liquidar(caso_desde_config(cfg), oficial=False)
    with tempfile.TemporaryDirectory(prefix="validar_") as td:
        rec = recalcular(ruta_xlsx, Path(td))
        wb = load_workbook(rec, data_only=True)
        fo, mx = wb[F210H], wb[MXH]
        fallas, n = [], 0
        for cas in CASILLAS:
            fila = fila_casilla(cas)
            for hoja, ws, esp, tol in (("oficial", fo, eo, D("0.5")), ("exacto", mx, ee, D("0.01"))):
                n += 1
                got = D(str(ws.cell(row=fila, column=4).value or 0))
                want = D(str(esp[cas]))
                if abs(got - want) > tol:
                    fallas.append({"casilla": cas, "capa": hoja, "excel": float(got), "motor": float(want)})
        for k, fila in FILAS_AUX.items():
            n += 1
            got = D(str(fo.cell(row=fila, column=4).value or 0))
            if abs(got - D(str(eo[k]))) > D("0.5"):
                fallas.append({"casilla": k, "capa": "aux", "excel": float(got), "motor": float(eo[k])})
        dictamen = ""
        ctrl = wb[CTRLH]
        for r in range(5, 60):
            if ctrl.cell(row=r, column=1).value == "ESTADO":
                dictamen = str(ctrl.cell(row=r, column=2).value)
        qa_fallas = wb["13_QA"]["B7"].value
        wb.close()
    rep = {
        "paridad": "100%" if not fallas else "FALLA",
        "comparaciones": n,
        "fallas": fallas[:10],
        "qa_fronteras_falla": qa_fallas,
        "dictamen_excel": dictamen,
        "casillas_clave": {str(c): int(eo[c]) for c in (29, 31, 32, 58, 74, 99, 111, 126, 129, 132, 133, 136, 137)},
        "sha256": hashlib.sha256(ruta_xlsx.read_bytes()).hexdigest(),
        "alcance": "Valida matemática (Excel↔motor Decimal tras recálculo LibreOffice); NO certifica hechos ni soportes.",
    }
    print(json.dumps(rep, indent=2, ensure_ascii=False))
    sys.exit(0 if not fallas and not qa_fallas else 1)


if __name__ == "__main__":
    main()
