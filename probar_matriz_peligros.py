"""Script de prueba aislada para la Matriz de Peligros IPVR.

Uso:
    python probar_matriz_peligros.py <ruta_yaml_cliente> [<carpeta_manuales>]

Ejemplos:
    # Usa la carpeta definida en manuales_cargos_dir del YAML
    python probar_matriz_peligros.py clientes_config/CRISTOVISION.yaml

    # Override de la carpeta de manuales
    python probar_matriz_peligros.py clientes_config/CRISTOVISION.yaml "C:/Users/bts19/Downloads/MANUALES CRISTOVISION"

Solo copia la plantilla SST-MTZ-012, la llena segun los manuales/cargos del YAML
y la guarda en ./prueba_matriz_peligros/SST-MTZ-012 MATRIZ DE IPVR.xlsx
sin tocar el resto del SG-SST.
"""
import sys
import shutil
from pathlib import Path

import yaml

from engine.matriz_peligros import generar_matriz_peligros, RUTA_MATRIZ_IPVR
from engine.config import PLANTILLA_PATH


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    yaml_path = Path(sys.argv[1])
    if not yaml_path.exists():
        print(f"[ERROR] No existe el YAML: {yaml_path}")
        sys.exit(1)

    with open(yaml_path, "r", encoding="utf-8") as f:
        cliente = yaml.safe_load(f)

    # Override opcional de la carpeta de manuales
    if len(sys.argv) >= 3:
        cliente["manuales_cargos_dir"] = sys.argv[2]

    print(f"=== PRUEBA MATRIZ DE PELIGROS ===")
    print(f"Cliente: {cliente.get('razon_social', 'N/A')}")
    print(f"Manuales: {cliente.get('manuales_cargos_dir', '(ninguno)')}")
    print(f"Cargos en YAML: {len(cliente.get('cargos') or [])}")

    # Copiar plantilla a carpeta temporal
    output_root = Path("prueba_matriz_peligros").resolve()
    if output_root.exists():
        shutil.rmtree(output_root)
    plantilla_src = Path(PLANTILLA_PATH) / RUTA_MATRIZ_IPVR
    if not plantilla_src.exists():
        print(f"[ERROR] No se encontro la plantilla: {plantilla_src}")
        sys.exit(1)
    destino = output_root / RUTA_MATRIZ_IPVR
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(str(plantilla_src), str(destino))
    print(f"\nPlantilla copiada a: {destino}")

    # Generar
    print(f"\nGenerando matriz... (puede tardar 30-60s con Claude)\n")
    ruta = generar_matriz_peligros(
        cliente,
        output_root,
        config_dir=yaml_path.parent,
    )

    if ruta:
        print(f"\n[OK] Matriz generada en:\n  {ruta}\n")
        print(f"Abrela con Excel para revisar el resultado.")
    else:
        print("\n[ERROR] No se pudo generar la matriz. Revisa los logs anteriores.")


if __name__ == "__main__":
    import logging
    logging.basicConfig(level=logging.INFO, format="%(levelname)s | %(message)s")
    main()
