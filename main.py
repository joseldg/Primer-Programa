import argparse
import sys
from pathlib import Path

try:
    from docx2pdf import convert
except ImportError:
    print("Error: falta la dependencia docx2pdf. Instala con 'pip install -r requirements.txt'.")
    sys.exit(1)


def convertir_docx_a_pdf(origen: Path, destino: Path | None = None) -> None:
    """Convierte un archivo .docx a PDF."""
    if not origen.exists():
        raise FileNotFoundError(f"No existe el archivo de origen: {origen}")
    if origen.suffix.lower() not in {".docx", ".doc"}:
        raise ValueError("El archivo de origen debe ser .docx o .doc")

    if destino is None:
        destino = origen.with_suffix(".pdf")

    if destino.exists():
        print(f"Aviso: el archivo de destino ya existe y se sobrescribirá: {destino}")

    convert(str(origen), str(destino))
    print(f"Conversión completada: {destino}")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Convierte un archivo Word (.docx/.doc) a PDF usando docx2pdf en Windows."
    )
    parser.add_argument(
        "origen",
        type=Path,
        help="Ruta del archivo Word de entrada (.docx o .doc)."
    )
    parser.add_argument(
        "destino",
        type=Path,
        nargs="?",
        help="Ruta opcional del PDF de salida. Si no se indica, se usa el mismo nombre con extensión .pdf."
    )

    args = parser.parse_args()

    try:
        convertir_docx_a_pdf(args.origen, args.destino)
    except Exception as error:
        print(f"Error: {error}")
        sys.exit(1)


if __name__ == "__main__":
    main()
