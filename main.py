#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aplicación de consola para convertir archivos Excel .xls a .xlsx."""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import List, Optional

from colorama import Fore, Style, init as colorama_init

from converter import ConvertResult, convert_many, discover_xls, is_xls_file, summarize

APP_NAME = "XLS -> XLSX CONVERTER"
APP_VERSION = "1.0.0"


def main(argv: Optional[List[str]] = None) -> int:
    colorama_init(autoreset=True)
    _enable_windows_vt()
    args = _parse_args(argv if argv is not None else sys.argv[1:])

    if args.path:
        return _run_batch(args.path, recursive=args.recursive, overwrite=args.overwrite, output=args.output)

    return _run_interactive()


def _parse_args(argv: List[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Convierte archivos Excel .xls (formato antiguo) a .xlsx.",
    )
    parser.add_argument("path", nargs="?", help="Archivo .xls o carpeta a convertir")
    parser.add_argument("-r", "--recursive", action="store_true", help="Incluir subcarpetas")
    parser.add_argument("-y", "--overwrite", action="store_true", help="Sobrescribir .xlsx existentes")
    parser.add_argument("-o", "--output", help="Carpeta de salida (opcional)")
    return parser.parse_args(argv)


def _run_batch(raw_path: str, recursive: bool, overwrite: bool, output: Optional[str]) -> int:
    _print_banner()
    target = Path(raw_path).expanduser()
    if not target.exists():
        _error("No existe la ruta: {}".format(target))
        return 1

    sources = discover_xls(target, recursive=recursive or target.is_file())
    if not sources:
        _warn("No se encontraron archivos .xls.")
        return 1

    output_dir = Path(output).expanduser().resolve() if output else None
    _info("Archivos detectados: {}".format(len(sources)))
    results = convert_many(sources, output_dir=output_dir, overwrite=overwrite)
    _print_results(results)
    ok, failed = summarize(results)
    return 0 if failed == 0 else 2 if ok else 1


def _run_interactive() -> int:
    while True:
        _clear()
        _print_banner()
        print(Fore.WHITE + "  1." + Style.RESET_ALL + " Convertir un archivo .xls")
        print(Fore.WHITE + "  2." + Style.RESET_ALL + " Convertir todos los .xls de una carpeta")
        print(Fore.WHITE + "  3." + Style.RESET_ALL + " Convertir una carpeta incluyendo subcarpetas")
        print(Fore.WHITE + "  4." + Style.RESET_ALL + " Ayuda")
        print(Fore.WHITE + "  0." + Style.RESET_ALL + " Salir")
        print()
        choice = input(Fore.GREEN + "  Seleccione una opcion > " + Style.RESET_ALL).strip()

        if choice == "0":
            print()
            _info("Hasta luego.")
            return 0
        if choice == "1":
            _menu_file()
        elif choice == "2":
            _menu_folder(recursive=False)
        elif choice == "3":
            _menu_folder(recursive=True)
        elif choice == "4":
            _menu_help()
        else:
            _warn("Opcion no valida.")
            _pause()


def _menu_file() -> None:
    print()
    raw = input("  Ruta del archivo .xls > ").strip().strip('"')
    if not raw:
        _warn("No se ingreso ninguna ruta.")
        _pause()
        return

    source = Path(raw).expanduser()
    if not source.exists():
        _error("No existe el archivo.")
        _pause()
        return
    if not is_xls_file(source):
        _error("El archivo no tiene extension .xls")
        _pause()
        return

    overwrite = _ask_yes_no("  Sobrescribir si ya existe el .xlsx", default=False)
    results = convert_many([source], overwrite=overwrite)
    _print_results(results)
    _pause()


def _menu_folder(recursive: bool) -> None:
    print()
    raw = input("  Ruta de la carpeta > ").strip().strip('"')
    if not raw:
        _warn("No se ingreso ninguna ruta.")
        _pause()
        return

    folder = Path(raw).expanduser()
    if not folder.is_dir():
        _error("La ruta no es una carpeta valida.")
        _pause()
        return

    sources = discover_xls(folder, recursive=recursive)
    if not sources:
        _warn("No se encontraron archivos .xls.")
        _pause()
        return

    _info("Se encontraron {} archivo(s) .xls.".format(len(sources)))
    for item in sources[:20]:
        print("    - {}".format(item))
    if len(sources) > 20:
        print("    ... y {} mas".format(len(sources) - 20))

    overwrite = _ask_yes_no("  Sobrescribir .xlsx existentes", default=False)
    custom = _ask_yes_no("  Guardar los .xlsx en otra carpeta", default=False)
    output_dir = None
    if custom:
        out_raw = input("  Carpeta de salida > ").strip().strip('"')
        if out_raw:
            output_dir = Path(out_raw).expanduser()

    print()
    _info("Convirtiendo...")
    results = convert_many(sources, output_dir=output_dir, overwrite=overwrite)
    _print_results(results)
    _pause()


def _menu_help() -> None:
    print()
    print(Fore.CYAN + "  Que hace esta aplicacion" + Style.RESET_ALL)
    print("  Convierte libros Excel del formato antiguo .xls (Excel 97-2003)")
    print("  al formato moderno .xlsx (Office Open XML).")
    print()
    print(Fore.CYAN + "  Como trabaja" + Style.RESET_ALL)
    print("  1. Si Microsoft Excel esta instalado, lo usa para una conversion")
    print("     mas fiel (hojas, formulas, graficos y formato).")
    print("  2. Si Excel no esta disponible, usa xlrd + openpyxl para copiar")
    print("     datos, hojas, celdas combinadas y estilos basicos.")
    print()
    print(Fore.CYAN + "  Uso directo desde CMD" + Style.RESET_ALL)
    print("    py main.py archivo.xls")
    print("    py main.py C:\\carpeta -r")
    print("    py main.py C:\\carpeta -r -y -o C:\\salida")
    print()
    _pause()


def _print_results(results: List[ConvertResult]) -> None:
    print()
    print(Fore.CYAN + "  " + "-" * 62 + Style.RESET_ALL)
    for item in results:
        status = Fore.GREEN + "OK" if item.ok else Fore.RED + "ERROR"
        print("  [{}] {}".format(status + Style.RESET_ALL, item.source.name))
        if item.ok:
            print("       Motor : {}".format(item.engine))
            print("       Salida: {}".format(item.destination))
        if item.message:
            print("       {}".format(item.message))
    ok, failed = summarize(results)
    print(Fore.CYAN + "  " + "-" * 62 + Style.RESET_ALL)
    print(
        "  Resumen: "
        + Fore.GREEN
        + "{} convertidos".format(ok)
        + Style.RESET_ALL
        + " | "
        + (Fore.RED if failed else Fore.WHITE)
        + "{} con error".format(failed)
        + Style.RESET_ALL
    )
    print()


def _print_banner() -> None:
    line = "=" * 62
    print()
    print(Fore.GREEN + "  " + line)
    print("  {:^62}".format(APP_NAME))
    print("  {:^62}".format("v{}  |  Excel 97-2003  ->  Office Open XML".format(APP_VERSION)))
    print(Fore.CYAN + "  {:^62}".format("Developed by: Gomez Luis Nicolas"))
    print(Fore.GREEN + "  " + line + Style.RESET_ALL)
    print()


def _ask_yes_no(prompt: str, default: bool) -> bool:
    suffix = " [S/n]: " if default else " [s/N]: "
    raw = input(prompt + suffix).strip().lower()
    if not raw:
        return default
    return raw in {"s", "si", "sí", "y", "yes"}


def _pause() -> None:
    print()
    input("  Pulse ENTER para continuar...")


def _info(message: str) -> None:
    print(Fore.CYAN + "  [i] " + Style.RESET_ALL + message)


def _warn(message: str) -> None:
    print(Fore.YELLOW + "  [!] " + Style.RESET_ALL + message)


def _error(message: str) -> None:
    print(Fore.RED + "  [x] " + Style.RESET_ALL + message)


def _clear() -> None:
    os.system("cls" if os.name == "nt" else "clear")


def _enable_windows_vt() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        kernel32 = ctypes.windll.kernel32
        handle = kernel32.GetStdHandle(-11)
        mode = ctypes.c_uint()
        if kernel32.GetConsoleMode(handle, ctypes.byref(mode)):
            kernel32.SetConsoleMode(handle, mode.value | 0x0004)
    except Exception:
        pass


if __name__ == "__main__":
    sys.exit(main())
