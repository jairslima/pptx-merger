"""
PPTX Merger by Jair Lima - versão CLI
Uso direto no terminal: pptx-merger-cli arquivo1.pptx arquivo2.pptx [-o saida.pptx]
"""

import argparse
import os
import sys

from merger import get_slide_count, merge_pptx, unique_output_path
from main import detect_part_number, auto_detect_pairs


def scan_folder(folder: str) -> list[str]:
    """Retorna lista de .pptx na pasta, sem arquivos temporários."""
    return [
        os.path.join(folder, f)
        for f in sorted(os.listdir(folder))
        if f.lower().endswith(".pptx") and not f.startswith("~")
    ]


def print_files(files: list[str], selected: list[str] = None):
    """Lista os arquivos encontrados com slide count e marcação de selecionado."""
    sel_set = set(selected or [])
    for f in files:
        n = get_slide_count(f)
        part = detect_part_number(f)
        mark = "[x]" if f in sel_set else "[ ]"
        parte_tag = f"  (Parte {part})" if part is not None else ""
        print(f"  {mark} {os.path.basename(f)}  [{n} slides]{parte_tag}")


def progress(done: int, total: int):
    pct = int(done / total * 100) if total else 100
    bar = "#" * (pct // 5)
    sys.stdout.write(f"\r  [{bar:<20}] {pct}%  ({done}/{total} slides copiados)")
    sys.stdout.flush()
    if done == total:
        print()


def main():
    parser = argparse.ArgumentParser(
        prog="pptx-merger-cli",
        description="Mescla múltiplos arquivos PPTX em um único. "
                    "Se nenhum arquivo for informado, varre a pasta atual e "
                    "detecta automaticamente Parte 1 / Parte 2.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""Exemplos:
  pptx-merger-cli                                      # auto-detecta na pasta atual
  pptx-merger-cli pasta/                               # auto-detecta em outra pasta
  pptx-merger-cli parte1.pptx parte2.pptx             # arquivos explícitos
  pptx-merger-cli parte1.pptx parte2.pptx -o saida.pptx
"""
    )
    parser.add_argument(
        "files", nargs="*",
        help="Arquivos PPTX a mesclar, em ordem. "
             "Se omitido (ou for uma pasta), usa auto-detecção."
    )
    parser.add_argument(
        "-o", "--output", default=None,
        help="Nome do arquivo de saída (padrão: apresentacao_completa.pptx)"
    )
    parser.add_argument(
        "--list", "-l", action="store_true",
        help="Apenas lista os arquivos encontrados, sem mesclar."
    )

    args = parser.parse_args()

    # ── Resolver lista de arquivos ──────────────────────────────
    if not args.files:
        # Sem argumentos: auto-detectar na pasta atual
        folder = os.getcwd()
        all_files = scan_folder(folder)
        ordered = auto_detect_pairs(all_files)
        to_merge = ordered if len(ordered) >= 2 else []

        print(f"\nPasta: {folder}")
        print(f"Arquivos .pptx encontrados ({len(all_files)}):")
        print_files(all_files, to_merge)

        if args.list:
            return

        if not to_merge:
            print("\nNenhum par Parte 1/Parte 2 detectado automaticamente.")
            print("Informe os arquivos explicitamente: pptx-merger-cli a.pptx b.pptx")
            sys.exit(1)

        print(f"\nOrdem de mesclagem detectada:")
        for i, f in enumerate(to_merge, 1):
            print(f"  {i}. {os.path.basename(f)}  [{get_slide_count(f)} slides]")

    elif len(args.files) == 1 and os.path.isdir(args.files[0]):
        # Único argumento é uma pasta: auto-detectar lá
        folder = args.files[0]
        all_files = scan_folder(folder)
        ordered = auto_detect_pairs(all_files)
        to_merge = ordered if len(ordered) >= 2 else []

        print(f"\nPasta: {folder}")
        print(f"Arquivos .pptx encontrados ({len(all_files)}):")
        print_files(all_files, to_merge)

        if args.list:
            return

        if not to_merge:
            print("\nNenhum par detectado. Informe os arquivos explicitamente.")
            sys.exit(1)

        folder = os.path.dirname(to_merge[0])
        print(f"\nOrdem de mesclagem detectada:")
        for i, f in enumerate(to_merge, 1):
            print(f"  {i}. {os.path.basename(f)}  [{get_slide_count(f)} slides]")

    else:
        # Arquivos explícitos
        to_merge = []
        for f in args.files:
            if not os.path.isfile(f):
                print(f"Erro: arquivo não encontrado: {f}", file=sys.stderr)
                sys.exit(1)
            if not f.lower().endswith(".pptx"):
                print(f"Aviso: {f} não é um .pptx, ignorando.")
                continue
            to_merge.append(os.path.abspath(f))

        if args.list:
            print_files(to_merge, to_merge)
            return

        if len(to_merge) < 2:
            print("Erro: informe pelo menos 2 arquivos .pptx.", file=sys.stderr)
            sys.exit(1)

        folder = os.path.dirname(to_merge[0])

    # ── Definir arquivo de saída ────────────────────────────────
    out_name = args.output or "apresentacao_completa.pptx"
    if not out_name.lower().endswith(".pptx"):
        out_name += ".pptx"
    if not os.path.isabs(out_name):
        out_name = os.path.join(folder, out_name)
    out_path = unique_output_path(out_name)

    total_in = sum(get_slide_count(f) for f in to_merge)
    print(f"\nMesclando {len(to_merge)} arquivo(s) — {total_in} slides no total...")
    print(f"Saída: {out_path}\n")

    # ── Executar mesclagem ──────────────────────────────────────
    try:
        total_out = merge_pptx(to_merge, out_path, progress_callback=progress)
        print(f"\nConcluído. {total_out} slide(s) salvos em:")
        print(f"  {out_path}")
    except Exception as e:
        print(f"\nErro durante a mesclagem: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
