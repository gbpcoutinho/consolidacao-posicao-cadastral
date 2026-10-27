r"""Gera app/recursos/logo_topo.png e app/recursos/icone.ico a partir de
logo_branco_GVS.png (raiz do projeto). Precisa do Pillow (requirements-dev.txt).

O logo e marca da empresa e o repositorio e publico: o PNG de origem e a pasta
app/recursos/ ficam no .gitignore. O app abre normalmente sem esses recursos.

Uso: .venv/Scripts/python.exe scripts/gerar_recursos.py [caminho_do_logo]
"""
import sys
from pathlib import Path

from PIL import Image

RAIZ = Path(__file__).resolve().parents[1]
DESTINO = RAIZ / "app" / "recursos"
PETROLEO = (0x0E, 0x4A, 0x55, 255)
ALTURA_TOPO = 40
TAMANHOS_ICONE = [(16, 16), (24, 24), (32, 32), (48, 48), (256, 256)]


def _recortar(logo: Image.Image) -> Image.Image:
    return logo.crop(logo.getchannel("A").getbbox())


def _simbolo(logo: Image.Image) -> Image.Image:
    """O simbolo circular e o primeiro bloco de colunas opacas, a esquerda."""
    alfa = logo.getchannel("A")
    colunas = alfa.resize((logo.width, 1), Image.Resampling.BOX)
    fim, vazio = 0, 0
    for x in range(logo.width):
        if colunas.getpixel((x, 0)) > 0:
            fim, vazio = x, 0
        else:
            vazio += 1
            if fim and vazio > logo.width // 200:
                break
    simbolo = logo.crop((0, 0, fim + 1, logo.height))
    return simbolo.crop(simbolo.getchannel("A").getbbox())


def gerar(origem: Path) -> None:
    DESTINO.mkdir(parents=True, exist_ok=True)
    logo = _recortar(Image.open(origem).convert("RGBA"))

    largura = round(logo.width * ALTURA_TOPO / logo.height)
    logo.resize((largura, ALTURA_TOPO), Image.Resampling.LANCZOS).save(DESTINO / "logo_topo.png")

    simbolo = _simbolo(logo)
    lado = 1024
    miolo = round(lado * 0.72)
    escala = miolo / max(simbolo.size)
    simbolo = simbolo.resize((round(simbolo.width * escala), round(simbolo.height * escala)),
                             Image.Resampling.LANCZOS)
    quadro = Image.new("RGBA", (lado, lado), PETROLEO)
    quadro.alpha_composite(simbolo, ((lado - simbolo.width) // 2, (lado - simbolo.height) // 2))
    quadro.resize((256, 256), Image.Resampling.LANCZOS).save(
        DESTINO / "icone.ico", sizes=TAMANHOS_ICONE)


if __name__ == "__main__":
    origem = Path(sys.argv[1]) if len(sys.argv) > 1 else RAIZ / "logo_branco_GVS.png"
    gerar(origem)
    print(f"Recursos gerados em {DESTINO}")
