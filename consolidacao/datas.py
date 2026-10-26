# Leitura controlada e datas (secao 3 do TRATAMENTO.R).
import math
import re
from datetime import date, datetime, time, timedelta

import pandas as pd

from .cabecalhos import chave_locale, normalizar_nome, trimws
from .erros import ErroValidacao

# Retorno de converter_data_com_ano_curto quando o texto nao tem ano curto
# (equivale ao NULL do R; None equivale a Date NA).
SEM_FORMATO = object()

FORMATOS = [
    "%d/%m/%Y", "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M",
    "%d-%m-%Y", "%Y-%m-%d", "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%dT%H:%M:%S", "%Y%m%d", "%d%m%Y",
]

_ANO_CURTO = re.compile(r"([0-9]{1,2})[/.-]([0-9]{1,2})[/.-]([0-9]{2})(?:[ T].*)?")
_SERIAL_TEXTO = re.compile(r"[0-9]+([.,][0-9]+)?")
_LIMITES_HORA = {"H": 24, "M": 59, "S": 61}
_CACHE_FORMATOS = {}


def _vazio_na(v) -> bool:
    return v is None or (isinstance(v, float) and math.isnan(v))


def converter_serial_excel(serial):
    if serial is None or not math.isfinite(serial):
        return None
    dia = math.floor(serial)
    if dia < 1 or dia > 2958465 or dia == 60:
        return None
    # O Excel tem um dia ficticio (serial 60 = 29/02/1900). Para dias anteriores
    # a ele, o ajuste de origem e diferente; para 60, a data e invalida.
    origem = date(1899, 12, 31) if dia < 60 else date(1899, 12, 30)
    return origem + timedelta(days=dia)


def _criar_data(ano: int, mes: int, dia: int):
    try:
        return date(ano, mes, dia)
    except ValueError:
        return None


def converter_data_com_ano_curto(valor: str):
    # Regra explicita para datas como 05/05/71 e 05/05/02:
    # 00-29 => 2000-2029; 30-99 => 1930-1999.
    m = _ANO_CURTO.fullmatch(valor)
    if not m:
        return SEM_FORMATO
    dia, mes, ano_curto = int(m.group(1)), int(m.group(2)), int(m.group(3))
    ano = 2000 + ano_curto if ano_curto <= 29 else 1900 + ano_curto
    return _criar_data(ano, mes, dia)


def _regex_formato(formato: str):
    if formato not in _CACHE_FORMATOS:
        partes, campos, i = [], [], 0
        while i < len(formato):
            c = formato[i]
            if c == "%":
                campo = formato[i + 1]
                largura = 4 if campo == "Y" else 2
                # Quantificador possessivo: como o strptime do R, nao devolve
                # digitos ja consumidos para tentar outra divisao.
                partes.append(r"\s*(\d{1,%d}+)" % largura)
                campos.append(campo)
                i += 2
            elif c == " ":
                partes.append(r"\s*")
                i += 1
            else:
                partes.append(re.escape(c))
                i += 1
        # re.ASCII: como o strptime do R, digitos e espacos Unicode (largura
        # total, NBSP) nao valem.
        _CACHE_FORMATOS[formato] = (re.compile("".join(partes), re.ASCII), campos)
    return _CACHE_FORMATOS[formato]


def parse_formato_r(valor: str, formato: str):
    """Imita as.Date(valor, format=formato): ignora o que sobra no fim."""
    regex, campos = _regex_formato(formato)
    m = regex.match(valor)
    if not m:
        return None
    valores = dict(zip(campos, (int(g) for g in m.groups())))
    for campo, limite in _LIMITES_HORA.items():
        if valores.get(campo, 0) > limite:
            return None
    if valores["Y"] == 0:
        return None
    return _criar_data(valores["Y"], valores["m"], valores["d"])


def _formatar_numero_r(x: float) -> str:
    # Imita format(x, scientific = FALSE): 7 digitos significativos, sem
    # truncar a parte inteira.
    if math.isnan(x):
        return "NaN"
    if math.isinf(x):
        return "Inf" if x > 0 else "-Inf"
    if float(x).is_integer():
        return str(int(x))
    inteiros = len(str(int(abs(x)))) if abs(x) >= 1 else 0
    if inteiros:
        casas = max(0, 7 - inteiros)
    else:
        casas = 7 - 1 - math.floor(math.log10(abs(x)))
    texto = f"{x:.{casas}f}"
    if "." in texto:
        texto = texto.rstrip("0").rstrip(".")
    return texto


def texto_celula(v):
    if _vazio_na(v):
        return None
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        return _formatar_numero_r(float(v))
    return str(v)


def _observacao(contexto, nome_coluna, linha, valor_original):
    return (
        f"{contexto} | coluna '{nome_coluna}' | registro {linha} | "
        f"valor invalido '{valor_original}' (convertido para vazio)"
    )


def converter_celula_data(v, nome_coluna: str, linha: int, contexto: str):
    if _vazio_na(v):
        return None, None
    if isinstance(v, datetime):
        return v.date(), None
    if isinstance(v, date):
        return v, None
    if isinstance(v, time):
        # readxl le horas sem data como 1899-12-31.
        return date(1899, 12, 31), None

    if isinstance(v, (int, float)) and not isinstance(v, bool):
        data_serial = converter_serial_excel(float(v))
        if data_serial is not None:
            return data_serial, None
        valor_original = texto_celula(v)
    else:
        valor_original = trimws(texto_celula(v))
        if valor_original == "":
            return None, None

        data_ano_curto = converter_data_com_ano_curto(valor_original)
        if data_ano_curto is not SEM_FORMATO:
            if data_ano_curto is not None:
                return data_ano_curto, None
            # O formato era de ano curto, mas a data nao existe no calendario.
            return None, _observacao(contexto, nome_coluna, linha, valor_original)

        for formato in FORMATOS:
            convertido = parse_formato_r(valor_original, formato)
            if convertido is not None:
                return convertido, None

        # Algumas origens armazenam a data como serial Excel, mas como texto.
        if _SERIAL_TEXTO.fullmatch(valor_original):
            data_serial = converter_serial_excel(float(valor_original.replace(",", ".", 1)))
            if data_serial is not None:
                return data_serial, None

    return None, _observacao(contexto, nome_coluna, linha, valor_original)


def validar_colunas_data(df: pd.DataFrame, contexto: str):
    observacoes, divergencias = [], []
    df = df.copy()
    nomes = list(df.columns)
    nomes_norm = [normalizar_nome(n) for n in nomes]

    for nome, norm in zip(nomes, nomes_norm):
        if "DATA" not in norm:
            continue
        datas = []
        for j, celula in enumerate(df[nome].tolist(), start=1):
            data, obs = converter_celula_data(celula, nome, j, contexto)
            datas.append(data)
            if obs is not None:
                observacoes.append(obs)
        df[nome] = pd.Series(datas, index=df.index, dtype=object)

    # Quando a data principal e a correspondente Y2K representam o mesmo dia e
    # mes, a Y2K fornece o ano completo. Se dia/mes nao coincidirem,
    # preservamos a data principal e registramos a divergencia.
    for idx_y, norm_y in enumerate(nomes_norm):
        if not norm_y.endswith("_Y2K"):
            continue
        base_norm = norm_y[: -len("_Y2K")]
        idx_base = [
            i for i, n in enumerate(nomes_norm)
            if n == base_norm and "DATA" in n and i != idx_y
        ]
        if len(idx_base) > 1:
            raise ErroValidacao(
                f"[{contexto}] Mais de uma coluna de data corresponde a '{nomes[idx_y]}'."
            )
        if not idx_base:
            continue
        nome_base, nome_y = nomes[idx_base[0]], nomes[idx_y]
        datas_base = df[nome_base].tolist()
        datas_y2k = df[nome_y].tolist()
        for j, (b, y) in enumerate(zip(datas_base, datas_y2k)):
            if b is None or y is None:
                continue
            if (b.month, b.day) == (y.month, y.day):
                datas_base[j] = y
            else:
                divergencias.append(
                    f"{contexto} | registro {j + 1} | divergencia entre '{nome_base}' "
                    f"({b.strftime('%d/%m/%Y')}) e '{nome_y}' ({y.strftime('%d/%m/%Y')}); "
                    "mantida a data principal"
                )
        df[nome_base] = pd.Series(datas_base, index=df.index, dtype=object)

    return df, observacoes, divergencias


def resumir_observacoes_datas(observacoes: list, max_amostras: int = 4):
    if not observacoes:
        return None
    # O Excel limita uma celula a 32.767 caracteres. O log mantem a contagem
    # completa por aba/coluna e exemplos.
    grupos = {}
    for obs in observacoes:
        # re.S: no R o "." tambem casa quebra de linha dentro do valor.
        chave = re.sub(r" \| registro [0-9]+ \| .*$", "", obs, count=1, flags=re.S)
        grupos.setdefault(chave, []).append(obs)
    resumos = []
    for chave in sorted(grupos, key=chave_locale):
        itens = grupos[chave]
        exemplos = [
            re.sub(r"^.* \| (registro [0-9]+ \| .*)$", r"\1", e, flags=re.S)
            for e in itens[:max_amostras]
        ]
        complemento = (
            f"; +{len(itens) - max_amostras} caso(s)" if len(itens) > max_amostras else ""
        )
        resumos.append(
            f"{chave}: {len(itens)} valor(es); amostras: " + "; ".join(exemplos) + complemento
        )
    return " || ".join(resumos)
