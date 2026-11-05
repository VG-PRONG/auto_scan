"""
AUTO SCAN (com interface) - Digitaliza automaticamente quando detecta papel.

Requisitos (Windows):
    pip install pywin32 pillow numpy

Executar:
    python auto_scan_gui.py

Para virar um programa .exe (sem precisar do Python no outro computador):
    pip install pyinstaller
    pyinstaller --onefile --noconsole auto_scan_gui.py
"""
import gc
import io
import json
import os
import platform
import re
import sys
import tempfile
import threading
import time
import traceback
import uuid
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pythoncom
import win32com.client
from PIL import Image, ImageFilter

# ------------------------- constantes WIA -------------------------
WIA_SCANNER = 1
FORMATO_PNG = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
ID_HANDLING_CAPS = 3086
ID_HANDLING_STATUS = 3087
ID_HANDLING_SELECT = 3088
ID_PAGES = 3096
FEED_READY = 0x01
FEEDER = 0x01
FLATBED = 0x02
ERRO_SEM_PAPEL = 0x80210003  # WIA_ERROR_PAPER_EMPTY

QUALIDADES = {
    "Rascunho (150 dpi)": 150,
    "Normal (200 dpi)": 200,
    "Alta (300 dpi)": 300,
    "Máxima (600 dpi)": 600,
}
# similaridade mínima (0 a 1) para considerar a folha repetida: quanto menor, mais avisos
SENSIBILIDADES = {"Baixa": 0.985, "Média": 0.96, "Alta": 0.92}
AUTOMATICO = "(Automático - primeiro encontrado)"

CONFIG_PADRAO = {
    "pasta": os.path.join(os.path.expanduser("~"), "Documents", "Scans"),
    "scanner": "",
    "qualidade": "Normal (200 dpi)",
    "colorido": False,
    "avisar_repetida": True,
    "sensibilidade": "Média",
    "perguntar_fim": True,
    "intervalo": 3,
}
ARQ_CONFIG = os.path.join(os.environ.get("APPDATA", "."), "AutoScan", "config.json")
ARQ_LOG = os.path.join(os.path.dirname(ARQ_CONFIG), "autoscan_log.txt")


def gravar_log(texto):
    """Anota tudo (inclusive os erros completos) em %APPDATA%\\AutoScan\\autoscan_log.txt."""
    try:
        os.makedirs(os.path.dirname(ARQ_LOG), exist_ok=True)
        if os.path.exists(ARQ_LOG) and os.path.getsize(ARQ_LOG) > 1_000_000:
            os.replace(ARQ_LOG, ARQ_LOG + ".antigo")
        with open(ARQ_LOG, "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now():%Y-%m-%d %H:%M:%S}] {texto}\n")
    except Exception:
        pass


# ------------------------- funções do scanner -------------------------
def _prop(obj, nome, padrao=""):
    try:
        return obj.Properties(nome).Value
    except Exception:
        return padrao


def listar_dispositivos():
    """Todos os dispositivos WIA que ESTA thread enxerga: lista de (info, nome, tipo, id)."""
    dm = win32com.client.Dispatch("WIA.DeviceManager")
    infos = dm.DeviceInfos
    try:
        itens = [infos.Item(i) for i in range(1, int(infos.Count) + 1)]
    except Exception:
        itens = list(infos)
    saida = []
    for info in itens:
        try:
            tipo = int(info.Type)
        except Exception:
            tipo = 0
        saida.append((info, str(_prop(info, "Name")), tipo, str(_prop(info, "Device ID"))))
    return saida


def _so_scanners(dispositivos):
    # tipo 1 = scanner; 0 = não informado. Câmeras (2) e vídeo (3) ficam de fora.
    return [d for d in dispositivos if d[2] in (0, WIA_SCANNER)]


def _norm(texto):
    return re.sub(r"[^a-z0-9]", "", (texto or "").lower())


def nomes_scanners():
    return [d[1] for d in _so_scanners(listar_dispositivos())]


def descrever_dispositivos():
    try:
        devs = listar_dispositivos()
    except Exception as e:
        return f"(não consegui listar os dispositivos do Windows: {e})"
    if not devs:
        return ("O Windows não mostra nenhum scanner agora. Ligue a impressora, confira o cabo/Wi-Fi "
                "e clique em Atualizar.")
    return "O Windows mostra: " + "; ".join(f"{n} (tipo {t})" for _, n, t, _ in devs)


def conectar(nome_parte=""):
    """Conecta ao scanner escolhido. Aceita pequenas diferenças no nome; se só existir
    um scanner, usa ele mesmo."""
    candidatos = _so_scanners(listar_dispositivos())
    if not candidatos:
        return None
    escolhido = None
    alvo = _norm(nome_parte)
    if not alvo:
        escolhido = candidatos[0]
    else:
        for d in candidatos:
            if _norm(d[1]) == alvo:
                escolhido = d
                break
        if escolhido is None:
            for d in candidatos:
                n = _norm(d[1])
                if n and (alvo in n or n in alvo):
                    escolhido = d
                    break
        if escolhido is None and len(candidatos) == 1:
            escolhido = candidatos[0]
    return escolhido[0].Connect() if escolhido else None


def definir_prop(props, pid, valor):
    for p in props:
        if p.PropertyID == pid:
            try:
                p.Value = valor
            except Exception:
                pass
            return


def ler_prop(props, pid):
    for p in props:
        if p.PropertyID == pid:
            try:
                return p.Value
            except Exception:
                return None
    return None


def tem_papel(device):
    status = ler_prop(device.Properties, ID_HANDLING_STATUS)
    return bool(status is not None and status & FEED_READY)


def tem_alimentador(device):
    """True se o scanner tem alimentador automático (ADF). Scanners só de vidro retornam False."""
    caps = ler_prop(device.Properties, ID_HANDLING_CAPS)
    return bool(caps is not None and caps & FEEDER)


PREVIA_DPI = 75
INTERVALO_VIDRO = 10.0    # segundos entre uma verificação do vidro e a próxima
CONTAGEM_SEG = 10         # contagem regressiva antes de digitalizar (e entre tentativas)
TENTATIVAS_SCAN = 3       # quantas vezes tenta a digitalização completa antes de desistir do ciclo
ESPERA_ERRO = 5.0         # espera depois de um erro (ex.: scanner ocupado) antes de tentar de novo
ESPERA_OCUPADO = 8.0
TINTA_MIN = 0.0010        # fração mínima da página com marcas para considerar que há uma folha
ESCURECE_MIN = 12         # quanto uma marca precisa ficar mais escura que o papel (0-255)
BRILHO_MIN = 120          # fundo mais escuro que isso = tampa aberta / sem tampa
ESTAVEL_MIN = 0.85        # semelhança entre duas prévias seguidas para "a folha parou"
MAX_PREVIAS_INSTAVEIS = 4  # depois disso digitaliza mesmo assim (nunca fica preso na prévia)
REARMA_ABAIXO = 0.50      # só considera "folha diferente" se a imagem mudou MUITO (folha fraca varia por ruído)


def analisar_previa(caminho_png):
    """Olha uma prévia rápida do vidro. Retorna (tem_folha, mapa_de_tinta, info).
    Folha = fundo claro (papel) com marcas (lápis/caneta/impressão) em pelo menos 0,1% da área.
    Vidro vazio com tampa fechada = fundo uniforme; tampa aberta = fundo escuro.
    As bordas são ignoradas (sombra da tampa)."""
    with Image.open(caminho_png) as im:
        g = im.convert("L")
        g.thumbnail((640, 640))
        a = np.asarray(g, dtype=np.float32)
        pequena = g.copy()
    h, w = a.shape
    mh, mw = max(1, int(h * 0.04)), max(1, int(w * 0.04))
    miolo = a[mh:h - mh, mw:w - mw] if h > 2 * mh + 8 and w > 2 * mw + 8 else a
    fundo = float(np.median(miolo))
    tinta = float((miolo < fundo - ESCURECE_MIN).mean())
    pequena.thumbnail((160, 160))
    mapa = 255.0 - np.asarray(pequena.filter(ImageFilter.GaussianBlur(1.0)), dtype=np.float32)
    tem = bool(fundo > BRILHO_MIN and TINTA_MIN < tinta < 0.5)
    return tem, mapa, {"fundo": fundo, "tinta": tinta}


def dpi_suportado(prop, desejado):
    """Resolução que o scanner realmente aceita, a mais próxima da desejada.
    Cada modelo aceita valores diferentes (ex.: só 100/200/300/600); se não der para
    descobrir, usa o valor pedido."""
    try:
        tipo = int(prop.SubType)
        if tipo == 2:  # lista de valores
            valores = [int(v) for v in prop.SubTypeValues]
            return min(valores, key=lambda v: abs(v - desejado)) if valores else desejado
        if tipo == 1:  # faixa mín/máx/passo
            lo, hi = int(prop.SubTypeMin), int(prop.SubTypeMax)
            passo = max(1, int(prop.SubTypeStep or 1))
            cand = lo + round((desejado - lo) / passo) * passo
            return max(lo, min(hi, cand))
    except Exception:
        pass
    return desejado


def preparar_item(device, dpi, colorido):
    item = device.Items[1]
    definir_prop(item.Properties, 6146, 1 if colorido else 2)  # cor / cinza
    for pid in (6147, 6148):
        for p in item.Properties:
            if p.PropertyID == pid:
                definir_prop(item.Properties, pid, dpi_suportado(p, dpi))
                break
    return item


def _salvar_tmp(img, n):
    caminho = os.path.join(tempfile.gettempdir(), f"autoscan_{n}_{time.time_ns()}.png")
    img.SaveFile(caminho)
    return caminho


def _codigo_erro(e):
    """Código HRESULT do WIA (sem sinal). Via IDispatch o Windows entrega o erro embrulhado
    (0x80020009) e o código verdadeiro fica em excepinfo[5]."""
    exc = getattr(e, "excepinfo", None)
    if exc and len(exc) >= 6 and isinstance(exc[5], int) and exc[5] != 0:
        return exc[5] & 0xFFFFFFFF
    codigo = getattr(e, "hresult", None)
    if codigo is None and getattr(e, "args", None):
        codigo = e.args[0]
    return (codigo & 0xFFFFFFFF) if isinstance(codigo, int) else None


ERROS_WIA = {
    0x80210001: "Erro geral no scanner.",
    0x80210002: "Papel preso no alimentador.",
    0x80210003: "Sem papel no alimentador.",
    0x80210004: "Problema com o papel no alimentador.",
    0x80210005: "Scanner desligado ou fora da rede/USB (offline). Ligue a impressora e confira o cabo ou o Wi-Fi.",
    0x80210006: "Scanner ocupado. Espere terminar ou feche o HP Smart/outro programa que esteja usando o scanner.",
    0x80210007: "Scanner aquecendo. Tente de novo em alguns segundos.",
    0x80210008: "O scanner pede uma ação sua (veja o visor da impressora).",
    0x8021000A: "Falha de comunicação com o scanner. Desligue e ligue a impressora e tente de novo.",
    0x8021000D: "Scanner bloqueado por outro programa.",
    0x80210015: "Nenhum scanner disponível. Ligue a impressora e confira a conexão.",
    0x80210016: "Tampa do scanner aberta.",
    0x80210017: "Lâmpada do scanner apagada/aquecendo.",
}


def explicar_erro(e):
    """Mensagem clara em português + o texto técnico original."""
    cod = _codigo_erro(e)
    texto = ERROS_WIA.get(cod)
    tecnico = str(e)
    exc = getattr(e, "excepinfo", None)
    if exc and len(exc) >= 3 and exc[2]:
        tecnico = str(exc[2])
    sufixo = f" (código 0x{cod:08X})" if cod else ""
    if texto:
        return f"{texto} [{tecnico}]{sufixo}"
    return f"{tecnico}{sufixo}"


def gerar_diagnostico():
    """Relatório do que o Windows/WIA enxerga (para descobrir por que um scanner não responde)."""
    L = [f"Python {sys.version.split()[0]} | {platform.platform()}",
         f"Pillow {getattr(Image, '__version__', '?')} | numpy {np.__version__}"]
    try:
        devs = listar_dispositivos()
    except Exception as e:
        L.append(f"FALHA ao listar dispositivos WIA: {explicar_erro(e)}")
        return "\n".join(L)
    L.append(f"Dispositivos WIA vistos: {len(devs)}")
    for info, nome, tipo, id_ in devs:
        L.append(f"- nome={nome!r} tipo={tipo} id={id_}")
        try:
            dev = info.Connect()
        except Exception as e:
            L.append(f"    Conectar: ERRO -> {explicar_erro(e)}")
            continue
        L.append("    Conectar: OK")
        try:
            L.append(f"    Tem alimentador (ADF): {tem_alimentador(dev)}")
            for p in dev.Properties:
                try:
                    L.append(f"    [{p.PropertyID}] {p.Name} = {p.Value}")
                except Exception:
                    pass
            try:
                n_itens = int(dev.Items.Count)
            except Exception:
                n_itens = len(dev.Items)
            L.append(f"    Itens de digitalização: {n_itens}")
        except Exception as e:
            L.append(f"    Leitura das propriedades falhou: {explicar_erro(e)}")
    return "\n".join(L)


def digitalizar_alimentador(device, dpi, colorido, log):
    definir_prop(device.Properties, ID_HANDLING_SELECT, FEEDER)
    definir_prop(device.Properties, ID_PAGES, 1)
    item = preparar_item(device, dpi, colorido)
    arquivos = []
    while True:
        try:
            img = item.Transfer(FORMATO_PNG)
        except pythoncom.com_error as e:
            if _codigo_erro(e) == ERRO_SEM_PAPEL:  # só "alimentador vazio"; papel preso é erro de verdade
                break
            if arquivos:  # travou no meio do lote: não perde as folhas que já saíram
                log(f"Erro na folha {len(arquivos) + 1} (papel preso?): {e}. "
                    f"Mantendo as {len(arquivos)} folha(s) já digitalizadas.")
                break
            raise
        arquivos.append(_salvar_tmp(img, len(arquivos)))
        log(f"Folha {len(arquivos)} digitalizada.")
        if not tem_papel(device):
            break
    return arquivos


def digitalizar_vidro(device, dpi, colorido, log):
    definir_prop(device.Properties, ID_HANDLING_SELECT, FLATBED)
    item = preparar_item(device, dpi, colorido)
    img = item.Transfer(FORMATO_PNG)
    log("Folha digitalizada pelo vidro.")
    return [_salvar_tmp(img, 0)]


def teste_passo_a_passo(cfg, log, esperar):
    """Roda cada etapa do modo automático separadamente no scanner REAL e diz qual falha e por quê.
    log(msg) mostra o resultado; esperar(seg, rotulo) faz a contagem regressiva entre digitalizações."""
    ctx, resumo = {}, []

    def etapa(nome, fn):
        t0 = time.time()
        try:
            r = fn()
            log(f"OK     {nome} ({time.time() - t0:.1f}s)" + (f" - {r}" if r else ""))
            resumo.append((nome, True))
            return True
        except Exception as e:
            log(f"FALHOU {nome} ({time.time() - t0:.1f}s): {explicar_erro(e)}")
            gravar_log("Traceback do teste:\n" + "".join(traceback.format_exception(type(e), e, e.__traceback__)))
            resumo.append((nome, False))
            return False

    def e1():
        ctx["nomes"] = nomes_scanners()
        if not ctx["nomes"]:
            raise RuntimeError(descrever_dispositivos())
        return "; ".join(ctx["nomes"])

    def e2():
        ctx["dev"] = conectar(cfg["scanner"])
        if ctx["dev"] is None:
            raise RuntimeError("não conectou. " + descrever_dispositivos())

    def e3():
        dev = ctx["dev"]
        info = [f"alimentador={'sim' if tem_alimentador(dev) else 'não (só vidro)'}"]
        for p in dev.Items[1].Properties:
            if p.PropertyID == 6147:
                info.append(f"resolução pedida {PREVIA_DPI}->{dpi_suportado(p, PREVIA_DPI)}, "
                            f"{cfg['dpi']}->{dpi_suportado(p, cfg['dpi'])}")
        return "; ".join(info)

    def e4():
        arqs = digitalizar_vidro(ctx["dev"], PREVIA_DPI, False, lambda m: None)
        ctx["previa"] = arqs[0]
        return f"{os.path.getsize(arqs[0]) // 1024} KB"

    def e5():
        tem, _, info = analisar_previa(ctx["previa"])
        return (f"brilho {info['fundo']:.0f}, marcas {info['tinta'] * 100:.2f}% -> "
                + ("TEM folha" if tem else "vidro vazio / sem folha detectada"))

    def e6():
        ctx["dev"] = None
        gc.collect()
        ctx["dev"] = conectar(cfg["scanner"])
        arqs = digitalizar_vidro(ctx["dev"], cfg["dpi"], cfg["colorido"], lambda m: None)
        ctx["scan"] = arqs
        return f"{os.path.getsize(arqs[0]) // 1024} KB"

    def e7():  # só informativo: o programa reconecta sozinho, então isso NÃO precisa funcionar
        try:
            arqs = digitalizar_vidro(ctx["dev"], PREVIA_DPI, False, lambda m: None)
            apagar(arqs)
            return "a mesma conexão aceitou uma 2ª digitalização"
        except Exception as e:
            return f"a mesma conexão NÃO aceita 2ª digitalização ({explicar_erro(e)}) - normal: o programa reconecta antes de cada uma"

    def e8():
        pasta_tmp = tempfile.mkdtemp(prefix="autoscan_teste_")
        destino = salvar_pdf(ctx["scan"], pasta_tmp, cfg["dpi"])
        tam = os.path.getsize(destino)
        os.remove(destino)
        os.rmdir(pasta_tmp)
        return f"PDF de {tam // 1024} KB gerado e apagado"

    ok = etapa("1. Listar scanners do Windows", e1)
    ok = ok and etapa("2. Conectar ao scanner", e2)
    ok = ok and etapa("3. Ler propriedades", e3)
    ok = ok and etapa(f"4. Prévia rápida ({PREVIA_DPI} dpi)", e4)
    if ok:
        etapa("5. Analisar a prévia (detecção da folha)", e5)
        if not esperar(CONTAGEM_SEG, "Aguardando entre digitalizações"):
            return resumo
        ok = etapa(f"6. RECONECTAR e digitalizar ({cfg['dpi']} dpi)", e6)
        if ok:
            if not esperar(CONTAGEM_SEG, "Aguardando entre digitalizações"):
                return resumo
            etapa("7. Digitalizar de novo na MESMA conexão (informativo)", e7)
            etapa("8. Montar o PDF", e8)
    apagar([ctx.get("previa")] if ctx.get("previa") else [])
    apagar(ctx.get("scan") or [])
    falhas = [n for n, o in resumo if not o]
    log("RESULTADO: tudo funcionou." if not falhas else "RESULTADO: falharam -> " + " | ".join(falhas))
    return resumo


# ------------------------- detecção de folha repetida -------------------------
# Cada folha vira uma miniatura de "tinta" (160x224). Duas digitalizações da mesma folha
# têm correlação altíssima mesmo com pequenos deslocamentos/inclinação do alimentador.
MINI = (160, 224)
GROSSO = (8, 11)
PRE_FILTRO = 0.70   # só confere em detalhe as folhas do histórico com perfil parecido
DESLOC = 2          # tolerância de deslocamento (em pixels da miniatura)
PASTA_DADOS = ".autoscan_dados"


def assinatura(caminho_png):
    """Retorna (miniatura, vetor_grosso) ou None se a folha estiver em branco."""
    with Image.open(caminho_png) as im:
        g = im.convert("L").resize(MINI, Image.BOX).filter(ImageFilter.GaussianBlur(1.5))
    mini = 255.0 - np.asarray(g, dtype=np.float32)
    mini = mini - np.median(mini)  # remove cor/sombra do fundo do papel
    if float(mini.std()) < 2.0:
        return None
    gr = np.asarray(
        Image.fromarray(np.clip(mini + 128, 0, 255).astype(np.uint8)).resize(GROSSO, Image.BOX),
        dtype=np.float32).flatten()
    gr = gr - gr.mean()
    n = float(np.linalg.norm(gr))
    return mini, (gr / n if n > 1e-6 else gr)


def correlacao(a, b, desloc=DESLOC):
    H, W = a.shape
    melhor = -1.0
    for dy in range(-desloc, desloc + 1):
        for dx in range(-desloc, desloc + 1):
            A = a[max(dy, 0):H + min(dy, 0), max(dx, 0):W + min(dx, 0)]
            B = b[max(-dy, 0):H + min(-dy, 0), max(-dx, 0):W + min(-dx, 0)]
            A = A - A.mean()
            B = B - B.mean()
            den = float(np.sqrt((A * A).sum() * (B * B).sum()))
            if den > 1e-6:
                melhor = max(melhor, float((A * B).sum()) / den)
    return melhor


def _dir_dados(pasta):
    return os.path.join(pasta, PASTA_DADOS)


def carregar_indice(pasta):
    try:
        with open(os.path.join(_dir_dados(pasta), "indice.json"), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def gravar_indice(pasta, indice):
    os.makedirs(_dir_dados(pasta), exist_ok=True)
    with open(os.path.join(_dir_dados(pasta), "indice.json"), "w", encoding="utf-8") as f:
        json.dump(indice, f, ensure_ascii=False)


def guardar_assinatura(pasta, sig, arquivo_pdf, indice):
    """Grava a miniatura e acrescenta a folha ao índice (em memória; use gravar_indice depois)."""
    os.makedirs(_dir_dados(pasta), exist_ok=True)
    id_ = uuid.uuid4().hex
    mini, grosso = sig
    Image.fromarray(np.clip(mini + 128, 0, 255).astype(np.uint8)).save(
        os.path.join(_dir_dados(pasta), id_ + ".png"))
    indice.append({"id": id_, "arquivo": arquivo_pdf, "data": f"{datetime.now():%d/%m/%Y %H:%M}",
                   "grosso": [round(float(x), 4) for x in grosso]})


def buscar_parecida(sig, indice, pasta, limiar):
    """Procura no histórico uma folha com similaridade >= limiar. Retorna o item ou None."""
    if not indice:
        return None
    mini, grosso = sig
    matriz = np.asarray([i["grosso"] for i in indice], dtype=np.float32)
    pontos = matriz @ grosso
    candidatos = [k for k in np.argsort(-pontos)[:30] if pontos[k] >= PRE_FILTRO]
    melhor, melhor_item = limiar, None
    for k in candidatos:
        try:
            with Image.open(os.path.join(_dir_dados(pasta), indice[k]["id"] + ".png")) as im:
                outra = np.asarray(im, dtype=np.float32) - 128.0
        except Exception:
            continue
        c = correlacao(mini, outra)
        if c >= melhor:
            melhor, melhor_item = c, indice[k]
    return melhor_item


def _pagina_jpeg(caminho, dpi, qualidade=90):
    """Lê UMA folha, devolve (bytes_jpeg, largura, altura, espaço_de_cor). Usa pouca memória."""
    with Image.open(caminho) as im:
        if im.mode in ("RGB", "L"):
            pronto = im
        elif im.mode in ("1", "LA"):
            pronto = im.convert("L")
        else:  # P, RGBA, CMYK...
            pronto = im.convert("RGB")
        buf = io.BytesIO()
        pronto.save(buf, "JPEG", quality=qualidade, dpi=(dpi, dpi))
        cor = "/DeviceGray" if pronto.mode == "L" else "/DeviceRGB"
        return buf.getvalue(), pronto.size[0], pronto.size[1], cor


def salvar_pdf(arquivos, pasta, dpi):
    """Monta o PDF folha por folha (sem carregar o lote inteiro na memória)."""
    os.makedirs(pasta, exist_ok=True)
    base = f"scan_{datetime.now():%Y-%m-%d_%H-%M-%S}"
    n = 1
    while True:  # cria sempre um arquivo novo; nunca sobrescreve outro
        destino = os.path.join(pasta, base + (".pdf" if n == 1 else f"_{n}.pdf"))
        try:
            f = open(destino, "xb")
            break
        except FileExistsError:
            n += 1

    total = len(arquivos)
    offsets = {}
    try:
        with f:
            def objeto(num, corpo):
                offsets[num] = f.tell()
                f.write(f"{num} 0 obj\n".encode() + corpo + b"\nendobj\n")

            f.write(b"%PDF-1.4\n%\xe2\xe3\xcf\xd3\n")
            for i, arq in enumerate(arquivos):
                jpg, w, h, cor = _pagina_jpeg(arq, dpi)
                pw, ph = w * 72.0 / dpi, h * 72.0 / dpi
                n_pag, n_cont, n_img = 3 + 3 * i, 4 + 3 * i, 5 + 3 * i
                objeto(n_img, (f"<< /Type /XObject /Subtype /Image /Width {w} /Height {h} "
                               f"/ColorSpace {cor} /BitsPerComponent 8 /Filter /DCTDecode "
                               f"/Length {len(jpg)} >>\nstream\n").encode() + jpg + b"\nendstream")
                conteudo = f"q {pw:.2f} 0 0 {ph:.2f} 0 0 cm /Im0 Do Q".encode()
                objeto(n_cont, f"<< /Length {len(conteudo)} >>\nstream\n".encode() + conteudo + b"\nendstream")
                objeto(n_pag, (f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {pw:.2f} {ph:.2f}] "
                               f"/Resources << /XObject << /Im0 {n_img} 0 R >> >> "
                               f"/Contents {n_cont} 0 R >>").encode())
                del jpg
            kids = " ".join(f"{3 + 3 * i} 0 R" for i in range(total))
            objeto(2, f"<< /Type /Pages /Kids [{kids}] /Count {total} >>".encode())
            objeto(1, b"<< /Type /Catalog /Pages 2 0 R >>")
            tam = 3 + 3 * total
            inicio_xref = f.tell()
            f.write(f"xref\n0 {tam}\n".encode() + b"0000000000 65535 f \n")
            for num in range(1, tam):
                f.write(f"{offsets[num]:010d} 00000 n \n".encode())
            f.write(f"trailer\n<< /Size {tam} /Root 1 0 R >>\nstartxref\n{inicio_xref}\n%%EOF\n".encode())
    except Exception:
        try:
            os.remove(destino)  # não deixa PDF incompleto para trás
        except OSError:
            pass
        raise
    return destino


def apagar(arquivos):
    for a in arquivos:
        try:
            os.remove(a)
        except OSError:
            pass


# ------------------------- interface -------------------------
class App:
    def __init__(self, root):
        self.root = root
        root.title("Auto Scan - Digitalização automática")
        root.resizable(False, False)

        self.parar = threading.Event()
        self.trava = threading.Lock()
        self.monitorando = False
        self.fechando = False
        self.msg_espera = None  # mensagem que fica na tela até a próxima folha chegar
        self.onde = "no alimentador"   # ou "no vidro" (scanner sem alimentador)
        self.vidro_armado = True       # pronto para detectar uma NOVA folha no vidro
        self.vidro_prev = None         # prévia anterior (para confirmar que a folha parou)
        self.avisou_vidro = False
        self._ultimo_aviso = None      # evita repetir a mesma mensagem de erro a cada ciclo
        self.forcar = threading.Event()  # pedido de digitalização manual durante o monitoramento
        self.cont_folha = 0            # prévias seguidas em que apareceu folha
        self.vidro_ultima = None       # prévia da última folha digitalizada
        self.n_previa = 0
        self._ov = None                # aviso grande da contagem regressiva (canto da tela)
        self._ov_lbl = None
        self._status_base = ""
        self._erros_iguais = 0

        cfg = dict(CONFIG_PADRAO)
        try:
            with open(ARQ_CONFIG, "r", encoding="utf-8") as f:
                cfg.update(json.load(f))
        except Exception:
            pass

        self.v_pasta = tk.StringVar(value=cfg["pasta"])
        self.v_scanner = tk.StringVar(value=cfg["scanner"] or AUTOMATICO)
        self.v_qualidade = tk.StringVar(value=cfg["qualidade"])
        self.v_colorido = tk.BooleanVar(value=cfg["colorido"])
        self.v_repetida = tk.BooleanVar(value=cfg["avisar_repetida"])
        self.v_sens = tk.StringVar(value=cfg["sensibilidade"])
        self.v_fim = tk.BooleanVar(value=cfg["perguntar_fim"])
        self.v_status = tk.StringVar(value="Parado.")

        self._montar_tela()
        self.atualizar_scanners()
        root.protocol("WM_DELETE_WINDOW", self.fechar)

    # ---------- layout ----------
    def _montar_tela(self):
        pad = {"padx": 8, "pady": 4}
        f = ttk.Frame(self.root, padding=12)
        f.grid()

        cx = ttk.LabelFrame(f, text="Onde salvar", padding=8)
        cx.grid(row=0, column=0, sticky="ew", **pad)
        ttk.Entry(cx, textvariable=self.v_pasta, width=48).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(cx, text="Escolher pasta...", command=self.escolher_pasta).grid(row=0, column=1)

        cx = ttk.LabelFrame(f, text="Scanner / Impressora", padding=8)
        cx.grid(row=1, column=0, sticky="ew", **pad)
        self.cb_scanner = ttk.Combobox(cx, textvariable=self.v_scanner, width=45, state="readonly")
        self.cb_scanner.grid(row=0, column=0, padx=(0, 6))
        ttk.Button(cx, text="Atualizar", command=self.atualizar_scanners).grid(row=0, column=1)

        cx = ttk.LabelFrame(f, text="Opções", padding=8)
        cx.grid(row=2, column=0, sticky="ew", **pad)
        ttk.Label(cx, text="Qualidade:").grid(row=0, column=0, sticky="w")
        ttk.Combobox(cx, textvariable=self.v_qualidade, values=list(QUALIDADES), width=20,
                     state="readonly").grid(row=0, column=1, sticky="w", padx=6)
        ttk.Checkbutton(cx, text="Digitalizar em cores (arquivo maior)",
                        variable=self.v_colorido).grid(row=1, column=0, columnspan=3, sticky="w", pady=2)
        ttk.Checkbutton(cx, text="Avisar quando a folha for repetida",
                        variable=self.v_repetida).grid(row=2, column=0, columnspan=2, sticky="w", pady=2)
        ttk.Label(cx, text="Sensibilidade:").grid(row=2, column=2, sticky="e", padx=(12, 4))
        ttk.Combobox(cx, textvariable=self.v_sens, values=list(SENSIBILIDADES), width=8,
                     state="readonly").grid(row=2, column=3)
        ttk.Checkbutton(cx, text="Ao terminar, perguntar se quero trocar de folha ou repetir",
                        variable=self.v_fim).grid(row=3, column=0, columnspan=4, sticky="w", pady=2)

        ttk.Label(cx, foreground="#555555", wraplength=440, justify="left",
                  text="Scanner só de vidro (sem alimentador)? Sem problema: o programa percebe a folha "
                       "sozinho. Coloque a folha e feche a tampa.").grid(
            row=4, column=0, columnspan=4, sticky="w", pady=(6, 0))

        bt = ttk.Frame(f)
        bt.grid(row=3, column=0, sticky="ew", **pad)
        self.bt_iniciar = ttk.Button(bt, text="▶ Iniciar monitoramento", command=self.alternar)
        self.bt_iniciar.grid(row=0, column=0, padx=(0, 6))
        ttk.Button(bt, text="Digitalizar agora (vidro)", command=self.digitalizar_agora).grid(row=0, column=1, padx=6)
        ttk.Button(bt, text="Abrir pasta", command=self.abrir_pasta).grid(row=0, column=2, padx=6)
        ttk.Button(bt, text="Diagnóstico", command=self.diagnostico).grid(row=0, column=3, padx=6)
        bt2 = ttk.Frame(f)
        bt2.grid(row=3, column=0, sticky="e", padx=8, pady=(34, 0))
        ttk.Button(bt2, text="Testar passo a passo", command=self.testar_passos).grid(row=0, column=0, padx=6)
        ttk.Button(bt2, text="Abrir log", command=self.abrir_log).grid(row=0, column=1, padx=6)

        ttk.Label(f, textvariable=self.v_status, font=("Segoe UI", 10, "bold")).grid(
            row=4, column=0, sticky="w", **pad)

        lf = ttk.Frame(f)
        lf.grid(row=5, column=0, sticky="ew", **pad)
        self.txt = tk.Text(lf, height=9, width=70, state="disabled", wrap="word")
        sb = ttk.Scrollbar(lf, command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.grid(row=0, column=0)
        sb.grid(row=0, column=1, sticky="ns")

    # ---------- utilidades de interface (seguras entre threads) ----------
    def log(self, msg):
        linha = f"[{datetime.now():%H:%M:%S}] {msg}\n"
        gravar_log(msg)

        def _add():
            self.txt.configure(state="normal")
            self.txt.insert("end", linha)
            self.txt.see("end")
            self.txt.configure(state="disabled")

        self._ui(_add)

    def status(self, msg):
        self._status_base = msg
        self._ui(lambda: self.v_status.set(msg))

    def _overlay(self, texto):
        """Aviso grande no canto da tela (fica por cima de tudo). texto=None esconde."""
        def _f():
            try:
                if texto is None:
                    if self._ov is not None:
                        self._ov.withdraw()
                    return
                if self._ov is None:
                    ov = tk.Toplevel(self.root)
                    ov.overrideredirect(True)
                    ov.attributes("-topmost", True)
                    ov.configure(bg="#202124")
                    self._ov_lbl = tk.Label(ov, font=("Segoe UI", 20, "bold"), fg="white", bg="#202124",
                                            padx=26, pady=14, justify="center")
                    self._ov_lbl.pack()
                    self._ov = ov
                self._ov_lbl.configure(text=texto)
                self._ov.update_idletasks()
                self._ov.geometry(f"+{self._ov.winfo_screenwidth() - self._ov.winfo_reqwidth() - 30}+30")
                self._ov.deiconify()
                self._ov.lift()
            except tk.TclError:
                pass
        self._ui(_f)

    def contagem(self, segundos, rotulo, overlay=True):
        """Contagem regressiva visível na janela e no canto da tela. False se o usuário parou."""
        for i in range(int(segundos), 0, -1):
            if self.parar.is_set():
                break
            msg = f"{rotulo} em {i}s..."
            self._ui(lambda m=msg: self.v_status.set(m))
            if overlay:
                self._overlay(f"{rotulo}\n{i}")
            self.parar.wait(1.0)
        self._overlay(None)
        return not self.parar.is_set()

    def registrar_erro(self, contexto, e, janela=True):
        """Mostra o erro com o código e grava o traceback completo no arquivo de log."""
        msg = explicar_erro(e)
        gravar_log(f"ERRO ({contexto}): {msg}\n" +
                   "".join(traceback.format_exception(type(e), e, e.__traceback__)))
        if janela:
            self.log(f"ERRO - {contexto}: {msg}")

    def _ui(self, fn):
        """Agenda fn na thread da interface; ignora se a janela já foi fechada."""
        if self.fechando:
            return False
        try:
            self.root.after(0, fn)
            return True
        except (tk.TclError, RuntimeError):
            return False

    def na_tela(self, fn):
        """Executa fn na thread da interface e espera o resultado (para janelas de aviso)."""
        res, ev = {}, threading.Event()

        def run():
            try:
                res["v"] = fn()
            finally:
                ev.set()

        if not self._ui(run):
            return None
        while not ev.wait(0.2):  # se a janela for fechada, não fica esperando para sempre
            if self.fechando:
                return None
        return res.get("v")

    def perguntar(self, titulo, texto, botoes):
        """Janela com botões personalizados. O primeiro botão é o padrão (também ao fechar no X)."""
        resp = {"v": botoes[0][0]}
        win = tk.Toplevel(self.root)
        win.title(titulo)
        win.resizable(False, False)
        win.attributes("-topmost", True)
        ttk.Label(win, text=texto, justify="left", wraplength=400, padding=16).pack()
        fr = ttk.Frame(win, padding=(16, 0, 16, 16))
        fr.pack()

        def escolher(k):
            resp["v"] = k
            win.destroy()

        for k, rotulo in botoes:
            ttk.Button(fr, text=rotulo, command=lambda k=k: escolher(k)).pack(side="left", padx=4)
        win.protocol("WM_DELETE_WINDOW", lambda: escolher(botoes[0][0]))
        win.update_idletasks()
        x = (win.winfo_screenwidth() - win.winfo_width()) // 2
        y = (win.winfo_screenheight() - win.winfo_height()) // 3
        win.geometry(f"+{x}+{y}")
        win.lift()
        win.focus_force()
        win.bell()
        win.grab_set()
        self.root.wait_window(win)
        return resp["v"]

    # ---------- ações dos botões ----------
    def escolher_pasta(self):
        p = filedialog.askdirectory(initialdir=self.v_pasta.get() or None, title="Escolha a pasta")
        if p:
            self.v_pasta.set(os.path.normpath(p))

    def atualizar_scanners(self):
        try:
            nomes = nomes_scanners()
        except Exception as e:
            nomes = []
            self.log(f"Não consegui listar scanners: {e}")
        self.cb_scanner["values"] = [AUTOMATICO] + nomes
        if self.v_scanner.get() not in self.cb_scanner["values"]:
            self.v_scanner.set(AUTOMATICO)
        if not nomes:
            self.log("Nenhum scanner encontrado. Ligue a impressora e clique em Atualizar.")

    def abrir_pasta(self):
        pasta = self.v_pasta.get().strip()
        if pasta:
            os.makedirs(pasta, exist_ok=True)
            os.startfile(pasta)

    def abrir_log(self):
        try:
            if not os.path.exists(ARQ_LOG):
                gravar_log("(log criado)")
            os.startfile(ARQ_LOG)
        except Exception as e:
            self.log(f"Não consegui abrir o log: {e}. Ele fica em: {ARQ_LOG}")

    def testar_passos(self):
        """Executa cada etapa do modo automático, uma por uma, no scanner real."""
        if self.monitorando:
            messagebox.showinfo("Auto Scan", "Pare o monitoramento antes de rodar o teste.")
            return
        cfg = self.coletar_cfg()
        if not cfg:
            return
        if not self.trava.acquire(blocking=False):
            messagebox.showinfo("Auto Scan", "Aguarde: já existe uma operação em andamento.")
            return
        messagebox.showinfo("Auto Scan", "O teste vai digitalizar 2 vezes (com 10 segundos de intervalo).\n"
                            "Coloque uma folha no vidro e feche a tampa. Clique OK para começar.")
        self.log("==== TESTE PASSO A PASSO ====")
        threading.Thread(target=self._testar_passos, args=(cfg,), daemon=True).start()

    def _testar_passos(self, cfg):
        pythoncom.CoInitialize()
        linhas = []

        def log(m):
            linhas.append(m)
            self.log(m)

        try:
            teste_passo_a_passo(cfg, log, lambda seg, rot: self.contagem(seg, rot))
        except Exception as e:
            self.registrar_erro("Teste passo a passo", e)
        finally:
            gc.collect()
            pythoncom.CoUninitialize()
            self.trava.release()
            self.status("Teste terminado - veja o histórico.")
            texto = "\n".join(linhas)
            self._ui(lambda: (self.root.clipboard_clear(), self.root.clipboard_append(texto)))
            self.log("(resultado copiado para a área de transferência)")

    def coletar_cfg(self):
        pasta = self.v_pasta.get().strip()
        if not pasta:
            messagebox.showwarning("Auto Scan", "Escolha a pasta onde os arquivos serão salvos.")
            return None
        try:
            os.makedirs(pasta, exist_ok=True)
        except OSError as e:
            messagebox.showerror("Auto Scan", f"Não foi possível usar essa pasta:\n{e}")
            return None
        scanner = self.v_scanner.get()
        cfg = {
            "pasta": pasta,
            "scanner": "" if scanner == AUTOMATICO else scanner,
            "qualidade": self.v_qualidade.get(),
            "colorido": self.v_colorido.get(),
            "avisar_repetida": self.v_repetida.get(),
            "sensibilidade": self.v_sens.get(),
            "perguntar_fim": self.v_fim.get(),
            "intervalo": CONFIG_PADRAO["intervalo"],
        }
        try:
            os.makedirs(os.path.dirname(ARQ_CONFIG), exist_ok=True)
            with open(ARQ_CONFIG, "w", encoding="utf-8") as f:
                json.dump(cfg, f, ensure_ascii=False, indent=1)
        except OSError:
            pass
        cfg["dpi"] = QUALIDADES.get(cfg["qualidade"], 200)
        cfg["limiar"] = SENSIBILIDADES.get(cfg["sensibilidade"], 0.96)
        return cfg

    def alternar(self):
        if self.monitorando:
            self.parar.set()
            self.status("Parando...")
            return
        cfg = self.coletar_cfg()
        if not cfg:
            return
        self.parar.clear()
        self.monitorando = True
        self.bt_iniciar.configure(text="■ Parar monitoramento")
        self.log(f"Monitoramento iniciado. Salvando em: {cfg['pasta']}")
        threading.Thread(target=self.monitorar, args=(cfg,), daemon=True).start()

    def digitalizar_agora(self):
        cfg = self.coletar_cfg()
        if not cfg:
            return
        if self.monitorando:  # o monitoramento já usa o scanner: entra na fila do próximo ciclo
            self.forcar.set()
            self.status("Digitalização manual na fila...")
            self.log("Digitalização manual pedida - será feita em poucos segundos.")
            return
        if not self.trava.acquire(blocking=False):
            messagebox.showinfo("Auto Scan", "Aguarde: uma digitalização já está em andamento.")
            return
        threading.Thread(target=self._manual, args=(cfg,), daemon=True).start()

    def diagnostico(self):
        """Mostra e copia um relatório do que o Windows enxerga do scanner."""
        if not self.trava.acquire(blocking=False):
            messagebox.showinfo("Auto Scan", "Aguarde: uma digitalização está em andamento.")
            return
        self.log("Gerando diagnóstico...")
        threading.Thread(target=self._diagnostico, daemon=True).start()

    def _diagnostico(self):
        pythoncom.CoInitialize()
        try:
            texto = gerar_diagnostico()
        except Exception as e:
            texto = f"Falha ao gerar o diagnóstico: {e}"
        finally:
            pythoncom.CoUninitialize()
            self.trava.release()
        caminho = os.path.join(os.path.expanduser("~"), "Documents", "diagnostico_autoscan.txt")
        try:
            os.makedirs(os.path.dirname(caminho), exist_ok=True)
            with open(caminho, "w", encoding="utf-8") as f:
                f.write(texto)
        except OSError:
            caminho = "(não consegui gravar o arquivo)"

        def mostrar():
            self.root.clipboard_clear()
            self.root.clipboard_append(texto)
            self.log("---- DIAGNÓSTICO (também copiado para a área de transferência) ----")
            for linha in texto.splitlines():
                self.log(linha)
            self.log(f"Salvo em: {caminho}")

        self._ui(mostrar)

    def fechar(self):
        self.fechando = True
        self.parar.set()
        self.root.destroy()

    # ---------- trabalho em segundo plano ----------
    def _manual(self, cfg):
        pythoncom.CoInitialize()
        final = None  # mensagem de status que deve ficar na tela ao terminar
        try:
            device = conectar(cfg["scanner"])
            if device is None:
                self.log("Scanner não encontrado. " + descrever_dispositivos())
                final = "Scanner não encontrado - veja o histórico."
                return
            self.status("Digitalizando pelo vidro...")
            arquivos = digitalizar_vidro(device, cfg["dpi"], cfg["colorido"], self.log)
            device = None
            self.vidro_armado = False  # evita o modo automático digitalizar a mesma folha de novo
            self.processar(arquivos, cfg)
        except Exception as e:
            self.log(f"Erro: {explicar_erro(e)}")
            final = "Erro - veja o histórico."
        finally:
            pythoncom.CoUninitialize()
            self.trava.release()
            self.status(final or ("Parado." if not self.monitorando else f"Aguardando folha {self.onde}..."))

    def _esperar(self, segundos, mostrar_contagem=False):
        """Espera, mas acorda na hora se o usuário parar ou pedir digitalização manual."""
        fim = time.time() + segundos
        base = self._status_base
        ultimo = None
        while not self.parar.is_set() and not self.forcar.is_set() and time.time() < fim:
            resta = int(fim - time.time()) + 1
            if mostrar_contagem and resta != ultimo:
                ultimo = resta
                self._ui(lambda m=f"{base}  | próxima verificação em {resta}s": self.v_status.set(m))
            self.parar.wait(0.2)

    def monitorar(self, cfg):
        pythoncom.CoInitialize()
        try:
            while not self.parar.is_set():
                espera = INTERVALO_VIDRO if self.onde == "no vidro" else cfg["intervalo"]
                device = None
                with self.trava:
                    try:
                        device = conectar(cfg["scanner"])
                        if device is None:
                            self.status("Scanner não encontrado. Tentando de novo...")
                            espera = ESPERA_ERRO
                            if self._ultimo_aviso != "sem_scanner":
                                self._ultimo_aviso = "sem_scanner"
                                self.log("Scanner não encontrado. " + descrever_dispositivos())
                        elif self.forcar.is_set():  # botão "Digitalizar agora" durante o monitoramento
                            self.forcar.clear()
                            self.msg_espera = None
                            self.status("Digitalização manual - digitalizando...")
                            arquivos = digitalizar_vidro(device, cfg["dpi"], cfg["colorido"], self.log)
                            device = None
                            self.vidro_armado = False
                            self.vidro_ultima = None
                            self.processar(arquivos, cfg)
                            self._ultimo_aviso = None
                        elif tem_alimentador(device):
                            self.onde = "no alimentador"
                            if tem_papel(device):
                                self.msg_espera = None
                                self.status("Papel detectado - digitalizando...")
                                arquivos = digitalizar_alimentador(device, cfg["dpi"], cfg["colorido"], self.log)
                                device = None
                                self.processar(arquivos, cfg)
                            else:
                                self.status(self.msg_espera or "Aguardando folha no alimentador...")
                            self._ultimo_aviso = None
                        else:
                            self.onde = "no vidro"
                            suporte = [device]
                            device = None  # só o ciclo segura a conexão, para poder soltá-la antes da contagem
                            self.ciclo_vidro(suporte, cfg)
                            self._ultimo_aviso = None
                    except Exception as e:
                        msg = explicar_erro(e)
                        ocupado = _codigo_erro(e) in (0x80210006, 0x8021000D)
                        espera = ESPERA_OCUPADO if ocupado else ESPERA_ERRO
                        nova = msg != self._ultimo_aviso
                        self._erros_iguais = 0 if nova else self._erros_iguais + 1
                        self.registrar_erro("Monitoramento", e, janela=nova or self._erros_iguais % 5 == 0)
                        if nova:
                            self._ultimo_aviso = msg
                            self.log(f"Vou tentar de novo sozinho a cada {espera:.0f} segundos.")
                        elif self._erros_iguais % 5 == 0:
                            self.log(f"(o mesmo erro se repetiu {self._erros_iguais + 1} vezes)")
                        self.status(f"Erro - tentando de novo em {espera:.0f}s (veja o histórico).")
                    finally:
                        device = None
                        gc.collect()  # solta a conexão com o scanner ANTES da próxima (evita \"ocupado\")
                self._esperar(espera, mostrar_contagem=(self.onde == "no vidro"))
        finally:
            pythoncom.CoUninitialize()
            self.monitorando = False
            self.forcar.clear()
            self.root.after(0, lambda: self.bt_iniciar.configure(text="▶ Iniciar monitoramento"))
            self.status("Parado.")
            self.log("Monitoramento parado.")

    def _digitalizar_com_tentativas(self, cfg):
        """Digitalização completa do vidro. RECONECTA ao scanner (a conexão da prévia já foi usada) e,
        se falhar, espera a contagem e tenta de novo. Cada erro aparece com o código."""
        for tent in range(1, TENTATIVAS_SCAN + 1):
            dev = None
            try:
                dev = conectar(cfg["scanner"])
                if dev is None:
                    raise RuntimeError("Scanner não encontrado ao reconectar. " + descrever_dispositivos())
                self.status(f"Digitalizando (tentativa {tent}/{TENTATIVAS_SCAN})...")
                return digitalizar_vidro(dev, cfg["dpi"], cfg["colorido"], self.log)
            except Exception as e:
                self.registrar_erro(f"digitalização completa, tentativa {tent}/{TENTATIVAS_SCAN}", e)
                if tent < TENTATIVAS_SCAN:
                    dev = None
                    gc.collect()
                    if not self.contagem(CONTAGEM_SEG, f"Nova tentativa {tent + 1}/{TENTATIVAS_SCAN}"):
                        return None
            finally:
                dev = None
                gc.collect()
        self.log("Não consegui digitalizar depois de várias tentativas. Veja o erro acima "
                 "(use \"Testar passo a passo\" e \"Abrir log\"). Vou continuar tentando nos próximos ciclos.")
        return None

    def ciclo_vidro(self, device, cfg):
        """Scanner só de vidro: prévia; achou folha -> contagem regressiva -> reconecta -> digitaliza."""
        device = suporte = device.pop()  # (recebe uma lista para a conexão poder ser solta aqui dentro)
        if not self.avisou_vidro:
            self.avisou_vidro = True
            self.log(f"Modo vidro: verifico o vidro a cada {INTERVALO_VIDRO:.0f} segundos; ao achar uma folha, "
                     f"faço uma contagem de {CONTAGEM_SEG} segundos (feche a tampa) e digitalizo.")
        arq = None
        try:
            arq = digitalizar_vidro(device, PREVIA_DPI, False, lambda m: None)[0]
            tem, tinta, info = analisar_previa(arq)
        finally:
            apagar([arq] if arq else [])
        device = suporte = None  # solta a conexão da prévia
        gc.collect()
        self.n_previa += 1
        medida = f"[prévia {self.n_previa}: brilho {info['fundo']:.0f}, marcas {info['tinta'] * 100:.2f}%]"

        if not tem:  # vidro vazio / tampa aberta: pronto para a próxima folha
            self.vidro_armado = True
            self.vidro_prev = None
            self.cont_folha = 0
            self.status((self.msg_espera or "Coloque a folha no vidro e feche a tampa...") + " " + medida)
            return
        if not self.vidro_armado:
            # a mesma folha ainda está lá? Se a imagem mudou bastante, é uma folha nova: rearma
            ultima = self.vidro_ultima
            if ultima is None or ultima.shape != tinta.shape or correlacao(tinta, ultima, 1) < REARMA_ABAIXO:
                self.vidro_armado = True
                self.log("Folha nova detectada no vidro.")
            else:
                self.status("Retire a folha já digitalizada e coloque a próxima no vidro... " + medida)
                return

        self.msg_espera = None
        self.log(f"Folha detectada {medida}. Contagem de {CONTAGEM_SEG}s para digitalizar "
                 "(feche a tampa e não mexa na folha).")
        if not self.contagem(CONTAGEM_SEG, "Digitalizando"):
            return
        arquivos = self._digitalizar_com_tentativas(cfg)
        if not arquivos:
            return  # erro já mostrado; o próximo ciclo tenta de novo
        _, _, info2 = analisar_previa(arquivos[0])
        if info2["fundo"] <= BRILHO_MIN:  # imagem escura = tampa aberta / folha tirada durante a contagem
            self.log("A imagem digitalizada saiu escura (tampa aberta ou folha retirada). Descartada - "
                     "feche a tampa com a folha no vidro.")
            apagar(arquivos)
            return
        self.vidro_armado = False
        self.vidro_ultima = tinta
        self.vidro_prev = None
        self.cont_folha = 0
        self.processar(arquivos, cfg)

    def processar(self, arquivos, cfg):
        """Verifica repetidas, pergunta o que fazer ao final e salva o PDF."""
        self._salvou_ok = True
        try:
            self._processar(arquivos, cfg)
        finally:
            if self._salvou_ok:
                apagar(arquivos)  # garante que nenhum arquivo temporário fique para trás

    def _processar(self, arquivos, cfg):
        if not arquivos:
            return
        pasta = cfg["pasta"]
        indice = carregar_indice(pasta)
        aceitas = []  # lista de (arquivo, assinatura ou None)

        for i, arq in enumerate(arquivos, start=1):
            sig = assinatura(arq)  # None = folha em branco (nunca marcada como repetida)
            parecida = None
            if cfg["avisar_repetida"] and sig is not None:
                item = buscar_parecida(sig, indice, pasta, cfg["limiar"])
                if item:
                    parecida = f"{item['arquivo']} ({item['data']})"
                else:
                    for _, outra in aceitas:
                        if outra is not None and float(sig[1] @ outra[1]) >= PRE_FILTRO \
                                and correlacao(sig[0], outra[0]) >= cfg["limiar"]:
                            parecida = "outra folha deste mesmo lote"
                            break
            if parecida:
                self.status("Folha repetida - aguardando sua resposta...")
                r = self.na_tela(lambda i=i, p=parecida: self.perguntar(
                    "Folha repetida",
                    f"A folha {i} parece IGUAL a uma já digitalizada:\n\n  {p}\n\n"
                    "Deseja salvar essa folha de novo?",
                    [("salvar", "Salvar mesmo assim"), ("descartar", "Descartar esta folha")]))
                if r == "descartar":
                    self.log(f"Folha {i} descartada (repetida).")
                    apagar([arq])
                    continue
            aceitas.append((arq, sig))

        if not aceitas:
            self.log("Nenhuma folha para salvar.")
            return

        escolha = "trocar"
        if cfg["perguntar_fim"]:
            self.status("Digitalização concluída - aguardando sua resposta...")
            n = len(aceitas)
            escolha = self.na_tela(lambda: self.perguntar(
                "Digitalização concluída",
                f"{n} folha(s) digitalizada(s).\n\nO que deseja fazer?\n\n"
                "• Salvar e trocar de folha: salva o arquivo; retire as folhas e "
                "coloque as próximas no alimentador.\n"
                "• Repetir digitalização: descarta esta digitalização; recoloque a(s) "
                "folha(s) para digitalizar de novo.\n"
                "• Salvar e parar: salva o arquivo e encerra o monitoramento.",
                [("trocar", "Salvar e trocar de folha"),
                 ("repetir", "Repetir digitalização"),
                 ("parar", "Salvar e parar")]))

        if escolha == "repetir":
            apagar([a for a, _ in aceitas])
            if self.onde == "no vidro":  # a folha continua no vidro: digitaliza de novo sozinho
                self.vidro_armado = True
                self.vidro_prev = None
                self.vidro_ultima = None
                self.cont_folha = 0
                self.log("Digitalização descartada. A folha continua no vidro: vou digitalizar de novo.")
                self.msg_espera = "Digitalizando a mesma folha de novo..."
            else:
                self.log("Digitalização descartada. Recoloque a folha para digitalizar de novo.")
                self.msg_espera = "Recoloque a folha no alimentador para digitalizar de novo..."
            self.status(self.msg_espera)
            return

        try:
            destino = salvar_pdf([a for a, _ in aceitas], pasta, cfg["dpi"])
        except Exception as e:
            self._salvou_ok = False  # mantém as imagens temporárias para você não perder a digitalização
            pasta_tmp = os.path.dirname(aceitas[0][0])
            self.log(f"ERRO ao salvar o PDF: {e}")
            self.na_tela(lambda: self.perguntar(
                "Não foi possível salvar",
                f"Não consegui salvar em:\n{pasta}\n\n{e}\n\n"
                f"As imagens digitalizadas foram mantidas na pasta temporária:\n{pasta_tmp}\n"
                "(arquivos autoscan_*.png). Verifique a pasta de destino e digitalize novamente se preciso.",
                [("ok", "OK")]))
            self.status("Erro ao salvar - veja o aviso.")
            return
        for _, sig in aceitas:
            if sig is not None:
                guardar_assinatura(pasta, sig, os.path.basename(destino), indice)
        gravar_indice(pasta, indice)
        apagar([a for a, _ in aceitas])
        self.log(f"Salvo: {destino}")

        if escolha == "parar":
            self.parar.set()
        else:
            self.msg_espera = f"Troque a folha {self.onde}..."
            self.status(self.msg_espera)


if __name__ == "__main__":
    janela = tk.Tk()
    App(janela)
    janela.mainloop()