import io
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog

import streamlit as st

from verificador import CANDIDATOS_DIR, classificar_candidatos

# ── Configuração da página ────────────────────────────────────────────────────

st.set_page_config(
    page_title="Verificador de Bolsas — PPG CMH / UFRGS",
    page_icon="🎓",
    layout="centered",
)

# ── Cabeçalho ─────────────────────────────────────────────────────────────────

col_logo, col_titulo = st.columns([1, 3])
with col_logo:
    logo = Path(__file__).parent / "logo.jpeg"
    if logo.exists():
        st.image(str(logo), width=110)
with col_titulo:
    st.title("Verificador de Bolsas")
    st.caption(
        "Programa de Pós-Graduação em Ciências do Movimento Humano · UFRGS"
    )
   

st.divider()

# ── Transfere valores selecionados via diálogo ANTES de criar os widgets ──────
# Streamlit não permite modificar session_state[key] após o widget ser criado.
# A solução: salvar em chave temporária (_novo) e transferir no início do ciclo.

for _chave in ("pasta_candidatos", "pasta_saida"):
    _novo = st.session_state.pop(f"_{_chave}_novo", None)
    if _novo:
        st.session_state[_chave] = _novo

if "pasta_candidatos" not in st.session_state:
    padrao = CANDIDATOS_DIR if CANDIDATOS_DIR.exists() else Path.cwd() / "candidatos"
    st.session_state["pasta_candidatos"] = str(padrao)

if "pasta_saida" not in st.session_state:
    st.session_state["pasta_saida"] = str(Path.cwd())

# ── Seleção de pastas ─────────────────────────────────────────────────────────

def _selecionar_pasta(chave_temp: str, titulo: str) -> None:
    """Abre diálogo nativo do SO e guarda o resultado em chave temporária."""
    resultado = [None]

    def _run():
        root = tk.Tk()
        root.withdraw()
        root.wm_attributes("-topmost", 1)
        pasta = filedialog.askdirectory(title=titulo)
        root.destroy()
        resultado[0] = pasta

    t = threading.Thread(target=_run)
    t.start()
    t.join()

    if resultado[0]:
        st.session_state[chave_temp] = resultado[0]


# Campo: pasta dos candidatos
st.subheader("1. Pasta dos candidatos")
st.caption(
    "Pasta que contém uma subpasta por candidato, cada uma com lattes.xml e os PDFs comprovantes."
)
col_input, col_btn = st.columns([4, 1])
with col_input:
    st.text_input(
        "Pasta dos candidatos",
        key="pasta_candidatos",
        label_visibility="collapsed",
    )
with col_btn:
    st.write("")
    if st.button("📁 Selecionar", key="btn_candidatos"):
        _selecionar_pasta("_pasta_candidatos_novo", "Selecionar pasta dos candidatos")
        st.rerun()

st.divider()

# Campo: pasta de saída
st.subheader("2. Pasta de saída")
st.caption("Onde os arquivos ranking.xlsx e relatórios individuais serão salvos.")
col_input2, col_btn2 = st.columns([4, 1])
with col_input2:
    st.text_input(
        "Pasta de saída",
        key="pasta_saida",
        label_visibility="collapsed",
    )
with col_btn2:
    st.write("")
    if st.button("📁 Selecionar", key="btn_saida"):
        _selecionar_pasta("_pasta_saida_novo", "Selecionar pasta de saída")
        st.rerun()

st.divider()

# ── Execução ──────────────────────────────────────────────────────────────────

if st.button("▶  Executar verificação", type="primary", use_container_width=True):
    pasta_cand = Path(st.session_state["pasta_candidatos"])
    pasta_out  = Path(st.session_state["pasta_saida"])

    if not pasta_cand.is_dir():
        st.error(f"Pasta de candidatos não encontrada:\n{pasta_cand}")
        st.stop()
    if not pasta_out.is_dir():
        st.error(f"Pasta de saída não encontrada:\n{pasta_out}")
        st.stop()

    buffer = io.StringIO()
    saida_ranking = None
    erro = None

    with st.spinner("Processando candidatos... Aguarde."):
        sys.stdout = buffer
        try:
            saida_ranking = classificar_candidatos(
                candidatos_dir=pasta_cand,
                pasta_saida=pasta_out,
            )
        except Exception as e:
            erro = e
        finally:
            sys.stdout = sys.__stdout__

    if erro:
        st.error(f"Erro durante a verificação:\n{erro}")
    else:
        st.success("Verificação concluída!")
        if saida_ranking:
            st.info(f"Ranking salvo em:\n`{saida_ranking}`")

    log = buffer.getvalue()
    if log:
        with st.expander("Ver log detalhado"):
            st.text(log)
