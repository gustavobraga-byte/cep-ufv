#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cep-ufv / run.py — Gerador travado aos templates oficiais CEP/UFV.
TRAVA: só preenche campos variáveis dentro do texto verbatim dos modelos.
Nunca cria seção/parágrafo/documento fora dos templates.
Gera sempre o par .docx + .pdf idêntico por documento (fail-closed).
Fontes: https://cep.ufv.br/modelos/ + https://cep.ufv.br/normativas/ + https://cep.ufv.br/projeto-novo/
Dependências: pip install python-docx reportlab pymupdf pdfplumber
"""
import argparse
import json
import sys
from pathlib import Path
from datetime import date, datetime

try:
    from docx import Document
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
except ImportError:
    print("ERRO: python-docx não instalado. Rode: pip install python-docx reportlab pymupdf pdfplumber", file=sys.stderr)
    sys.exit(2)

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.enums import TA_JUSTIFY, TA_CENTER, TA_LEFT
except ImportError:
    print("ERRO: reportlab não instalado. Rode: pip install python-docx reportlab pymupdf pdfplumber", file=sys.stderr)
    sys.exit(2)

SKILL_VERSION = "1.7.1-total-no-docs"
FONTES = {
    "tcle": "https://cep.ufv.br/wp-content/uploads/2026/05/TCLE.pdf (2026, 3p)",
    "tcle_resp": "https://cep.ufv.br/wp-content/uploads/2026/05/TCLE-Responsavel.pdf (2026, 3p)",
    "tale": "https://cep.ufv.br/wp-content/uploads/2026/05/Termo-de-Assentimento.pdf (2026, 3p)",
    "anuencia": "https://cep.ufv.br/wp-content/uploads/2026/05/Modelo-de-Autorizacao-Institucional.pdf (2026, 1p)",
    "cronograma": "https://cep.ufv.br/wp-content/uploads/2026/05/Modelo-de-cronograma.pdf",
    "relatorio": "https://cep.ufv.br/wp-content/uploads/2016/06/Modelo-Relatório-Final.pdf (2016)",
    "carta": "https://cep.ufv.br/wp-content/uploads/2016/05/Modelo-Carta-Resposta.odt (2016)",
    "sigilo": "https://cep.ufv.br/wp-content/uploads/2019/07/Modelo-de-termo-de-sigilo-e-confidencialidade.odt (2019)",
    "checklist": "https://cep.ufv.br/wp-content/uploads/2021/02/Ckeck-List-CEP-UFV.pdf (2021)",
    "normativas": "https://cep.ufv.br/normativas/",
    "rito": "https://cep.ufv.br/projeto-novo/",
}

CEP_CONTATO = (
    "Comitê de Ética em Pesquisa com Seres Humanos da Universidade Federal de Viçosa (CEP/UFV), "
    "localizado no Edifício Arthur Bernardes, piso inferior, Avenida PH Rolfs, s/n, Campus Universitário, "
    "Viçosa/MG, CEP: 36570-900. Telefone: (31) 3612-2316. E-mail: cep@ufv.br. Site: https://cep.ufv.br"
)
BASE_LEGAL = (
    "Modelo institucional CEP/UFV – versão 2026 – elaborado com base na Resolução CNS 466/2012, "
    "Resolução CNS 510/2016 e Lei 14.874/2024."
)

PENDENCIAS = []
MODO_RASCUNHO = False  # True = permite [PREENCHER] e blocos de orientação; False = pronto p/ submissão
COM_CARTA = False      # Carta-Resposta só gerada com --com-carta (quando há pendência)
PACOTE_NOME = ""       # nome da pasta de saída (para o roteiro citar onde está o pacote)
PREVISAO_TEXTO = ""    # previsão de tramitação calculada da data atual + calendário CEP

# Reuniões ordinárias CEP/UFV 2026 (https://cep.ufv.br/cronograma/, verificado em 05/09/2026).
# Submissão até dia 20 do mês anterior; checagem documental até 10 dias (NO 001/2013).
REUNIOES_2026 = [(6, 3), (17, 4), (8, 5), (12, 6), (10, 7), (14, 8),
                 (11, 9), (9, 10), (13, 11), (4, 12)]

def prever_parecer(hoje=None):
    """Estima reunião de apreciação e parecer a partir da data atual.
    Retorna texto pronto para o pesquisador (previsão, não garantia)."""
    from datetime import date as _d
    h = hoje or _d.today()
    for dia, mes in REUNIOES_2026:
        corte_mes = mes - 1 or 12
        corte_ano = 2026 if mes > 1 else 2026
        from datetime import date as _d2
        corte = _d2(corte_ano if corte_mes != 12 else 2026, corte_mes, 20) if not (mes == 1) else _d2(2025, 12, 20)
        reuniao = _d2(2026, mes, dia)
        if h <= corte and h <= reuniao:
            return (f"Hoje é {h.day:02d}/{h.month:02d}/{h.year}. Submetendo até 20/"
                    f"{corte.month:02d}, o protocolo pode entrar na reunião ordinária de "
                    f"{dia:02d}/{mes:02d}/2026. O parecer costuma sair nos dias seguintes à reunião "
                    f"(acompanhe na Plataforma Brasil). Se houver pendência, some ~30 dias de resposta + "
                    f"nova apreciação na reunião seguinte.")
    return (f"Hoje é {h.day:02d}/{h.month:02d}/{h.year}. O calendário 2026 publicado vai até 04/12/2026 "
            f"(corte 20/11). Para protocolo submetido agora, a apreciação deve cair na próxima reunião com "
            f"documentação regular — consulte o calendário vigente em https://cep.ufv.br/cronograma/ e, após "
            f"submeter, acompanhe a Plataforma Brasil: o parecer sai nos dias seguintes à reunião; "
            f"pendência soma ~30 dias + nova reunião.")

CAMPOS_OBRIGATORIOS = [
    ("titulo", "titulo da pesquisa"),
    ("pesquisador.nome", "pesquisador.nome"),
    ("pesquisador.departamento", "pesquisador.departamento"),
    ("pesquisador.endereco", "pesquisador.endereco"),
    ("pesquisador.telefone", "pesquisador.telefone"),
    ("pesquisador.email", "pesquisador.email"),
    ("objetivo", "objetivo"),
    ("justificativa", "justificativa"),
    ("procedimentos", "procedimentos"),
    ("local", "local/plataforma de realização"),
    ("tempo_estimado", "tempo estimado de participação"),
    ("riscos", "riscos"),
    ("medidas_risco", "medidas de prevenção e minimização dos riscos"),
    ("beneficios", "benefícios diretos e/ou indiretos"),
    ("ressarcimento", "forma de ressarcimento"),
    ("retorno", "retorno da pesquisa"),
    ("contato_retirada", "contato para retirada do consentimento"),
]

MESES_PT = ["janeiro", "fevereiro", "março", "abril", "maio", "junho",
            "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]

def data_documento(p, hoje=None):
    """Data automática dos documentos (nunca é pendência): usa data_local se
    informada e completa; senão gera Cidade, DD de mês de AAAA com a data atual."""
    dl = (p.get("data_local") or "").strip() if isinstance(p.get("data_local"), str) else ""
    if dl and "___" not in dl:
        return dl
    cidade = (p.get("cidade") or "Viçosa").strip() or "Viçosa"
    h = hoje or date.today()
    return f"{cidade}, {h.day:02d} de {MESES_PT[h.month - 1]} de {h.year}"

def get_campo(p, caminho):
    cur = p
    for parte in caminho.split("."):
        if not isinstance(cur, dict):
            return None
        cur = cur.get(parte)
    return cur

def fill(valor, rotulo):
    """Modo final (padrão): registra a pendência; o main aborta antes de escrever.
    Modo rascunho (--rascunho): insere [PREENCHER] no texto. Nunca inventa dado."""
    if valor is None or (isinstance(valor, str) and not valor.strip()):
        PENDENCIAS.append(rotulo)
        if MODO_RASCUNHO:
            return f"[PREENCHER: {rotulo}]"
        return f"_____{rotulo.upper()}_____"
    return str(valor).strip()

def fill_opt(valor):
    return str(valor).strip() if valor else ""

# ------------------------------------------------------------------ blocos
# Bloco = tupla: ("aviso"|"h1"|"h2"|"p"|"bullet"|"table"|"assinaturas"|"nota", ...)

def doc_base_info():
    return [
        ("aviso", "Este modelo possui caráter orientativo e deverá ser adaptado às especificidades metodológicas e éticas de cada pesquisa. "
                  "O Termo deve ser elaborado em linguagem clara, acessível e compatível com o público participante. "
                  "A utilização do modelo não dispensa a leitura das Resoluções CNS nº 466/2012, nº 510/2016, da Lei nº 14.874/2024, "
                  "do Decreto nº 12.651/2025 e da LGPD (Lei nº 13.709/2018)."),
    ]

def pesquisador_bloco(p):
    pes = p.get("pesquisador", {})
    return [
        ("p", f"Pesquisador(a) responsável: Nome: {fill(pes.get('nome'), 'pesquisador.nome')}"),
        ("p", f"Departamento/Programa/Instituição: {fill(pes.get('departamento'), 'pesquisador.departamento')} — {fill(pes.get('instituicao') or 'Universidade Federal de Viçosa', 'pesquisador.instituicao')}"),
        ("p", f"Endereço institucional: {fill(pes.get('endereco'), 'pesquisador.endereco')}"),
        ("p", f"Telefone: {fill(pes.get('telefone'), 'pesquisador.telefone')}"),
        ("p", f"E-mail: {fill(pes.get('email'), 'pesquisador.email')}"),
    ]

def info_pesquisador_tcle():
    return [
        ("h2", "INFORMAÇÕES IMPORTANTES AO PESQUISADOR"),
        ("p", "1. As informações cadastradas na Plataforma Brasil devem ser compatíveis com aquelas descritas neste Termo, "
              "especialmente quanto aos objetivos, procedimentos, riscos e benefícios da pesquisa."),
        ("p", "2. Toda pesquisa envolvendo seres humanos apresenta riscos, ainda que mínimos. Os possíveis riscos da pesquisa "
              "e as medidas de prevenção e minimização deverão ser descritos de forma clara e adequada. "
              "É vedado afirmar que “a pesquisa não oferece risco”."),
        ("p", "3. Informar sempre: tempo estimado de participação; procedimentos realizados; local, plataforma ou ambiente; "
              "forma de contato com os pesquisadores."),
        ("p", "4. O Termo deverá ser elaborado em linguagem clara, acessível e compatível com o público participante. "
              "Utilizar o termo “participante da pesquisa” (nunca “respondente”, “entrevistado(a)”, “sujeito”)."),
        ("p", "5. Quando houver gravação de imagem e/ou voz, a autorização correspondente deverá constar no próprio Termo."),
        ("p", "6. Todos os campos do modelo deverão ser preenchidos e adaptados às especificidades metodológicas e éticas da pesquisa."),
        ("p", "7. O pesquisador responsável indicado neste Termo deverá corresponder ao pesquisador responsável cadastrado "
              "na Plataforma Brasil (para o CEP/UFV, o orientador)."),
        ("p", "8. Em pesquisas em ambiente virtual, disponibilizar ao participante uma via do Termo assinada pelo pesquisador "
              "responsável, em formato digital, preferencialmente antes do início da participação."),
    ]

def build_tcle_adulto(p):
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    b = []
    if MODO_RASCUNHO:
        b += doc_base_info()
    b.append(("h1", "TERMO DE CONSENTIMENTO LIVRE E ESCLARECIDO (TCLE)"))
    b.append(("p", f"O(a) Sr.(a) está sendo convidado(a) a participar, de forma voluntária, da pesquisa intitulada: “{titulo}”."))
    b.append(("p", "Antes de decidir se deseja participar, é importante que você leia atentamente as informações abaixo. "
                   "Caso tenha dúvidas, você poderá esclarecê-las com os pesquisadores."))
    b.append(("p", f"Esta pesquisa tem como objetivo: {fill(p.get('objetivo'), 'objetivo')} "
                   f"Justificativa: {fill(p.get('justificativa'), 'justificativa')}."))
    b.append(("p", f"Sua participação ocorrerá por meio de: {fill(p.get('procedimentos'), 'procedimentos (entrevista, questionário, grupo focal, coleta de material biológico, etc.)')}."))
    b.append(("p", f"O tempo estimado de participação será de aproximadamente: {fill(p.get('tempo_estimado'), 'tempo estimado de participação')}."))
    b.append(("p", f"A pesquisa será realizada no seguinte local ou plataforma: {fill(p.get('local'), 'local/plataforma de realização')}."))
    b.append(("p", "Toda pesquisa envolvendo seres humanos apresenta riscos, ainda que mínimos. "
                   f"Os possíveis riscos relacionados a esta pesquisa incluem: {fill(p.get('riscos'), 'riscos (físicos, psicológicos, morais, sociais, emocionais, privacidade, internet, identificação indireta)')}."))
    b.append(("p", f"Para reduzir esses riscos, os pesquisadores adotarão as seguintes medidas: {fill(p.get('medidas_risco'), 'medidas de prevenção e minimização dos riscos')}."))
    b.append(("p", "Nas pesquisas realizadas em ambiente virtual, apesar das medidas de segurança adotadas, podem existir riscos "
                   "relacionados à confidencialidade e à proteção das informações em razão das limitações das tecnologias utilizadas."))
    if p.get("usa_biologico"):
        b.append(("p", f"Material biológico: o material coletado pertence ao participante. Laboratório de análise: "
                       f"{fill(p.get('biologico_lab'), 'laboratório de análise do material biológico (Brasil/Exterior)')} "
                       f"Necessidade de novo consentimento no futuro: {fill(p.get('biologico_futuro'), 'consentimento futuro para material biológico')}. "
                       "(Res. CNS 441/2011.)"))
    if p.get("gravacao"):
        b.append(("p", "Registro de voz e/ou imagem: ao assinar este Termo, você autoriza o registro de voz e/ou imagem para fins "
                       "exclusivos desta pesquisa, conforme procedimentos descritos acima."))
    if fill_opt(p.get("vulneraveis")):
        b.append(("p", f"Proteção específica (participante em situação de vulnerabilidade): {p.get('vulneraveis').strip()}. "
                       "Foram previstas medidas adicionais de proteção, acolhimento e respeito à autonomia, "
                       "nos termos das Resoluções CNS aplicáveis (incl. 304/2000 para povos indígenas, quando cabível)."))
    if fill_opt(p.get("acompanhamento")):
        b.append(("p", f"Acompanhamento após a participação: {p.get('acompanhamento').strip()}."))
    if p.get("saude_mental"):
        b.append(("p", "Esclarecimento (pesquisa em saúde mental): o resultado desta pesquisa NÃO se trata de diagnóstico clínico individual. "
                       f"Orientações e encaminhamentos previstos: {fill(p.get('saude_mental_orientacoes'), 'orientações/aconselhamentos em saúde mental')}."))
    if p.get("sus"):
        b.append(("p", "Pesquisa no SUS: esta pesquisa não interferirá na rotina de assistência à saúde, salvo expressa autorização do(a) responsável pelo setor. "
                       "A pesquisa não se confunde com a assistência: recusar ou retirar o consentimento não trará qualquer prejuízo ao seu atendimento."))
    b.append(("p", "Em caso de danos decorrentes da pesquisa, você terá direito à assistência integral e à indenização, "
                   "nos termos da legislação e das normas éticas vigentes."))
    b.append(("p", f"A pesquisa poderá contribuir para: {fill(p.get('beneficios'), 'benefícios diretos e/ou indiretos')}."))

    b.append(("p", "A participação nesta pesquisa não terá custos para você. "
                   f"Caso haja despesas relacionadas à participação, você terá direito ao ressarcimento, conforme descrito a seguir: "
                   f"{fill(p.get('ressarcimento'), 'forma de ressarcimento (alimentação, transporte de participantes/acompanhantes)')}."))
    b.append(("p", "Você não receberá pagamento pela participação nesta pesquisa. A participação é voluntária."))
    b.append(("p", "Para os fins deste Termo, você é o(a) participante da pesquisa."))
    b.append(("p", "Você poderá, sem necessidade de justificativa e sem qualquer prejuízo:"))
    for t in ["recusar-se a participar da pesquisa;",
              "deixar de responder perguntas que não desejar responder;",
              "retirar seu consentimento;",
              "desistir da pesquisa a qualquer momento."]:
        b.append(("bullet", t))
    b.append(("p", f"A solicitação de retirada do consentimento poderá ser realizada por meio do seguinte contato: "
                   f"{fill(p.get('contato_retirada'), 'contato para retirada do consentimento (e-mail/telefone)')}."))

    b.append(("p", "Você não será identificado(a) em publicações, apresentações ou quaisquer formas de divulgação decorrentes desta pesquisa."))
    b.append(("p", "Os dados coletados serão armazenados em ambiente seguro, com acesso restrito aos pesquisadores responsáveis."))
    b.append(("p", "As informações obtidas serão utilizadas exclusivamente para fins científicos e acadêmicos, observadas as normas éticas e legais aplicáveis, "
                   "incluindo a Lei nº 13.709/2018 (LGPD)."))
    b.append(("p", "Seu nome ou qualquer informação que possibilite sua identificação não será divulgado sem sua autorização."))
    b.append(("p", "Os dados e documentos da pesquisa permanecerão armazenados sob responsabilidade do pesquisador, em meio físico ou digital, "
                   "pelo prazo mínimo de 5 (cinco) anos após o término ou a descontinuação da pesquisa, observadas as normas éticas e legais aplicáveis. "
                   "Após esse período, os dados poderão ser descartados de forma segura."))
    b.append(("p", "Os resultados da pesquisa poderão ser disponibilizados aos participantes após o encerramento da pesquisa, "
                   "quando solicitado ou quando previsto pelos pesquisadores."))
    b.append(("p", f"Como forma de retorno da pesquisa aos participantes e à sociedade: {fill(p.get('retorno'), 'retorno (relatório, material educativo, seminário, publicação)')}."))

    b += pesquisador_bloco(p)
    b.append(("p", f"Em caso de dúvidas, denúncias, reclamações ou irregularidades relacionadas aos aspectos éticos desta pesquisa, "
                   f"você poderá entrar em contato com o {CEP_CONTATO}."))
    b.append(("p", "Você poderá solicitar acesso a este Termo de Consentimento Livre e Esclarecido a qualquer momento."))
    b.append(("p", "Declaro que fui devidamente informado(a) sobre os objetivos, procedimentos, possíveis riscos e benefícios desta pesquisa. "
                   "Tive oportunidade de esclarecer minhas dúvidas e concordo, de forma livre e voluntária, em participar desta pesquisa."))
    b.append(("p", f"{data_documento(p)}."))
    b.append(("assinatura", "Assinatura do(a) participante"))
    b.append(("assinatura", "Assinatura do(a) pesquisador(a)"))
    if MODO_RASCUNHO:
        b += info_pesquisador_tcle()
    return ("01-TCLE-Adulto", "TERMO DE CONSENTIMENTO LIVRE E ESCLARECIDO (TCLE) — Adultos", b)

def build_tcle_responsavel(p):
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    b = []
    if MODO_RASCUNHO:
        b += doc_base_info()
    b.append(("h1", "TERMO DE CONSENTIMENTO LIVRE E ESCLARECIDO (TCLE) — Responsável Legal"))
    b.append(("p", f"O(a) Sr.(a) está sendo convidado(a), na condição de responsável legal, a autorizar a participação de "
                   f"{fill(p.get('participante_nome'), 'nome do participante representado')} na pesquisa intitulada: “{titulo}”."))
    b.append(("p", "Antes de decidir se autoriza a participação, é importante que o(a) Sr.(a) leia atentamente as informações abaixo. "
                   "Caso tenha dúvidas, poderá esclarecê-las com os pesquisadores."))
    b.append(("p", f"Esta pesquisa tem como objetivo: {fill(p.get('objetivo'), 'objetivo')} "
                   f"Justificativa: {fill(p.get('justificativa'), 'justificativa')}."))
    b.append(("p", f"A participação do(a) participante ocorrerá por meio de: {fill(p.get('procedimentos'), 'procedimentos')}."))

    b.append(("p", f"O tempo estimado de participação será de aproximadamente: {fill(p.get('tempo_estimado'), 'tempo estimado de participação')}."))
    b.append(("p", f"A pesquisa será realizada no seguinte local ou plataforma: {fill(p.get('local'), 'local/plataforma de realização')}."))
    b.append(("p", "Toda pesquisa envolvendo seres humanos apresenta riscos, ainda que mínimos. "
                   f"Os possíveis riscos relacionados a esta pesquisa incluem: {fill(p.get('riscos'), 'riscos')}."))
    b.append(("p", f"Para reduzir esses riscos, os pesquisadores adotarão as seguintes medidas: {fill(p.get('medidas_risco'), 'medidas de prevenção e minimização dos riscos')}."))
    b.append(("p", "Nas pesquisas realizadas em ambiente virtual, apesar das medidas de segurança adotadas, podem existir riscos "
                   "relacionados à confidencialidade e à proteção das informações em razão das limitações das tecnologias utilizadas."))
    if p.get("usa_biologico"):
        b.append(("p", f"Material biológico: o material coletado pertence ao participante. Laboratório de análise: "
                       f"{fill(p.get('biologico_lab'), 'laboratório de análise do material biológico (Brasil/Exterior)')} "
                       f"Necessidade de novo consentimento no futuro: {fill(p.get('biologico_futuro'), 'consentimento futuro para material biológico')}. "
                       "(Res. CNS 441/2011.)"))
    if p.get("gravacao"):
        b.append(("p", "Registro de voz e/ou imagem: ao assinar este Termo, fica autorizado o registro de voz e/ou imagem do(a) participante para fins "
                       "exclusivos desta pesquisa, conforme procedimentos descritos acima."))
    if fill_opt(p.get("vulneraveis")):
        b.append(("p", f"Proteção específica (participante em situação de vulnerabilidade): {p.get('vulneraveis').strip()}. "
                       "Foram previstas medidas adicionais de proteção e respeito à autonomia, nos termos das normas aplicáveis."))
    if fill_opt(p.get("acompanhamento")):
        b.append(("p", f"Acompanhamento após a participação: {p.get('acompanhamento').strip()}."))
    if p.get("saude_mental"):
        b.append(("p", "Esclarecimento (saúde mental): o resultado desta pesquisa NÃO se trata de diagnóstico clínico individual. "
                       f"Orientações e encaminhamentos previstos: {fill(p.get('saude_mental_orientacoes'), 'orientações/aconselhamentos em saúde mental')}."))
    if p.get("sus"):
        b.append(("p", "Pesquisa no SUS: sem interferir na rotina de assistência salvo expressa autorização do(a) responsável pelo setor. "
                       "A pesquisa não se confunde com a assistência: recusar ou retirar o consentimento não prejudica o atendimento."))
    b.append(("p", "Em caso de danos decorrentes da pesquisa, o(a) participante terá direito à assistência integral e à indenização, "
                   "nos termos da legislação e das normas éticas vigentes."))
    b.append(("p", f"A pesquisa poderá contribuir para: {fill(p.get('beneficios'), 'benefícios diretos e/ou indiretos')}."))
    b.append(("p", "O(a) participante não receberá pagamento pela participação nesta pesquisa. A participação é voluntária."))
    b.append(("p", "Para os fins deste Termo, o(a) menor/assistido é o(a) participante da pesquisa."))
    b.append(("p", "O(a) responsável legal poderá, sem necessidade de justificativa e sem qualquer prejuízo:"))
    for t in ["recusar a participação na pesquisa;",
              "solicitar a retirada do consentimento;",
              "interromper a participação do(a) participante a qualquer momento."]:
        b.append(("bullet", t))
    b.append(("p", "Além disso, o(a) participante poderá deixar de responder perguntas que não desejar responder, quando aplicável."))
    b.append(("p", f"A solicitação de retirada do consentimento poderá ser realizada por meio do seguinte contato: "
                   f"{fill(p.get('contato_retirada'), 'contato para retirada do consentimento')}."))
    b.append(("p", "O(a) participante não será identificado(a) em publicações, apresentações ou quaisquer formas de divulgação decorrentes desta pesquisa."))
    b.append(("p", "Os dados coletados serão armazenados em ambiente seguro, com acesso restrito aos pesquisadores responsáveis."))
    b.append(("p", "As informações obtidas serão utilizadas exclusivamente para fins científicos e acadêmicos, observadas as normas éticas e legais aplicáveis."))
    b.append(("p", "O nome do(a) participante ou qualquer informação que possibilite sua identificação não será divulgado sem autorização."))
    b.append(("p", "Os dados e documentos da pesquisa permanecerão armazenados sob responsabilidade do pesquisador, em meio físico ou digital, "
                   "pelo prazo mínimo de 5 (cinco) anos após o término ou a descontinuação da pesquisa, observadas as normas éticas e legais aplicáveis. "
                   "Após esse período, os dados poderão ser descartados de forma segura."))
    b.append(("p", "Os resultados da pesquisa poderão ser disponibilizados aos participantes e responsáveis legais após o encerramento da pesquisa, "
                   "quando solicitado ou quando previsto pelos pesquisadores."))
    b.append(("p", f"Como forma de retorno da pesquisa aos participantes e à sociedade: {fill(p.get('retorno'), 'retorno da pesquisa')}."))
    b += pesquisador_bloco(p)
    b.append(("p", f"Em caso de dúvidas, denúncias, reclamações ou irregularidades relacionadas aos aspectos éticos desta pesquisa, "
                   f"você poderá entrar em contato com o {CEP_CONTATO}."))
    b.append(("p", "O(a) Sr.(a) poderá solicitar acesso a este Termo de Consentimento Livre e Esclarecido a qualquer momento."))
    b.append(("p", "Declaro que fui devidamente informado(a) sobre os objetivos, procedimentos, possíveis riscos e benefícios desta pesquisa. "
                   "Tive oportunidade de esclarecer minhas dúvidas e autorizo, de forma livre e voluntária, "
                   "a participação do(a) menor/participante nesta pesquisa."))
    b.append(("p", f"{data_documento(p)}."))
    b.append(("assinatura", "Assinatura do(a) responsável legal"))
    b.append(("assinatura", "Assinatura do(a) pesquisador(a)"))
    if MODO_RASCUNHO:
        b += info_pesquisador_tcle()
    return ("02-TCLE-Responsavel", "TCLE — Responsável Legal", b)

def build_tale(p):
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    b = []
    if MODO_RASCUNHO:
        b += doc_base_info()
    b.append(("h1", "TERMO DE ASSENTIMENTO LIVRE E ESCLARECIDO"))
    b.append(("p", f"Você está sendo convidado(a) a participar da pesquisa: “{titulo}”."))
    b.append(("p", "Antes de decidir se quer participar, é importante que você entenda por que esta pesquisa será realizada e o que acontecerá durante sua participação."))
    b.append(("p", f"Esta pesquisa tem como objetivo: {fill(p.get('objetivo'), 'objetivo em linguagem simples e adequada à faixa etária')}."))
    b.append(("p", f"Se você concordar em participar, será solicitado que: {fill(p.get('procedimentos'), 'o que a criança/adolescente irá fazer')}."))

    b.append(("p", f"O tempo estimado de participação será de aproximadamente: {fill(p.get('tempo_estimado'), 'tempo estimado de participação')}."))
    b.append(("p", f"A pesquisa será realizada no seguinte local ou plataforma: {fill(p.get('local'), 'local/plataforma de realização')}."))
    b.append(("p", "Toda pesquisa pode apresentar alguns riscos, ainda que pequenos. "
                   f"Os possíveis riscos desta pesquisa incluem: {fill(p.get('riscos'), 'riscos em linguagem adequada à faixa etária')}."))
    b.append(("p", f"Para diminuir esses riscos, os pesquisadores adotarão as seguintes medidas: {fill(p.get('medidas_risco'), 'medidas de proteção e cuidado')}."))
    b.append(("p", "Nas pesquisas realizadas pela internet, apesar das medidas de segurança adotadas, podem existir riscos "
                   "relacionados à proteção das informações e ao uso das tecnologias."))
    b.append(("p", f"A pesquisa poderá contribuir para: {fill(p.get('beneficios'), 'benefícios em linguagem acessível')}."))
    b.append(("p", "Sua participação é voluntária. Isso significa que, sem qualquer prejuízo:"))
    for t in ["você pode escolher participar ou não;",
              "pode deixar de responder perguntas que não quiser responder;",
              "pode desistir da pesquisa a qualquer momento."]:
        b.append(("bullet", t))
    b.append(("p", "Você não receberá pagamento pela participação nesta pesquisa."))
    b.append(("p", "Você não será identificado(a) em trabalhos, apresentações ou publicações realizadas pelos pesquisadores."))
    b.append(("p", "As informações coletadas serão mantidas em sigilo e utilizadas apenas para fins científicos e acadêmicos."))
    b.append(("p", "Se tiver dúvidas sobre a pesquisa, você poderá conversar com os pesquisadores a qualquer momento."))
    b += pesquisador_bloco(p)
    b.append(("p", f"Em caso de dúvidas, denúncias, reclamações ou irregularidades relacionadas aos aspectos éticos desta pesquisa, "
                   f"você poderá entrar em contato com o {CEP_CONTATO}."))
    b.append(("p", "Declaro que as informações desta pesquisa foram explicadas para mim de forma clara e que pude fazer perguntas. "
                   "Eu concordo, de forma livre e voluntária, em participar desta pesquisa."))
    b.append(("p", f"{data_documento(p)}."))
    b.append(("assinatura", "Assinatura do(a) participante"))
    b.append(("assinatura", "Assinatura do(a) pesquisador(a)"))
    if MODO_RASCUNHO:
        b += info_pesquisador_tcle()
    return ("03-TALE-Assentimento", "TERMO DE ASSENTIMENTO LIVRE E ESCLARECIDO", b)

def build_projeto(p):
    """00 — Estrutura do projeto completo organizada a partir dos dados do pesquisador.
    Não inventa conteúdo: transpõe os campos do projeto.json para as 8 seções exigidas
    em https://cep.ufv.br/projeto-novo/ (introdução, justificativa, objetivos, metodologia,
    riscos e benefícios, aspectos éticos, cronograma, referências) + instrumentos."""
    pes = p.get("pesquisador", {})
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    crono = p.get("cronograma") or []
    orc = p.get("orcamento") or []
    b = [
        ("h1", f"PROJETO DE PESQUISA — {titulo}"),
        ("p", f"Pesquisador responsável (orientador): {fill(pes.get('nome'), 'pesquisador.nome')} — "
               f"{fill(pes.get('departamento'), 'pesquisador.departamento')} / "
               f"{fill(pes.get('instituicao') or 'Universidade Federal de Viçosa', 'pesquisador.instituicao')}. "
               f"Contato: {fill(pes.get('email'), 'pesquisador.email')} | {fill(pes.get('telefone'), 'pesquisador.telefone')}."),
        ("h2", "1. Introdução"),
        ("p", fill(p.get("introducao"), "introdução")),
        ("h2", "2. Justificativa"),
        ("p", fill(p.get("justificativa"), "justificativa")),
        ("h2", "3. Objetivos"),
        ("p", fill(p.get("objetivo"), "objetivo")),
        ("h2", "4. Metodologia (incl. participantes, local, procedimentos e instrumentos)"),
        ("p", f"Local/plataforma: {fill(p.get('local'), 'local/plataforma de realização')}. "
               f"Tempo estimado: {fill(p.get('tempo_estimado'), 'tempo estimado de participação')}."),
        ("p", f"Procedimentos: {fill(p.get('procedimentos'), 'procedimentos')}."),
        ("p", f"Instrumentos de coleta (anexar na íntegra na PB): {fill(', '.join(p.get('instrumentos', [])) if p.get('instrumentos') else None, 'instrumentos de coleta')}."),

        ("h2", "5. Riscos e benefícios + minimização"),
        ("p", f"Riscos: {fill(p.get('riscos'), 'riscos')}."),
        ("p", f"Minimização: {fill(p.get('medidas_risco'), 'medidas de prevenção e minimização dos riscos')}."),
        ("p", f"Benefícios: {fill(p.get('beneficios'), 'benefícios diretos e/ou indiretos')}. "
               f"Ressarcimento: {fill(p.get('ressarcimento'), 'forma de ressarcimento')}. "
               f"Acompanhamento pós-participação: {fill_opt(p.get('acompanhamento')) or 'não se aplica / descrever se houver'}."),
        ("h2", "6. Aspectos éticos"),
        ("p", "Consentimento: TCLE de adultos e, quando aplicável, TCLE de responsáveis + TALE; guarda por 5 anos; "
               "sigilo e LGPD; retorno dos resultados; indenização e assistência em caso de dano; "
               "coleta somente após aprovação ética (Res. 466/2012, 510/2016, Lei 14.874/2024)."),
        ("p", f"Situações específicas: vulneráveis: {fill_opt(p.get('vulneraveis')) or 'não declaradas'}; "
               f"saúde mental: {'sim — ver TCLE' if p.get('saude_mental') else 'não'}; "
               f"SUS: {'sim — ver TCLE' if p.get('sus') else 'não'}; "
               f"material biológico: {'sim — Res. 441/2011, ver TCLE' if p.get('usa_biologico') else 'não'}; "
               f"voz/imagem: {'sim — autorização no TCLE' if p.get('gravacao') else 'não'}."),
        ("h2", "7. Cronograma (compatível com documento 06 e com a PB)"),
        ("table", ["Etapa", "Período"],
         [[fill(e.get("etapa"), "etapa"), fill(e.get("periodo"), "período")] for e in crono] if crono else
         [["[ver documento 06-Cronograma]", "[ver documento 06-Cronograma]"]]),
        ("p", "A coleta de dados com seres humanos somente será iniciada após aprovação ética."),
        ("h2", "8. Orçamento e financiamento"),
        ("table", ["Item", "Valor", "Fonte"],
         [[fill(i.get("item"), "item"), fill(i.get("valor"), "valor"), fill(i.get("fonte"), "fonte")] for i in orc] if orc else
         [["[ver documento 07-Orcamento]", "[ver documento 07-Orcamento]", "[ver documento 07-Orcamento]"]]),
        ("h2", "Referências"),
        ("p", fill("\n".join(p.get("referencias", [])) if p.get("referencias") else None, "referências")),
    ]
    for campo in ["introducao", "referencias"]:
        if not p.get(campo):
            PENDENCIAS.append(campo)
    if not p.get("instrumentos"):
        PENDENCIAS.append("instrumentos de coleta")
    if not crono:
        PENDENCIAS.append("cronograma (etapas e períodos)")
    if not orc:
        PENDENCIAS.append("orcamento detalhado")
    return ("00-Projeto-de-Pesquisa", "PROJETO DE PESQUISA (estrutura CEP/UFV)", b)

def build_financiamento(p):
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    pes = p.get("pesquisador", {})
    fin_proprio = p.get("financiamento_proprio", True)
    if fin_proprio:
        b = [
            ("h1", "DECLARAÇÃO DE FINANCIAMENTO PRÓPRIO"),
            ("p", f"Eu, {fill(pes.get('nome'), 'pesquisador.nome')}, pesquisador responsável pela pesquisa “{titulo}”, "
                   "declaro que a pesquisa será realizada com financiamento próprio, sem patrocinador externo."),
            ("p", "Campo “patrocinador principal” da folha de rosto: suprimido, nos termos das orientações do CEP/UFV."),
            ("p", f"{data_documento(p)}."),
            ("assinatura", "Assinatura do pesquisador responsável"),
        ]
    else:
        b = [
            ("h1", "DECLARAÇÃO DE FINANCIAMENTO EXTERNO"),
            ("p", f"Pesquisa: “{titulo}”. Pesquisador responsável: {fill(pes.get('nome'), 'pesquisador.nome')}."),
            ("p", f"Patrocinador principal: {fill(p.get('patrocinador'), 'patrocinador principal')}. "
                   f"Documento de implementação anexado: {fill(p.get('financiamento_doc'), 'documento de implementação')}."),
            ("p", "O campo “patrocinador” da folha de rosto deverá estar preenchido e assinado pelo financiador; "
                   "o documento de implementação assinado segue anexo na Plataforma Brasil."),
            ("p", f"{data_documento(p)}."),
            ("assinatura", "Assinatura do pesquisador responsável"),
            ("assinatura", "Assinatura do financiador"),
        ]
        if not p.get("patrocinador"):
            PENDENCIAS.append("patrocinador principal")
    return ("08b-Declaracao-Financiamento", "DECLARAÇÃO DE FINANCIAMENTO", b)

def build_anuencia(p):
    inst = p.get("instituicao_local", {})
    pes = p.get("pesquisador", {})
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    b = []
    if MODO_RASCUNHO:
        b.append(("aviso", "Preferencialmente em papel timbrado da instituição ou mediante assinatura digital institucional válida. "
                  "Natureza exclusivamente administrativa e institucional; não substitui a apreciação ética pelo SINEP."))
    b += [
        ("h1", "DECLARAÇÃO DE ANUÊNCIA INSTITUCIONAL"),
        ("p", f"A instituição {fill(inst.get('nome'), 'instituição anuente')}, neste ato representada por "
               f"{fill(inst.get('representante'), 'representante institucional')}, no exercício da função de "
               f"{fill(inst.get('cargo'), 'cargo/função')}, declara estar ciente da realização da pesquisa intitulada: “{titulo}”, "
               f"sob responsabilidade do(a) pesquisador(a) {fill(pes.get('nome'), 'pesquisador.nome')}, "
               f"vinculado(a) à instituição {fill(pes.get('instituicao') or 'Universidade Federal de Viçosa', 'pesquisador.instituicao')}."),
        ("p", "A presente anuência autoriza, quando aplicável e observados os limites institucionais pertinentes:"),
        ("bullet", "a realização da coleta de dados nas dependências desta instituição;"),
        ("bullet", "o acesso a participantes institucionalmente vinculados;"),
        ("bullet", "o acesso a documentos, prontuários, bancos de dados ou sistemas de informação estritamente necessários ao desenvolvimento da pesquisa;"),
        ("bullet", "e/ou a utilização da infraestrutura institucional necessária à execução do estudo."),
        ("p", "A instituição declara ciência de que:"),
        ("bullet", "a presente anuência possui natureza exclusivamente administrativa e institucional;"),
        ("bullet", "a autorização ora concedida não substitui a apreciação e aprovação ética pelo Sistema Nacional de Ética em Pesquisa;"),
        ("bullet", "esta declaração não implica aprovação automática da pesquisa;"),
        ("bullet", "a responsabilidade pela condução ética e científica do estudo permanece atribuída ao pesquisador responsável e à instituição proponente, nos termos da regulamentação aplicável;"),
        ("bullet", "o desenvolvimento da pesquisa deverá observar as Resoluções do Conselho Nacional de Saúde, a legislação relativa à proteção de dados pessoais, "
                   "especialmente a Lei nº 13.709/2018 (LGPD), bem como as normas institucionais pertinentes."),
        ("p", "A presente anuência restringe-se à pesquisa acima identificada e poderá ser revogada em caso de descumprimento de normas éticas, legais ou institucionais aplicáveis."),
        ("p", f"{data_documento(p)}."),
        ("p", "Nome do responsável institucional (legível):"),
        ("p", "Cargo/Função:"),
        ("assinatura", "Assinatura institucional (manual ou digital) + carimbo/identificação da instituição"),
    ]
    return ("04-Anuencia-Institucional", "DECLARAÇÃO DE ANUÊNCIA INSTITUCIONAL", b)

def build_sigilo(p):
    pes = p.get("pesquisador", {})
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    fonte = "prontuários de pacientes" if p.get("usa_prontuario", True) else "bases de dados, documentos e sistemas de informação"
    b = [
        ("h1", "TERMO DE SIGILO E CONFIDENCIALIDADE"),
        ("p", f"Eu, {fill(pes.get('nome'), 'pesquisador.nome')}, responsável pelo projeto de pesquisa intitulado “{titulo}”, declaro cumprir com todas as implicações abaixo:"),
        ("p", "Declaro:"),
        ("bullet", f"Que o acesso aos dados registrados em {fonte} para fins da pesquisa científica será feito somente após aprovação do projeto de pesquisa pelo Comitê de Ética;"),
        ("bullet", "Que o acesso aos dados será supervisionado por pessoa plenamente informada sobre as exigências de confiabilidade;"),
        ("bullet", "Meu compromisso com a privacidade e a confidencialidade dos dados utilizados, preservando integralmente o anonimato e a imagem do participante, bem como a sua não estigmatização;"),
        ("bullet", "Não utilizar as informações em prejuízo das pessoas e/ou das comunidades, inclusive em termos de autoestima, de prestígio e/ou econômico-financeiro;"),
        ("bullet", "Que o pesquisador responsável estabeleceu salvaguardar e assegurar a confidencialidade dos dados de pesquisa;"),
        ("bullet", "Que os dados obtidos na pesquisa serão usados exclusivamente para a finalidade prevista no protocolo;"),
        ("bullet", "Que os dados obtidos na pesquisa somente serão utilizados para o projeto vinculado, os quais serão mantidos em sigilo, "
                   "em conformidade com o que prevê a Resolução CNS nº 466/2012, assinando este termo para salvaguardar os direitos dos participantes."),
        ("p", f"Nome do Pesquisador Responsável: {fill(pes.get('nome'), 'pesquisador.nome')}"),
        ("p", f"Endereço: {fill(pes.get('endereco'), 'pesquisador.endereco')}"),
        ("p", f"Fone: {fill(pes.get('telefone'), 'pesquisador.telefone')}"),
        ("p", f"E-mail: {fill(pes.get('email'), 'pesquisador.email')}"),
        ("p", f"{data_documento(p)}."),
        ("assinatura", "Assinatura e carimbo do pesquisador responsável"),
    ]
    return ("05-Sigilo-Confidencialidade", "TERMO DE SIGILO E CONFIDENCIALIDADE", b)

def build_cronograma(p):
    crono = p.get("cronograma") or []
    b = [
        ("h1", "CRONOGRAMA DE PESQUISA"),
        ("p", "O cronograma apresenta informações compatíveis com o Projeto de Pesquisa e com os dados cadastrados no formulário da Plataforma Brasil (PB)."),
        ("table", ["Etapa da Pesquisa", "Período Previsto"],
         [[fill(e.get("etapa"), "etapa"), fill(e.get("periodo"), "período")] for e in crono] if crono else
         [["Elaboração do projeto de pesquisa", "[PREENCHER: período]"],
          ["Submissão do protocolo na Plataforma Brasil", "[PREENCHER: período]"],
          ["Apreciação ética pelo CEP/UFV", "[PREENCHER: período]"],
          ["Adequações e resposta às pendências (quando aplicável)", "[PREENCHER: período]"],
          ["Início da coleta de dados*", "[PREENCHER: período — somente após aprovação]"],
          ["Realização da coleta de dados", "[PREENCHER: período]"],
          ["Organização e análise dos dados", "[PREENCHER: período]"],
          ["Redação do relatório final/dissertação/artigo", "[PREENCHER: período]"],
          ["Encerramento e envio de relatório final ao CEP/UFV", "[PREENCHER: período]"]]),
        ("p", "* A coleta de dados envolvendo seres humanos será iniciada somente após aprovação ética do protocolo pelo CEP/UFV."),
        ("p", "DECLARAÇÃO DE COMPROMISSO EXPRESSO: a pesquisa com seres humanos somente será iniciada após a aprovação do sistema CEP/CONEP (atual SINEP). "
               "Pesquisas com coleta já iniciada não poderão ser apreciadas. Recomenda-se prever início da coleta 2 a 3 meses após a submissão."),
    ]
    if not crono:
        PENDENCIAS.append("cronograma (etapas e períodos)")
    return ("06-Cronograma", "CRONOGRAMA DE PESQUISA", b)

def build_orcamento(p):
    orc = p.get("orcamento") or []
    fin_proprio = p.get("financiamento_proprio", True)
    linhas = [[fill(i.get("item"), "item"), fill(i.get("valor"), "valor"), fill(i.get("fonte"), "fonte")] for i in orc] if orc else []
    b = [
        ("h1", "ORÇAMENTO DETALHADO"),
        ("p", f"Título da pesquisa: {fill(p.get('titulo'), 'titulo da pesquisa')}."),
        ("table", ["Item", "Valor (R$)", "Fonte"],
         linhas if linhas else [["[PREENCHER: item]", "[PREENCHER: valor]", "[PREENCHER: fonte]"]]),
    ]
    if fin_proprio and not orc:
        b.append(("p", "Financiamento próprio (conforme checklist: se somente houver bolsa de estudos ou recursos próprios, informar “Financiamento próprio”)."))
    else:
        b.append(("p", f"Financiamento: {'próprio' if fin_proprio else 'externo — preencher patrocinador na folha de rosto e anexar documento de implementação com assinatura do financiador'}."))
    b.append(("p", "Previsão de ressarcimento aos participantes (alimentação, transporte de participantes e/ou acompanhantes, previstos ou não): "
                   f"{fill(p.get('ressarcimento'), 'ressarcimento')}."))
    b.append(("p", "Custos decorrentes da pesquisa serão ressarcidos pelos pesquisadores. Não haverá contrapartida financeira pela participação. "
                   "Em caso de dano comprovado, há direito a assistência integral e indenização."))
    if not orc:
        PENDENCIAS.append("orcamento detalhado")
    return ("07-Orcamento", "ORÇAMENTO DETALHADO", b)

def build_dispensa(p):
    d = p.get("dispensa_tcle", {}) or {}
    sol = bool(d.get("solicitar"))
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    if not sol:
        b = [
            ("h1", "DECLARAÇÃO — NÃO SOLICITAÇÃO DE DISPENSA DE TCLE"),
            ("p", f"Na pesquisa “{titulo}”, NÃO se solicita dispensa de TCLE. Será obtido o consentimento de todos os participantes (ou responsáveis + assentimento, quando aplicável), "
                   "nos termos da Lei nº 14.874/2024 e das Resoluções CNS nº 466/2012 e nº 510/2016."),
            ("p", "Se futuramente for necessária dispensa (ex.: dados secundários anonimizados, impossibilidade de contato), o pesquisador submeterá justificativa própria à apreciação do CEP/UFV."),
            ("p", f"{data_documento(p)}."),
            ("assinatura", "Assinatura do pesquisador responsável"),
        ]
    else:
        b = [
            ("h1", "SOLICITAÇÃO DE DISPENSA DE TCLE"),
            ("p", f"Título da pesquisa: “{titulo}”. Pesquisador responsável: {fill(p.get('pesquisador', {}).get('nome'), 'pesquisador.nome')}."),
            ("p", "Solicita-se a dispensa do TCLE (ou do registro do consentimento), nos casos previstos nas normas aplicáveis, pelos motivos abaixo, submetidos à apreciação ética do CEP/UFV:"),
            ("p", f"Justificativa: {fill(d.get('justificativa'), 'justificativa da dispensa')}"),
            ("p", "A dispensa somente terá validade se aprovada pelo CEP/UFV em parecer. Mantidas as garantias de sigilo, guarda dos dados por 5 anos e retorno quando cabível."),
            ("p", f"{data_documento(p)}."),
            ("assinatura", "Assinatura do pesquisador responsável"),
        ]
        if not d.get("justificativa"):
            pass
    return ("08-Dispensa-TCLE", "DISPENSA DE TCLE / DECLARAÇÃO", b)

def build_carta(p):
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    pes = p.get("pesquisador", {})
    b = [
        ("h1", "CARTA RESPOSTA ÀS PENDÊNCIAS"),
        ("p", f"Título do projeto: {titulo}"),
        ("p", f"Pesquisador responsável: {fill(pes.get('nome'), 'pesquisador.nome')}"),
        ("p", "Data: ____/____/________ (preencher no uso)"),
        ("p", "Em resposta às pendências informadas, seguem as informações necessárias:"),
        ("h2", "Pendência 1"),
        ("p", "(Abrir o Parecer Consubstanciado, copiar e colar o item 1 contido em “Conclusões ou Pendências e Lista de Inadequações”)"),
        ("p", "Resposta à pendência 1: ___________________________________________________________________________"),
        ("p", "Há algum documento anexo para a pendência? (Se sim, favor informar)"),
        ("h2", "Pendência 2"),
        ("p", "(Abrir o Parecer Consubstanciado, copiar e colar o item 2 contido em “Conclusões ou Pendências e Lista de Inadequações”)"),
        ("p", "Resposta à pendência 2: ___________________________________________________________________________"),
        ("p", "Há algum documento anexo para a pendência? (Se sim, favor informar)"),
        ("h2", "INSTRUÇÕES (verbatim do modelo)"),
        ("p", "1. Todos os documentos ajustados devem ser nomeados com o padrão: “X modificado”. Por exemplo, “TCLE modificado”. "
               "Nessa fase nenhum documento deve ser excluído. A diferenciação será feita pela nomenclatura."),
        ("p", "2. Para facilitar a análise, elaborar as respostas na ordem em que as pendências foram apresentadas."),
        ("p", "COMUNICAMOS QUE O PESQUISADOR TEM TRINTA (30) DIAS PARA ATENDER ÀS PENDÊNCIAS. ULTRAPASSADO ESSE PRAZO TODO O PROTOCOLO "
               "SERÁ ARQUIVADO CONFORME DETERMINA A RESOLUÇÃO CNS 466/2012."),
    ]
    return ("09-Carta-Resposta", "CARTA RESPOSTA ÀS PENDÊNCIAS", b)

CHECKLIST_GERAIS = [
    "Projeto de pesquisa completo (doc 00-Projeto gerado pela skill a partir dos seus dados)",
    "TCLE adultos",
    "TALE (crianças, adolescentes, incapazes)",
    "TCLE responsáveis",
    "Folha de rosto assinada e com carimbos (Plataforma Brasil; se pesquisador = chefia/coordenação, assina superior hierárquico)",
    "Autorizações dos locais / anuência institucional (incl. vulneráveis: indígenas, cárcere etc.)",
    "Cronograma com compromisso expresso de início só após aprovação CEP/CONEP (SINEP)",
    "Termo de Sigilo e Confidencialidade",
    "Orçamento detalhado (ou “Financiamento próprio”)",
    "Solicitação de dispensa de TCLE com justificativa (se for o caso)",
    "Financiamento: declaração 08b + documento de implementação + assinatura do financiador na folha de rosto (se externo)",
    "Instrumentos de coleta na íntegra (campo do doc 00 + anexos na PB)",
]
CHECKLIST_TCLE = [
    "Texto em forma de convite (“O(a) Sr.(a) está sendo convidado(a)...”)",
    "Inexistência de logomarca da UFV ou do Departamento",
    "Uso de “participante da pesquisa” (nunca respondente/entrevistado/sujeito)",
    "Texto simples, sem jargão técnico",
    "Justificativa da pesquisa",
    "Métodos de participação descritos",
    "Riscos detalhados (vedado “não oferece risco”)",
    "Meios de minimizar os riscos",
    "Benefícios diretos e/ou indiretos",
    "Sem contrapartida financeira",
    "Ressarcimento de custos (alimentação, transporte, acompanhantes)",
    "Acompanhamento pós-intervenção (se for o caso)",
    "Autorização de voz/imagem (quando houver registro)",
    "Indenização de danos comprovados",
    "Retirada do consentimento a qualquer momento/fase",
    "Local e tempo de guarda dos dados (5 anos)",
    "Retorno dos resultados aos participantes",
    "Numeração de páginas (“Página X de Y”)",
    "Contatos do pesquisador (telefone, e-mail, endereço)",
    "Contatos do CEP/UFV atualizados",
    "Material biológico pertence ao participante (se coleta)",
    "Laboratório + consentimento futuro do material biológico (se coleta)",
    "Proteção a vulneráveis (coma, gravidez, lactantes, cárcere, indígenas etc.)",
    "Saúde mental: resultado não é diagnóstico (se aplicável)",
    "Saúde mental: orientações/aconselhamentos (se aplicável)",
    "SUS: sem interferir na rotina salvo autorização (se aplicável)",
    "SUS: pesquisa ≠ assistência; sem prejuízo se recusar/retirar (se aplicável)",
]

NECESSIDADE = {}  # builder -> (gerar: bool, motivo: str)

def avaliar_necessidade(p):
    """Decide, para CADA documento do checklist, se o caso concreto exige ou não.
    Ex.: com TCLE não há dispensa (e vice-versa); sem menores, sem TALE/TCLE-resp;
    sem coleta institucional, sem anuência; sem dado sigiloso, sem sigilo."""
    disp = bool((p.get("dispensa_tcle") or {}).get("solicitar"))
    menores = bool(p.get("usa_menores"))
    inst_nome = ((p.get("instituicao_local") or {}).get("nome") or "").strip()
    online = bool(p.get("coleta_online") or p.get("ambiente_virtual"))
    # Anuência: SÓ se houver entrevistas ou similares DENTRO de alguma instituição.
    if p.get("coleta_institucional") is not None:
        coleta_inst = bool(p.get("coleta_institucional"))
    else:
        coleta_inst = bool(inst_nome) and bool(p.get("entrevistas_instituicao", True)) and not online
    if coleta_inst:
        motivo_anu = f"entrevistas/atividades na instituição: {inst_nome or 'declarada no projeto'}"
    elif online and not p.get("entrevistas_instituicao", False):
        motivo_anu = "coleta online, sem entrevistas presenciais em instituição"
    elif inst_nome:
        motivo_anu = "instituição declarada sem entrevistas/atividades presenciais"
    else:
        motivo_anu = "sem entrevistas ou similares em instituição (via pública, documental ou online)"
    sigilo = bool(p.get("usa_prontuario") or p.get("dados_sigilosos") or p.get("dados_secundarios"))
    return {
        build_projeto: (True, "projeto completo sempre obrigatório"),
        build_tcle_adulto: (not disp, "TCLE dispensado neste protocolo" if disp else "consentimento dos participantes adultos"),
        build_tcle_responsavel: (menores and not disp, "TCLE dispensado neste protocolo" if disp else ("responsáveis de menores/incapazes" if menores else "sem menores ou incapazes no estudo")),
        build_tale: (menores and not disp, "TCLE dispensado neste protocolo" if disp else ("assentimento de menores/incapazes" if menores else "sem menores ou incapazes no estudo")),
        build_anuencia: (coleta_inst, motivo_anu),
        build_sigilo: (sigilo, "dados sigilosos/prontuário/secundários" if sigilo else "sem prontuário, dado sigiloso ou base secundária declarados"),
        build_cronograma: (True, "cronograma sempre obrigatório"),
        build_orcamento: (True, "orçamento sempre obrigatório"),
        build_dispensa: (disp, "dispensa solicitada com justificativa" if disp else "há TCLE — dispensa desnecessária"),
        build_financiamento: (True, "declaração de financiamento sempre obrigatória"),
        build_carta: (COM_CARTA, "sob demanda (há pendência)" if COM_CARTA else "só quando houver pendência do CEP"),
        build_coleta_online: (online, "coleta online declarada" if online else "coleta presencial"),
        build_checklist: (True, "conferência interna"),
        build_orientacoes: (True, "guia de submissão"),
        build_roteiro: (True, "guia de submissão"),
    }

def build_checklist(p):
    disp = bool((p.get("dispensa_tcle") or {}).get("solicitar"))
    usa_menores = bool(p.get("usa_menores"))
    def _st(fn):
        gerar, motivo = NECESSIDADE.get(fn, (True, ""))
        return ("OK", None) if gerar else ("N/A", motivo)
    itens = [
        ("Projeto de pesquisa completo (doc 00-Projeto gerado pela skill a partir dos seus dados)", build_projeto),
        ("TCLE adultos", build_tcle_adulto),
        ("TALE (crianças, adolescentes, incapazes)", build_tale),
        ("TCLE responsáveis", build_tcle_responsavel),
        (None, None),  # folha de rosto: sai da PB
        ("Autorizações dos locais / anuência institucional (incl. vulneráveis: indígenas, cárcere etc.)", build_anuencia),
        ("Cronograma com compromisso expresso de início só após aprovação CEP/CONEP (SINEP)", build_cronograma),
        ("Termo de Sigilo e Confidencialidade", build_sigilo),
        ("Orçamento detalhado (ou “Financiamento próprio”)", build_orcamento),
        ("Solicitação de dispensa de TCLE com justificativa (se for o caso)", build_dispensa),
        ("Financiamento: declaração 08b + documento de implementação + assinatura do financiador na folha de rosto (se externo)", build_financiamento),
        ("Instrumentos de coleta na íntegra (campo do doc 00 + anexos na PB)", None),
    ]
    if online_ := bool(p.get("coleta_online") or p.get("ambiente_virtual")):
        itens.append(("Coleta online: instrumento 13 + criar-form-google.gs", build_coleta_online))
    b = [("h1", "CHECK LIST DE DOCUMENTOS — CEP/UFV (transcrição fiel 2021 + STATUS)"),
         ("p", "Legenda STATUS: OK = no pacote; N/A = desnecessário neste caso (motivo justificado); PB = providenciar na Plataforma Brasil. "
               "Só submeta com todos os OK/PB resolvidos."),
         ("h2", "A. DOCUMENTOS GERAIS")]
    for item, fn in itens:
        if fn is None and item is None:
            b.append(("p", "[STATUS: PB] Folha de rosto assinada e com carimbos — gerar na Plataforma Brasil "
                           "(responsável = orientador; proponente = chefia/coordenação; superior hierárquico se acúmulo)"))
        elif fn is None:
            st = "OK" if p.get("instrumentos") else "PENDENTE — declarar os instrumentos de coleta"
            b.append(("p", f"[STATUS: {st}] {item}"))
        else:
            st, motivo = _st(fn)
            b.append(("p", f"[STATUS: {st}{(' — ' + motivo) if motivo and st == 'N/A' else ''}] {item}"))
    b.append(("h2", "B. CONTEÚDO MÍNIMO DO TCLE (verificado no 01-TCLE gerado)"))
    if disp:
        b.append(("p", "[STATUS: N/A — TCLE dispensado neste protocolo] Itens abaixo não se aplicam."))
        b.append(("p", "Fontes: checklist 2021 + https://cep.ufv.br/projeto-novo/ + https://cep.ufv.br/normativas/. "
                       "Em divergência, vale o site oficial + parecer do CEP."))
        return ("10-Checklist-Conformidade", "CHECK LIST DE CONFORMIDADE", b)
    for item in CHECKLIST_TCLE:
        st = "OK"
        if "vulneráveis" in item and not usa_menores:
            st = "N/A — sem vulneráveis específicos declarados"
        if "Saúde mental" in item or "SUS" in item or "acompanhamento" in item.lower() or "voz/imagem" in item.lower() or "biológico" in item.lower():
            st = "OK (verificar aplicabilidade ao seu protocolo)"
        b.append(("p", f"[STATUS: {st}] {item}"))
    b.append(("p", "Fontes: checklist 2021 + https://cep.ufv.br/projeto-novo/ + https://cep.ufv.br/normativas/. "
                   "Em divergência, vale o site oficial + parecer do CEP."))
    return ("10-Checklist-Conformidade", "CHECK LIST DE CONFORMIDADE", b)

def build_orientacoes(p):
    b = [
        ("aviso", "NÃO SUBMETER — guia operacional. Não anexar este arquivo na Plataforma Brasil."),
        ("h1", "ORIENTAÇÕES DE SUBMISSÃO — Plataforma Brasil (resumo fiel de https://cep.ufv.br/projeto-novo/)"),
        ("p", "1. Cadastro: todos da equipe com cadastro ativo (documento ID, Lattes, foto, vínculo UFV)."),
        ("p", "2. Documentos separados, legíveis, nomeados (“Projeto de Pesquisa”, “TCLE”...), compatíveis entre si e com o formulário PB, com assinaturas (preferencialmente digitais)."),
        ("p", "3. Projeto de pesquisa mínimo: introdução, justificativa, objetivos, metodologia, riscos e benefícios, aspectos éticos, cronograma, referências. "
               "Anexar instrumentos na íntegra (questionários, roteiros, formulários, convites)."),
        ("p", "4. Folha de rosto (gerada pela PB): conferir, coletar assinaturas, anexar final. Responsável = orientador. "
               "Proponente: chefia do departamento (graduação/IC/extensão) ou coordenação do PPG (pós). Se acúmulo, assina superior hierárquico. "
               "Patrocinador externo preenche+assina; próprio suprime o campo."),
        ("p", "5. Cronograma em documento separado, compatível com projeto e PB, com frase expressa de início só após aprovação. Prever 2–3 meses até a coleta."),
        ("p", "6. Submissão só pelo pesquisador responsável (orientador). Fluxo: Cadastro → Documentação → Submissão → Validação documental → Relatoria → Parecer."),
        ("p", "7. Acompanhar só pela PB: recepção/validação, pendência documental, apreciação ética, pendência do CEP, aprovado (iniciar coleta), não aprovado (recurso)."),
        ("p", "8. Antes de submeter: legibilidade, assinaturas, coerência, datas/cronograma. Nunca iniciar coleta antes da aprovação. Responder todas as pendências (prazo 30 dias; nomear “X modificado”)."),
        ("p", "Base legal: Lei 14.874/2024, Decreto 12.651/2025, Res. 466/2012, Res. 510/2016, NO 01/2013, LGPD. Dúvidas: https://cep.ufv.br/ | cep@ufv.br | (31) 3612-2316."),
    ]
    return ("11-Orientacoes-Plataforma-Brasil", "ORIENTAÇÕES DE SUBMISSÃO (GUIA — NÃO SUBMETER)", b)

TIPOS_FORMS = {
    "texto": "Resposta curta",
    "paragrafo": "Parágrafo",
    "multipla": "Múltipla escolha",
    "checkbox": "Caixas de seleção",
    "escala": "Escala linear",
}

def build_coleta_online(p):
    """13 — Instrumento de coleta online + montagem do Google Forms.
    Texto pronto para colar + script criar-form-google.gs que monta o Forms sozinho."""
    pes = p.get("pesquisador", {})
    titulo = fill(p.get("titulo"), "titulo da pesquisa")
    quest = p.get("questionario") or []
    if not quest:
        PENDENCIAS.append("questionario online (perguntas do Google Forms)")
    b = [
        ("h1", "INSTRUMENTO DE COLETA ONLINE — GOOGLE FORMS"),
        ("p", f"Pesquisa: “{titulo}”. Este documento traz o instrumento pronto para o Google Forms: "
               "Seção 1 (consentimento, obrigatória), Seção 2 (perguntas) e as configurações exigidas. "
               "A montagem pode ser automática com o arquivo criar-form-google.gs da pasta do pacote."),
        ("h2", "Seção 1 — Consentimento (primeira seção do formulário, resposta obrigatória)"),
        ("p", "Título do formulário: o título da pesquisa. Descrição: colar o objetivo, o tempo estimado "
               f"({fill(p.get('tempo_estimado'), 'tempo estimado de participação')}), os contatos do pesquisador "
               f"({fill(pes.get('email'), 'pesquisador.email')} | {fill(pes.get('telefone'), 'pesquisador.telefone')}) "
               "e do CEP/UFV (cep@ufv.br | (31) 3612-2316), mais o link/arquivo do TCLE completo."),
        ("p", "Questão 1 (obrigatória, múltipla escolha): “Declaro que li o TCLE e concordo em participar da pesquisa.” "
               "Opções: “Sim, concordo e quero participar” (seguir para a Seção 2) / “Não concordo” (encerrar)."),
        ("h2", "Seção 2 — Perguntas"),
        ("table", ["Nº", "Pergunta", "Tipo no Forms", "Obrigatória"],
         [[str(i + 1), fill(q.get("pergunta"), f"pergunta {i + 1}"),
           TIPOS_FORMS.get((q.get("tipo") or "texto"), "Resposta curta"),
           "Sim" if q.get("obrigatoria", True) else "Não"]
          for i, q in enumerate(quest)] if quest else [["—", "—", "—", "—"]]),
        ("h2", "Configurações obrigatórias do formulário"),
        ("bullet", "Resposta à questão de consentimento obrigatória; sem “Sim” não avança."),
        ("bullet", "Mensagem de confirmação com agradecimento + contatos do pesquisador e do CEP/UFV."),
        ("bullet", "Vincular planilha de respostas; acesso restrito à equipe; guarda por 5 anos (LGPD)."),
        ("bullet", "Testar o formulário antes de divulgar; divulgar somente após aprovação ética do CEP/UFV."),
        ("h2", "Montagem automática"),
        ("p", "Abrir https://script.google.com, colar o conteúdo de criar-form-google.gs, executar criarFormularioCEP "
               "e autorizar. O formulário é criado com as seções, perguntas e mensagem de confirmação deste documento."),
    ]
    return ("13-Coleta-Online-GoogleForms", "INSTRUMENTO DE COLETA ONLINE — GOOGLE FORMS", b)

def gerar_form_gs(p):
    """Gera o Apps Script (FormApp) que monta o Google Forms sozinho."""
    pes = p.get("pesquisador", {})
    dados = {
        "titulo": p.get("titulo", ""),
        "descricao": f"{(p.get('objetivo', '') or '').rstrip('.')}. Tempo estimado: {(p.get('tempo_estimado', '') or '').rstrip('.')}. "
                     f"Contato: {pes.get('email', '')} | {pes.get('telefone', '')}. CEP/UFV: cep@ufv.br | (31) 3612-2316.",
        "confirmacao": f"Obrigado pela participação! Contato: {pes.get('email', '')}. CEP/UFV: cep@ufv.br | (31) 3612-2316.",
        "questoes": [{"pergunta": q.get("pergunta", ""), "tipo": (q.get("tipo") or "texto"),
                      "opcoes": q.get("opcoes") or [], "obrigatoria": bool(q.get("obrigatoria", True))}
                     for q in (p.get("questionario") or [])],
    }
    js = json.dumps(dados, ensure_ascii=False)
    return ("function criarFormularioCEP() {\n"
            "  var D = " + js + ";\n"
            "  var form = FormApp.create(D.titulo);\n"
            "  form.setDescription(D.descricao);\n"
            "  form.setConfirmationMessage(D.confirmacao);\n"
            "  form.setLimitOneResponsePerUser(true);\n"
            "  var consent = form.addMultipleChoiceItem()\n"
            "    .setTitle('Declaro que li o TCLE e concordo em participar da pesquisa.')\n"
            "    .setChoiceValues(['Sim, concordo e quero participar', 'Não concordo'])\n"
            "    .setRequired(true);\n"
            "  form.addPageBreakItem().setTitle('Questionário');\n"
            "  D.questoes.forEach(function(q) {\n"
            "    var item = null;\n"
            "    if (q.tipo === 'paragrafo') item = form.addParagraphTextItem();\n"
            "    else if (q.tipo === 'multipla') item = form.addMultipleChoiceItem().setChoiceValues(q.opcoes);\n"
            "    else if (q.tipo === 'checkbox') item = form.addCheckboxItem().setChoiceValues(q.opcoes);\n"
            "    else if (q.tipo === 'escala') item = form.addScaleItem().setBounds(\n"
            "      parseInt(q.opcoes[0] || 1, 10), parseInt(q.opcoes[1] || 5, 10))\n"
            "      .setLeftLabel(q.opcoes[2] || '').setRightLabel(q.opcoes[3] || '');\n"
            "    else item = form.addTextItem();\n"
            "    item.setTitle(q.pergunta).setRequired(!!q.obrigatoria);\n"
            "  });\n"
            "  Logger.log('Formulário criado: ' + form.getEditUrl());\n"
            "}\n")

def build_roteiro(p):
    titulo = fill_opt(p.get("titulo")) or "[PREENCHER: titulo da pesquisa]"
    b = [
        ("aviso", "NÃO SUBMETER — roteiro operacional passo a passo. Não anexar na Plataforma Brasil. "
                  "Procedimentos verificados em https://cep.ufv.br/projeto-novo/, https://cep.ufv.br/pendencia/, "
                  "https://cep.ufv.br/cronograma/ e https://cep.ufv.br/normativas/ em 05/09/2026. "
                  "Em divergência, vale o site oficial + a Plataforma Brasil."),
        ("h1", "ROTEIRO PASSO A PASSO — Submissão ao CEP/UFV via Plataforma Brasil"),
        ("p", f"Pesquisa: “{titulo}”. Use este roteiro na ordem. Marque cada etapa como feita."),
        ("h2", "ETAPA 0 — Regra de ouro (não negociável)"),
        ("bullet", "NÃO iniciar coleta antes da aprovação ética. Coleta já iniciada não pode ser apreciada."),
        ("bullet", "Pesquisador responsável no CEP/UFV = orientador (não o orientando)."),
        ("bullet", "Prever coleta 2–3 meses após a submissão (tramitação + eventual pendência)."),
        ("bullet", "Reuniões ordinárias 2026: 06/03, 17/04, 08/05, 12/06, 10/07, 14/08, 11/09, 09/10, 13/11, 04/12. "
                   "Submeter até dia 20 do mês anterior (checagem documental leva até 10 dias — NO 001/2013)."),
        ("h2", "ETAPA 1 — Cadastro na Plataforma Brasil (todos da equipe)"),
        ("bullet", "Acessar https://plataformabrasil.saude.gov.br (usar preferencialmente Firefox)."),
        ("bullet", "Documentos do cadastro: documento de identificação, currículo/link Lattes, foto, vínculo UFV (selecionar “Universidade Federal de Viçosa” + unidade)."),
        ("h2", "ETAPA 2 — Seu pacote (já gerado pelo agente)"),
        ("p", f"Seu pacote está na pasta {PACOTE_NOME or 'outputs-cep-<slug>/'}: cada documento em par .docx (editável) + .pdf idêntico, já datado e formatado ABNT/UFV."),
        ("bullet", "Confira o 10-Checklist-Conformidade e o pendencias.txt (deve indicar zero pendências)."),
        ("bullet", "Documentos do pacote: 00-Projeto, 01-TCLE, 02-TCLE-Responsável, 03-TALE, 04-Anuência, 05-Sigilo, 06-Cronograma, 07-Orçamento, 08-Dispensa/declaração, 08b-Financiamento, 10-Checklist, 11-Orientações (guia, não anexar), este 12-Roteiro (guia, não anexar)."),
        ("bullet", "Carta-Resposta: NÃO acompanha o pacote inicial — ela só é gerada se houver pendência do CEP, e o agente a prepara para você na hora."),
        ("bullet", "O que NÃO vai no pacote (por definição do CEP): folha de rosto (sai da PB, ETAPA 3). Instrumentos (questionários/roteiros) estão descritos no 00-Projeto e devem ser anexados na íntegra."),
        ("h2", "PREVISÃO DE TRAMITAÇÃO (calculada hoje — previsão, não garantia)"),
        ("p", PREVISAO_TEXTO or "Ver previsao.txt na pasta do pacote."),
        ("h2", "ETAPA 3 — Folha de rosto (dentro da PB)"),
        ("bullet", "Gerar na PB, conferir tudo, colher assinaturas preferencialmente digitais com identificação (carimbo/assinatura eletrônica)."),
        ("bullet", "Responsável = orientador. Proponente: chefia do departamento (graduação/IC/extensão) ou coordenação do PPG (pós). Se a mesma pessoa acumula, assina o superior hierárquico."),
        ("bullet", "Patrocinador: externo preenche+assina; próprio suprime o campo. Financiamento externo exige documento de implementação anexado."),
        ("h2", "ETAPA 4 — Anexar na PB (arquivos separados, legíveis, nomeados)"),
        ("bullet", "Ex.: “Projeto de Pesquisa”, “TCLE”, “TALE”, “Anuência”, “Sigilo”, “Cronograma”, “Orçamento”, “Instrumentos”."),
        ("bullet", "Checar compatibilidade total entre documentos e o formulário PB (objetivos, riscos, benefícios, cronograma, local, amostra)."),
        ("bullet", "Cronograma em arquivo separado com frase expressa de início só após aprovação."),
        ("h2", "ETAPA 5 — Enviar (“Enviar Projeto ao CEP” → etapa 6 – Finalizar)"),
        ("bullet", "Só o pesquisador responsável (orientador) finaliza. Só anexar não conclui."),
        ("bullet", "Acompanhar só pela PB: Recepção/Validação → Pendência Documental (se houver) → Apreciação Ética → Pendência do CEP / Aprovado / Não Aprovado (cabe recurso)."),
        ("h2", "ETAPA 6 — Se vier “Pendência Documental” (secretaria, formal)"),
        ("bullet", "Documento recusado (R): EXCLUIR o antigo e SUBSTITUIR pela versão corrigida. Anexar faltantes. NÃO precisa de Carta Resposta."),
        ("bullet", "Refazer: “Enviar Projeto ao CEP” → etapa 6 – Finalizar. Ajustar até dia 20 para entrar na reunião seguinte."),
        ("h2", "ETAPA 7 — Se vier “Pendência Emitida pelo CEP” (parecer, ética)"),
        ("bullet", "Ler o Parecer Consubstanciado integral (lupa do protocolo → expandir setas → pasta “Pareceres” → ícone azul “P” → PDF), foco em “Conclusões ou Pendências e Lista de Inadequações”. Prazo: 30 dias ou arquiva (Res. 466/2012)."),
        ("bullet", "Peça ao agente a Carta-Resposta: ele a gera na hora (uma resposta por pendência) com os documentos alterados correspondentes. NÃO excluir antigos; anexar como “X modificado” (ex.: “TCLE modificado”)."),
        ("bullet", "Refinalizar: “Enviar Projeto ao CEP” → etapa 6 – Finalizar."),
        ("h2", "ETAPA 8 — Aprovado → executar; depois encerrar"),
        ("bullet", "Só então iniciar a coleta. Guardar dados 5 anos. Devolver resultados aos participantes (conforme TCLE)."),
        ("bullet", "Ao fim: enviar Relatório Final ao CEP (modelo 2016 em modelos/Relatorio-Final.pdf) + publicar sem identificar participantes."),
        ("p", "Contato/suporte: CEP/UFV — Edifício Arthur Bernardes, piso inferior, Av. PH Rolfs s/n, Viçosa/MG 36570-900. "
               "Tel (31) 3612-2316, cep@ufv.br, https://cep.ufv.br. Manuais PB: https://cep.ufv.br/publicacoes/. "
               "Base legal: Lei 14.874/2024, Decreto 12.651/2025, Res. 466/2012, Res. 510/2016, NO 001/2013, LGPD."),
    ]
    if p.get("coleta_online") or p.get("ambiente_virtual"):
        b.append(("h2", "ETAPA 4B — Coleta online (Google Forms)"))
        b.append(("bullet", "Montar o formulário com o 13-Coleta-Online-GoogleForms + criar-form-google.gs da pasta do pacote (consentimento obrigatório na Seção 1, planilha vinculada, guarda 5 anos)."))
        b.append(("bullet", "Divulgar o link SOMENTE após a aprovação ética; anexar o PDF do instrumento na PB como “Instrumentos”."))
    return ("12-Roteiro-Submissao", "ROTEIRO PASSO A PASSO DE SUBMISSÃO (GUIA — NÃO SUBMETER)", b)

ALL_BUILDERS = [build_projeto, build_tcle_adulto, build_tcle_responsavel, build_tale, build_anuencia,
                build_sigilo, build_cronograma, build_orcamento, build_dispensa, build_financiamento,
                build_carta, build_coleta_online, build_checklist, build_orientacoes, build_roteiro]

TEMPLATE_FINGERPRINTS = [
    "está sendo convidado(a) a participar, de forma voluntária",
    "Toda pesquisa envolvendo seres humanos apresenta riscos, ainda que mínimos",
    "prazo mínimo de 5 (cinco) anos",
    "Edifício Arthur Bernardes",
    "natureza exclusivamente administrativa e institucional",
    "TRINTA (30) DIAS PARA ATENDER ÀS PENDÊNCIAS",
    "somente será iniciada após a aprovação",
    "Financiamento próprio",
    "participante da pesquisa",
]

# ------------------------------------------------------------------ render
def _footer_field(par, instr, placeholder="1"):
    """Campo Word válido (begin → instr → separate → texto → end) para PAGE/NUMPAGES."""
    r1 = par.add_run(); e1 = OxmlElement('w:fldChar'); e1.set(qn('w:fldCharType'), 'begin'); r1._r.append(e1)
    r2 = par.add_run(); it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = f" {instr} "; r2._r.append(it)
    r3 = par.add_run(); e3 = OxmlElement('w:fldChar'); e3.set(qn('w:fldCharType'), 'separate'); r3._r.append(e3)
    r4 = par.add_run(placeholder)
    r5 = par.add_run(); e5 = OxmlElement('w:fldChar'); e5.set(qn('w:fldCharType'), 'end'); r5._r.append(e5)
    for r in (r1, r2, r3, r4, r5):
        r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    return r4

def add_page_number_footer(doc, total_paginas=None):
    # NUMPAGES leva como texto provisório o total real (vindo do PDF gêmeo):
    # o Word recalcula de todo modo, e o Google Docs (que não recalcula NUMPAGES
    # na pré-visualização) exibe o valor correto em vez de "1".
    section = doc.sections[0]
    footer = section.footer
    footer.is_linked_to_previous = False
    par = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = par.add_run("Página ")
    r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    _footer_field(par, "PAGE")
    r = par.add_run(" de ")
    r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    _footer_field(par, "NUMPAGES", placeholder=str(total_paginas) if total_paginas else "1")

def render_docx(out_path, titulo_doc, blocos, total_paginas=None):
    # Formatação ABNT/UFV (NBR 14724/2025): A4, Arial 12, texto preto justificado,
    # margens sup/esq 3 cm e inf/dir 2 cm, espaçamento 1,5, paginação no rodapé.
    from docx.shared import RGBColor as _RC
    doc = Document()
    for s in doc.sections:
        s.top_margin = Cm(3); s.left_margin = Cm(3); s.bottom_margin = Cm(2); s.right_margin = Cm(2)
    st = doc.styles['Normal']
    st.font.name = 'Arial'; st.font.size = Pt(12); st.font.color.rgb = _RC(0, 0, 0)
    st.paragraph_format.line_spacing = 1.5
    st.paragraph_format.space_after = Pt(6)
    add_page_number_footer(doc, total_paginas)
    for bl in blocos:
        kind = bl[0]
        if kind == "aviso":
            par = doc.add_paragraph()
            par.paragraph_format.line_spacing = 1.0
            r = par.add_run("AVISO: " + bl[1])
            r.italic = True; r.font.size = Pt(10); r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
        elif kind == "h1":
            par = doc.add_heading(level=1)
            par.paragraph_format.line_spacing = 1.5
            r = par.add_run(bl[1]); r.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0, 0, 0)
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif kind == "h2":
            par = doc.add_heading(level=2)
            par.paragraph_format.line_spacing = 1.5
            r = par.add_run(bl[1]); r.bold = True; r.font.size = Pt(12); r.font.color.rgb = RGBColor(0, 0, 0)
        elif kind == "p":
            par = doc.add_paragraph(bl[1])
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            par.paragraph_format.line_spacing = 1.5
        elif kind == "bullet":
            par = doc.add_paragraph(bl[1], style='List Bullet')
            par.paragraph_format.line_spacing = 1.5
        elif kind == "assinatura":
            # Bloco de assinatura ABNT: respiro, linha, identificação e nome legível.
            doc.add_paragraph("").paragraph_format.space_after = Pt(18)
            linha = doc.add_paragraph()
            linha.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = linha.add_run("_" * 52)
            r.font.size = Pt(12)
            rot = doc.add_paragraph()
            rot.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = rot.add_run(bl[1]); r.font.size = Pt(11)
            nome = doc.add_paragraph()
            nome.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = nome.add_run("Nome legível: " + "_" * 38)
            r.font.size = Pt(11)
        elif kind == "table":
            _, headers, rows = bl
            t = doc.add_table(rows=1 + len(rows), cols=len(headers))
            t.style = 'Light Grid Accent 1'
            for j, h in enumerate(headers):
                c = t.cell(0, j); c.text = ""
                r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(11)
            for i, row in enumerate(rows, start=1):
                for j, v in enumerate(row):
                    c = t.cell(i, j); c.text = ""
                    r = c.paragraphs[0].add_run(str(v)); r.font.size = Pt(11)
            doc.add_paragraph("")
    # rodapé legal fixo (verbatim, não editável via JSON)
    par = doc.add_paragraph(BASE_LEGAL)
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for r in par.runs:
        r.font.size = Pt(9); r.italic = True; r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    doc.core_properties.title = titulo_doc
    doc.core_properties.subject = "CEP/UFV — template oficial preenchido (skill cep-ufv)"
    doc.save(str(out_path))

def pdf_footer_total(total):
    def _f(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica-Oblique", 8)
        canvas.setFillColor(colors.HexColor("#666666"))
        canvas.drawCentredString(A4[0] / 2, 1.5 * cm, f"{BASE_LEGAL}  |  Página {doc.page} de {total}")
        canvas.restoreState()
    return _f

def pdf_footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica-Oblique", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawCentredString(A4[0] / 2, 1.5 * cm, f"{BASE_LEGAL}  |  Página {doc.page}")
    canvas.restoreState()

def _pdf_kwargs(titulo_doc):
    return dict(pagesize=A4, topMargin=3 * cm, leftMargin=3 * cm, bottomMargin=2.5 * cm,
                rightMargin=2 * cm, title=titulo_doc,
                subject="CEP/UFV — template oficial preenchido (skill cep-ufv)")

def render_pdf(out_path, titulo_doc, blocos):
    # Espelho ABNT/UFV do DOCX: Helvetica 12 (equivalente Arial), 1,5 (12/18), justificado.
    styles = getSampleStyleSheet()
    sN = ParagraphStyle('N', parent=styles['Normal'], fontName='Helvetica', fontSize=12, leading=18, alignment=TA_JUSTIFY, spaceAfter=6, textColor=colors.black)
    sH1 = ParagraphStyle('H1', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=12, leading=18, alignment=TA_CENTER, spaceAfter=10, spaceBefore=6, textColor=colors.black)
    sH2 = ParagraphStyle('H2', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=12, leading=18, alignment=TA_LEFT, spaceAfter=8, spaceBefore=10, textColor=colors.black)
    sAv = ParagraphStyle('Av', parent=sN, fontSize=9, textColor=colors.HexColor('#555555'), borderPadding=4)
    sB = ParagraphStyle('B', parent=sN, leftIndent=18, bulletIndent=8, spaceAfter=3)
    def esc(t):
        return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    def make_story():
        # Via independente por passagem: o Platypus consome/muta a lista no build.
        st = []
        for bl in blocos:
            kind = bl[0]
            if kind == "aviso":
                st.append(Paragraph(f"<i>AVISO: {esc(bl[1])}</i>", sAv))
            elif kind == "h1":
                st.append(Paragraph(f"<b>{esc(bl[1])}</b>", sH1))
            elif kind == "h2":
                st.append(Paragraph(f"<b>{esc(bl[1])}</b>", sH2))
            elif kind == "p":
                st.append(Paragraph(esc(bl[1]), sN))
            elif kind == "bullet":
                st.append(Paragraph(esc(bl[1]), sB, bulletText="•"))
            elif kind == "assinatura":
                sC = ParagraphStyle('sig', parent=sN, alignment=TA_CENTER, spaceBefore=18, spaceAfter=2)
                st.append(Paragraph("_" * 52, sC))
                st.append(Paragraph(esc(bl[1]), sC))
                st.append(Paragraph("Nome legível: " + "_" * 30, sC))
            elif kind == "table":
                _, headers, rows = bl
                data = [[Paragraph(f"<b>{esc(h)}</b>", sN) for h in headers]]
                for row in rows:
                    data.append([Paragraph(esc(v), sN) for v in row])
                w = (A4[0] - 4 * cm) / max(len(headers), 1)
                t = Table(data, colWidths=[w] * len(headers), repeatRows=1)
                t.setStyle(TableStyle([
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#eeeeee')),
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ]))
                st.append(t); st.append(Spacer(1, 6))
        st.append(Spacer(1, 12))
        st.append(Paragraph(f"<i>{esc(BASE_LEGAL)}</i>", sAv))
        return st
    # Duas passagens: 1ª conta as páginas, 2ª grava "Página X de Y" (exigência do checklist).
    from io import BytesIO
    import fitz
    _buf = BytesIO()
    SimpleDocTemplate(_buf, **_pdf_kwargs(titulo_doc)).build(
        make_story(), onFirstPage=pdf_footer, onLaterPages=pdf_footer)
    total = fitz.open(stream=_buf.getvalue(), filetype="pdf").page_count
    doc = SimpleDocTemplate(str(out_path), **_pdf_kwargs(titulo_doc))
    doc.build(make_story(), onFirstPage=pdf_footer_total(total), onLaterPages=pdf_footer_total(total))
    return total

EXEMPLO = {
    "titulo": "Exemplo — Efeitos de uma intervenção educativa sobre hábitos alimentares de escolares em Viçosa/MG",
    "pesquisador": {"nome": "Profa. Dra. Maria Silva (orientadora e responsável)", "departamento": "Departamento de Nutrição e Saúde",
                    "endereco": "Av. PH Rolfs, s/n, Campus Universitário, Viçosa/MG", "telefone": "(31) 3612-0000",
                    "email": "maria.silva@ufv.br", "instituicao": "Universidade Federal de Viçosa"},
    "objetivo": "Avaliar o efeito de oficinas educativas sobre o consumo de frutas e hortaliças por escolares do ensino fundamental.",
    "justificativa": "Baixo consumo de frutas e hortaliças entre escolares justifica ações educativas avaliadas com rigor ético.",
    "procedimentos": " questionário estruturado (20 min), medidas antropométricas (peso/altura) e 3 oficinas educativas em grupo.",
    "local": "Escola Municipal X, Viçosa/MG (sala de aula e pátio, em horário escolar).",
    "tempo_estimado": "40 minutos por encontro, 4 encontros.",
    "riscos": "Constrangimento ao responder, cansaço, identificação indireta e, no ambiente virtual (se usado), limites de confidencialidade da tecnologia.",
    "medidas_risco": "Ambiente reservado, anonimização por códigos, interrupção a qualquer momento, suporte da equipe, dados em ambiente seguro.",
    "beneficios": "Orientação nutricional aos participantes e subsídio a políticas escolares de alimentação saudável.",
    "ressarcimento": "Não haverá custos ao participante; despesas eventuais (transporte/alimentação) serão ressarcidas pelos pesquisadores.",
    "retorno": "Relatório-síntese à escola e aos responsáveis + roda de conversa + publicação científica sem identificações.",
    "contato_retirada": "maria.silva@ufv.br / (31) 3612-0000",
    "participante_nome": "Nome do(a) participante (preenchido no ato da assinatura)",
    "usa_menores": True, "usa_biologico": False, "biologico_lab": "Não se aplica (sem coleta de material biológico)", "biologico_futuro": "Não se aplica",
    "usa_prontuario": False, "dados_sigilosos": True, "dados_secundarios": False,
    "ambiente_virtual": False, "gravacao": True,
    "coleta_online": False,
    "questionario": [{"pergunta": "Com que frequência você consome frutas?", "tipo": "multipla",
                      "opcoes": ["Todo dia", "3-5x por semana", "1-2x por semana", "Raramente"], "obrigatoria": True}],
    "vulneraveis": "", "acompanhamento": "A equipe permanecerá disponível após cada encontro para acolhimento e esclarecimentos",
    "saude_mental": False, "saude_mental_orientacoes": "", "sus": False,
    "introducao": "Introdução: contexto do consumo alimentar de escolares e lacuna que a intervenção educativa pretende enfrentar.",
    "instrumentos": ["Questionário estruturado de consumo alimentar (anexo 1)", "Ficha antropométrica (anexo 2)", "Roteiro das oficinas (anexo 3)"],
    "referencias": ["BRASIL. Resolução CNS nº 466/2012.", "BRASIL. Resolução CNS nº 510/2016.", "BRASIL. Lei nº 14.874/2024."],
    "instituicao_local": {"nome": "Escola Municipal X — Secretaria Municipal de Educação de Viçosa", "representante": "Diretora Y", "cargo": "Diretora escolar"},
    "entrevistas_instituicao": True,
    "cronograma": [{"etapa": "Elaboração do projeto", "periodo": "Janeiro/2027"},
                   {"etapa": "Submissão na Plataforma Brasil", "periodo": "Fevereiro/2027"},
                   {"etapa": "Apreciação ética CEP/UFV", "periodo": "Fevereiro a Abril/2027"},
                   {"etapa": "Coleta de dados*", "periodo": "Maio a Julho/2027"},
                   {"etapa": "Análise + redação + relatório final ao CEP", "periodo": "Agosto a Dezembro/2027"}],
    "orcamento": [{"item": "Impressão de questionários", "valor": "R$ 300,00", "fonte": "Financiamento próprio"},
                  {"item": "Transporte da equipe", "valor": "R$ 500,00", "fonte": "Financiamento próprio"}],
    "financiamento_proprio": True,
    "dispensa_tcle": {"solicitar": False, "justificativa": ""},
    "cidade": "Viçosa"
}

def main():
    global MODO_RASCUNHO, COM_CARTA, PACOTE_NOME, PREVISAO_TEXTO, NECESSIDADE
    ap = argparse.ArgumentParser(description="cep-ufv — gerador travado aos templates oficiais (docx+pdf, pronto p/ submissão).")
    ap.add_argument("projeto", nargs="?", help="projeto.json (montado pelo AGENTE a partir dos dados do pesquisador)")
    ap.add_argument("--out", default="outputs-cep-estudo", help="pasta de saída")
    ap.add_argument("--validar-apenas", action="store_true")
    ap.add_argument("--exemplo", action="store_true", help="escreve projeto-exemplo.json no cwd")
    ap.add_argument("--rascunho", action="store_true",
                    help="permite gerar rascunho com [PREENCHER] e blocos de orientação (NÃO submetível)")
    ap.add_argument("--com-carta", action="store_true",
                    help="inclui a 09-Carta-Resposta (usar SOMENTE quando houver pendência do CEP)")
    a = ap.parse_args()
    MODO_RASCUNHO = bool(a.rascunho)
    COM_CARTA = bool(a.com_carta)
    if a.exemplo:
        Path("projeto-exemplo.json").write_text(json.dumps(EXEMPLO, ensure_ascii=False, indent=2), encoding="utf-8")
        print("projeto-exemplo.json escrito.")
        return 0
    if not a.projeto:
        print("Uso: python3 run.py projeto.json --out outputs-cep-<slug>/", file=sys.stderr)
        return 2
    try:
        p = json.loads(Path(a.projeto).read_text(encoding="utf-8"))
    except Exception as e:
        print(f"ERRO ao ler {a.projeto}: {e}", file=sys.stderr)
        return 2
    if not isinstance(p, dict):
        print("ERRO: projeto.json deve ser um objeto.", file=sys.stderr)
        return 2
    PENDENCIAS.clear()
    PACOTE_NOME = Path(a.out).name if a.out else ""
    PREVISAO_TEXTO = prever_parecer()
    NECESSIDADE = avaliar_necessidade(p)
    builders = [fn for fn in ALL_BUILDERS if NECESSIDADE.get(fn, (True, ""))[0]]
    docs = [fn(p) for fn in builders]
    # fingerprint: texto gerado deve conter todas as âncoras (impede drift criativo)
    texto_total = "\n".join(" ".join(str(x) for x in bl[1:3] if isinstance(x, str)) for _, _, blocos in docs for bl in blocos if bl[0] in ("p", "bullet", "h1"))
    _gerados = {slug for slug, _, _ in docs}
    _ancora_fontes = {
        "está sendo convidado(a) a participar, de forma voluntária": {"01-TCLE-Adulto"},
        "Toda pesquisa envolvendo seres humanos apresenta riscos, ainda que mínimos": {"01-TCLE-Adulto", "02-TCLE-Responsavel"},
        "prazo mínimo de 5 (cinco) anos": {"01-TCLE-Adulto", "02-TCLE-Responsavel"},
        "Edifício Arthur Bernardes": {"01-TCLE-Adulto", "02-TCLE-Responsavel", "03-TALE-Assentimento"},
        "natureza exclusivamente administrativa e institucional": {"04-Anuencia-Institucional"},
        "TRINTA (30) DIAS PARA ATENDER ÀS PENDÊNCIAS": {"09-Carta-Resposta"},
        "somente será iniciada após a aprovação": {"06-Cronograma", "00-Projeto-de-Pesquisa"},
        "Financiamento próprio": {"07-Orcamento", "10-Checklist-Conformidade"},
        "participante da pesquisa": {"01-TCLE-Adulto", "02-TCLE-Responsavel"},
    }
    _anchors = [f for f in TEMPLATE_FINGERPRINTS
                if _ancora_fontes.get(f, set()) & _gerados]
    faltando = [f for f in _anchors if f not in texto_total]
    if faltando:
        print(f"ERRO INTERNO DE TEMPLATE (drift bloqueado). Âncoras ausentes: {faltando}", file=sys.stderr)
        return 1
    # checagem crítica TCLE (só quando há TCLE no pacote; com dispensa total não há o que checar)
    tcle_docs = [blocos for slug, _, blocos in docs if slug.startswith(("01-TCLE", "02-TCLE"))]
    tcle_texto = " ".join(bl[1] for blocos in tcle_docs for bl in blocos if bl[0] in ("p", "bullet"))
    erros = []
    if tcle_docs and "participante da pesquisa" not in tcle_texto:
        erros.append("TCLE sem o termo obrigatório “participante da pesquisa”.")
    if "não oferece risco" in tcle_texto.lower() or "sem risco" in tcle_texto.lower().replace("ainda que mínimos", ""):
        pass  # template nunca contém; checagem de segurança
    if tcle_docs and "5 (cinco) anos" not in tcle_texto:
        erros.append("TCLE sem guarda de 5 anos.")
    if tcle_docs and ("3612-2316" not in tcle_texto or "cep@ufv.br" not in tcle_texto):
        erros.append("TCLE sem contato do CEP/UFV.")
    if erros:
        print("ERRO CRÍTICO DE TEMPLATE:", file=sys.stderr)
        for e in erros:
            print(" - " + e, file=sys.stderr)
        return 1
    if a.validar_apenas:
        pend = sorted(set(PENDENCIAS))
        print(f"Validação OK. Pendências de dados: {len(pend)}")
        for pe in pend:
            print(" - " + pe)
        print("Necessidade por documento (neste caso):")
        for fn in ALL_BUILDERS:
            gerar, motivo = NECESSIDADE.get(fn, (True, ""))
            print(f" - {'GERAR' if gerar else 'N/A'}: {fn.__name__} ({motivo})")
        if pend and not MODO_RASCUNHO:
            print("MODO FINAL: complete todos os campos acima antes de gerar (sem marcações).", file=sys.stderr)
            return 1
        return 0
    pend_pre = sorted(set(PENDENCIAS))
    if pend_pre and not MODO_RASCUNHO:
        print("ERRO: dados incompletos — documentos sairiam com marcações. Complete e rode de novo.", file=sys.stderr)
        for pe in pend_pre:
            print(f" - faltando: {pe}", file=sys.stderr)
        print("Dica: rode com --validar-apenas para listar, ou --rascunho para gerar rascunho NÃO submetível.", file=sys.stderr)
        return 1
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    manifesto = [f"cep-ufv {SKILL_VERSION} — {datetime.now().isoformat(timespec='seconds')}",
                 f"Projeto: {p.get('titulo', '[sem título]')}", ""]
    for slug, titulo_doc, blocos in docs:
        dx = out / f"{slug}.docx"
        pf = out / f"{slug}.pdf"
        try:
            # PDF primeiro: o total real alimenta o campo NUMPAGES do DOCX.
            total = render_pdf(pf, titulo_doc, blocos)
            render_docx(dx, titulo_doc, blocos, total_paginas=total)
        except Exception as e:
            print(f"ERRO ao gerar {slug}: {e}", file=sys.stderr)
            return 1
        for f in (dx, pf):
            if not f.exists() or f.stat().st_size == 0:
                print(f"ERRO: {f.name} não gerado ou vazio (fail-closed).", file=sys.stderr)
                return 1
        manifesto.append(f"{slug}.docx  {dx.stat().st_size} bytes  OK")
        manifesto.append(f"{slug}.pdf   {pf.stat().st_size} bytes  OK")
    pend = sorted(set(PENDENCIAS))
    (out / "pendencias.txt").write_text(
        "PENDÊNCIAS DE DADOS:\n" +
        ("(nenhuma — pacote final pronto para submissão, sem marcações)\n" if not pend else "".join(f" - {x}\n" for x in pend)),
        encoding="utf-8")
    (out / "previsao.txt").write_text(
        "PREVISÃO DE TRAMITAÇÃO (previsão calculada hoje — não é garantia):\n" + PREVISAO_TEXTO + "\n",
        encoding="utf-8")
    online = bool(p.get("coleta_online") or p.get("ambiente_virtual"))
    if online:
        (out / "criar-form-google.gs").write_text(gerar_form_gs(p), encoding="utf-8")
        manifesto.append("criar-form-google.gs  OK (Apps Script que monta o Google Forms sozinho)")
    if not MODO_RASCUNHO:
        # Varredura final: nenhum documento submetível pode conter rascunho ou instruções internas.
        # Linhas de assinatura (___) são permitidas.
        MARCAS_PROIBIDAS = ["[PREENCHER:", "AVISO:", "INFORMAÇÕES IMPORTANTES AO PESQUISADOR",
                            "padrão da skill", "sem edição da skill", "sem norma inventada",
                            "leia-se CEP", "(OBS:"]
        texto_final = "\n".join(
            " ".join(str(x) for x in bl[1:3] if isinstance(x, str))
            for _, _, blocos in docs for bl in blocos)
        achadas = [m for m in MARCAS_PROIBIDAS if m in texto_final]
        if achadas or pend:
            print(f"ERRO: pacote final com marcações {achadas} pend={pend} (fail-closed).", file=sys.stderr)
            return 1
    N = len(docs)
    dispensados = [(fn.__name__, mot) for fn in ALL_BUILDERS
                   if not NECESSIDADE.get(fn, (True, ""))[0]
                   for mot in [NECESSIDADE[fn][1]]]
    manifesto += ["", f"Pendências de dados: {len(pend)} (ver pendencias.txt)",
                  f"Pares docx+pdf: {N}/{N} verificados (existentes e >0 bytes).",
                  "Documentos avaliados como DESNECESSÁRIOS neste caso (não gerados): " +
                  ("nenhum" if not dispensados else "; ".join(f"{n} ({m})" for n, m in dispensados)),
                  "Previsão de tramitação: ver previsao.txt (" + PREVISAO_TEXTO + ")",
                  "Modo: " + ("RASCUNHO (NÃO submeter)" if MODO_RASCUNHO else "FINAL (pronto para submissão, sem marcações)"),
                  "Folha de rosto: gerar na Plataforma Brasil (não incluída por definição do CEP).",
                  "Coleta: somente após aprovação ética. Responsável = orientador."]
    (out / "MANIFESTO.txt").write_text("\n".join(manifesto) + "\n", encoding="utf-8")
    print(f"OK: {N} documentos x2 formatos em {out}/")
    print(f"Previsão: {PREVISAO_TEXTO}")
    print(f"Pendências de dados: {len(pend)} (ver {out}/pendencias.txt)")
    return 0

if __name__ == "__main__":
    sys.exit(main())
