#!/usr/bin/env python3
"""Generate an Excel template to assess technical and functional documentation quality."""

from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile


OUTPUT_FILE = Path(__file__).with_name("avaliacao_qualidade_documentacao.xlsx")

CRITERIA = [
    (
        "C01",
        "Governança",
        "Clareza de objetivo e âmbito",
        "Explica para que serve o documento, o sistema abrangido, os limites e o público-alvo.",
        "Ambas",
        5,
    ),
    (
        "C02",
        "Governança",
        "Estrutura e navegabilidade",
        "Tem índice, secções coerentes, referências cruzadas e leitura fácil.",
        "Ambas",
        5,
    ),
    (
        "C03",
        "Governança",
        "Atualização e controlo de alterações",
        "Inclui versão, data, autores, histórico de alterações e estado de aprovação.",
        "Ambas",
        5,
    ),
    (
        "C04",
        "Governança",
        "Consistência terminológica",
        "Usa os mesmos termos de negócio e tecnologia em todo o documento.",
        "Ambas",
        4,
    ),
    (
        "C05",
        "Governança",
        "Rastreabilidade",
        "Liga requisitos, processos, decisões técnicas, testes e evidências.",
        "Ambas",
        8,
    ),
    (
        "C06",
        "Funcional",
        "Cobertura de requisitos e cenários",
        "Descreve os requisitos principais, fluxos alternativos e casos limite do negócio.",
        "Funcional",
        9,
    ),
    (
        "C07",
        "Funcional",
        "Regras de negócio e exceções",
        "Documenta regras, validações, exceções e critérios de decisão relevantes.",
        "Funcional",
        7,
    ),
    (
        "C08",
        "Funcional",
        "Perfis, permissões e impacto operacional",
        "Identifica atores, permissões, responsabilidades e efeitos no processo operativo.",
        "Funcional",
        4,
    ),
    (
        "C09",
        "Funcional",
        "Fluxos, casos de uso e exemplos",
        "Usa exemplos claros, fluxogramas ou casos de uso para facilitar entendimento.",
        "Funcional",
        5,
    ),
    (
        "C10",
        "Técnica",
        "Arquitetura e componentes",
        "Explica módulos, dependências, responsabilidade de cada componente e decisões-chave.",
        "Técnica",
        8,
    ),
    (
        "C11",
        "Técnica",
        "APIs, integrações e contratos",
        "Detalha interfaces, endpoints, eventos, formatos de dados e dependências externas.",
        "Técnica",
        8,
    ),
    (
        "C12",
        "Técnica",
        "Modelo de dados e dicionário",
        "Inclui entidades, campos, regras de integridade, origem e consumo dos dados.",
        "Técnica",
        6,
    ),
    (
        "C13",
        "Técnica",
        "Instalação, configuração e dependências",
        "Permite preparar o ambiente e operar a solução sem conhecimento tácito.",
        "Técnica",
        6,
    ),
    (
        "C14",
        "Técnica",
        "Operação, monitorização e suporte",
        "Documenta logs, alertas, troubleshooting, backup e procedimentos operacionais.",
        "Técnica",
        4,
    ),
    (
        "C15",
        "Técnica",
        "Segurança e requisitos não funcionais",
        "Cobre segurança, performance, disponibilidade, auditoria e conformidade.",
        "Técnica",
        8,
    ),
    (
        "C16",
        "Validação",
        "Critérios de aceitação e evidência de teste",
        "Mostra como validar o conteúdo com testes, evidências ou critérios objetivos.",
        "Ambas",
        8,
    ),
]


def xml_text(value: str) -> str:
    return escape(value).replace("\n", "&#10;")


def inline_text_cell(ref: str, text: str, style: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    return f'<c r="{ref}" t="inlineStr"{style_attr}><is><t>{xml_text(text)}</t></is></c>'


def blank_cell(ref: str, style: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    return f'<c r="{ref}"{style_attr}/>'


def number_cell(ref: str, value: int | float, style: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    return f'<c r="{ref}"{style_attr}><v>{value}</v></c>'


def formula_cell(ref: str, formula: str, style: int | None = None) -> str:
    style_attr = f' s="{style}"' if style is not None else ""
    return f'<c r="{ref}"{style_attr}><f>{escape(formula)}</f></c>'


def row_xml(index: int, cells: list[str], height: int | None = None) -> str:
    if height is None:
        return f'<row r="{index}">{"".join(cells)}</row>'
    return (
        f'<row r="{index}" ht="{height}" customHeight="1">'
        f'{"".join(cells)}'
        f"</row>"
    )


def build_instructions_sheet() -> str:
    rows = [
        row_xml(1, [inline_text_cell("A1", "Modelo de avaliação da qualidade da documentação técnica e funcional", 1)], 26),
        row_xml(3, [inline_text_cell("A3", "Como usar", 2)]),
        row_xml(
            4,
            [
                inline_text_cell(
                    "A4",
                    "1. Preencha o contexto da avaliação na folha 'Avaliacao'.\n"
                    "2. Atribua uma nota de 1 a 5 a cada critério.\n"
                    "3. Registe evidências, riscos e ações recomendadas.\n"
                    "4. Use a pontuação global e as médias técnica/funcional para priorizar melhorias.",
                    4,
                )
            ],
            56,
        ),
        row_xml(6, [inline_text_cell("A6", "Escala de pontuação", 2)]),
        row_xml(
            7,
            [
                inline_text_cell("A7", "Nota", 3),
                inline_text_cell("B7", "Interpretação", 3),
                inline_text_cell("C7", "Ação sugerida", 3),
            ],
            22,
        ),
        row_xml(8, [number_cell("A8", 1, 5), inline_text_cell("B8", "Inexistente ou incorreto", 4), inline_text_cell("C8", "Criar de raiz e tratar como risco crítico", 4)], 26),
        row_xml(9, [number_cell("A9", 2, 5), inline_text_cell("B9", "Muito fraco", 4), inline_text_cell("C9", "Corrigir urgentemente antes de depender do documento", 4)], 26),
        row_xml(10, [number_cell("A10", 3, 5), inline_text_cell("B10", "Aceitável, mas incompleto", 4), inline_text_cell("C10", "Planear melhorias com prioridade média", 4)], 26),
        row_xml(11, [number_cell("A11", 4, 5), inline_text_cell("B11", "Bom", 4), inline_text_cell("C11", "Ajustar pontos residuais e manter atualizado", 4)], 26),
        row_xml(12, [number_cell("A12", 5, 5), inline_text_cell("B12", "Excelente", 4), inline_text_cell("C12", "Manter como referência e replicar padrão", 4)], 26),
        row_xml(14, [inline_text_cell("A14", "Leitura da pontuação global", 2)]),
        row_xml(
            15,
            [
                inline_text_cell("A15", "Faixa", 3),
                inline_text_cell("B15", "Leitura", 3),
                inline_text_cell("C15", "Decisão sugerida", 3),
            ],
            22,
        ),
        row_xml(16, [inline_text_cell("A16", "85-100", 5), inline_text_cell("B16", "Excelente", 4), inline_text_cell("C16", "Documentação robusta e reutilizável", 4)], 24),
        row_xml(17, [inline_text_cell("A17", "70-84,9", 5), inline_text_cell("B17", "Boa", 4), inline_text_cell("C17", "Apta para uso com pequenas melhorias", 4)], 24),
        row_xml(18, [inline_text_cell("A18", "50-69,9", 5), inline_text_cell("B18", "Suficiente", 4), inline_text_cell("C18", "Usável, mas com lacunas relevantes", 4)], 24),
        row_xml(19, [inline_text_cell("A19", "<50", 5), inline_text_cell("B19", "Crítica", 4), inline_text_cell("C19", "Necessita intervenção estruturada", 4)], 24),
        row_xml(21, [inline_text_cell("A21", "Boas práticas", 2)]),
        row_xml(
            22,
            [
                inline_text_cell(
                    "A22",
                    "Registe links, versões e evidências objetivas. Sempre que uma nota seja 1 ou 2, defina um responsável e uma ação concreta para eliminar a lacuna.",
                    4,
                )
            ],
            40,
        ),
    ]
    merges = [
        "A1:F1",
        "A3:C3",
        "A4:F4",
        "A6:C6",
        "A14:C14",
        "A21:C21",
        "A22:F22",
    ]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:F22"/>
  <sheetViews>
    <sheetView workbookViewId="0"/>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="18" customWidth="1"/>
    <col min="2" max="2" width="36" customWidth="1"/>
    <col min="3" max="3" width="46" customWidth="1"/>
    <col min="4" max="6" width="18" customWidth="1"/>
  </cols>
  <sheetData>
    {''.join(rows)}
  </sheetData>
  <mergeCells count="{len(merges)}">
    {''.join(f'<mergeCell ref="{ref}"/>' for ref in merges)}
  </mergeCells>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_evaluation_sheet() -> str:
    rows = [
        row_xml(1, [inline_text_cell("A1", "Avaliação da qualidade documental", 1)], 26),
        row_xml(
            2,
            [
                inline_text_cell("A2", "Projeto / sistema", 2),
                blank_cell("B2", 6),
                inline_text_cell("J2", "Pontuação global (0-100)", 8),
                formula_cell("L2", "ROUND(SUM($H$9:$H$24),1)", 9),
            ],
            22,
        ),
        row_xml(
            3,
            [
                inline_text_cell("A3", "Data da avaliação", 2),
                blank_cell("B3", 6),
                inline_text_cell("J3", "Média técnica (0-100)", 8),
                formula_cell(
                    "L3",
                    'IFERROR(ROUND(SUMIFS($H$9:$H$24,$E$9:$E$24,"Técnica")/SUMIFS($F$9:$F$24,$E$9:$E$24,"Técnica")*100,1),"")',
                    9,
                ),
            ],
            22,
        ),
        row_xml(
            4,
            [
                inline_text_cell("A4", "Avaliador(es)", 2),
                blank_cell("B4", 6),
                inline_text_cell("J4", "Média funcional (0-100)", 8),
                formula_cell(
                    "L4",
                    'IFERROR(ROUND(SUMIFS($H$9:$H$24,$E$9:$E$24,"Funcional")/SUMIFS($F$9:$F$24,$E$9:$E$24,"Funcional")*100,1),"")',
                    9,
                ),
            ],
            22,
        ),
        row_xml(
            5,
            [
                inline_text_cell("J5", "Nível de qualidade", 8),
                formula_cell(
                    "L5",
                    'IF(L2="","Por avaliar",IF(L2>=85,"Excelente",IF(L2>=70,"Boa",IF(L2>=50,"Suficiente","Crítica"))))',
                    9,
                ),
            ],
            22,
        ),
        row_xml(
            6,
            [
                inline_text_cell("J6", "% de critérios críticos", 8),
                formula_cell(
                    "L6",
                    'IF(COUNT($G$9:$G$24)=0,"",ROUND(COUNTIF($G$9:$G$24,"<=2")/COUNT($G$9:$G$24)*100,0))',
                    9,
                ),
            ],
            22,
        ),
        row_xml(
            8,
            [
                inline_text_cell("A8", "ID", 3),
                inline_text_cell("B8", "Dimensão", 3),
                inline_text_cell("C8", "Critério", 3),
                inline_text_cell("D8", "O que validar", 3),
                inline_text_cell("E8", "Tipo", 3),
                inline_text_cell("F8", "Peso", 3),
                inline_text_cell("G8", "Nota (1-5)", 3),
                inline_text_cell("H8", "Pontuação", 3),
                inline_text_cell("I8", "Evidências / referências", 3),
                inline_text_cell("J8", "Riscos identificados", 3),
                inline_text_cell("K8", "Ações recomendadas", 3),
                inline_text_cell("L8", "Responsável", 3),
                inline_text_cell("M8", "Prioridade", 3),
                inline_text_cell("N8", "Estado", 3),
            ],
            32,
        ),
    ]

    for row_index, (code, dimension, criterion, validation, doc_type, weight) in enumerate(CRITERIA, start=9):
        rows.append(
            row_xml(
                row_index,
                [
                    inline_text_cell(f"A{row_index}", code, 5),
                    inline_text_cell(f"B{row_index}", dimension, 4),
                    inline_text_cell(f"C{row_index}", criterion, 4),
                    inline_text_cell(f"D{row_index}", validation, 4),
                    inline_text_cell(f"E{row_index}", doc_type, 5),
                    number_cell(f"F{row_index}", weight, 5),
                    blank_cell(f"G{row_index}", 10),
                    formula_cell(
                        f"H{row_index}",
                        f'IF(G{row_index}="","",ROUND(F{row_index}*G{row_index}/5,1))',
                        7,
                    ),
                    blank_cell(f"I{row_index}", 6),
                    blank_cell(f"J{row_index}", 6),
                    blank_cell(f"K{row_index}", 6),
                    blank_cell(f"L{row_index}", 6),
                    formula_cell(
                        f"M{row_index}",
                        f'IF(G{row_index}="","",IF(G{row_index}<=2,"Alta",IF(G{row_index}=3,"Média","Baixa")))',
                        7,
                    ),
                    blank_cell(f"N{row_index}", 6),
                ],
                42,
            )
        )

    merges = ["A1:N1", "B2:H2", "B3:H3", "B4:H4", "J2:K2", "J3:K3", "J4:K4", "J5:K5", "J6:K6"]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:N24"/>
  <sheetViews>
    <sheetView workbookViewId="0">
      <pane ySplit="8" topLeftCell="A9" activePane="bottomLeft" state="frozen"/>
      <selection pane="bottomLeft" activeCell="B2" sqref="B2"/>
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="18"/>
  <cols>
    <col min="1" max="1" width="10" customWidth="1"/>
    <col min="2" max="2" width="16" customWidth="1"/>
    <col min="3" max="3" width="28" customWidth="1"/>
    <col min="4" max="4" width="44" customWidth="1"/>
    <col min="5" max="5" width="14" customWidth="1"/>
    <col min="6" max="6" width="10" customWidth="1"/>
    <col min="7" max="7" width="11" customWidth="1"/>
    <col min="8" max="8" width="12" customWidth="1"/>
    <col min="9" max="11" width="28" customWidth="1"/>
    <col min="12" max="12" width="18" customWidth="1"/>
    <col min="13" max="13" width="12" customWidth="1"/>
    <col min="14" max="14" width="14" customWidth="1"/>
  </cols>
  <sheetData>
    {''.join(rows)}
  </sheetData>
  <autoFilter ref="A8:N24"/>
  <mergeCells count="{len(merges)}">
    {''.join(f'<mergeCell ref="{ref}"/>' for ref in merges)}
  </mergeCells>
  <dataValidations count="2">
    <dataValidation type="whole" operator="between" allowBlank="1" showErrorMessage="1" sqref="G9:G24">
      <formula1>1</formula1>
      <formula2>5</formula2>
    </dataValidation>
    <dataValidation type="list" allowBlank="1" showInputMessage="1" sqref="N9:N24">
      <formula1>"Não iniciado,Em curso,Planeado,Concluído"</formula1>
    </dataValidation>
  </dataValidations>
  <pageMargins left="0.5" right="0.5" top="0.7" bottom="0.7" header="0.3" footer="0.3"/>
</worksheet>
"""


def build_styles() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1">
    <numFmt numFmtId="164" formatCode="0.0"/>
  </numFmts>
  <fonts count="4">
    <font>
      <sz val="11"/>
      <color theme="1"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="11"/>
      <color rgb="FF1F2937"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
    <font>
      <b/>
      <sz val="14"/>
      <color rgb="FFFFFFFF"/>
      <name val="Calibri"/>
      <family val="2"/>
    </font>
  </fonts>
  <fills count="7">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFD9EAF7"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2F0D9"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFF2F2F2"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border>
      <left/><right/><top/><bottom/><diagonal/>
    </border>
    <border>
      <left style="thin"><color rgb="FFB7C3D0"/></left>
      <right style="thin"><color rgb="FFB7C3D0"/></right>
      <top style="thin"><color rgb="FFB7C3D0"/></top>
      <bottom style="thin"><color rgb="FFB7C3D0"/></bottom>
      <diagonal/>
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="11">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="top" wrapText="1"/>
    </xf>
    <xf numFmtId="164" fontId="0" fillId="6" borderId="1" xfId="0" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="164" fontId="2" fillId="5" borderId="1" xfId="0" applyFont="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
    <xf numFmtId="0" fontId="0" fillId="4" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1">
      <alignment horizontal="center" vertical="center" wrapText="1"/>
    </xf>
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
"""


def build_workbook() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl"/>
  <workbookPr defaultThemeVersion="166925"/>
  <bookViews>
    <workbookView xWindow="0" yWindow="0" windowWidth="28800" windowHeight="17280"/>
  </bookViews>
  <sheets>
    <sheet name="Instrucoes" sheetId="1" r:id="rId1"/>
    <sheet name="Avaliacao" sheetId="2" r:id="rId2"/>
  </sheets>
  <calcPr calcId="191029" calcMode="auto" fullCalcOnLoad="1"/>
</workbook>
"""


def build_workbook_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
"""


def build_root_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


def build_content_types() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""


def build_app_properties() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
            xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Microsoft Excel</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>2</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="2" baseType="lpstr">
      <vt:lpstr>Instrucoes</vt:lpstr>
      <vt:lpstr>Avaliacao</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0300</AppVersion>
</Properties>
"""


def build_core_properties() -> str:
    timestamp = datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
                   xmlns:dc="http://purl.org/dc/elements/1.1/"
                   xmlns:dcterms="http://purl.org/dc/terms/"
                   xmlns:dcmitype="http://purl.org/dc/dcmitype/"
                   xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>Modelo de avaliação da qualidade da documentação</dc:title>
  <dc:subject>Avaliação de documentação técnica e funcional</dc:subject>
  <dc:creator>Cursor</dc:creator>
  <cp:keywords>documentação,qualidade,excel,avaliação</cp:keywords>
  <dc:description>Template Excel para avaliar documentação técnica e funcional com pesos e pontuação automática.</dc:description>
  <cp:lastModifiedBy>Cursor</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">{timestamp}</dcterms:modified>
</cp:coreProperties>
"""


def generate_workbook(destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(destination, "w", compression=ZIP_DEFLATED) as workbook:
        workbook.writestr("[Content_Types].xml", build_content_types())
        workbook.writestr("_rels/.rels", build_root_rels())
        workbook.writestr("docProps/app.xml", build_app_properties())
        workbook.writestr("docProps/core.xml", build_core_properties())
        workbook.writestr("xl/workbook.xml", build_workbook())
        workbook.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels())
        workbook.writestr("xl/styles.xml", build_styles())
        workbook.writestr("xl/worksheets/sheet1.xml", build_instructions_sheet())
        workbook.writestr("xl/worksheets/sheet2.xml", build_evaluation_sheet())


if __name__ == "__main__":
    generate_workbook(OUTPUT_FILE)
    print(f"Workbook criado em: {OUTPUT_FILE}")
