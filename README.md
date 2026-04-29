# Modelo Excel para avaliar documentação

Este repositório inclui um gerador de um ficheiro Excel para avaliar a qualidade de documentação técnica e funcional.

## Gerar o ficheiro

```bash
python3 generate_documentation_quality_excel.py
```

O comando cria o ficheiro `avaliacao_qualidade_documentacao.xlsx` na raiz do repositório.

## O que o ficheiro inclui

- folha `Instrucoes` com escala de avaliação e interpretação do resultado;
- folha `Avaliacao` com critérios ponderados;
- cálculo automático de pontuação global;
- médias separadas para documentação técnica e funcional;
- campos para evidências, riscos, ações, responsável, prioridade e estado.
