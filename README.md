# SeismoLift

**SeismoLift** é uma ferramenta gráfica para cálculo da aceleração de projecto e da categoria sísmica de ascensores em Portugal, com base no Eurocódigo 8, na EN 81-77 e nas ET 11/2020 para edifícios hospitalares sujeitos a condições sísmicas.

Repositório: https://github.com/lutondatomalela/SeismoLift-2025

## Estado da versão

Versão actual: **v5.0.0-rc3**

## Funcionalidades principais

- Consulta de zonamento sísmico para Portugal Continental, Madeira e Açores.
- Avaliação automática das acções sísmicas Tipo 1 e Tipo 2 no Continente.
- Adopção automática da acção sísmica condicionante.
- Cálculo da categoria sísmica do ascensor segundo a EN 81-77.
- Dois modos de cálculo:
  - `Geral EC8 / EN 81-77`;
  - `ET 11/2020 - edifício de base fixa`.
- Classes de importância e tipos de terreno apresentados com descrição sucinta.
- Exportação de relatórios da categoria sísmica em DOCX e PDF.
- Exportação dos dados de cálculo em XLSX, com metadados.
- Aba de espetros de resposta com:
  - visualização simultânea dos espetros Tipo 1 e Tipo 2, quando aplicável;
  - marcação de `TB`, `TC`, `TD`, `T1` e `Ta`;
  - unidade em `m/s²` ou `g`;
  - cálculo automático de `T1 = Ct·H^0,75` ou introdução manual;
  - exportação de espetros em XLSX, CSV e PNG;
  - relatório dos espetros em DOCX e PDF.
- Interface gráfica redimensionável, com painéis ajustáveis e scroll nas áreas principais.
- Ícone próprio com fundo transparente.

## Estrutura da pasta

```text
SeismoLift/
├── SeismoLift.py
├── Zonas_Sismicas_PT.xlsx
├── assets/
│   ├── seismolift_icon.ico
│   ├── seismolift_icon_32.png
│   └── seismolift_icon_128.png
├── requirements.txt
├── requirements-build.txt
├── build_windows.bat
├── VERSION.txt
├── CHANGELOG.md
├── LICENSE.md
├── .gitignore
└── README.md
```

## Instalação para correr com Python

Requer Python 3.10 ou superior.

```bash
pip install -r requirements.txt
python SeismoLift.py
```

## Com executável Windows

## Base sísmica

A base `Zonas_Sismicas_PT.xlsx` é carregada automaticamente e não é apresentada na interface principal. A referência normativa ao zonamento sísmico é mantida nos relatórios e na documentação.

Em modo Python, manter estes elementos na mesma pasta:

```text
SeismoLift.py
Zonas_Sismicas_PT.xlsx
assets/
```

## Notas técnicas

- Para Portugal Continental, o programa avalia a acção sísmica Tipo 1 e Tipo 2 e retém a condicionante.
- O modo `ET 11/2020 - edifício de base fixa` é uma formulação prática para edifícios de base fixa.
- A determinação da categoria sísmica não substitui a verificação completa dos requisitos construtivos e funcionais da EN 81-77.
- Os relatórios incluem metadados com autoria de `Engº Lutonda Tomalela`.

## Autor

**Engº Lutonda Tomalela**

## Licença

MIT License. Ver `LICENSE.md`.
