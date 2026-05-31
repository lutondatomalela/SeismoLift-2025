# Changelog

## v5.0.0-rc3

### Melhorado
- Removida a indicação "Base sísmica: automática" da GUI principal, deixando a interface mais limpa.
- Mantido o carregamento automático da base `Zonas_Sismicas_PT.xlsx` e a selecção manual apenas como mecanismo de recuperação em caso de falha.
- A referência ao zonamento sísmico permanece na documentação e nos relatórios técnicos, não na janela principal.

## v5.0.0-rc2

### Melhorado
- A base sísmica `Zonas_Sismicas_PT.xlsx` passa a ser carregada automaticamente, sem botão visível na interface principal.
- Adicionado suporte mais robusto para caminhos em executável PyInstaller `onefile` e `onedir`.
- Mantida selecção manual da base apenas como mecanismo de recuperação em caso de falha de carregamento automático.

## v5.0.0-rc1

Versão candidata para publicação no GitHub.

### Adicionado
- Script `build_windows.bat` para gerar executável Windows com PyInstaller.
- Ficheiro `.gitignore` para evitar commits de builds, caches e relatórios exportados.
- Ficheiro `VERSION.txt`.
- Ficheiro `requirements-build.txt` para dependências de empacotamento.
- Relatórios específicos dos espetros de resposta em DOCX e PDF.
- Exportação de espetros em XLSX, CSV e PNG.

### Melhorado
- GUI redimensionável com painéis ajustáveis.
- Aba de espetros de resposta com resumo técnico, Tipo 1/Tipo 2, marcadores TB/TC/TD/T1/Ta e unidade em m/s² ou g.
- Textos compactos para classes de importância, tipo de terreno e critério γa.
- Relatórios DOCX/PDF com formatação consistente e metadados.
- Exportações XLSX com metadados.
- Ícone da aplicação com fundo transparente.

### Corrigido
- Nome da ferramenta uniformizado como `SeismoLift`.
- Avaliação do Tipo 1 e Tipo 2 para Portugal Continental, com adopção automática da acção condicionante.
- Utilização dos coeficientes de importância do Anexo Nacional português, conforme região e acção sísmica.

### Nota
Esta versão é adequada para teste externo e publicação como release candidate. Recomenda-se validação adicional em vários concelhos, classes de importância, tipos de terreno e formatos de exportação antes de marcar a versão como estável.
