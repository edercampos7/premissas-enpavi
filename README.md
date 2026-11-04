# Encontro de Contas Enpavi — SCP Tríade × Enpavi

Apresentação didática da aba **PREMISSAS ENPAVI** da planilha de fechamento: como a parte da Enpavi é calculada obra a obra, a administração central, a Usina Caçapava, o custo 029, a matriz, as obras antigas, as retiradas, o encontro de contas mês a mês com juros e um simulador.

**Site:** https://edercampos7.github.io/premissas-enpavi/

Os números vêm direto da planilha `... ENCONTRO ENPAVI.xlsx`. A planilha **não** vai para o GitHub: só os valores que aparecem na apresentação.

---

## Como o site se atualiza

```
Você salva a planilha no Excel
        ↓  (o vigia percebe em até ~1 minuto)
scripts\vigiar_planilha.ps1 lê a planilha e gera o index.html
        ↓
envia ao GitHub (commit + push)
        ↓  (~1 a 2 minutos)
o site mostra os valores novos
```

Quando termina, aparece um aviso no canto da tela ("Site Enpavi atualizado" ou "NÃO atualizado", com o motivo).

O vigia usa a planilha **mais recente** cujo nome termina em **`ENCONTRO ENPAVI.xlsx`** na pasta **Área de Trabalho\Custo Enpavi**. Pode trocar o mês no começo do nome (082026, 092026...). Cópias com outro final, como `... ENCONTRO ENPAVI (versao antes de 30-10).xlsx`, são ignoradas.

## Instalar a atualização automática (uma vez só)

1. Aperte a tecla **Windows**, digite **PowerShell** e aperte **Enter**.
2. Cole a linha abaixo e aperte **Enter**:

   ```powershell
   irm https://raw.githubusercontent.com/edercampos7/premissas-enpavi/main/scripts/instalar.ps1 | iex
   ```

3. Se abrir uma janela do GitHub, entre com a conta **edercampos7** (só na primeira vez).
4. No fim aparece **=== Pronto ===** com o endereço do site.

O instalador:
- confere o Git (se não tiver, tenta instalar pelo winget; se não conseguir, abre a página de download);
- baixa o projeto para **Documentos\premissas-enpavi**;
- publica o site uma vez com a planilha atual;
- cria a tarefa agendada **"Atualizar site Premissas Enpavi"**, que inicia o vigia sempre que você entra no Windows (se não puder criar a tarefa, usa um atalho na pasta Inicializar).

Pode rodar o instalador de novo a qualquer momento; ele só atualiza a instalação.

## Atualizar na hora

Dois cliques em **Documentos\premissas-enpavi\atualizar_agora.bat**. Ele publica com a planilha atual, mesmo que nada tenha mudado.

## Conferir se atualizou

- No rodapé do site: **"Planilha salva em ..."** (horário em que a planilha foi salva) e **"Apresentação gerada em ..."**.
- No simulador: a faixa verde **"Conferido"** confirma que a conta refeita no site bate com o resumo K4–K11 da planilha.
- No topo do site, o seletor **Juros aplicados** mostra tudo com juros simples ou compostos. O padrão é a forma da célula C6 da PREMISSAS; a outra forma é recalculada no próprio site (endereço com `?juros=composto` ou `?juros=simples`) e a planilha não muda.
- Histórico do que o vigia fez: **Documentos\premissas-enpavi\scripts\log.txt**.

## Se aparecer "Site Enpavi NÃO atualizado"

| Motivo no aviso / log | O que fazer |
|---|---|
| Problema na planilha (ex.: "não tem a aba 'PREMISSAS ENPAVI'", "cabeçalho do item 1 não encontrado") | A estrutura da aba mudou. O site continua com a última versão boa. Ajuste a planilha e salve de novo. |
| Falha ao enviar ao GitHub | Normalmente é internet ou login. O vigia tenta de novo a cada 5 minutos. Se continuar, rode o **atualizar_agora.bat** (ele abre a janela de login do GitHub, se for preciso). |

O que o script lê da planilha:
- **PREMISSAS ENPAVI**: parâmetros C3–C8, resumo K4–K11, item 1 (participações, cabeçalho "CÓDIGO") e item 2 (provisões);
- **ENCONTRO ENPAVI**: meses na coluna A a partir da linha 12, fluxos nas colunas B–AA e totais AC–AN (cabeçalhos na linha 7);
- **RECEITAS ENPAVI-USINAS SP-ESO** e os **CAIXA OBRA** usados na administração central.

Ele usa os valores que o Excel grava ao salvar, então a planilha precisa ter sido salva pelo Excel.

## Mudar os textos do site

Os textos fixos (por exemplo **Pontos de Observação**) ficam em `scripts/template.html`. Dá para editar direto no GitHub (abrir o arquivo → lápis → *Commit changes*). Em até 15 minutos o vigia baixa a mudança e publica o site de novo com a planilha atual.

## Desligar a atualização automática

Na pasta **Documentos\premissas-enpavi\scripts**, botão direito em **desinstalar.ps1** → **Executar com o PowerShell**. O site continua no ar com a última versão.

## Arquivos

| Arquivo | Para que serve |
|---|---|
| `index.html` | O site (gerado automaticamente; não editar à mão) |
| `scripts/atualizar_painel.ps1` | Lê a planilha e gera o `index.html` |
| `scripts/template.html` | Layout, textos e gráficos do site |
| `scripts/vigiar_planilha.ps1` | Vigia a planilha e publica no GitHub |
| `scripts/instalar.ps1` / `desinstalar.ps1` | Liga / desliga a atualização automática |
| `atualizar_agora.bat` | Publica na hora |
| `scripts/testar_simulador.js` | Teste opcional (Node.js): confere se o simulador reproduz o resumo da planilha |
| `scripts/logo_base64.txt` | Logo da Tríade usada no topo |
