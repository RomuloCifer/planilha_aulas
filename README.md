# Sistema de Controle de Alunos e Aulas (Excel + VBA) 📊

Descrição do projeto
- Objetivo: Criar um sistema em Microsoft Excel com VBA para gerenciar alunos particulares, pacotes de aulas, aulas ministradas, pagamentos e indicadores financeiros e operacionais. O sistema foi pensado para uso contínuo ao longo do ano, sem dependência de referências fixas de mês.

Principais funcionalidades
- Cadastro e gestão de alunos (ID automático).
- Controle de pacotes e saldo de aulas (status positivo/negativo).
- Registro de aulas com data e conteúdo, por aluno.
- Histórico de pagamentos e sincronização com saldo de aulas.
- Dashboard com cards e gráficos: alunos ativos, receita esperada, receita recebida e evolução mensal.
- Formatação condicional e indicadores visuais (barra de progresso e cores) para facilitar cobranças.
- Interface com botões e rotinas VBA para cadastro rápido e carregamento dinâmico.

Estrutura do projeto (abas principais)

- `Controle_Alunos`
  - Tabela: `tb_alunos`
  - Campos principais:
    - `ID` (numérico, gerado automaticamente)
    - `Nome`
    - `Pacote` (quantidade de aulas contratadas por mês)
    - `Status` (saldo de aulas: pode ser negativo ou positivo)
    - `Dia da aula`
    - `Telefone`
    - `Valor` (mensalidade)
    - `Objetivo`
    - `Ativo` (SIM / NÃO)
  - Visual: coluna com barra de progresso e cores (verde / amarelo / vermelho) indicando necessidade de cobrança.

- `Controle_Aulas`
  - Interface mensal (abas ou área de seleção: Janeiro a Dezembro)
  - Seleção do aluno e carregamento automático das aulas do mês
  - Registro das aulas com `Data` + `Conteúdo`
  - Coluna `Pago?` (SIM / NÃO) por mês
  - Botão para cadastrar aulas (rotina VBA)
  - Limite visual configurado para mostrar até 10 aulas por mês

- `BASES`
  - `tb_aulas`: histórico de aulas (AlunoID, DataAula, Conteúdo, MesRef)
  - `tb_pagamentos`: histórico de pagamentos (AlunoID, MesRef, Pago, DataPgto, ValorPago)

- `Dashboard`
  - Cards e gráficos alimentados pelas tabelas estruturadas
  - Indicadores principais:
    - Total de alunos ativos
    - Valor esperado no mês
    - Valor recebido por mês
    - Evolução mensal de receita
    - Base preparada para horas trabalhadas e variação mês a mês

Detalhes das funcionalidades implementadas
- Cadastro de alunos via VBA (ex.: `InputBox`) com geração automática de `ID`.
- `Status` do aluno calculado como saldo de aulas (aulas dadas – aulas pagas):
  - Aumenta +1 a cada aula registrada.
  - Diminui automaticamente ao registrar pagamento do pacote.
  - Aceita valores negativos (débito) ou positivos (crédito de aulas).
- Registro e sincronização de pagamentos com `tb_pagamentos`.
- Carregamento dinâmico das aulas ao trocar de aluno na interface (`Controle_Aulas`).
- Ordenação automática das aulas por data.
- Formatação condicional inteligente para indicar necessidade de cobrança.
- Dashboard dinâmico sem referências fixas de mês (uso de fórmulas relativas/estruturadas).

Lógica de negócio (resumo)
- `Status` representa o saldo de aulas: aulas ministradas menos aulas pagas (pode ser ajustado por pagamentos fora do ciclo).
- Pagamentos podem ocorrer antes ou depois das aulas — o sistema aceita ambos os fluxos.
- Cobrança é visual e informativa, não impede o registro de aulas (não é bloqueante).
- O sistema foi desenhado para refletir a operação real de aulas particulares e facilitar acompanhamento mensal e anual.

Tecnologias e recursos usados
- Microsoft Excel (.xlsx/.xlsm)
- VBA (Visual Basic for Applications) — rotinas para cadastro, carregamento e sincronização
- Tabelas estruturadas (ListObjects)
- Fórmulas: `SOMASES`, `CONT.SES`, funções de data, `PROCV`/`XLOOKUP` (conforme versão) e outras auxiliares
- Formatação condicional e elementos visuais (barras, cores)
- Botões e controles para operações rápidas

Boas práticas e segurança
- O arquivo utiliza macros (formato `.xlsm`). Habilite macros apenas se confiar na origem.
- Faça sempre uma cópia de segurança antes de alterações significativas.
- Para revisar/editar código VBA: `Alt+F11` para abrir o Editor VBA.
- Mantenha versões nomeadas (ex.: `aulas-particulares_v1.xlsm`, `aulas-particulares_v2.xlsm`).


Como contribuir / editar
1. Trabalhe em uma cópia do arquivo original.
2. Documente alterações no histórico de versões.
3. Ao alterar rotinas VBA, teste em ambiente controlado e verifique ordenação e sincronização de tabelas.

Contato
- Responsável: (preencher nome / e-mail do autor)

Histórico
- v1.0 — README profissional criado (data: 2026-01-06).

---
Arquivo: `aulas_particulares.xlsm` — README gerado automaticamente conforme escopo fornecido.
