# Conciliação entre Ativo e Passivo — VBA
Macro para organizar relatórios e identificar divergências entre as abas `Ativo` e `Passivo`.

## Regra de cruzamento
- No `Ativo`, a chave é `Filial-Agente`.
- No `Passivo`, a chave é `Agente-Filial`.
- A coluna J busca o saldo da coluna H da outra aba.
- A coluna K soma os dois saldos. O resultado esperado é zero quando os valores se compensam.
  
## Como utilizar
1. Prepare as abas `Ativo` e `Passivo` no formato esperado pelo código.
2. Mantenha os identificadores no padrão `Filial: código - descrição` e `Agente: código - descrição`.
3. Execute a macro `nd`.
4. Confira as linhas destacadas e a mensagem final de divergências.

## Cuidados
- **Faça uma cópia dos dados antes de executar:** a macro exclui linhas e colunas.
- Não execute novamente sobre relatórios já tratados.
- O cruzamento não soma chaves repetidas: utiliza o resultado retornado pelo `XLOOKUP`.
- Diferenças são comparadas diretamente com zero, sem arredondamento.
