# Conferência de DRE e BP com a ECD — VBA
Macros para cruzar os valores das abas `DRE` e `BP` com os dados da aba `ECD`, destacando diferenças e erros de busca.
## Macros disponíveis
| Macro | Finalidade |
|---|---|
| `ecd_dre` | Confere a aba `DRE` com a `ECD` |
| `ecd_bp` | Confere a aba `BP` com a `ECD` |
| `ReplaceAcentos` | Substitui os caracteres acentuados previstos no código |
## Estrutura necessária
A pasta de trabalho ativa deve conter a aba `ECD` e a aba correspondente à macro executada: `DRE` ou `BP`.
### Aba ECD
- Se `A1` for diferente de zero, a macro separa o intervalo `A1:A100` pelo delimitador `|`.
- Após a separação, a coluna H deve conter as descrições utilizadas na busca.
- A coluna K deve conter os valores de comparação.
## Como utilizar
1. Prepare as abas no formato esperado, colando o arquivo txt que vem do validador
2. Execute `ecd_dre` para conferir a DRE ou `ecd_bp` para conferir o BP.
3. Revise os valores das colunas J e K e as linhas destacadas.
## Cuidados
- **Faça uma cópia dos dados antes de executar:** as macros inserem colunas, alteram descrições e convertem valores negativos em positivos.
- Não execute novamente sobre um relatório já tratado.
- O cruzamento utiliza `XLOOKUP`, sem somar descrições repetidas.
- As diferenças são comparadas diretamente com zero, sem arredondamento.
