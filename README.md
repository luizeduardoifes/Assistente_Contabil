# Automação Nassau

Assistente de automação contábil para conciliar títulos e conferir os valores correspondentes. Ao final do processo, gera um relatório em Excel com os valores conciliados.

## Funcionalidades

- Conciliação de títulos.
- Conferência de títulos e valores.
- Geração de relatório em Excel com os resultados da conciliação.

## Estrutura do projeto

- `app.py`: entrada principal da aplicação e referência para acompanhar como o fluxo é iniciado.
- Demais arquivos e pastas: módulos, recursos e configurações utilizados pelo fluxo, conforme as dependências chamadas por `app.py`.

## Como entender o fluxo

Para conhecer a execução, comece por `app.py` e siga as importações e chamadas feitas por ele. Assim, é possível identificar como a aplicação realiza a conciliação, confere os dados e gera o relatório.
