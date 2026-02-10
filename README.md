## Automacao DCTF (e-CAC) - Guia Simples

Este sistema ajuda a baixar guias DARF (DCTFWeb) no e-CAC para varios CNPJs usando uma planilha Excel.

## Antes de comecar

- Este projeto **nao faz login automatico** no e-CAC.
- O login sempre sera feito por voce, manualmente, no navegador.
- Use apenas conforme as regras da Receita e da sua empresa.

## O que voce precisa

- Windows
- Google Chrome instalado
- Python 3.11 ou superior
- Acesso ao e-CAC

## Instalacao (primeira vez)

Abra o terminal na pasta do projeto e rode os comandos abaixo:

```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

Se der erro de permissao, feche o terminal e abra novamente como administrador.

## Como abrir o sistema

No terminal (dentro da pasta do projeto):

```bash
venv\Scripts\activate
python main.py
```

Isso abre a tela principal (interface grafica).

## Como preparar a planilha

A planilha deve ser `.xlsx` e ter estas colunas:

- `CNPJ` (obrigatoria)
- `COD` (obrigatoria)
- `STATUS` (opcional)

Colunas de nome da empresa sao opcionais:
`NOME`, `RAZAO`, `RAZAO_SOCIAL`, `RAZAO SOCIAL`, `EMPRESA`.

## Passo a passo de uso (na tela)

1. Clique em **Selecionar** e escolha a planilha.
2. Clique em **Carregar Dados**.
3. Confira as configuracoes (datas, competencia, timeout e tentativas).
4. Clique em **Iniciar Automacao**.
5. Quando o navegador abrir, faca o login no e-CAC.
6. Volte para o sistema e clique em **Confirmar Login**.
7. Aguarde o processamento terminar.

## Onde ficam os resultados

- PDFs baixados: pasta `Competencias executadas/` (organizados por competencia)
- Log de execucao: arquivo `AUTOMACAO-DCTF.log`
- Configuracoes salvas: arquivo `config.json`

## Erros comuns e como resolver

### 1) "No module named ... "
Faltam dependencias. Rode:

```bash
venv\Scripts\activate
pip install -r requirements.txt
```

### 2) Chrome nao abre ou fecha sozinho

- Feche todas as janelas do Chrome e tente novamente.
- Atualize o Chrome para a versao mais recente.
- Reinicie o computador, se necessario.

### 3) "Arquivo nao encontrado" / "Nenhum CNPJ encontrado"

- Verifique se escolheu a planilha correta.
- Confirme se as colunas `CNPJ` e `COD` existem e estao escritas exatamente assim.

### 4) O sistema para durante a execucao

- Veja a mensagem no log da tela.
- Abra o arquivo `AUTOMACAO-DCTF.log` para mais detalhes.
- Tente novamente com timeout maior.

## Comandos uteis

- Abrir interface: `python main.py`
- Modo texto (avancado): `python main.py --cli`
- Ver ajuda: `python main.py --help`

## Suporte interno

Ao pedir suporte para TI, envie:

- print da tela com o erro
- trecho final do `AUTOMACAO-DCTF.log`
- planilha usada (sem dados sensiveis, se possivel)


