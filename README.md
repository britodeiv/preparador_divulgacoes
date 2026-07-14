# Observação

Este projeto foi desenvolvido para automatizar um processo interno específico, considerando regras de negócio, estrutura de dados e fluxo de trabalho próprios do ambiente em que foi utilizado.

Embora o código seja totalmente funcional, sua utilização em outros contextos pode exigir adaptações conforme as necessidades de cada organização.

Dentro do cenário para o qual foi desenvolvido, a ferramenta desempenha sua finalidade com eficiência, estabilidade e confiabilidade.


# Preparador de Divulgações

Aplicativo desktop para Windows que limpa listas de e-mail, remove duplicados, compara com uma base de bloqueio e prepara lotes em CCO no Outlook clássico.

**Desenvolvido por Deivid Brito**

## Funcionalidades

- Lê uma planilha consolidada da campanha (`.xlsx` ou `.xlsm`).
- Converte endereços para minúsculas.
- Remove espaços, tabulações, aspas, `http`, `www`, `mailto` e resíduos comuns.
- Corrige alguns erros simples de domínio.
- Separa e descarta endereços inválidos.
- Remove duplicados depois da normalização.
- Compara exclusivamente com a aba `Resumo` da planilha de bloqueados.
- Permite configurar um endereço adicional para excluir em campanhas especiais.
- Divide a lista final em lotes de até 600 destinatários.
- Preenche os destinatários somente em CCO.
- Reaproveita assunto, corpo, imagens e anexos do e-mail selecionado no Outlook clássico.
- Gera relatório Excel, lista TXT e resumo final.

## Privacidade

Este repositório não contém:

- listas reais de destinatários;
- endereços institucionais reais;
- arquivos de campanhas;
- credenciais;
- dados pessoais;
- relatórios produzidos em ambiente de trabalho.

Os arquivos da pasta `examples` usam dados inteiramente fictícios.

## Requisitos para desenvolvimento

- Windows 10 ou 11
- Python 3.11 ou superior
- Outlook clássico instalado e configurado

Instale as dependências:

```bash
python -m pip install -r requirements.txt
```

Execute:

```bash
python main.py
```

## Gerar executável

No Windows, execute:

```text
1_GERAR_EXECUTAVEL.bat
```

O executável será criado em:

```text
PARA_ENVIAR\Preparador_de_Divulgacoes.exe
```

O usuário final não precisa ter Python instalado quando o programa é empacotado com `--onefile`.

## Formato esperado

### Lista da campanha

Pode conter endereços em qualquer célula. É recomendável usar uma planilha consolidada contendo somente os contatos selecionados para a campanha.

### Lista de bloqueados

Deve conter uma aba chamada:

```text
Resumo
```

O programa usa apenas essa aba para remover endereços que não devem receber novas mensagens.

## Segurança operacional

O programa não envia e-mails automaticamente. Ele abre cada mensagem pronta no Outlook para revisão manual.

## Aviso

Faça testes com poucos destinatários antes de usar em produção. Políticas de envio, limites e permissões podem variar conforme a organização e a configuração do Outlook.

## Licença

MIT.

