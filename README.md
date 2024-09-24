# Blood Antigen Bot

Este bot automatiza o preenchimento de dados em um sistema, lendo informações de um arquivo Excel e preenchendo campos específicos na tela com o uso de `pyautogui`.

## Índice

- [Visão Geral](#visão-geral)
- [Instalação](#instalação)
- [Como Usar](#como-usar)

## Visão Geral

O Blood Antigen Bot é um script automatizado que lê dados de um arquivo Excel e insere essas informações em campos específicos de um sistema de preenchimento de antígenos sanguíneos. O bot é útil para evitar inserções manuais repetitivas, garantindo consistência e eficiência.

## Instalação

### Pré-requisitos

Certifique-se de ter o Python 3.x instalado na sua máquina e as seguintes bibliotecas Python:

- `pyautogui`: para controlar o mouse e o teclado.
- `openpyxl`: para manipular arquivos Excel.

Instale as bibliotecas necessárias com o seguinte comando:

pip install pyautogui openpyxl

### Como usar

1. Clonar o Repositório:

   Clone este repositório na sua máquina local usando o seguinte comando:

   git clone https://github.com/seu-usuario/blood-antigen-bot.git

2. Configuração do Arquivo Excel:

   Coloque o arquivo Excel com os dados na pasta do projeto. O caminho padrão esperado é C:\project\bot-antigenos_sanguineos\Tableofbloodgroupsystems.xlsx.

   O arquivo Excel deve ter 8 colunas, com as informações que serão preenchidas no sistema.

3. Defina as Coordenadas do Sistema:

   Verifique as coordenadas do sistema onde os dados serão inseridos e ajuste no código, se necessário.

   Use pyautogui.position() para capturar as coordenadas corretas ao posicionar o mouse sobre o campo desejado.

4. Execute o Script:

Execute o script com o seguinte comando:

python bot.py

   O bot esperará 3 segundos para que você abra e prepare o sistema antes de começar a preencher os dados.