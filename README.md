# Gerador de Certificado de Visitação

Este projeto é um gerador de certificados de visitação que permite preencher um modelo de documento Word com informações personalizadas e enviar o documento por e-mail com uma assinatura digital. O aplicativo foi desenvolvido em Python utilizando a biblioteca `tkinter` para a interface gráfica, `python-docx` para manipulação de documentos do Word e `Pillow` para suporte a imagens.

## Funcionalidades

- Preenchimento automático de um modelo de certificado.
- Inserção de uma imagem de assinatura no documento.
- Envio do certificado gerado por e-mail como anexo.
- Interface gráfica amigável para facilitar o uso.
- Opções para ajustar o tamanho da fonte e capturar a assinatura.

## Requisitos

Para executar este projeto, você precisa ter o Python instalado em sua máquina. Além disso, instale as seguintes dependências:

- `python-docx`
- `Pillow`
- `tkinter` (geralmente incluído na instalação do Python)

### Instalação de dependências

Você pode instalar as dependências necessárias usando o `pip`. Execute o seguinte comando:

```bash
pip install python-docx Pillow
