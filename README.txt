SISTEMA DE RELATÓRIOS PEDAGÓGICOS - VERSÃO INICIAL PARA RENDER

1. COMO RODAR LOCALMENTE

No terminal, dentro da pasta do projeto:

python3 -m pip install -r requirements.txt
python3 app.py

Abra no navegador:
http://127.0.0.1:5000


2. COMO SUBIR PARA O GITHUB

- Crie um repositório no GitHub
- Envie os arquivos deste projeto
- Faça o push


3. COMO PUBLICAR NO RENDER

- Entre no Render
- Clique em New +
- Escolha Web Service
- Conecte seu repositório do GitHub
- O Render vai ler o arquivo render.yaml
- Faça o deploy


4. OBSERVAÇÃO IMPORTANTE

Esta primeira versão usa listas em memória.
Isso significa que, ao reiniciar a aplicação, os dados serão perdidos.

Para a próxima versão, o ideal é migrar para PostgreSQL no Render.
