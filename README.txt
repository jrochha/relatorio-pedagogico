SISTEMA DE RELATÓRIOS PEDAGÓGICOS - VERSÃO 2

Esta versão já inclui:
- lista de estudantes importada da planilha
- exibição automática de turma e responsável
- lista suspensa de disciplinas
- opções em checkbox para aprendizagem, comportamento, intervenções e encaminhamentos
- geração automática do relatório formal

COMO RODAR LOCALMENTE

python3 -m pip install -r requirements.txt
python3 app.py

Abra:
http://127.0.0.1:5000

OBSERVAÇÃO
Esta versão ainda usa registros em memória. Se o Render reiniciar, os registros serão perdidos.
Na próxima etapa, o ideal é migrar os registros para PostgreSQL.
Foram carregados 79 estudantes da planilha.


ATUALIZAÇÃO V3
- Campo de data do registro
- Campo de hora do registro
- Tipo de relatório com duas modalidades:
  * Relatório de Acompanhamento Pedagógico
  * Relatório de Ocorrência Disciplinar


ATUALIZAÇÃO V4
- Suporte visual para logo da escola em tela e nos arquivos exportados
- Download do relatório em TXT, DOCX e PDF
- Campos de assinatura do estudante, professor/pedagogo e responsável
- Para usar a logo oficial, substitua o arquivo static/logo_escola.png pela imagem da escola
