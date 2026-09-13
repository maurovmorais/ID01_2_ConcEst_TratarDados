UPDATE <tabela_dados_execucao>
SET 
    fim_exec = ?,
    qtd_itens_sucesso = ?,
    qtd_itens_business = ?,
    qtd_itens_app = ?,
    tempo_execucao = ?,
    excecao_inicializacao = COALESCE(?, excecao_inicializacao),
    tipo_excecao_inicializacao = COALESCE(?, tipo_excecao_inicializacao),
    screenshot_excecao = COALESCE(?, screenshot_excecao)
WHERE id_execucao = ?;
