SELECT nome_processo,
inicio_exec,
fim_exec,
tempo_execucao,
qtd_itens_sucesso+qtd_itens_business+qtd_itens_app 'soma_processados',
qtd_itens_sucesso,
qtd_itens_business,
qtd_itens_app,
nome_maquina
FROM <tabela_dados_execucao>
WHERE guid_execucao = ?;