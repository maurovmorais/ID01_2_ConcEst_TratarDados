SELECT dif.data_hora_inicio,
dif.data_hora_fim,
dif.detalhes_item_fila,
dif.referencia,
dex.nome_maquina,
dif.status,
dif.descricao_excecao
FROM <tabela_dados_itens> as dif
INNER JOIN <tabela_dados_execucao> as dex
ON dif.id_execucao = dex.id_execucao
WHERE dex.guid_execucao = ?;
