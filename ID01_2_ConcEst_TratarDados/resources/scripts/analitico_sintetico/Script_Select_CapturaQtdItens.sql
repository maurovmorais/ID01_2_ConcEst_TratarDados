
WITH RankedItems AS (
    SELECT 
        id_item,
        status,
        tipo_excecao,
        DENSE_RANK() OVER (
            PARTITION BY id_execucao, id_item_fila
            ORDER BY ultima_atualizacao DESC
        ) AS ranks
    FROM <tabela_dados_itens>
    WHERE id_execucao = ?
),
FilteredItems AS (
    SELECT 
        id_item,
        status,
        tipo_excecao
    FROM RankedItems
    WHERE ranks = 1
)
SELECT 
    COUNT(CASE WHEN UPPER(status) = 'SUCESSO' THEN 1 END) AS qtd_itens_sucesso,
    COUNT(CASE WHEN UPPER(tipo_excecao) = 'NEGOCIO' THEN 1 END) AS qtd_itens_business,
    COUNT(CASE WHEN UPPER(tipo_excecao) = 'SISTEMA' THEN 1 END) AS qtd_itens_app
FROM FilteredItems;