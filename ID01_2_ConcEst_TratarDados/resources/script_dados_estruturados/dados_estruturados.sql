INSERT INTO tbl_dados_estruturados (
    "Nomenclatura SoftCase", adquirente, empresa, SIGLA, data_processamento,
    forma_pagto, Dia_Comp, bandeira, valor_lancamento, valor_taxa,
    taxa_adquirente, data_atualizacao, ID_UNICO, "Razao Social", CNPJ,
    ultima_atualizacao
)

WITH base AS (
    SELECT 
        emp."Nomenclatura SoftCase" AS nomenclatura_softcase,
        aux.adquirente,
        aux.empresa,
        emp.SIGLA,
        aux.data_processamento,
        aux.forma_pagto AS forma_pagto_raw,
        CASE
            WHEN aux.adquirente = 'Veloe' AND aux.forma_pagto = 'CREDITO' THEN 'TAG'
            WHEN aux.forma_pagto IS NULL
                 AND aux.adquirente IN ('ConectCar', 'Greenpass', 'SemParar', 'Veloe')
            THEN 'TAG'
            WHEN aux.forma_pagto IN ('Débito à vista', 'Débito pré-pago') THEN 'DEBITO'
            WHEN aux.adquirente = 'Bradesco'
                 AND aux.forma_pagto = 'CRED PIX QR CODE DINAMIC' THEN 'PIX'
            ELSE aux.forma_pagto
        END AS forma_pagto,
        CASE
            WHEN aux.forma_pagto = 'Crédito à vista' THEN 31
            WHEN aux.forma_pagto = 'Crédito pré-pago' THEN 2
            WHEN aux.forma_pagto = 'Crédito conversor de moedas' THEN 5
            ELSE NULL
        END AS Dia_Comp,
        aux.bandeira,
        aux.valor_lancamento,
        ABS(aux.valor_taxa) AS valor_taxa,
        CASE
            WHEN aux.adquirente = 'Cielo'
                THEN REPLACE(printf('%.2f', CAST(REPLACE(tcielo.Taxa, ',', '.') AS REAL) * 100), '.', ',')
            WHEN aux.adquirente = 'ConectCar' THEN REPLACE(printf('%.2f', CAST(REPLACE(tax.CONECTCAR, ',', '.') AS REAL) * 100), '.', ',')
            WHEN aux.adquirente = 'Veloe' THEN REPLACE(printf('%.2f', CAST(REPLACE(tax.VELOE, ',', '.') AS REAL) * 100), '.', ',')
            WHEN aux.adquirente = 'Greenpass' THEN REPLACE(printf('%.2f', CAST(REPLACE(tax.GREENPASS, ',', '.') AS REAL) * 100), '.', ',')
            WHEN aux.adquirente = 'SemParar' THEN REPLACE(printf('%.2f', CAST(REPLACE(tax.SEMPARAR, ',', '.') AS REAL) * 100), '.', ',')
            ELSE NULL
        END AS taxa_adquirente,
        aux.data_atualizacao,
        emp.ID_UNICO,
        emp."Razao Social" AS razao_social,
        emp.CNPJ
    FROM tbl_empresas AS emp
    INNER JOIN tbl_aux_dados AS aux
        ON aux.empresa = emp.Cielo
        OR aux.empresa = emp.Veloe
        OR aux.empresa = emp.SemParar
        OR aux.empresa = emp.Greenpass
        OR aux.empresa = emp.ConectCar
        OR aux.empresa = emp.Bradesco
    LEFT JOIN tbl_TaxaAdquirentes AS tax
        ON tax."Sigla Empresa" = emp.SIGLA
    LEFT JOIN tbl_TaxaCielo AS tcielo
        ON aux.adquirente = 'Cielo'
        AND tcielo."TAXAS CIELO" = CASE
            WHEN aux.forma_pagto IN ('Débito à vista', 'Débito pré-pago') THEN 'DÉBITO'
            WHEN aux.forma_pagto IN ('Crédito à vista', 'Crédito pré-pago') THEN 'CRÉDITO'
            WHEN aux.forma_pagto = 'Pix' THEN 'PIX'
            ELSE aux.forma_pagto
        END
        AND UPPER(tcielo.Bandeira) = UPPER(aux.bandeira)
    WHERE NOT (aux.adquirente = 'Veloe' AND aux.forma_pagto = 'DEBITO')
),
pix_pairs AS (
    SELECT empresa, SIGLA, data_processamento
    FROM base
    WHERE adquirente = 'Bradesco'
      AND forma_pagto_raw IN ('CRED PIX QR CODE DINAMIC', 'DEVOLUCOES PIX')
    GROUP BY empresa, SIGLA, data_processamento
    HAVING COUNT(DISTINCT forma_pagto_raw) = 2
),
pix_combined AS (
    SELECT 
        cred.nomenclatura_softcase,
        cred.adquirente,
        cred.empresa,
        cred.SIGLA,
        cred.data_processamento,
        'PIX' AS forma_pagto,
        cred.Dia_Comp,
        cred.bandeira,
        (cred.valor_lancamento + dev.valor_lancamento) AS valor_lancamento,
        cred.valor_taxa,
        cred.taxa_adquirente,
        cred.data_atualizacao,
        cred.ID_UNICO,
        cred.razao_social,
        cred.CNPJ
    FROM base cred
    JOIN base dev
        ON cred.empresa = dev.empresa
       AND cred.SIGLA = dev.SIGLA
       AND cred.data_processamento = dev.data_processamento
       AND cred.forma_pagto_raw = 'CRED PIX QR CODE DINAMIC'
       AND dev.forma_pagto_raw = 'DEVOLUCOES PIX'
    WHERE (cred.empresa, cred.SIGLA, cred.data_processamento) IN (SELECT empresa, SIGLA, data_processamento FROM pix_pairs)
)
SELECT 
    nomenclatura_softcase, adquirente, empresa, SIGLA, data_processamento,
    forma_pagto, Dia_Comp, bandeira, valor_lancamento, valor_taxa,
    taxa_adquirente, data_atualizacao, ID_UNICO, razao_social, CNPJ,
    datetime('now', 'localtime') AS ultima_atualizacao
FROM base
WHERE NOT (
    adquirente = 'Bradesco'
    AND forma_pagto_raw IN ('CRED PIX QR CODE DINAMIC', 'DEVOLUCOES PIX')
    AND (empresa, SIGLA, data_processamento) IN (SELECT empresa, SIGLA, data_processamento FROM pix_pairs)
)
UNION ALL
SELECT 
    nomenclatura_softcase, adquirente, empresa, SIGLA, data_processamento,
    forma_pagto, Dia_Comp, bandeira, valor_lancamento, valor_taxa,
    taxa_adquirente, data_atualizacao, ID_UNICO, razao_social, CNPJ,
    datetime('now', 'localtime') AS ultima_atualizacao
FROM pix_combined;