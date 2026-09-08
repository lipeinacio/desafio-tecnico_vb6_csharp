-- Estrutura utilizada para demonstração

CREATE TABLE IF NOT EXISTS public.clientes (
    id INTEGER PRIMARY KEY,
    nome TEXT NOT NULL
);

-- Dados de exemplo

INSERT INTO public.clientes (id, nome)
VALUES
    (1, 'Ana Costa'),
    (2, 'Bruno Costa'),
    (3, 'Mariana Souza'),
    (4, 'João Lima')
ON CONFLICT (id) DO NOTHING;

-- Função para buscar clientes por parte do nome

CREATE OR REPLACE FUNCTION public.buscar_clientes(p_nome TEXT)
RETURNS TABLE (
    cliente_id INTEGER,
    nome_cliente TEXT
)
LANGUAGE sql
STABLE
AS $$
    SELECT
        c.id,
        c.nome
    FROM public.clientes AS c
    WHERE NULLIF(BTRIM(p_nome), '') IS NOT NULL
      AND c.nome ILIKE '%' || BTRIM(p_nome) || '%'
    ORDER BY c.id;
$$;


-- Exemplos de utilização

SELECT * FROM public.buscar_clientes('ana');
SELECT * FROM public.buscar_clientes('ANA');
SELECT * FROM public.buscar_clientes('Costa');
SELECT * FROM public.buscar_clientes('');