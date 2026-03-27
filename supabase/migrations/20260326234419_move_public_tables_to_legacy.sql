CREATE SCHEMA IF NOT EXISTS legacy;

DO $$
DECLARE
    table_row record;
    grant_row record;
BEGIN
    FOR table_row IN
        SELECT tablename
        FROM pg_tables
        WHERE schemaname = 'public'
          AND tablename <> 'Bizerba_ID'
        ORDER BY tablename
    LOOP
        EXECUTE format('ALTER TABLE public.%I SET SCHEMA legacy', table_row.tablename);

        EXECUTE format(
            'CREATE OR REPLACE VIEW public.%I WITH (security_invoker = true) AS SELECT * FROM legacy.%I',
            table_row.tablename,
            table_row.tablename
        );

        FOR grant_row IN
            SELECT
                grantee,
                string_agg(privilege_type, ', ' ORDER BY privilege_type) AS privileges
            FROM information_schema.role_table_grants
            WHERE table_schema = 'legacy'
              AND table_name = table_row.tablename
              AND privilege_type IN ('SELECT', 'INSERT', 'UPDATE', 'DELETE')
            GROUP BY grantee
        LOOP
            IF grant_row.grantee = 'PUBLIC' THEN
                EXECUTE format(
                    'GRANT %s ON TABLE public.%I TO PUBLIC',
                    grant_row.privileges,
                    table_row.tablename
                );
            ELSE
                EXECUTE format(
                    'GRANT %s ON TABLE public.%I TO %I',
                    grant_row.privileges,
                    table_row.tablename,
                    grant_row.grantee
                );
            END IF;
        END LOOP;
    END LOOP;
END $$;

DO $$
DECLARE
    role_row record;
BEGIN
    FOR role_row IN
        SELECT DISTINCT grantee
        FROM information_schema.role_table_grants
        WHERE table_schema = 'legacy'
    LOOP
        IF role_row.grantee = 'PUBLIC' THEN
            EXECUTE 'GRANT USAGE ON SCHEMA legacy TO PUBLIC';
        ELSE
            EXECUTE format('GRANT USAGE ON SCHEMA legacy TO %I', role_row.grantee);
        END IF;
    END LOOP;
END $$;
