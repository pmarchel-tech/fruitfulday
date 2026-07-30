-- Fix database profiles and align task data with correct Auth IDs
-- Paste this script into your Supabase Dashboard SQL Editor and click Run.

DO $$
DECLARE
    u_rec RECORD;
    p_rec RECORD;
    v_username TEXT;
BEGIN
    -- Loop through all users in auth.users
    FOR u_rec IN SELECT id, email, raw_user_meta_data FROM auth.users LOOP
        -- Resolve username from metadata or email prefix
        v_username := coalesce(u_rec.raw_user_meta_data->>'username', split_part(u_rec.email, '@', 1));
        
        -- Check if there is an existing profile with the same username but a different ID
        SELECT * INTO p_rec FROM public.profiles 
        WHERE username = v_username AND id <> u_rec.id::text;
        
        IF FOUND THEN
            RAISE NOTICE 'Migrating user % from old ID % to new ID %', v_username, p_rec.id, u_rec.id;
            
            -- Rename old profile temporarily to avoid unique username constraint during insertion
            UPDATE public.profiles SET username = username || '_old_migrated' WHERE id = p_rec.id;
            
            -- Insert/upsert the new profile using the correct Auth ID
            INSERT INTO public.profiles (id, username, pin, role)
            VALUES (u_rec.id::text, p_rec.username, p_rec.pin, p_rec.role)
            ON CONFLICT (id) DO UPDATE 
            SET username = p_rec.username, pin = p_rec.pin, role = p_rec.role;
            
            -- Update all related records in other tables to the new Auth ID
            UPDATE public.tasks SET user_id = u_rec.id::text WHERE user_id = p_rec.id;
            UPDATE public.task_updates SET user_id = u_rec.id::text WHERE user_id = p_rec.id;
            UPDATE public.team_members SET admin_id = u_rec.id::text WHERE admin_id = p_rec.id;
            UPDATE public.team_members SET member_id = u_rec.id::text WHERE member_id = p_rec.id;
            UPDATE public.ai_token_records SET user_id = u_rec.id::text WHERE user_id = p_rec.id;
            
            -- Clean up the old profile record
            DELETE FROM public.profiles WHERE id = p_rec.id;
        ELSE
            -- Ensure a profile exists for the user
            INSERT INTO public.profiles (id, username, pin, role)
            VALUES (
                u_rec.id::text, 
                v_username, 
                coalesce(u_rec.raw_user_meta_data->>'pin', '1234'), 
                coalesce(u_rec.raw_user_meta_data->>'role', 'MEMBER')
            )
            ON CONFLICT (id) DO NOTHING;
        END IF;
    END LOOP;
END $$;
