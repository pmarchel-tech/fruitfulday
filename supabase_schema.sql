-- Relational Schema for migrating FruitfulDay (DailyPulse) from Firestore to Supabase (PostgreSQL)

-- Enable extension for uuid generation if needed
CREATE EXTENSION IF NOT EXISTS "uuid-ossp";

-- ==========================================
-- TABLE DEFINITIONS
-- ==========================================

-- 1. PROFILES (formerly 'users' collection)
CREATE TABLE IF NOT EXISTS public.profiles (
    id TEXT PRIMARY KEY, -- Using TEXT to support legacy Firebase UIDs and Supabase Auth string IDs
    username TEXT UNIQUE NOT NULL,
    pin TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('ADMIN', 'MEMBER')) DEFAULT 'MEMBER',
    created_at TIMESTAMPTZ DEFAULT now() NOT NULL
);

-- 2. TEAM MEMBERS (replacing teamMemberIds array in user documents)
CREATE TABLE IF NOT EXISTS public.team_members (
    admin_id TEXT REFERENCES public.profiles(id) ON DELETE CASCADE,
    member_id TEXT REFERENCES public.profiles(id) ON DELETE CASCADE,
    PRIMARY KEY (admin_id, member_id)
);

-- 3. TAGS
CREATE TABLE IF NOT EXISTS public.tags (
    id TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
    name TEXT UNIQUE NOT NULL
);

-- 4. TASKS
CREATE TABLE IF NOT EXISTS public.tasks (
    id TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
    user_id TEXT REFERENCES public.profiles(id) ON DELETE CASCADE NOT NULL,
    title TEXT NOT NULL,
    category TEXT NOT NULL,
    status TEXT NOT NULL CHECK (status IN ('NOT_YET', 'PROGRESS', 'DONE', 'CANCEL', 'PENDING', 'REPETITIVE', 'FOLLOW_UP')) DEFAULT 'NOT_YET',
    target_date DATE NOT NULL,
    created_at TIMESTAMPTZ DEFAULT now() NOT NULL,
    created_at_ms BIGINT -- Stores Firestore raw epoch ms timestamp for backward compatibility
);

-- 5. TASK TAGS (replacing tags array in task documents)
CREATE TABLE IF NOT EXISTS public.task_tags (
    task_id TEXT REFERENCES public.tasks(id) ON DELETE CASCADE,
    tag_id TEXT REFERENCES public.tags(id) ON DELETE CASCADE,
    PRIMARY KEY (task_id, tag_id)
);

-- 6. TASK UPDATES (formerly 'task_updates' collection)
CREATE TABLE IF NOT EXISTS public.task_updates (
    id TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
    task_id TEXT REFERENCES public.tasks(id) ON DELETE CASCADE NOT NULL,
    user_id TEXT REFERENCES public.profiles(id) ON DELETE SET NULL,
    date DATE NOT NULL,
    content TEXT NOT NULL,
    timestamp TIMESTAMPTZ DEFAULT now() NOT NULL,
    timestamp_ms BIGINT, -- Stores Firestore raw epoch ms timestamp
    is_archived BOOLEAN DEFAULT FALSE NOT NULL,
    status_from TEXT CHECK (status_from IN ('NOT_YET', 'PROGRESS', 'DONE', 'CANCEL', 'PENDING', 'REPETITIVE', 'FOLLOW_UP')),
    status_to TEXT CHECK (status_to IN ('NOT_YET', 'PROGRESS', 'DONE', 'CANCEL', 'PENDING', 'REPETITIVE', 'FOLLOW_UP'))
);

-- 7. AI TOKEN RECORDS (formerly 'ai_token_records' collection)
CREATE TABLE IF NOT EXISTS public.ai_token_records (
    id TEXT PRIMARY KEY DEFAULT gen_random_uuid()::text,
    user_id TEXT REFERENCES public.profiles(id) ON DELETE CASCADE NOT NULL,
    date DATE NOT NULL,
    timestamp TIMESTAMPTZ DEFAULT now() NOT NULL,
    timestamp_ms BIGINT, -- Stores Firestore raw epoch ms timestamp
    used_for TEXT NOT NULL,
    input_tokens INTEGER NOT NULL,
    output_tokens INTEGER NOT NULL
);

-- ==========================================
-- ROW LEVEL SECURITY (RLS) POLICIES
-- ==========================================

-- Enable RLS on all tables
ALTER TABLE public.profiles ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.team_members ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.tags ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.tasks ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.task_tags ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.task_updates ENABLE ROW LEVEL SECURITY;
ALTER TABLE public.ai_token_records ENABLE ROW LEVEL SECURITY;

-- Helper Function: Check if the current user is an Admin (Optimized STABLE with JWT role lookup)
CREATE OR REPLACE FUNCTION public.is_admin()
RETURNS BOOLEAN SECURITY DEFINER STABLE AS $$
BEGIN
    -- Check JWT metadata first for maximum speed (zero database queries)
    IF coalesce(auth.jwt() -> 'user_metadata' ->> 'role', '') = 'ADMIN' THEN
        RETURN TRUE;
    END IF;
    IF auth.jwt()->>'email' = 'pmarchel@gmail.com' THEN
        RETURN TRUE;
    END IF;
    -- Fallback to database check
    RETURN EXISTS (
        SELECT 1 FROM public.profiles
        WHERE id = auth.uid()::text AND role = 'ADMIN'
    );
END;
$$ LANGUAGE plpgsql;

-- Helper Function: Check if a member is in an admin's team (Optimized STABLE with short-circuit check)
CREATE OR REPLACE FUNCTION public.is_team_member(member_id TEXT)
RETURNS BOOLEAN SECURITY DEFINER STABLE AS $$
BEGIN
    -- If the current user is not an admin, they cannot have team members (skip database scan)
    IF coalesce(auth.jwt() -> 'user_metadata' ->> 'role', '') != 'ADMIN' THEN
        RETURN auth.uid()::text = member_id;
    END IF;

    RETURN auth.uid()::text = member_id OR EXISTS (
        SELECT 1 FROM public.team_members
        WHERE admin_id = auth.uid()::text AND member_id = is_team_member.member_id
    );
END;
$$ LANGUAGE plpgsql;

-- 1. Profiles Policies
CREATE POLICY "Allow authenticated users read access to profiles"
ON public.profiles FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "Allow users to update their own profile"
ON public.profiles FOR UPDATE
TO authenticated
USING (auth.uid()::text = id)
WITH CHECK (auth.uid()::text = id);

CREATE POLICY "Allow profile creation for authenticated users"
ON public.profiles FOR INSERT
TO authenticated
WITH CHECK (auth.uid()::text = id OR public.is_admin());

-- 2. Team Members Policies
CREATE POLICY "Allow authenticated users to read team members"
ON public.team_members FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "Allow admins to manage team members"
ON public.team_members FOR ALL
TO authenticated
USING (public.is_admin());

-- 3. Tags Policies
CREATE POLICY "Allow authenticated users to read tags"
ON public.tags FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "Allow authenticated users to insert tags"
ON public.tags FOR INSERT
TO authenticated
WITH CHECK (true);

-- 4. Tasks Policies
CREATE POLICY "Allow users to read their own tasks and admins to read all"
ON public.tasks FOR SELECT
TO authenticated
USING (auth.uid()::text = user_id OR public.is_admin());

CREATE POLICY "Allow users to insert their own tasks"
ON public.tasks FOR INSERT
TO authenticated
WITH CHECK (auth.uid()::text = user_id OR public.is_admin());

CREATE POLICY "Allow users to update their own tasks and admins to update all"
ON public.tasks FOR UPDATE
TO authenticated
USING (auth.uid()::text = user_id OR public.is_admin())
WITH CHECK (auth.uid()::text = user_id OR public.is_admin());

CREATE POLICY "Allow users to delete their own tasks and admins to delete all"
ON public.tasks FOR DELETE
TO authenticated
USING (auth.uid()::text = user_id OR public.is_admin());

-- 5. Task Tags Policies
CREATE POLICY "Allow authenticated users to read task tags"
ON public.task_tags FOR SELECT
TO authenticated
USING (true);

CREATE POLICY "Allow users to manage task tags for their own tasks"
ON public.task_tags FOR ALL
TO authenticated
USING (
    EXISTS (
        SELECT 1 FROM public.tasks
        WHERE id = task_id AND (user_id = auth.uid()::text OR public.is_admin())
    )
);

-- 6. Task Updates Policies
CREATE POLICY "Allow users to read updates of tasks they have access to"
ON public.task_updates FOR SELECT
TO authenticated
USING (
    EXISTS (
        SELECT 1 FROM public.tasks
        WHERE tasks.id = task_id
    )
);

CREATE POLICY "Allow users and admins to create updates"
ON public.task_updates FOR INSERT
TO authenticated
WITH CHECK (user_id = auth.uid()::text OR public.is_admin());

CREATE POLICY "Allow users and admins to update updates"
ON public.task_updates FOR UPDATE
TO authenticated
USING (user_id = auth.uid()::text OR public.is_admin())
WITH CHECK (user_id = auth.uid()::text OR public.is_admin());

CREATE POLICY "Allow users and admins to delete updates"
ON public.task_updates FOR DELETE
TO authenticated
USING (user_id = auth.uid()::text OR public.is_admin());

-- 7. AI Token Records Policies
CREATE POLICY "Allow users to read their own token records and admins to read all"
ON public.ai_token_records FOR SELECT
TO authenticated
USING (user_id = auth.uid()::text OR public.is_admin());

CREATE POLICY "Allow users to insert their own token records"
ON public.ai_token_records FOR INSERT
TO authenticated
WITH CHECK (user_id = auth.uid()::text);

-- ==========================================
-- AUTOMATIC PROFILE CREATION TRIGGER
-- ==========================================
-- Automatically inserts a record into public.profiles when a new user signs up in Supabase Auth

CREATE OR REPLACE FUNCTION public.handle_new_user()
RETURNS TRIGGER AS $$
BEGIN
    INSERT INTO public.profiles (id, username, pin, role)
    VALUES (
        new.id::text,
        coalesce(new.raw_user_meta_data->>'username', split_part(new.email, '@', 1)),
        coalesce(new.raw_user_meta_data->>'pin', '1234'), -- Default/fallback pin
        coalesce(new.raw_user_meta_data->>'role', 'MEMBER')
    )
    ON CONFLICT (id) DO NOTHING;
    RETURN NEW;
END;
$$ LANGUAGE plpgsql SECURITY DEFINER;

-- Trigger to run the function on auth.users insert
CREATE OR REPLACE TRIGGER on_auth_user_created
    AFTER INSERT ON auth.users
    FOR EACH ROW EXECUTE FUNCTION public.handle_new_user();

-- ==========================================
-- PERFORMANCE INDEXES (Optimizes queries and RLS policy scans)
-- ==========================================
CREATE INDEX IF NOT EXISTS idx_tasks_user_id ON public.tasks(user_id);
CREATE INDEX IF NOT EXISTS idx_task_updates_task_id ON public.task_updates(task_id);
CREATE INDEX IF NOT EXISTS idx_task_updates_user_id ON public.task_updates(user_id);
CREATE INDEX IF NOT EXISTS idx_task_tags_tag_id ON public.task_tags(tag_id);
CREATE INDEX IF NOT EXISTS idx_ai_token_records_user_id ON public.ai_token_records(user_id);
CREATE INDEX IF NOT EXISTS idx_team_members_member_id ON public.team_members(member_id);
CREATE INDEX IF NOT EXISTS idx_task_updates_timestamp_ms ON public.task_updates(timestamp_ms DESC NULLS LAST);
CREATE INDEX IF NOT EXISTS idx_ai_token_records_timestamp_ms ON public.ai_token_records(timestamp_ms DESC NULLS LAST);
