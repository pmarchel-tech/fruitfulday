import { Task, TaskUpdate, TaskStatus, User, UserRole, Tag, AiTokenRecord } from '../types';
import { supabase } from './supabaseClient';

const CURRENT_USER_KEY = 'dailypulse_current_user_id';

// --- AUTH SERVICE ---

export const getUsers = async (): Promise<User[]> => {
  try {
    const { data: profiles, error: profileErr } = await supabase
      .from('profiles')
      .select('*');
    
    if (profileErr) throw profileErr;

    const { data: teamMembers, error: teamErr } = await supabase
      .from('team_members')
      .select('*');
    
    if (teamErr) {
      console.error("Error fetching team members:", teamErr);
    }

    const teamMap = new Map<string, string[]>();
    if (teamMembers) {
      teamMembers.forEach(tm => {
        const list = teamMap.get(tm.admin_id) || [];
        list.push(tm.member_id);
        teamMap.set(tm.admin_id, list);
      });
    }

    return (profiles || []).map(p => ({
      id: p.id,
      username: p.username,
      pin: p.pin,
      role: p.role as UserRole,
      teamMemberIds: teamMap.get(p.id) || []
    }));
  } catch (error) {
    console.error('Error fetching users from Supabase:', error);
    return [];
  }
};

export const registerUser = async (username: string, email: string, pin: string, role: UserRole): Promise<User | null> => {
  try {
    const { data, error } = await supabase.auth.signUp({
      email,
      password: pin,
      options: {
        data: {
          username,
          pin,
          role
        }
      }
    });

    if (error) {
      throw new Error(error.message);
    }

    if (!data.user) throw new Error("Registration failed: no user returned");

    const newUser: User = {
      id: data.user.id,
      username,
      pin,
      role,
      teamMemberIds: []
    };

    localStorage.setItem(CURRENT_USER_KEY, data.user.id);
    return newUser;
  } catch (error: any) {
    console.error('Error registering user in Supabase:', error);
    throw error;
  }
};

export const loginUser = async (usernameOrEmail: string, pin: string): Promise<User | null> => {
  try {
    let email = usernameOrEmail;

    // If it's a username (doesn't contain '@'), resolve it to the registered email via RPC
    if (!usernameOrEmail.includes('@')) {
      const { data: resolvedEmail, error: rpcError } = await supabase
        .rpc('get_email_by_username', { p_username: usernameOrEmail });

      if (rpcError) {
        console.error("Error resolving email by username (RPC may not be created yet):", rpcError.message);
        // Fallback to legacy format as a last resort
        email = `${usernameOrEmail.toLowerCase()}@fruitfulday.local`;
      } else if (resolvedEmail) {
        email = resolvedEmail;
      } else {
        console.error("Username not found via RPC, trying legacy fallback:", usernameOrEmail);
        // Fallback to legacy format as a last resort
        email = `${usernameOrEmail.toLowerCase()}@fruitfulday.local`;
      }
    }

    const { data, error } = await supabase.auth.signInWithPassword({
      email,
      password: pin
    });

    if (error) {
      throw new Error(error.message);
    }

    if (!data.user) throw new Error("No user found in session");

    // Fetch matching user profile using maybeSingle to avoid PGRST116 error if missing
    let { data: profile, error: profileErr } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', data.user.id)
      .maybeSingle();

    if (profileErr) {
      console.error("Error fetching profile:", profileErr.message);
      throw new Error(profileErr.message);
    }

    if (!profile) {
      console.log("Profile missing. Dynamically creating profile for user:", data.user.id);
      // Extract from user metadata
      const username = data.user.user_metadata?.username || data.user.email?.split('@')[0] || 'User';
      const pin = data.user.user_metadata?.pin || '1234';
      const role = data.user.user_metadata?.role || 'MEMBER';

      const { data: newProfile, error: insertErr } = await supabase
        .from('profiles')
        .insert({
          id: data.user.id,
          username,
          pin,
          role
        })
        .select()
        .single();

      if (insertErr) {
        console.error("Error creating missing profile on-the-fly:", insertErr.message);
        throw new Error("Gagal membuat profil pengguna: " + insertErr.message);
      }
      profile = newProfile;
    }

    // Fetch team members
    const { data: team } = await supabase
      .from('team_members')
      .select('member_id')
      .eq('admin_id', data.user.id);

    const teamMemberIds = team ? team.map(t => t.member_id) : [];

    const loggedInUser: User = {
      id: profile.id,
      username: profile.username,
      pin: profile.pin,
      role: profile.role as UserRole,
      teamMemberIds
    };

    localStorage.setItem(CURRENT_USER_KEY, profile.id);
    return loggedInUser;
  } catch (error: any) {
    console.error('Error logging in user with Supabase:', error);
    throw error;
  }
};

export const getCurrentUser = async (): Promise<User | null> => {
  try {
    const { data: { session } } = await supabase.auth.getSession();
    const id = session?.user?.id || localStorage.getItem(CURRENT_USER_KEY);

    if (!id) return null;

    let { data: profile, error } = await supabase
      .from('profiles')
      .select('*')
      .eq('id', id)
      .maybeSingle();

    if (error) {
      console.error("Error fetching current user profile:", error.message);
      return null;
    }

    if (!profile && session?.user) {
      console.log("Profile missing in getCurrentUser. Creating on-the-fly:", id);
      const username = session.user.user_metadata?.username || session.user.email?.split('@')[0] || 'User';
      const pin = session.user.user_metadata?.pin || '1234';
      const role = session.user.user_metadata?.role || 'MEMBER';

      const { data: newProfile, error: insertErr } = await supabase
        .from('profiles')
        .insert({
          id,
          username,
          pin,
          role
        })
        .select()
        .single();

      if (insertErr) {
        console.error("Error creating profile on-the-fly in getCurrentUser:", insertErr.message);
        return null;
      }
      profile = newProfile;
    }

    if (!profile) return null;

    const { data: team } = await supabase
      .from('team_members')
      .select('member_id')
      .eq('admin_id', id);

    const teamMemberIds = team ? team.map(t => t.member_id) : [];

    localStorage.setItem(CURRENT_USER_KEY, id);
    return {
      id: profile.id,
      username: profile.username,
      pin: profile.pin,
      role: profile.role as UserRole,
      teamMemberIds
    };
  } catch (error) {
    console.error('Error fetching current user from Supabase:', error);
    return null;
  }
};

export const updateUser = async (updatedUser: User): Promise<boolean> => {
  try {
    // If the user changed their pin/password, let's also update it in Supabase Auth
    const { error: authError } = await supabase.auth.updateUser({
      password: updatedUser.pin
    });
    if (authError) {
      console.error('Error updating auth password in Supabase:', authError.message);
      return false;
    }

    const { error } = await supabase
      .from('profiles')
      .update({
        username: updatedUser.username,
        pin: updatedUser.pin,
        role: updatedUser.role
      })
      .eq('id', updatedUser.id);

    if (error) throw error;
    return true;
  } catch (error) {
    console.error('Error updating user in Supabase:', error);
    return false;
  }
};

export const logoutUser = async () => {
  localStorage.removeItem(CURRENT_USER_KEY);
  await supabase.auth.signOut();
};

export const addTeamMembers = async (adminId: string, memberIds: string[]): Promise<boolean> => {
  try {
    const records = memberIds.map(memberId => ({
      admin_id: adminId,
      member_id: memberId
    }));

    const { error } = await supabase
      .from('team_members')
      .insert(records);

    if (error) throw error;
    return true;
  } catch (error) {
    console.error('Error adding team members in Supabase:', error);
    return false;
  }
};

// --- DATA SERVICE ---

export const getTasks = async (user?: User): Promise<Task[]> => {
  try {
    let queryBuilder = supabase.from('tasks').select('*');

    if (user && user.role !== 'ADMIN') {
      queryBuilder = queryBuilder.eq('user_id', user.id);
    }

    const { data: tasksData, error: tasksErr } = await queryBuilder;
    if (tasksErr) throw tasksErr;

    // Fetch task_tags and tags in parallel using simple, fast queries
    const [taskTagsResult, tagsResult] = await Promise.all([
      supabase.from('task_tags').select('*'),
      supabase.from('tags').select('*')
    ]);

    const taskTags = taskTagsResult.data || [];
    const tags = tagsResult.data || [];

    // Map tag ID to tag name
    const tagMap = new Map<string, string>();
    tags.forEach(t => tagMap.set(t.id, t.name));

    // Group tag names by task ID
    const taskTagsMap = new Map<string, string[]>();
    taskTags.forEach(tt => {
      const tagName = tagMap.get(tt.tag_id);
      if (tagName) {
        const list = taskTagsMap.get(tt.task_id) || [];
        list.push(tagName);
        taskTagsMap.set(tt.task_id, list);
      }
    });

    return (tasksData || []).map(t => ({
      id: t.id,
      userId: t.user_id,
      title: t.title,
      category: t.category,
      status: t.status as TaskStatus,
      targetDate: t.target_date,
      createdAt: t.created_at_ms ? Number(t.created_at_ms) : new Date(t.created_at).getTime(),
      tags: taskTagsMap.get(t.id) || [],
      todos: t.todos || []
    }));
  } catch (error) {
    console.error('Error fetching tasks from Supabase:', error);
    return [];
  }
};

export const saveTasks = async (tasks: Task[]) => {
  try {
    for (const task of tasks) {
      await saveSingleTask(task);
    }
  } catch (error) {
    console.error('Error saving batch tasks in Supabase:', error);
  }
};

export const saveSingleTask = async (task: Task) => {
  try {
    const taskData = {
      id: task.id,
      user_id: task.userId,
      title: task.title,
      category: task.category,
      status: task.status,
      target_date: task.targetDate,
      created_at_ms: task.createdAt,
      todos: task.todos || []
    };

    const { error: taskErr } = await supabase
      .from('tasks')
      .upsert(taskData);

    if (taskErr) throw taskErr;

    // Handle tags (relational normalization)
    const { data: existingJunctions } = await supabase
      .from('task_tags')
      .select('tag_id, tags(name)')
      .eq('task_id', task.id);

    const existingTagNames = existingJunctions
      ? (existingJunctions as any[]).map(j => j.tags?.name).filter(Boolean)
      : [];

    const newTagNames = task.tags || [];

    // Compare tag arrays to detect changes
    const hasTagsChanged = 
      existingTagNames.length !== newTagNames.length ||
      !existingTagNames.every(t => newTagNames.includes(t)) ||
      !newTagNames.every(t => existingTagNames.includes(t));

    if (hasTagsChanged) {
      if (newTagNames.length > 0) {
        // 1. Ensure tags exist in public.tags
        for (const tagName of newTagNames) {
          await supabase
            .from('tags')
            .upsert({ name: tagName }, { onConflict: 'name' });
        }

        // 2. Fetch tag records to get their IDs
        const { data: tagRecords } = await supabase
          .from('tags')
          .select('id, name')
          .in('name', newTagNames);

        if (tagRecords) {
          // 3. Clear existing relations
          await supabase
            .from('task_tags')
            .delete()
            .eq('task_id', task.id);

          // 4. Insert new relations
          const junctionRecords = tagRecords.map(tr => ({
            task_id: task.id,
            tag_id: tr.id
          }));

          const { error: junctionErr } = await supabase
            .from('task_tags')
            .insert(junctionRecords);

          if (junctionErr) throw junctionErr;
        }
      } else {
        // Clear relations if no tags are provided
        await supabase
          .from('task_tags')
          .delete()
          .eq('task_id', task.id);
      }
    }
  } catch (error) {
    console.error('Error saving single task in Supabase:', error);
    throw error;
  }
};

export const deleteTask = async (id: string) => {
  try {
    // Foreign key cascading rules will handle deleting task_tags and task_updates automatically.
    const { error } = await supabase
      .from('tasks')
      .delete()
      .eq('id', id);

    if (error) throw error;
  } catch (error) {
    console.error('Error deleting task from Supabase:', error);
  }
};

export const getUpdates = async (taskIds?: string[]): Promise<TaskUpdate[]> => {
  try {
    if (taskIds && taskIds.length === 0) {
      return [];
    }

    let query = supabase.from('task_updates').select('*');
    if (taskIds) {
      query = query.in('task_id', taskIds);
    }

    const { data, error } = await query;

    if (error) throw error;

    return (data || []).map(u => ({
      id: u.id,
      taskId: u.task_id,
      userId: u.user_id || undefined,
      date: u.date,
      content: u.content,
      timestamp: u.timestamp_ms ? Number(u.timestamp_ms) : new Date(u.timestamp).getTime(),
      isArchived: u.is_archived,
      statusChange: (u.status_from && u.status_to) ? {
        from: u.status_from as TaskStatus,
        to: u.status_to as TaskStatus
      } : undefined
    }));
  } catch (error) {
    console.error('Error fetching updates from Supabase:', error);
    return [];
  }
};

export const saveUpdates = async (updates: TaskUpdate[]) => {
  try {
    for (const u of updates) {
      await saveSingleUpdate(u);
    }
  } catch (error) {
    console.error('Error saving batch updates in Supabase:', error);
  }
};

export const saveSingleUpdate = async (u: TaskUpdate) => {
  try {
    const updateData = {
      id: u.id,
      task_id: u.taskId,
      user_id: u.userId || null,
      date: u.date,
      content: u.content,
      timestamp_ms: u.timestamp,
      is_archived: u.isArchived,
      status_from: u.statusChange?.from || null,
      status_to: u.statusChange?.to || null
    };

    const { error } = await supabase
      .from('task_updates')
      .upsert(updateData);

    if (error) throw error;
  } catch (error) {
    console.error('Error saving single update in Supabase:', error);
    throw error;
  }
};

export const deleteUpdate = async (id: string) => {
  try {
    const { error } = await supabase
      .from('task_updates')
      .delete()
      .eq('id', id);

    if (error) throw error;
  } catch (error) {
    console.error('Error deleting update from Supabase:', error);
  }
};

// --- REAL-TIME SUBSCRIPTIONS ---

export const subscribeToTasks = (callback: (tasks: Task[]) => void, user?: User, skipInitialLoad = false) => {
  const channel = supabase
    .channel('realtime:tasks')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'tasks' }, async () => {
      const tasks = await getTasks(user);
      callback(tasks);
    })
    .subscribe();

  // Load initial data
  if (!skipInitialLoad) {
    getTasks(user).then(callback);
  }

  return () => {
    supabase.removeChannel(channel);
  };
};

export const subscribeToUpdates = (callback: (updates: TaskUpdate[]) => void, user?: User, skipInitialLoad = false) => {
  const channel = supabase
    .channel('realtime:task_updates')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'task_updates' }, async () => {
      const tasks = await getTasks(user);
      const taskIds = tasks.map(t => t.id);
      const updates = taskIds.length > 0 ? await getUpdates(taskIds) : [];
      callback(updates);
    })
    .subscribe();

  // Load initial data
  if (!skipInitialLoad) {
    getTasks(user).then(tasks => {
      const taskIds = tasks.map(t => t.id);
      if (taskIds.length > 0) {
        getUpdates(taskIds).then(callback);
      } else {
        callback([]);
      }
    });
  }

  return () => {
    supabase.removeChannel(channel);
  };
};

export const subscribeToUsers = (callback: (users: User[]) => void, skipInitialLoad = false) => {
  const channel = supabase
    .channel('realtime:profiles')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'profiles' }, async () => {
      const users = await getUsers();
      callback(users);
    })
    .subscribe();

  // Load initial data
  if (!skipInitialLoad) {
    getUsers().then(callback);
  }

  return () => {
    supabase.removeChannel(channel);
  };
};

export const saveSingleTag = async (tag: Tag) => {
  try {
    const { error } = await supabase
      .from('tags')
      .upsert({ id: tag.id, name: tag.name });

    if (error) throw error;
  } catch (error) {
    console.error('Error saving tag in Supabase:', error);
  }
};

export const subscribeToTags = (callback: (tags: Tag[]) => void, skipInitialLoad = false) => {
  const channel = supabase
    .channel('realtime:tags')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'tags' }, async () => {
      const { data } = await supabase.from('tags').select('*');
      callback(data || []);
    })
    .subscribe();

  // Load initial data
  if (!skipInitialLoad) {
    supabase.from('tags').select('*').then(({ data }) => {
      callback(data || []);
    });
  }

  return () => {
    supabase.removeChannel(channel);
  };
};

export const saveAiTokenRecord = async (record: AiTokenRecord) => {
  try {
    const recordData = {
      id: record.id,
      user_id: record.userId,
      date: record.date,
      timestamp_ms: record.timestamp,
      used_for: record.usedFor,
      input_tokens: record.inputTokens,
      output_tokens: record.outputTokens
    };

    const { error } = await supabase
      .from('ai_token_records')
      .upsert(recordData);

    if (error) throw error;
  } catch (error) {
    console.error('Error saving AI token record in Supabase:', error);
  }
};

export const subscribeToAiTokenRecords = (callback: (records: AiTokenRecord[]) => void, userId?: string, skipInitialLoad = false) => {
  const channel = supabase
    .channel('realtime:ai_token_records')
    .on('postgres_changes', { event: '*', schema: 'public', table: 'ai_token_records' }, async () => {
      let q = supabase.from('ai_token_records').select('*');
      if (userId) {
        q = q.eq('user_id', userId);
      }
      const { data } = await q.order('timestamp_ms', { ascending: false });
      
      const records = (data || []).map(r => ({
        id: r.id,
        userId: r.user_id,
        date: r.date,
        timestamp: r.timestamp_ms ? Number(r.timestamp_ms) : new Date(r.timestamp).getTime(),
        usedFor: r.used_for,
        inputTokens: r.input_tokens,
        outputTokens: r.output_tokens
      }));
      callback(records);
    })
    .subscribe();

  // Load initial data
  if (!skipInitialLoad) {
    let q = supabase.from('ai_token_records').select('*');
    if (userId) {
      q = q.eq('user_id', userId);
    }
    q.order('timestamp_ms', { ascending: false }).then(({ data }) => {
      const records = (data || []).map(r => ({
        id: r.id,
        userId: r.user_id,
        date: r.date,
        timestamp: r.timestamp_ms ? Number(r.timestamp_ms) : new Date(r.timestamp).getTime(),
        usedFor: r.used_for,
        inputTokens: r.input_tokens,
        outputTokens: r.output_tokens
      }));
      callback(records);
    });
  }

  return () => {
    supabase.removeChannel(channel);
  };
};

export const importRestoreDatabase = async (data: { tasks: Task[]; updates: TaskUpdate[]; users: User[] }) => {
  try {
    // 1. Restore profiles
    const profileRecords = data.users.map(u => ({
      id: u.id,
      username: u.username,
      pin: u.pin,
      role: u.role
    }));
    if (profileRecords.length > 0) {
      const { error: profileErr } = await supabase.from('profiles').upsert(profileRecords);
      if (profileErr) throw profileErr;
    }

    // 2. Restore team member links
    const teamRecords: any[] = [];
    data.users.forEach(u => {
      if (u.teamMemberIds && u.teamMemberIds.length > 0) {
        u.teamMemberIds.forEach(mId => {
          teamRecords.push({ admin_id: u.id, member_id: mId });
        });
      }
    });
    if (teamRecords.length > 0) {
      const { error: teamErr } = await supabase.from('team_members').upsert(teamRecords);
      if (teamErr) throw teamErr;
    }

    // 3. Restore tasks (and extract tags)
    for (const task of data.tasks) {
      await saveSingleTask(task);
    }

    // 4. Restore updates
    for (const u of data.updates) {
      await saveSingleUpdate(u);
    }
  } catch (error) {
    console.error('Error importing/restoring database in Supabase:', error);
    throw error;
  }
};

// Backfill matches legacy API
export const backfillTaskUpdates = async (userId: string, userTasks: Task[]) => {
  // Supabase RLS and normalization handles this automatically or via SQL script.
};
