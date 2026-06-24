
import { Task, TaskUpdate, TaskStatus, User, UserRole, Tag, AiTokenRecord } from '../types';
import { db, auth } from '../firebase';
import { 
  collection, 
  doc, 
  getDoc, 
  getDocs, 
  setDoc, 
  updateDoc, 
  deleteDoc, 
  query, 
  where, 
  orderBy,
  onSnapshot,
  writeBatch
} from 'firebase/firestore';

const CURRENT_USER_KEY = 'dailypulse_current_user_id';

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId: string | undefined;
    email: string | null | undefined;
    emailVerified: boolean | undefined;
    isAnonymous: boolean | undefined;
    tenantId: string | null | undefined;
    providerInfo: {
      providerId: string;
      displayName: string | null;
      email: string | null;
      photoUrl: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData.map(provider => ({
        providerId: provider.providerId,
        displayName: provider.displayName,
        email: provider.email,
        photoUrl: provider.photoURL
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

// --- AUTH SERVICE ---

export const getUsers = async (): Promise<User[]> => {
  const path = 'users';
  try {
    const snapshot = await getDocs(collection(db, path));
    return snapshot.docs.map(doc => doc.data() as User);
  } catch (error) {
    handleFirestoreError(error, OperationType.LIST, path);
    return [];
  }
};

export const registerUser = async (username: string, pin: string, role: UserRole): Promise<User | null> => {
  const path = 'users';
  try {
    // Check if username exists
    const q = query(collection(db, path), where("username", "==", username));
    const snapshot = await getDocs(q);
    if (!snapshot.empty) return null;
    
    const id = auth.currentUser?.uid || Date.now().toString();
    const newUser: User = {
      id,
      username,
      pin,
      role,
      teamMemberIds: []
    };

    await setDoc(doc(db, path, id), newUser);
    localStorage.setItem(CURRENT_USER_KEY, id);
    return newUser;
  } catch (error) {
    handleFirestoreError(error, OperationType.CREATE, path);
    return null;
  }
};

export const loginUser = async (username: string, pin: string): Promise<User | null> => {
  const path = 'users';
  try {
    const q = query(collection(db, path), where("username", "==", username), where("pin", "==", pin));
    const snapshot = await getDocs(q);
    
    if (snapshot.empty) return null;
    
    const user = snapshot.docs[0].data() as User;
    localStorage.setItem(CURRENT_USER_KEY, user.id);
    return user;
  } catch (error) {
    handleFirestoreError(error, OperationType.GET, path);
    return null;
  }
};

export const getCurrentUser = async (): Promise<User | null> => {
  const firebaseUser = auth.currentUser;
  
  // If we have a permanent firebase user (Google), always use that UID
  // If anonymous, we might still want to check localStorage for a "profile" ID
  const id = (firebaseUser && !firebaseUser.isAnonymous) 
    ? firebaseUser.uid 
    : (localStorage.getItem(CURRENT_USER_KEY) || firebaseUser?.uid);

  if (!id) return null;

  const path = `users/${id}`;
  try {
    const userDoc = await getDoc(doc(db, 'users', id));
    if (userDoc.exists()) {
      const userData = userDoc.data() as User;
      // Ensure localStorage is in sync
      localStorage.setItem(CURRENT_USER_KEY, id);
      return userData;
    }
    
    // If document doesn't exist but we have a firebase user, 
    // it might be a first-time Google login that hasn't registered yet
    return null;
  } catch (error) {
    handleFirestoreError(error, OperationType.GET, path);
    return null;
  }
};

export const updateUser = async (updatedUser: User): Promise<boolean> => {
  const path = `users/${updatedUser.id}`;
  try {
    await setDoc(doc(db, 'users', updatedUser.id), updatedUser, { merge: true });
    return true;
  } catch (error) {
    handleFirestoreError(error, OperationType.UPDATE, path);
    return false;
  }
};

export const logoutUser = () => {
  localStorage.removeItem(CURRENT_USER_KEY);
  auth.signOut();
};

export const addTeamMembers = async (adminId: string, memberIds: string[]): Promise<boolean> => {
  const path = `users/${adminId}`;
  try {
    const adminRef = doc(db, 'users', adminId);
    const adminDoc = await getDoc(adminRef);
    if (!adminDoc.exists()) return false;
    
    const currentTeam = (adminDoc.data() as User).teamMemberIds || [];
    const newTeam = Array.from(new Set([...currentTeam, ...memberIds]));
    
    await updateDoc(adminRef, { teamMemberIds: newTeam });
    return true;
  } catch (error) {
    handleFirestoreError(error, OperationType.UPDATE, path);
    return false;
  }
};

// --- DATA SERVICE ---

export const getTasks = async (user?: User): Promise<Task[]> => {
  const path = 'tasks';
  try {
    let q;
    const currentUid = auth.currentUser?.uid;
    
    if (user?.role === 'ADMIN') {
      q = collection(db, path);
    } else if (user) {
      q = query(collection(db, path), where("userId", "==", user.id));
    } else if (currentUid) {
      // Fallback for authenticated users without a profile yet
      q = query(collection(db, path), where("userId", "==", currentUid));
    } else {
      // This will likely fail due to rules, but it's a safe fallback
      q = collection(db, path);
    }
    
    const snapshot = await getDocs(q);
    return snapshot.docs.map(doc => doc.data() as Task);
  } catch (error) {
    handleFirestoreError(error, OperationType.LIST, path);
    return [];
  }
};

export const saveTasks = async (tasks: Task[]) => {
  const path = 'tasks';
  try {
    const batch = writeBatch(db);
    tasks.forEach(task => {
      const ref = doc(db, path, task.id);
      batch.set(ref, task);
    });
    await batch.commit();
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
};

export const saveSingleTask = async (task: Task) => {
  const path = `tasks/${task.id}`;
  try {
    await setDoc(doc(db, 'tasks', task.id), task);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
};

export const deleteTask = async (id: string) => {
  const path = `tasks/${id}`;
  try {
    await deleteDoc(doc(db, 'tasks', id));
    
    // Also delete updates for this task
    const updatesQ = query(collection(db, 'task_updates'), where("taskId", "==", id));
    const updatesSnapshot = await getDocs(updatesQ);
    const batch = writeBatch(db);
    updatesSnapshot.docs.forEach(d => batch.delete(d.ref));
    await batch.commit();
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
};

export const getUpdates = async (): Promise<TaskUpdate[]> => {
  const path = 'task_updates';
  try {
    const snapshot = await getDocs(collection(db, path));
    return snapshot.docs.map(doc => doc.data() as TaskUpdate);
  } catch (error) {
    handleFirestoreError(error, OperationType.LIST, path);
    return [];
  }
};

export const saveUpdates = async (updates: TaskUpdate[]) => {
  const path = 'task_updates';
  try {
    const batch = writeBatch(db);
    updates.forEach(u => {
      const ref = doc(db, path, u.id);
      batch.set(ref, u);
    });
    await batch.commit();
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
};

export const saveSingleUpdate = async (u: TaskUpdate) => {
  const path = `task_updates/${u.id}`;
  try {
    await setDoc(doc(db, 'task_updates', u.id), u);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
};

export const deleteUpdate = async (id: string) => {
  const path = `task_updates/${id}`;
  try {
    await deleteDoc(doc(db, 'task_updates', id));
  } catch (error) {
    handleFirestoreError(error, OperationType.DELETE, path);
  }
};

// Real-time listeners
export const subscribeToTasks = (callback: (tasks: Task[]) => void, user?: User) => {
  const path = 'tasks';
  let q;
  const currentUid = auth.currentUser?.uid;

  if (user?.role === 'ADMIN') {
    q = collection(db, path);
  } else if (user) {
    q = query(collection(db, path), where("userId", "==", user.id));
  } else if (currentUid) {
    q = query(collection(db, path), where("userId", "==", currentUid));
  } else {
    q = collection(db, path);
  }

  return onSnapshot(q, (snapshot) => {
    callback(snapshot.docs.map(doc => doc.data() as Task));
  }, (error) => {
    handleFirestoreError(error, OperationType.LIST, path);
  });
};

export const subscribeToUpdates = (callback: (updates: TaskUpdate[]) => void, user?: User) => {
  const path = 'task_updates';
  let q;
  const currentUid = auth.currentUser?.uid;

  if (user?.role === 'ADMIN') {
    const uids = [user.id, ...(user.teamMemberIds || [])];
    if (uids.length > 0 && uids.length <= 30) {
      q = query(collection(db, path), where("userId", "in", uids));
    } else {
      q = collection(db, path);
    }
  } else if (user) {
    q = query(collection(db, path), where("userId", "==", user.id));
  } else if (currentUid) {
    q = query(collection(db, path), where("userId", "==", currentUid));
  } else {
    q = collection(db, path);
  }

  return onSnapshot(q, (snapshot) => {
    callback(snapshot.docs.map(doc => doc.data() as TaskUpdate));
  }, (error) => {
    handleFirestoreError(error, OperationType.LIST, path);
  });
};

export const backfillTaskUpdates = async (userId: string, userTasks: Task[]) => {
  if (!userTasks || userTasks.length === 0) return;
  const path = 'task_updates';
  try {
    const batch = writeBatch(db);
    let count = 0;
    
    // Process up to 30 tasks to keep performance very high and reads light
    const targetTasks = userTasks.slice(0, 30);
    for (const task of targetTasks) {
      const q = query(collection(db, path), where('taskId', '==', task.id));
      const snap = await getDocs(q);
      snap.docs.forEach(d => {
        const data = d.data();
        if (!data.userId) {
          batch.update(d.ref, { userId });
          count++;
        }
      });
    }
    
    if (count > 0) {
      await batch.commit();
      console.log(`[GrowDaily Migration] Successfully backfilled userId for ${count} legacy task updates.`);
    }
  } catch (error) {
    console.warn("Backfill non-fatal error:", error);
  }
};

export const subscribeToUsers = (callback: (users: User[]) => void) => {
  const path = 'users';
  return onSnapshot(collection(db, path), (snapshot) => {
    callback(snapshot.docs.map(doc => doc.data() as User));
  }, (error) => {
    handleFirestoreError(error, OperationType.LIST, path);
  });
};

export const saveSingleTag = async (tag: Tag) => {
  const path = `tags/${tag.id}`;
  try {
    await setDoc(doc(db, 'tags', tag.id), tag);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
};

export const subscribeToTags = (callback: (tags: Tag[]) => void) => {
  const path = 'tags';
  return onSnapshot(collection(db, path), (snapshot) => {
    callback(snapshot.docs.map(doc => doc.data() as Tag));
  }, (error) => {
    handleFirestoreError(error, OperationType.LIST, path);
  });
};

export const importRestoreDatabase = async (data: { tasks: Task[]; updates: TaskUpdate[]; users: User[] }) => {
  const batch = writeBatch(db);
  
  // 1. Clear/Save tasks
  data.tasks.forEach(task => {
    batch.set(doc(db, 'tasks', task.id), task);
  });
  
  // 2. Clear/Save updates
  data.updates.forEach(u => {
    batch.set(doc(db, 'task_updates', u.id), u);
  });
  
  // 3. Clear/Save users merging pins safely
  const currentUsers = await getUsers();
  const currentUsersMap = new Map(currentUsers.map(u => [u.id, u]));
  
  data.users.forEach(user => {
    const existing = currentUsersMap.get(user.id);
    const pin = existing ? existing.pin : user.pin;
    batch.set(doc(db, 'users', user.id), { ...user, pin });
  });
  
  await batch.commit();
};

export const saveAiTokenRecord = async (record: AiTokenRecord) => {
  const path = `ai_token_records/${record.id}`;
  try {
    await setDoc(doc(db, 'ai_token_records', record.id), record);
  } catch (error) {
    handleFirestoreError(error, OperationType.WRITE, path);
  }
};

export const subscribeToAiTokenRecords = (callback: (records: AiTokenRecord[]) => void, userId?: string) => {
  const path = 'ai_token_records';
  let q;
  if (userId) {
    q = query(collection(db, path), where("userId", "==", userId));
  } else {
    q = collection(db, path);
  }
  return onSnapshot(q, (snapshot) => {
    const records = snapshot.docs.map(doc => doc.data() as AiTokenRecord);
    records.sort((a, b) => b.timestamp - a.timestamp);
    callback(records);
  }, (error) => {
    handleFirestoreError(error, OperationType.LIST, path);
  });
};


