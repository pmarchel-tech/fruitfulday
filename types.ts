
export enum TaskStatus {
  NOT_YET = 'NOT_YET',
  PROGRESS = 'PROGRESS',
  DONE = 'DONE',
  CANCEL = 'CANCEL',
  PENDING = 'PENDING',
  REPETITIVE = 'REPETITIVE',
  FOLLOW_UP = 'FOLLOW_UP'
}

export type UserRole = 'ADMIN' | 'MEMBER';

export interface User {
  id: string;
  username: string;
  pin: string; // Simple password/pin
  role: UserRole;
  teamMemberIds: string[]; // IDs of users this admin manages
}

export interface TodoItem {
  id: string;
  text: string;
  completed: boolean;
}

export interface Task {
  id: string;
  userId: string; // Owner of the task
  title: string;
  category: string;
  status: TaskStatus;
  targetDate: string; // ISO Date string YYYY-MM-DD
  createdAt: number;
  tags?: string[];
  todos?: TodoItem[];
}

export interface Tag {
  id: string;
  name: string;
}

export interface TaskUpdate {
  id: string;
  taskId: string;
  userId?: string;
  date: string; // ISO Date string YYYY-MM-DD
  content: string;
  timestamp: number;
  isArchived: boolean;
  statusChange?: {
    from: TaskStatus;
    to: TaskStatus;
  };
}

export type FilterTime = 'ALL' | 'OVERDUE' | 'TODAY' | 'THIS_WEEK' | 'FUTURE';
export type SortOrder = 'NEWEST' | 'OLDEST';

export interface AiTokenRecord {
  id: string;
  userId: string;
  date: string; // YYYY-MM-DD
  timestamp: number;
  usedFor: string;
  inputTokens: number;
  outputTokens: number;
}

